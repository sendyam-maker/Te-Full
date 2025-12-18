VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm090712 
   BorderStyle     =   1  '單線固定
   Caption         =   "繪圖分案作業"
   ClientHeight    =   8880
   ClientLeft      =   -1785
   ClientTop       =   960
   ClientWidth     =   15000
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8880
   ScaleWidth      =   15000
   Begin VB.CommandButton cmdData 
      Caption         =   "南、高所"
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   3555
      TabIndex        =   98
      Top             =   45
      Width           =   1260
   End
   Begin VB.CommandButton cmdData 
      Caption         =   "中所"
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   2295
      TabIndex        =   97
      Top             =   45
      Width           =   1260
   End
   Begin VB.CommandButton cmdData 
      Caption         =   "北所"
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   1035
      TabIndex        =   96
      Top             =   45
      Width           =   1260
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "確認分案"
      Height          =   400
      Index           =   0
      Left            =   11460
      TabIndex        =   95
      Top             =   -15
      Width           =   1260
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "確定"
      Height          =   400
      Index           =   2
      Left            =   12735
      TabIndex        =   18
      Top             =   -15
      Width           =   756
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "回前畫面(&U)"
      Height          =   400
      Index           =   1
      Left            =   13515
      TabIndex        =   19
      Top             =   -15
      Width           =   1200
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   8250
      Left            =   120
      TabIndex        =   20
      Top             =   540
      Width           =   14685
      _ExtentX        =   25903
      _ExtentY        =   14552
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "瀏覽資料"
      TabPicture(0)   =   "frm090712.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Combo3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "grd1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label5"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "詳細資料"
      TabPicture(1)   =   "frm090712.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label1(30)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "lbl1(30)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "lbl1(25)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "lbl1(20)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "lbl1(23)"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "lbl1(22)"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "lbl1(24)"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "lbl1(19)"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "lbl1(18)"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "Label1(31)"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "Label1(22)"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "Label1(6)"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "Label1(3)"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "Label1(26)"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "Label1(23)"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "Label1(17)"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "Label1(2)"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "Label1(28)"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).Control(18)=   "Label1(34)"
      Tab(1).Control(18).Enabled=   0   'False
      Tab(1).Control(19)=   "lbl1(21)"
      Tab(1).Control(19).Enabled=   0   'False
      Tab(1).Control(20)=   "Label1(5)"
      Tab(1).Control(20).Enabled=   0   'False
      Tab(1).Control(21)=   "Label1(24)"
      Tab(1).Control(21).Enabled=   0   'False
      Tab(1).Control(22)=   "lbl1(26)"
      Tab(1).Control(22).Enabled=   0   'False
      Tab(1).Control(23)=   "Label1(27)"
      Tab(1).Control(23).Enabled=   0   'False
      Tab(1).Control(24)=   "lbl1(28)"
      Tab(1).Control(24).Enabled=   0   'False
      Tab(1).Control(25)=   "lbl1(29)"
      Tab(1).Control(25).Enabled=   0   'False
      Tab(1).Control(26)=   "Label1(35)"
      Tab(1).Control(26).Enabled=   0   'False
      Tab(1).Control(27)=   "Label1(1)"
      Tab(1).Control(27).Enabled=   0   'False
      Tab(1).Control(28)=   "Label1(33)"
      Tab(1).Control(28).Enabled=   0   'False
      Tab(1).Control(29)=   "Label1(29)"
      Tab(1).Control(29).Enabled=   0   'False
      Tab(1).Control(30)=   "Label1(21)"
      Tab(1).Control(30).Enabled=   0   'False
      Tab(1).Control(31)=   "Label1(20)"
      Tab(1).Control(31).Enabled=   0   'False
      Tab(1).Control(32)=   "Label1(19)"
      Tab(1).Control(32).Enabled=   0   'False
      Tab(1).Control(33)=   "Label1(18)"
      Tab(1).Control(33).Enabled=   0   'False
      Tab(1).Control(34)=   "Label1(16)"
      Tab(1).Control(34).Enabled=   0   'False
      Tab(1).Control(35)=   "Label1(15)"
      Tab(1).Control(35).Enabled=   0   'False
      Tab(1).Control(36)=   "Label1(14)"
      Tab(1).Control(36).Enabled=   0   'False
      Tab(1).Control(37)=   "Label1(13)"
      Tab(1).Control(37).Enabled=   0   'False
      Tab(1).Control(38)=   "Label1(12)"
      Tab(1).Control(38).Enabled=   0   'False
      Tab(1).Control(39)=   "Label1(11)"
      Tab(1).Control(39).Enabled=   0   'False
      Tab(1).Control(40)=   "Label1(10)"
      Tab(1).Control(40).Enabled=   0   'False
      Tab(1).Control(41)=   "Label1(9)"
      Tab(1).Control(41).Enabled=   0   'False
      Tab(1).Control(42)=   "Label1(8)"
      Tab(1).Control(42).Enabled=   0   'False
      Tab(1).Control(43)=   "Label1(4)"
      Tab(1).Control(43).Enabled=   0   'False
      Tab(1).Control(44)=   "Label1(25)"
      Tab(1).Control(44).Enabled=   0   'False
      Tab(1).Control(45)=   "lbl1(27)"
      Tab(1).Control(45).Enabled=   0   'False
      Tab(1).Control(46)=   "lbl1(10)"
      Tab(1).Control(46).Enabled=   0   'False
      Tab(1).Control(47)=   "lbl1(1)"
      Tab(1).Control(47).Enabled=   0   'False
      Tab(1).Control(48)=   "lbl1(3)"
      Tab(1).Control(48).Enabled=   0   'False
      Tab(1).Control(49)=   "lbl1(4)"
      Tab(1).Control(49).Enabled=   0   'False
      Tab(1).Control(50)=   "lbl1(5)"
      Tab(1).Control(50).Enabled=   0   'False
      Tab(1).Control(51)=   "lbl1(8)"
      Tab(1).Control(51).Enabled=   0   'False
      Tab(1).Control(52)=   "lbl1(9)"
      Tab(1).Control(52).Enabled=   0   'False
      Tab(1).Control(53)=   "lbl1(11)"
      Tab(1).Control(53).Enabled=   0   'False
      Tab(1).Control(54)=   "lbl1(6)"
      Tab(1).Control(54).Enabled=   0   'False
      Tab(1).Control(55)=   "lbl1(16)"
      Tab(1).Control(55).Enabled=   0   'False
      Tab(1).Control(56)=   "lbl1(12)"
      Tab(1).Control(56).Enabled=   0   'False
      Tab(1).Control(57)=   "lbl1(14)"
      Tab(1).Control(57).Enabled=   0   'False
      Tab(1).Control(58)=   "lbl1(15)"
      Tab(1).Control(58).Enabled=   0   'False
      Tab(1).Control(59)=   "lbl1(13)"
      Tab(1).Control(59).Enabled=   0   'False
      Tab(1).Control(60)=   "lbl1(7)"
      Tab(1).Control(60).Enabled=   0   'False
      Tab(1).Control(61)=   "lbl1(2)"
      Tab(1).Control(61).Enabled=   0   'False
      Tab(1).Control(62)=   "lbl1(0)"
      Tab(1).Control(62).Enabled=   0   'False
      Tab(1).Control(63)=   "lblClose"
      Tab(1).Control(63).Enabled=   0   'False
      Tab(1).Control(64)=   "Label1(7)"
      Tab(1).Control(64).Enabled=   0   'False
      Tab(1).Control(65)=   "Label1(36)"
      Tab(1).Control(65).Enabled=   0   'False
      Tab(1).Control(66)=   "lbl1(17)"
      Tab(1).Control(66).Enabled=   0   'False
      Tab(1).Control(67)=   "Label1(37)"
      Tab(1).Control(67).Enabled=   0   'False
      Tab(1).Control(68)=   "lbl1(31)"
      Tab(1).Control(68).Enabled=   0   'False
      Tab(1).Control(69)=   "Label1(38)"
      Tab(1).Control(69).Enabled=   0   'False
      Tab(1).Control(70)=   "lbl1(32)"
      Tab(1).Control(70).Enabled=   0   'False
      Tab(1).Control(71)=   "txt1(8)"
      Tab(1).Control(71).Enabled=   0   'False
      Tab(1).Control(72)=   "txt1(9)"
      Tab(1).Control(72).Enabled=   0   'False
      Tab(1).Control(73)=   "txt1(0)"
      Tab(1).Control(73).Enabled=   0   'False
      Tab(1).Control(74)=   "txt1(1)"
      Tab(1).Control(74).Enabled=   0   'False
      Tab(1).Control(75)=   "txt1(2)"
      Tab(1).Control(75).Enabled=   0   'False
      Tab(1).Control(76)=   "txt1(3)"
      Tab(1).Control(76).Enabled=   0   'False
      Tab(1).Control(77)=   "txt1(4)"
      Tab(1).Control(77).Enabled=   0   'False
      Tab(1).Control(78)=   "txt1(5)"
      Tab(1).Control(78).Enabled=   0   'False
      Tab(1).Control(79)=   "txt1(6)"
      Tab(1).Control(79).Enabled=   0   'False
      Tab(1).Control(80)=   "txt1(7)"
      Tab(1).Control(80).Enabled=   0   'False
      Tab(1).Control(81)=   "txt1(10)"
      Tab(1).Control(81).Enabled=   0   'False
      Tab(1).Control(82)=   "txt1(11)"
      Tab(1).Control(82).Enabled=   0   'False
      Tab(1).Control(83)=   "txt1(12)"
      Tab(1).Control(83).Enabled=   0   'False
      Tab(1).Control(84)=   "txt1(13)"
      Tab(1).Control(84).Enabled=   0   'False
      Tab(1).Control(85)=   "txt1(14)"
      Tab(1).Control(85).Enabled=   0   'False
      Tab(1).Control(86)=   "Option1(0)"
      Tab(1).Control(86).Enabled=   0   'False
      Tab(1).Control(87)=   "Option1(1)"
      Tab(1).Control(87).Enabled=   0   'False
      Tab(1).Control(88)=   "Combo2"
      Tab(1).Control(88).Enabled=   0   'False
      Tab(1).ControlCount=   89
      Begin VB.ComboBox Combo3 
         Height          =   300
         ItemData        =   "frm090712.frx":0038
         Left            =   -73950
         List            =   "frm090712.frx":004E
         TabIndex        =   93
         Top             =   390
         Width           =   2430
      End
      Begin VB.ComboBox Combo2 
         Enabled         =   0   'False
         Height          =   300
         Left            =   5820
         TabIndex        =   17
         Text            =   "Combo2"
         Top             =   4845
         Width           =   1560
      End
      Begin VB.OptionButton Option1 
         Enabled         =   0   'False
         Height          =   300
         Index           =   1
         Left            =   5160
         TabIndex        =   16
         Top             =   4785
         Width           =   285
      End
      Begin VB.OptionButton Option1 
         Enabled         =   0   'False
         Height          =   300
         Index           =   0
         Left            =   5160
         TabIndex        =   14
         Top             =   4500
         Value           =   -1  'True
         Width           =   300
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grd1 
         Height          =   6990
         Left            =   -74925
         TabIndex        =   21
         Top             =   795
         Width           =   14505
         _ExtentX        =   25585
         _ExtentY        =   12330
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
      Begin MSForms.TextBox txt1 
         Height          =   300
         Index           =   14
         Left            =   5820
         TabIndex        =   5
         Top             =   1800
         Width           =   585
         VariousPropertyBits=   671107097
         MaxLength       =   1
         Size            =   "1032;529"
         FontName        =   "新細明體-ExtB"
         FontEffects     =   1073750016
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   300
         Index           =   13
         Left            =   5820
         TabIndex        =   15
         Top             =   4560
         Width           =   3810
         VariousPropertyBits=   -1467989991
         MaxLength       =   2000
         ScrollBars      =   2
         Size            =   "6720;529"
         FontName        =   "新細明體-ExtB"
         FontEffects     =   1073750016
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   300
         Index           =   12
         Left            =   5820
         TabIndex        =   13
         Top             =   4275
         Width           =   1200
         VariousPropertyBits=   671107097
         MaxLength       =   4
         Size            =   "2117;529"
         FontName        =   "新細明體-ExtB"
         FontEffects     =   1073750016
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   300
         Index           =   11
         Left            =   5820
         TabIndex        =   12
         Top             =   4005
         Width           =   1200
         VariousPropertyBits=   671107097
         MaxLength       =   4
         Size            =   "2117;529"
         FontName        =   "新細明體-ExtB"
         FontEffects     =   1073750016
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   300
         Index           =   10
         Left            =   5820
         TabIndex        =   11
         Top             =   3720
         Width           =   1200
         VariousPropertyBits=   671107097
         MaxLength       =   4
         Size            =   "2117;529"
         FontName        =   "新細明體-ExtB"
         FontEffects     =   1073750016
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   300
         Index           =   7
         Left            =   5820
         TabIndex        =   1
         Top             =   420
         Width           =   585
         VariousPropertyBits=   671107097
         MaxLength       =   1
         Size            =   "1032;529"
         FontName        =   "新細明體-ExtB"
         FontEffects     =   1073750016
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   300
         Index           =   6
         Left            =   5820
         TabIndex        =   8
         Top             =   2625
         Width           =   1200
         VariousPropertyBits=   671107097
         MaxLength       =   2
         Size            =   "2117;529"
         FontName        =   "新細明體-ExtB"
         FontEffects     =   1073750016
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   300
         Index           =   5
         Left            =   5820
         TabIndex        =   7
         Top             =   2340
         Width           =   1200
         VariousPropertyBits=   671107097
         MaxLength       =   7
         Size            =   "2117;529"
         FontName        =   "新細明體-ExtB"
         FontEffects     =   1073750016
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   300
         Index           =   4
         Left            =   5820
         TabIndex        =   6
         Top             =   2070
         Width           =   1200
         VariousPropertyBits=   671107097
         MaxLength       =   7
         Size            =   "2117;529"
         FontName        =   "新細明體-ExtB"
         FontEffects     =   1073750016
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   300
         Index           =   3
         Left            =   5820
         TabIndex        =   4
         Top             =   1245
         Width           =   1200
         VariousPropertyBits=   671107097
         MaxLength       =   2
         Size            =   "2117;529"
         FontName        =   "新細明體-ExtB"
         FontEffects     =   1073750016
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   300
         Index           =   2
         Left            =   5820
         TabIndex        =   3
         Top             =   975
         Width           =   1200
         VariousPropertyBits=   671107097
         MaxLength       =   7
         Size            =   "2117;529"
         FontName        =   "新細明體-ExtB"
         FontEffects     =   1073750016
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   300
         Index           =   1
         Left            =   5820
         TabIndex        =   2
         Top             =   690
         Width           =   1200
         VariousPropertyBits=   671107097
         MaxLength       =   7
         Size            =   "2117;529"
         FontName        =   "新細明體-ExtB"
         FontEffects     =   1073750016
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   300
         Index           =   0
         Left            =   975
         TabIndex        =   0
         Top             =   390
         Width           =   1200
         VariousPropertyBits=   671107099
         MaxLength       =   6
         Size            =   "2117;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   300
         Index           =   9
         Left            =   6030
         TabIndex        =   10
         Top             =   3450
         Width           =   1200
         VariousPropertyBits=   671107097
         MaxLength       =   4
         Size            =   "2117;529"
         FontName        =   "新細明體-ExtB"
         FontEffects     =   1073750016
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   300
         Index           =   8
         Left            =   6030
         TabIndex        =   9
         Top             =   3165
         Width           =   1200
         VariousPropertyBits=   671107097
         MaxLength       =   4
         Size            =   "2117;529"
         FontName        =   "新細明體-ExtB"
         FontEffects     =   1073750016
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label5 
         Caption         =   "顏色說明："
         Height          =   225
         Left            =   -74850
         TabIndex        =   94
         Top             =   435
         Width           =   915
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   32
         Left            =   8490
         TabIndex        =   92
         Top             =   2070
         Width           =   1740
         VariousPropertyBits=   27
         Caption         =   "123"
         Size            =   "3069;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         Caption         =   "墨圖承辦期限："
         Height          =   195
         Index           =   38
         Left            =   7140
         TabIndex        =   91
         Top             =   2070
         Width           =   1305
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   31
         Left            =   8520
         TabIndex        =   90
         Top             =   690
         Width           =   1530
         VariousPropertyBits=   27
         Caption         =   "123"
         Size            =   "2699;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         Caption         =   "草圖承辦期限："
         Height          =   195
         Index           =   37
         Left            =   7170
         TabIndex        =   89
         Top             =   690
         Width           =   1305
      End
      Begin MSForms.Label lbl1 
         Height          =   285
         Index           =   17
         Left            =   11205
         TabIndex        =   88
         Top             =   1800
         Visible         =   0   'False
         Width           =   1770
         VariousPropertyBits=   27
         Caption         =   "123"
         Size            =   "3122;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         Caption         =   "墨圖是否計件："
         Height          =   195
         Index           =   36
         Left            =   4440
         TabIndex        =   87
         Top             =   1800
         Width           =   1275
      End
      Begin VB.Label Label1 
         Caption         =   "(N：不計)"
         Height          =   180
         Index           =   7
         Left            =   6480
         TabIndex        =   86
         Top             =   1800
         Width           =   1065
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
         Height          =   180
         Left            =   3060
         TabIndex        =   66
         Top             =   1245
         Width           =   930
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   0
         Left            =   2250
         TabIndex        =   48
         Top             =   420
         Width           =   1905
         VariousPropertyBits=   27
         Caption         =   "123"
         Size            =   "3360;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   2
         Left            =   870
         TabIndex        =   49
         Top             =   975
         Width           =   1605
         VariousPropertyBits=   27
         Caption         =   "123"
         Size            =   "2831;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   7
         Left            =   690
         TabIndex        =   53
         Top             =   2355
         Width           =   1605
         VariousPropertyBits=   27
         Caption         =   "123"
         Size            =   "2831;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   13
         Left            =   1410
         TabIndex        =   62
         Top             =   4020
         Width           =   1605
         VariousPropertyBits=   27
         Caption         =   "123"
         Size            =   "2831;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   15
         Left            =   1440
         TabIndex        =   61
         Top             =   4575
         Width           =   1605
         VariousPropertyBits=   27
         Caption         =   "123"
         Size            =   "2831;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   14
         Left            =   1230
         TabIndex        =   60
         Top             =   4305
         Width           =   1605
         VariousPropertyBits=   27
         Caption         =   "123"
         Size            =   "2831;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   12
         Left            =   1440
         TabIndex        =   59
         Top             =   3750
         Width           =   1605
         VariousPropertyBits=   27
         Caption         =   "123"
         Size            =   "2831;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   16
         Left            =   1680
         TabIndex        =   58
         Top             =   4860
         Width           =   1575
         VariousPropertyBits=   27
         Caption         =   "123"
         Size            =   "2778;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   6
         Left            =   1050
         TabIndex        =   57
         Top             =   2085
         Width           =   2835
         VariousPropertyBits=   27
         Caption         =   "123"
         Size            =   "5001;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   11
         Left            =   1050
         TabIndex        =   56
         Top             =   3465
         Width           =   1605
         VariousPropertyBits=   27
         Caption         =   "123"
         Size            =   "2831;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   9
         Left            =   1050
         TabIndex        =   55
         Top             =   2910
         Width           =   1605
         VariousPropertyBits=   27
         Caption         =   "123"
         Size            =   "2831;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   8
         Left            =   1080
         TabIndex        =   54
         Top             =   2640
         Width           =   1605
         VariousPropertyBits=   27
         Caption         =   "123"
         Size            =   "2831;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   5
         Left            =   1470
         TabIndex        =   52
         Top             =   1800
         Width           =   1605
         VariousPropertyBits=   27
         Caption         =   "123"
         Size            =   "2831;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   4
         Left            =   1020
         TabIndex        =   51
         Top             =   1530
         Width           =   3375
         VariousPropertyBits=   27
         Caption         =   "123"
         Size            =   "5953;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   3
         Left            =   1020
         TabIndex        =   50
         Top             =   1245
         Width           =   1905
         VariousPropertyBits=   27
         Caption         =   "123"
         Size            =   "3360;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   1
         Left            =   1050
         TabIndex        =   47
         Top             =   690
         Width           =   1605
         VariousPropertyBits=   27
         Caption         =   "123"
         Size            =   "2831;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   10
         Left            =   870
         TabIndex        =   46
         Top             =   3195
         Width           =   1605
         VariousPropertyBits=   27
         Caption         =   "123"
         Size            =   "2831;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   285
         Index           =   27
         Left            =   11205
         TabIndex        =   85
         Top             =   3720
         Visible         =   0   'False
         Width           =   1590
         VariousPropertyBits=   27
         Caption         =   "123"
         Size            =   "2805;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         Caption         =   "2."
         Height          =   180
         Index           =   25
         Left            =   5400
         TabIndex        =   84
         Top             =   4005
         Width           =   345
      End
      Begin VB.Label Label1 
         Caption         =   "繪圖人員："
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   83
         Top             =   420
         Width           =   885
      End
      Begin VB.Label Label1 
         Caption         =   "承辦人："
         Height          =   195
         Index           =   8
         Left            =   120
         TabIndex        =   82
         Top             =   3195
         Width           =   765
      End
      Begin VB.Label Label1 
         Caption         =   "國外案承辦人："
         Height          =   195
         Index           =   9
         Left            =   120
         TabIndex        =   81
         Top             =   4575
         Width           =   1305
      End
      Begin VB.Label Label1 
         Caption         =   "國外案本所案號："
         Height          =   195
         Index           =   10
         Left            =   120
         TabIndex        =   80
         Top             =   4860
         Width           =   1485
      End
      Begin VB.Label Label1 
         Caption         =   "點數："
         Height          =   195
         Index           =   11
         Left            =   120
         TabIndex        =   79
         Top             =   2355
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "智權人員："
         Height          =   195
         Index           =   12
         Left            =   120
         TabIndex        =   78
         Top             =   3465
         Width           =   945
      End
      Begin VB.Label Label1 
         Caption         =   "法定期限："
         Height          =   195
         Index           =   13
         Left            =   120
         TabIndex        =   77
         Top             =   2910
         Width           =   945
      End
      Begin VB.Label Label1 
         Caption         =   "本所期限："
         Height          =   195
         Index           =   14
         Left            =   120
         TabIndex        =   76
         Top             =   2640
         Width           =   945
      End
      Begin VB.Label Label1 
         Caption         =   "案件性質："
         Height          =   195
         Index           =   15
         Left            =   120
         TabIndex        =   75
         Top             =   2085
         Width           =   945
      End
      Begin VB.Label Label1 
         Caption         =   "專利/商標種類："
         Height          =   195
         Index           =   16
         Left            =   120
         TabIndex        =   74
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "案件名稱："
         Height          =   195
         Index           =   18
         Left            =   120
         TabIndex        =   73
         Top             =   1530
         Width           =   945
      End
      Begin VB.Label Label1 
         Caption         =   "本所案號："
         Height          =   195
         Index           =   19
         Left            =   120
         TabIndex        =   72
         Top             =   1245
         Width           =   945
      End
      Begin VB.Label Label1 
         Caption         =   "收文日："
         Height          =   195
         Index           =   20
         Left            =   120
         TabIndex        =   71
         Top             =   975
         Width           =   765
      End
      Begin VB.Label Label1 
         Caption         =   "總收文號："
         Height          =   195
         Index           =   21
         Left            =   120
         TabIndex        =   70
         Top             =   690
         Width           =   945
      End
      Begin VB.Label Label1 
         Caption         =   "取消收文日："
         Height          =   195
         Index           =   29
         Left            =   120
         TabIndex        =   69
         Top             =   4305
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "草圖作業天數："
         Height          =   195
         Index           =   33
         Left            =   120
         TabIndex        =   68
         Top             =   3750
         Width           =   1305
      End
      Begin VB.Label Label1 
         Caption         =   "墨圖作業天數："
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   67
         Top             =   4020
         Width           =   1305
      End
      Begin VB.Label Label1 
         Caption         =   "(N：不計)"
         Height          =   180
         Index           =   35
         Left            =   6480
         TabIndex        =   65
         Top             =   420
         Width           =   1065
      End
      Begin MSForms.Label lbl1 
         Height          =   285
         Index           =   29
         Left            =   11205
         TabIndex        =   64
         Top             =   4275
         Visible         =   0   'False
         Width           =   1590
         VariousPropertyBits=   27
         Caption         =   "29"
         Size            =   "2805;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   285
         Index           =   28
         Left            =   11205
         TabIndex        =   63
         Top             =   4005
         Visible         =   0   'False
         Width           =   1590
         VariousPropertyBits=   27
         Caption         =   "123"
         Size            =   "2805;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         Caption         =   "3."
         Height          =   180
         Index           =   27
         Left            =   5400
         TabIndex        =   45
         Top             =   4275
         Width           =   345
      End
      Begin MSForms.Label lbl1 
         Height          =   285
         Index           =   26
         Left            =   11205
         TabIndex        =   44
         Top             =   3450
         Visible         =   0   'False
         Width           =   1590
         VariousPropertyBits=   27
         Caption         =   "123"
         Size            =   "2805;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         Caption         =   "墨圖："
         Height          =   180
         Index           =   24
         Left            =   5340
         TabIndex        =   43
         Top             =   3450
         Width           =   555
      End
      Begin VB.Label Label1 
         Caption         =   "草圖："
         Height          =   180
         Index           =   5
         Left            =   5340
         TabIndex        =   42
         Top             =   3165
         Width           =   555
      End
      Begin MSForms.Label lbl1 
         Height          =   285
         Index           =   21
         Left            =   11205
         TabIndex        =   41
         Top             =   2070
         Visible         =   0   'False
         Width           =   1590
         VariousPropertyBits=   27
         Caption         =   "123"
         Size            =   "2805;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         Caption         =   "墨圖張數："
         Height          =   195
         Index           =   34
         Left            =   4440
         TabIndex        =   40
         Top             =   2625
         Width           =   945
      End
      Begin VB.Label Label1 
         Caption         =   "墨圖完稿日："
         Height          =   195
         Index           =   28
         Left            =   4440
         TabIndex        =   39
         Top             =   2340
         Width           =   1125
      End
      Begin VB.Label Label1 
         Caption         =   "墨圖齊備日："
         Height          =   195
         Index           =   2
         Left            =   4440
         TabIndex        =   38
         Top             =   2070
         Width           =   1125
      End
      Begin VB.Label Label1 
         Caption         =   "草圖是否計件："
         Height          =   195
         Index           =   17
         Left            =   4440
         TabIndex        =   37
         Top             =   420
         Width           =   1275
      End
      Begin VB.Label Label1 
         Caption         =   "草圖齊備日："
         Height          =   195
         Index           =   23
         Left            =   4440
         TabIndex        =   36
         Top             =   690
         Width           =   1125
      End
      Begin VB.Label Label1 
         Caption         =   "草圖完稿日："
         Height          =   195
         Index           =   26
         Left            =   4440
         TabIndex        =   35
         Top             =   975
         Width           =   1125
      End
      Begin VB.Label Label1 
         Caption         =   "草圖張數："
         Height          =   195
         Index           =   3
         Left            =   4440
         TabIndex        =   34
         Top             =   1245
         Width           =   945
      End
      Begin VB.Label Label1 
         Caption         =   "修改時數："
         Height          =   180
         Index           =   6
         Left            =   4440
         TabIndex        =   33
         Top             =   3720
         Width           =   930
      End
      Begin VB.Label Label1 
         Caption         =   "承辦時數："
         Height          =   180
         Index           =   22
         Left            =   4440
         TabIndex        =   32
         Top             =   3165
         Width           =   930
      End
      Begin VB.Label Label1 
         Caption         =   "備註："
         Height          =   180
         Index           =   31
         Left            =   4440
         TabIndex        =   31
         Top             =   4560
         Width           =   690
      End
      Begin MSForms.Label lbl1 
         Height          =   285
         Index           =   18
         Left            =   11205
         TabIndex        =   30
         Top             =   690
         Visible         =   0   'False
         Width           =   1590
         VariousPropertyBits=   27
         Caption         =   "123"
         Size            =   "2805;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   285
         Index           =   19
         Left            =   11205
         TabIndex        =   29
         Top             =   975
         Visible         =   0   'False
         Width           =   1590
         VariousPropertyBits=   27
         Caption         =   "123"
         Size            =   "2805;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   285
         Index           =   24
         Left            =   11205
         TabIndex        =   28
         Top             =   420
         Visible         =   0   'False
         Width           =   1770
         VariousPropertyBits=   27
         Caption         =   "123"
         Size            =   "3122;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   285
         Index           =   22
         Left            =   11205
         TabIndex        =   27
         Top             =   2340
         Visible         =   0   'False
         Width           =   1590
         VariousPropertyBits=   27
         Caption         =   "123"
         Size            =   "2805;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   285
         Index           =   23
         Left            =   11205
         TabIndex        =   26
         Top             =   2625
         Visible         =   0   'False
         Width           =   1590
         VariousPropertyBits=   27
         Caption         =   "123"
         Size            =   "2805;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   285
         Index           =   20
         Left            =   11205
         TabIndex        =   25
         Top             =   1245
         Visible         =   0   'False
         Width           =   1590
         VariousPropertyBits=   27
         Caption         =   "123"
         Size            =   "2805;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   285
         Index           =   25
         Left            =   11205
         TabIndex        =   24
         Top             =   3165
         Visible         =   0   'False
         Width           =   1590
         VariousPropertyBits=   27
         Caption         =   "123"
         Size            =   "2805;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   270
         Index           =   30
         Left            =   12345
         TabIndex        =   23
         Top             =   4560
         Visible         =   0   'False
         Width           =   1350
         VariousPropertyBits=   27
         Caption         =   "123"
         Size            =   "2381;476"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         Caption         =   "1."
         Height          =   180
         Index           =   30
         Left            =   5400
         TabIndex        =   22
         Top             =   3720
         Width           =   345
      End
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "PS：舊案之後續若需繪圖，系統會自動做繪圖分案確認！"
      Height          =   180
      Left            =   5640
      TabIndex        =   99
      Top             =   120
      Width           =   4500
   End
End
Attribute VB_Name = "frm090712"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2022/1/17 改成Form2.0 (grd1,txt1,lbl1)
'Memo By Morgan 2012/12/10 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/17 日期欄已修改
Option Explicit
Public TextOk As Boolean, StrGrp090711 As String
Dim strSql As String, strSQL1 As String, i As Integer, j As Integer, s As Integer, SavDay1 As String, SavDay2 As String, SavDay3 As String, StrSQL7 As String, StrSQL4 As String, strSQL5 As String
Dim strSQL2 As String, iPrint As Integer, Page As Integer, strTemp(0 To 21) As String, strTemp3 As String, TestOk As Boolean, StrTemp99(0 To 21) As String
Dim PLeft(0 To 21) As Integer, strTemp1 As Variant, strTemp2 As Variant, StrSQL6 As String, Str020401SysKind As String, Seekok As Integer, SeekTemp As Integer, strDate1 As String, strDate2 As String
Dim ChkData2 As Boolean, SWPRow As String, strCP10 As String, k As Integer, ChkNoData As Boolean, TXT090711 As TextBox
Dim NickRS As ADODB.Recordset, StrColor1 As String, StrColor2 As String, StrColor3 As String, StrColor4 As String, StrColor5 As String, StrColor6 As String
Dim ll As Integer
Dim m_blnColOrderAsc As Boolean '欄位資料由小到大排序
Dim m_CP10 As String 'Add by Morgan 2009/11/4

Sub Process(strText As String)
'代第2畫面資料
With grd1
    '收文號
'    .col = 22
    .col = 24
    strSql = "SELECT CM01||'-'||CM02||'-'||CM03||'-'||CM04,ST02 FROM CASEMAP,CASEPROGRESS,STAFF WHERE CP01=CM01(+) AND CP02=CM02(+) AND CP03=CM03(+) AND CP04=CM04(+) AND CP14=ST01(+) AND CP31='Y' and cp09='" & .Text & "' order by CM01||'-'||CM02||'-'||CM03||'-'||CM04 "
    CheckOC2
    adoRecordset1.CursorLocation = adUseClient
    adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If adoRecordset1.RecordCount <> 0 And adoRecordset1.RecordCount > 0 Then
        '國外案本所案號
        lbl1(16) = CheckStr(adoRecordset1.Fields(0))
        '國外案承辦人
        lbl1(15) = CheckStr(adoRecordset1.Fields(1))
    Else
        lbl1(16) = ""
        lbl1(15) = ""
    End If
    CheckOC2
    '繪圖人員
    .col = 36
    'add by nickc 2006/05/29 紀錄舊繪圖人員
    txt1(0).Tag = .Text
    .col = 36
    txt1(0).Text = .Text
    .col = 8
    lbl1(0).Caption = .Text
    '收文號
'    .col = 22
    .col = 24
    lbl1(1).Caption = .Text
    '收文日
    .col = 2
    lbl1(2).Caption = .Text
    '本所案號
    .col = 3
    lbl1(3).Caption = .Text
    '是否閉卷
    If Right("" & .Text, 1) = "＊" Then
        Me.lblClose.Caption = "已閉卷"
    Else
        Me.lblClose.Caption = ""
    End If
    '案件名稱
    .col = 4
    lbl1(4).Caption = .Text
    '案件性質名稱
'    .col = 23
    .col = 25
    lbl1(5).Caption = .Text
    '案件性質
    .col = 5
    lbl1(6).Caption = .Text
    '點數
    .col = 7
    lbl1(7).Caption = .Text
    '本所期限
'    .col = 18
    .col = 20
    lbl1(8).Caption = .Text
    '法定期限
'    .col = 24
    .col = 26
    lbl1(9).Caption = .Text
    '承辦人
    .col = 6
    lbl1(10).Caption = .Text
    '智權人員
'    .col = 21
    .col = 23
    lbl1(11).Caption = .Text
    '取消收文日
'    .col = 25
    .col = 27
    lbl1(14).Caption = .Text
    '草齊日
'    .col = 9
    .col = 11
    If .Text = "******" Then
        txt1(1) = .Text
    Else
        txt1(1) = ChangeTDateStringToTString(.Text)
    End If
    '草圖承辦期限
'    .col = 10
    .col = 12
    Me.lbl1(31).Caption = .Text
    '草完日
'    .col = 11
    .col = 13
    If .Text = "******" Then
        txt1(2) = .Text
    Else
        txt1(2) = ChangeTDateStringToTString(.Text)
    End If
    '草圖張數
'    .col = 31
    .col = 33
    txt1(3) = .Text
    '墨齊日
'    .col = 14
    .col = 16
    If .Text = "******" Then
        txt1(4) = .Text
    Else
        txt1(4) = ChangeTDateStringToTString(.Text)
    End If
    '墨圖承辦期限
'    .col = 15
    .col = 17
    Me.lbl1(32).Caption = .Text
    '墨完日
'    .col = 16
    .col = 18
    If .Text = "******" Then
        txt1(5) = .Text
    Else
        txt1(5) = ChangeTDateStringToTString(.Text)
    End If
    If Len(txt1(1)) <> 0 And Len(txt1(2)) <> 0 Then
        lbl1(12).Caption = GetWorkDay(ChangeTStringToWString(txt1(2)), ChangeTStringToWString(txt1(1)))
    Else
        lbl1(12).Caption = "0"
    End If
    If Len(txt1(4)) <> 0 And Len(txt1(5)) <> 0 Then
        lbl1(13).Caption = GetWorkDay(ChangeTStringToWString(txt1(5)), ChangeTStringToWString(txt1(4)))
    Else
        lbl1(13).Caption = "0"
    End If
    '墨圖張數
'    .col = 32
    .col = 34
    txt1(6) = .Text
    '草計
'    .col = 8
    .col = 10
    txt1(7) = .Text
    '墨計
'    .col = 13
    .col = 15
    txt1(14) = .Text
    '草圖承辦時數
'    .col = 26
    .col = 28
    txt1(8) = .Text
    '墨圖承辦時數
'    .col = 27
    .col = 29
    txt1(9) = .Text
    '修改時數1
'    .col = 28
    .col = 30
    txt1(10) = .Text
    '修改時數2
'    .col = 29
    .col = 31
    txt1(11) = .Text
    '修改時數3
'    .col = 30
    .col = 32
    txt1(12) = .Text
    '備註
'    .col = 20
    .col = 22
    txt1(13) = .Text
    'Add by Morgan 2009/11/4
    m_CP10 = .TextMatrix(.row, 35)
End With
End Sub

Private Sub cmdData_Click(Index As Integer)
Select Case Index
Case 0
         cmdData(0).Enabled = False
         cmdData(1).Enabled = True
         cmdData(2).Enabled = True
Case 1
         cmdData(0).Enabled = True
         cmdData(1).Enabled = False
         cmdData(2).Enabled = True
Case 2
         cmdData(0).Enabled = True
         cmdData(1).Enabled = True
         cmdData(2).Enabled = False
Case Else
End Select
StrMenu
End Sub

Private Sub cmdOK_Click(Index As Integer)
Dim ii As Integer
'add by nickc 2005/03/07
Dim BolIsUpdate As Boolean
Select Case Index
'add by nickc 2005/03/07   加入主管確認註記
Case 0 '確認分案
'      BolIsUpdate = False
      For ii = 1 To grd1.Rows - 1
           grd1.row = ii
           grd1.col = 0
           If Trim(grd1.Text) = "V" Then
'               BolIsUpdate = True
               'add by nickc 2008/01/22 加入判斷，若是已經上了會完日，要提醒
               If grd1.TextMatrix(ii, 38) <> "" And Val(Me.grd1.TextMatrix(ii, 38)) <> 0 Then
                    If Me.grd1.TextMatrix(ii, 10) = "" And Me.grd1.TextMatrix(ii, 11) = "" And Val(Me.grd1.TextMatrix(ii, 11)) = 0 And Me.grd1.TextMatrix(ii, 15) = "" And Me.grd1.TextMatrix(ii, 16) = "" And Val(Me.grd1.TextMatrix(ii, 16)) = 0 Then
                        MsgBox "注意！" & vbCrLf & "此 " & Me.grd1.TextMatrix(ii, 3) & " 已有會稿完成日，卻未上任何齊備日，可能最近工程師才上''繪圖人員'' 或國內外關聯最近才建立，注意期限！！", vbExclamation, "請抄下案號，特別注意！！"
                    End If
                    If Me.grd1.TextMatrix(ii, 10) = "N" And Me.grd1.TextMatrix(ii, 15) = "" And Me.grd1.TextMatrix(ii, 16) = "" And Val(Me.grd1.TextMatrix(ii, 16)) = 0 Then
                        MsgBox "注意！" & vbCrLf & "此 " & Me.grd1.TextMatrix(ii, 3) & " 已有會稿完成日，卻未上墨圖齊備日，可能最近工程師才上''繪圖人員'' 或國內外關聯最近才建立，注意期限！！", vbExclamation, "請抄下案號，特別注意！！"
                    End If
                    
               End If
               grd1.col = 24
               strSql = "update caseprogress set cp107='Y' where cp09='" & grd1.Text & "' "
               cnnConnection.Execute strSql
         End If
      Next ii
'      If BolIsUpdate = True Then
         Screen.MousePointer = vbHourglass
         MoveFormToCenter Me
         Me.lblClose.Caption = ""
         SSTab1.Tab = 0
         TextOk = True
         StrMenu
         Option1(0).Value = True
         Combo2.Clear
         Combo2.AddItem "速件", 0
         Combo2.AddItem "未齊備", 1
         Combo2.AddItem "複雜", 2
         Combo2.AddItem "其他新案", 3
         Combo2.AddItem "ACAD", 4
         Combo2.Text = "請選擇...."
         Combo2.Enabled = False
         Me.Combo3.ListIndex = 0
         MouseClick (1)
         Screen.MousePointer = vbDefault
'      End If
Case 1 '回前畫面
    Unload Me
Case 2 '存檔
    If SSTab1.Tab = 1 Then
        If ChkNoData = False Then
            ChkData2 = True
'            For Each TXT090711 In txt1
'                If txt1(TXT090711.Index).Visible = True And txt1(TXT090711.Index).Enabled = True Then
'                   txt1_LostFocus (TXT090711.Index)
'                   If ChkData2 = False Then Exit Sub
'                End If
'            Next

            If txt1(0).Enabled = True Then
               If Trim(txt1(0)) = "" Then
                  If MsgBox("沒有繪圖人員，請確定！", vbYesNo, "警告！") = vbNo Then
                     Exit Sub
                  End If
               End If
            End If

            Screen.MousePointer = vbHourglass
            Me.Enabled = False
            Me.txt1(1).Text = Replace(Me.txt1(1).Text, "******", "")
            Me.txt1(2).Text = Replace(Me.txt1(2).Text, "******", "")
            Me.txt1(4).Text = Replace(Me.txt1(4).Text, "******", "")
            Me.txt1(5).Text = Replace(Me.txt1(5).Text, "******", "")
            On Error GoTo ErrorHandler
            cnnConnection.BeginTrans
'2009/9/11 cancel by sonia EP13由trigger更新
'            strSQL = "UPDATE ENGINEERPROGRESS SET EP13='" & txt1(0) & "' WHERE EP02='" & lbl1(1).Caption & "' "
'            cnnConnection.Execute strSQL
'2009/9/11 end
            'Add By Cheng 2003/11/18
            '更新進度檔的繪圖人員
            strSql = "Update CaseProgress Set CP29='" & Me.txt1(0).Text & "' WHERE CP09='" & lbl1(1).Caption & "' "
            cnnConnection.Execute strSql
            'add by nickc 2006/05/29 若是有更改繪圖人員時，一併更改其他關聯案
            If Me.txt1(0).Tag <> Me.txt1(0).Text Then
               'Add by Morgan 2009/11/4 條件要和承辦人工作進度資料維護一樣
               If InStr(lbl1(3).Caption, "P") > 0 And InStr(NewCasePtyList, m_CP10) > 0 Then
               'end 2009/11/4
                  'edit by nickc 2006/07/04 因為之前太慢，修正語法
                  'strSQL = "UPDATE caseprogress set cp29='" & Me.txt1(0).Text & "' where cp01||cp02||cp03||cp04 in (select cm01||cm02||cm03||cm04 from casemap where cm05||'-'||cm06||'-'||cm07||'-'||cm08='" & lbl1(3).Caption & "' and cm10='0'  union select cm05||cm06||cm07||cm08 from casemap where cm01||'-'||cm02||'-'||cm03||'-'||cm04='" & lbl1(3).Caption & "' and cm10='0' )"
                  '2009/9/11 modify by sonia casemap的國內案改繪圖人員,其他關聯案才可一併修改且僅限於新申請案,國外案修改時不可改回國內案CFP-021836
                  'strSQL = "UPDATE caseprogress set cp29='" & Me.txt1(0).Text & "' where cp09 in (select cp09 from casemap,caseprogress where cm05||'-'||cm06||'-'||cm07||'-'||cm08='" & lbl1(3).Caption & "' and cm10='0' and cm01=cp01(+) and cm02=cp02(+) and cm03=cp03(+) and cm04=cp04(+)  union select cp09 from casemap,caseprogress where cm01||'-'||cm02||'-'||cm03||'-'||cm04='" & lbl1(3).Caption & "' and cm10='0'  and cm05=cp01(+) and cm06=cp02(+) and cm07=cp03(+) and cm08=cp04(+) ) "
                  strSql = "UPDATE caseprogress set cp29='" & Me.txt1(0).Text & "' where cp09 in (select cp09 from casemap,caseprogress where cm05||'-'||cm06||'-'||cm07||'-'||cm08='" & lbl1(3).Caption & "' and cm10='0' and cm01=cp01(+) and cm02=cp02(+) and cm03=cp03(+) and cm04=cp04(+) and instr('" & NewCasePtyList & "',cp10)>0) and cp27 is null "
                  '2009/9/11 END
                  cnnConnection.Execute strSql
               End If
               
               'Add by Morgan 2009/11/4
               '多國案的主案才更新其他多國案未發文且繪圖主管未分案確認之新申請程序(參照承辦人工作進度資料維護)
               If SystemNumber(lbl1(3).Caption, 1) = "CFP" And InStr(NewCasePtyList, m_CP10) > 0 Then
                  strSql = "update CaseProgress set Cp29=" & CNULL(Me.txt1(0).Text) & _
                     " Where Cp09 in (select C1.cp09 from CASEPROGRESS C2,caseprogress C1,caserelation,engineerprogress" & _
                     " where C2.cp09='" & lbl1(1) & "' AND C2.CP21 IS NULL AND cr01(+)=C2.CP01 and cr02(+)=C2.CP02 and cr03(+)=C2.CP03 and cr04(+)=C2.CP04" & _
                     " and C1.cp01(+)=cr05 and C1.cp02(+)=cr06 and C1.cp03(+)=cr07 and C1.cp04(+)=cr08 and C1.cp21='Y'" & _
                     " and instr('" & NewCasePtyList & "',C1.cp10)>0 and C1.cp27 is null and C1.cp107 is null" & _
                     " and ep02(+)=C1.cp09 and ep14 is null )"
                  cnnConnection.Execute strSql, intI
               End If
            End If
            cnnConnection.CommitTrans
'            grd1.Clear
'            grd1.Rows = 2
'            StrMenu
'            SetGrd1
            'Add By Cheng 2004/03/22
            '用收文號尋找瀏覽資料
            For ii = 1 To Me.grd1.Rows - 1
                If Me.lbl1(1).Caption = Me.grd1.TextMatrix(ii, 23) Then
                    SWPRow = ii
                    Exit For
                End If
            Next ii
            'End
            If SWPRow >= 1 Then
                '更新繪圖人員到前畫面
                Me.grd1.TextMatrix(SWPRow, 36) = Me.txt1(0).Text
                Me.grd1.TextMatrix(SWPRow, 8) = Me.lbl1(0).Caption
            End If
        End If
        MouseClick IIf(Val("0" & SWPRow) < 1, 1, SWPRow)
        SSTab1.Tab = 0
        'Modify By Cheng 2004/04/19
'        cmdOK(2).Caption = "確定(&O)"
        cmdok(2).Caption = "確定"
        'End
        Me.Enabled = True
        Screen.MousePointer = vbDefault
    Else
        SSTab1.Tab = 1
        'Modify By Cheng 2004/04/19
'        cmdOK(2).Caption = "存檔(&O)"
        cmdok(2).Caption = "存檔"
        'End
    End If
Case Else
End Select
Exit Sub
ErrorHandler:
    cnnConnection.RollbackTrans
    MsgBox Err.Description
    Me.Enabled = True
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
Screen.MousePointer = vbHourglass
MoveFormToCenter Me
Me.lblClose.Caption = ""
SSTab1.Tab = 0
TextOk = True
'判斷所別，鎖按鈕
If PUB_GetST06(strUserNum) = "1" Then
   cmdData(0).Enabled = False
ElseIf PUB_GetST06(strUserNum) = "2" Then
   cmdData(1).Enabled = False
ElseIf PUB_GetST06(strUserNum) = "3" Or PUB_GetST06(strUserNum) = "4" Then
   cmdData(2).Enabled = False
Else  '其他算北所
   cmdData(0).Enabled = False
End If
StrMenu
Option1(0).Value = True
Combo2.Clear
Combo2.AddItem "速件", 0
Combo2.AddItem "未齊備", 1
Combo2.AddItem "複雜", 2
Combo2.AddItem "其他新案", 3
Combo2.AddItem "ACAD", 4
Combo2.Text = "請選擇...."
Combo2.Enabled = False
Me.Combo3.ListIndex = 0
MouseClick (1)
Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frm090712 = Nothing
End Sub

Sub StrMenu()       '代資料  當月資料
Screen.MousePointer = vbHourglass
grd1.MousePointer = flexHourglass
'Modify by Morgan 2004/6/9 CFP 的 105 也要
'Modify by Morgan 2004/5/6
'案件性質抓 101-105,109,110,112-115 ; CFP 的 105 除外

''Modify By Cheng 2003/07/22
''strSQL = "SELECT SUBSTR(CP09,1,1)," & SQLDate("CP05") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(PA57,'Y','＊',''),NVL(PA05,NVL(PA06,PA07)), Decode(pa09,'000',cpm03,cpm04),s1.st02, ROUND(cp18,2), EP13, DECODE(EP06,NULL,' ',0,' '," & SQLDate("EP06") & "), EP20, DECODE(EP14,NULL,'',0,''," & SQLDate("eP14") & "),'' As 草期限, " & SQLDate("eP15") & ",0, EP29, DECODE(EP17,NULL,DECODE(EP08,NULL,'',0,''," & SQLDate("EP08") & "),0,DECODE(EP08,NULL,'',0,''," & SQLDate("EP08") & ")," & SQLDate("eP17") & "), '' As 墨期限, " & SQLDate("EP18") & ",0," & SQLDate("cP06") & "," & SQLDate("cP27") & ",ep26,s3.st02,CP09,DECODE(PA09,'000',PTM03,PTM04)," & SQLDate("CP07") & "," & SQLDate("CP57") & ",ep21,ep22,ep23,ep24,ep25,ep16,ep19, CP10, S2.ST02 " & _
''            " FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,PATENT,CASEPROPERTYMAP,PATENTTRADEMARKMAP, Staff S4 " & _
''            " WHERE EP02=CP09(+) And CP01 In ('P','CFP','FCP') AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP14=S1.ST01(+) AND EP13=S2.ST01(+) AND CP13=S3.ST01(+) and CP01=CPM01(+) and CP10=CPM02(+) AND '1'=PTM01(+)  AND PA08=PTM02(+) And EP14 Is Not Null And EP13 Is Null And S4.ST06=S1.ST06 And S4.ST01='" & strUserNum & "' "
''93/2/7 MODIFY BY SONIA
''strSQL = "SELECT SUBSTR(CP09,1,1)," & SQLDate("CP05") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(PA57,'Y','＊',''),NVL(PA05,NVL(PA06,PA07)), Decode(pa09,'000',cpm03,cpm04),s1.st02, ROUND(cp18,2), EP13, DECODE(EP06,NULL,' ',0,' '," & SQLDate("EP06") & "), EP20, DECODE(EP14,NULL,'',0,''," & SQLDate("eP14") & "),'' As 草期限, " & SQLDate("eP15") & ",0, EP29, DECODE(EP17,NULL,DECODE(EP08,NULL,'',0,''," & SQLDate("EP08") & "),0,DECODE(EP08,NULL,'',0,''," & SQLDate("EP08") & ")," & SQLDate("eP17") & "), '' As 墨期限, " & SQLDate("EP18") & ",0," & SQLDate("cP06") & "," & SQLDate("cP27") & ",ep26,s3.st02,CP09,DECODE(PA09,'000',PTM03,PTM04)," & SQLDate("CP07") & "," & SQLDate("CP57") & ",ep21,ep22,ep23,ep24,ep25,ep16,ep19, CP10, S2.ST02 " & _
''            " FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,PATENT,CASEPROPERTYMAP,PATENTTRADEMARKMAP, Staff S4 " & _
''            " WHERE EP02=CP09(+) And CP01 In ('P','CFP','FCP') AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP14=S1.ST01(+) AND EP13=S2.ST01(+) AND CP13=S3.ST01(+) and CP01=CPM01(+) and CP10=CPM02(+) AND '1'=PTM01(+)  AND PA08=PTM02(+) And EP06 Is Not Null And EP13 Is Null And S4.ST06=S1.ST06 And S4.ST01='" & strUserNum & "' "
''若使用者為南所人員, 可看到高所的資料
'If PUB_GetST06(strUserNum) = "3" Then
'    strSQL = "SELECT SUBSTR(CP09,1,1)," & SQLDate("CP05") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(PA57,'Y','＊',''),NVL(PA05,NVL(PA06,PA07)), Decode(pa09,'000',cpm03,cpm04),s1.st02, ROUND(cp18,2), EP13, DECODE(EP06,NULL,' ',0,' '," & SQLDate("EP06") & "), EP20, DECODE(EP14,NULL,'',0,''," & SQLDate("eP14") & "),'' As 草期限, " & SQLDate("eP15") & ",0, EP29, " & SQLDate("eP17") & ", '' As 墨期限, " & SQLDate("EP18") & ",0," & SQLDate("cP06") & "," & SQLDate("cP27") & ",ep26,s3.st02,CP09,DECODE(PA09,'000',PTM03,PTM04)," & SQLDate("CP07") & "," & SQLDate("CP57") & ",ep21,ep22,ep23,ep24,ep25,ep16,ep19, CP10, S2.ST02 " & _
'                " FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,PATENT,CASEPROPERTYMAP,PATENTTRADEMARKMAP " & _
'                " WHERE EP02=CP09(+) And CP01 In ('P','CFP','FCP') AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP14=S1.ST01(+) AND EP13=S2.ST01(+) AND CP13=S3.ST01(+) and CP01=CPM01(+) and CP10=CPM02(+) AND '1'=PTM01(+)  AND PA08=PTM02(+) And EP06 Is Not Null And EP13 Is Null And EP10 IS NULL And S1.ST06 In ('3', '4') And CP57 Is Null And CP10 In ('101','102','103','104','105') "
''若使用者非南所人員, 只可看到自所所別的資料
'Else
'    strSQL = "SELECT SUBSTR(CP09,1,1)," & SQLDate("CP05") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(PA57,'Y','＊',''),NVL(PA05,NVL(PA06,PA07)), Decode(pa09,'000',cpm03,cpm04),s1.st02, ROUND(cp18,2), EP13, DECODE(EP06,NULL,' ',0,' '," & SQLDate("EP06") & "), EP20, DECODE(EP14,NULL,'',0,''," & SQLDate("eP14") & "),'' As 草期限, " & SQLDate("eP15") & ",0, EP29, " & SQLDate("eP17") & ", '' As 墨期限, " & SQLDate("EP18") & ",0," & SQLDate("cP06") & "," & SQLDate("cP27") & ",ep26,s3.st02,CP09,DECODE(PA09,'000',PTM03,PTM04)," & SQLDate("CP07") & "," & SQLDate("CP57") & ",ep21,ep22,ep23,ep24,ep25,ep16,ep19, CP10, S2.ST02 " & _
'                " FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,PATENT,CASEPROPERTYMAP,PATENTTRADEMARKMAP, Staff S4 " & _
'                " WHERE EP02=CP09(+) And CP01 In ('P','CFP','FCP') AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP14=S1.ST01(+) AND EP13=S2.ST01(+) AND CP13=S3.ST01(+) and CP01=CPM01(+) and CP10=CPM02(+) AND '1'=PTM01(+)  AND PA08=PTM02(+) And EP06 Is Not Null And EP13 Is Null And EP10 IS NULL And S4.ST06=S1.ST06 And '" & strUserNum & "'=S4.ST01 And CP57 Is Null And CP10 In ('101','102','103','104','105') "
'End If

'add by nickc 2005/04/12
'edit by nickc 2005/05/04 皆不管文齊日
'    strSQL = "SELECT '',SUBSTR(CP09,1,1)," & SQLDate("CP05") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(PA57,'Y','＊',''),NVL(PA05,NVL(PA06,PA07)), Decode(pa09,'000',cpm03,cpm04),s1.st02, ROUND(cp18,2), S2.ST02, DECODE(EP06,NULL,' ',0,' '," & SQLDate("EP06") & "), EP20, DECODE(EP14,NULL,'',0,''," & SQLDate("eP14") & "),'' As 草期限, " & SQLDate("eP15") & ",0, EP29, " & SQLDate("eP17") & ", '' As 墨期限, " & SQLDate("EP18") & ",0," & SQLDate("cP06") & "," & SQLDate("cP27") & ",ep26,s3.st02,CP09,DECODE(PA09,'000',PTM03,PTM04)," & SQLDate("CP07") & "," & SQLDate("CP57") & ",ep21,ep22,ep23,ep24,ep25,ep16,ep19, CP10, ep13, PA08 " & _
                " FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,PATENT,CASEPROPERTYMAP,PATENTTRADEMARKMAP " & _
                " WHERE EP02=CP09(+) And CP01 In ('P','CFP','FCP') AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP14=S1.ST01(+) AND EP13=S2.ST01(+) AND CP13=S3.ST01(+) and CP01=CPM01(+) and CP10=CPM02(+) AND '1'=PTM01(+)  AND PA08=PTM02(+) And EP06 Is Not Null  And EP10 IS NULL  And CP57 Is Null  and cp107 is null and (ep20 is null or ep29 is null)  and ep13 is not null "
'edit by nickc 2008/01/22
'    strSQL = "SELECT '',SUBSTR(CP09,1,1)," & SQLDate("CP05") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(PA57,'Y','＊',''),NVL(PA05,NVL(PA06,PA07)), Decode(pa09,'000',cpm03,cpm04),s1.st02, ROUND(cp18,2), S2.ST02, DECODE(EP06,NULL,' ',0,' '," & SQLDate("EP06") & "), EP20, DECODE(EP14,NULL,'',0,''," & SQLDate("eP14") & "),'' As 草期限, " & SQLDate("eP15") & ",0, EP29, " & SQLDate("eP17") & ", '' As 墨期限, " & SQLDate("EP18") & ",0," & SQLDate("cP06") & "," & SQLDate("cP27") & ",ep26,s3.st02,CP09,DECODE(PA09,'000',PTM03,PTM04)," & SQLDate("CP07") & "," & SQLDate("CP57") & ",ep21,ep22,ep23,ep24,ep25,ep16,ep19, CP10, ep13, PA08 " & _
                " FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,PATENT,CASEPROPERTYMAP,PATENTTRADEMARKMAP " & _
                " WHERE EP02=CP09(+) And CP01 In ('P','CFP','FCP') AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP14=S1.ST01(+) AND EP13=S2.ST01(+) AND CP13=S3.ST01(+) and CP01=CPM01(+) and CP10=CPM02(+) AND '1'=PTM01(+)  AND PA08=PTM02(+) And EP10 IS NULL  And CP57 Is Null  and cp107 is null and (ep20 is null or ep29 is null)  and ep13 is not null "
    
    'Modify by Morgan 2010/8/17 百年蟲 " & SQLDate("CP05") & "-->substrb(' '||sqldatet(cp05),-9)
    strSql = "SELECT '',SUBSTR(CP09,1,1),substrb(' '||sqldatet(cp05),-9),CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(PA57,'Y','＊',''),NVL(PA05,NVL(PA06,PA07)), Decode(pa09,'000',cpm03,cpm04),s1.st02, ROUND(cp18,2), S2.ST02, DECODE(EP06,NULL,' ',0,' '," & SQLDate("EP06") & "), EP20, DECODE(EP14,NULL,'',0,''," & SQLDate("eP14") & "),'' As 草期限, " & SQLDate("eP15") & ",0, EP29, " & SQLDate("eP17") & ", '' As 墨期限, " & SQLDate("EP18") & ",0," & SQLDate("cP06") & "," & SQLDate("cP27") & ",ep26,s3.st02,CP09,DECODE(PA09,'000',PTM03,PTM04)," & SQLDate("CP07") & "," & SQLDate("CP57") & ",ep21,ep22,ep23,ep24,ep25,ep16,ep19, CP10, ep13, PA08,ep08 " & _
                " FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,PATENT,CASEPROPERTYMAP,PATENTTRADEMARKMAP " & _
                " WHERE EP02=CP09(+) And CP01 In ('P','CFP','FCP') AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP14=S1.ST01(+) AND EP13=S2.ST01(+) AND CP13=S3.ST01(+) and CP01=CPM01(+) and CP10=CPM02(+) AND '1'=PTM01(+)  AND PA08=PTM02(+) And EP10 IS NULL  And CP57 Is Null  and cp107 is null and (ep20 is null or ep29 is null)  and ep13 is not null "

'add by nickc 2006/06/13 瓊玉說若承辦人是程序，就不要出來
'edit by nickc 2006/09/08
'strSQL = strSQL & " and s1.st15<>'P12' "
strSql = strSql & " and (s1.st15<>'P12' or s1.st15 is null ) "

'若使用者為南所人員, 可看到高所的資料
'edit by nickc 2005/04/12 改可查其他所
'If PUB_GetST06(strUserNum) = "3" Then
If cmdData(2).Enabled = False Then
   'edit by nickc 2005/03/01 加入未經主管確認的才出來，且草或墨有計件才出現
   'Modify by Morgan 2004/5/19
   '加專利種類
'    strSQL = "SELECT SUBSTR(CP09,1,1)," & SQLDate("CP05") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(PA57,'Y','＊',''),NVL(PA05,NVL(PA06,PA07)), Decode(pa09,'000',cpm03,cpm04),s1.st02, ROUND(cp18,2), EP13, DECODE(EP06,NULL,' ',0,' '," & SQLDate("EP06") & "), EP20, DECODE(EP14,NULL,'',0,''," & SQLDate("eP14") & "),'' As 草期限, " & SQLDate("eP15") & ",0, EP29, " & SQLDate("eP17") & ", '' As 墨期限, " & SQLDate("EP18") & ",0," & SQLDate("cP06") & "," & SQLDate("cP27") & ",ep26,s3.st02,CP09,DECODE(PA09,'000',PTM03,PTM04)," & SQLDate("CP07") & "," & SQLDate("CP57") & ",ep21,ep22,ep23,ep24,ep25,ep16,ep19, CP10, S2.ST02 " & _
'                " FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,PATENT,CASEPROPERTYMAP,PATENTTRADEMARKMAP " & _
'                " WHERE EP02=CP09(+) And CP01 In ('P','CFP','FCP') AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP14=S1.ST01(+) AND EP13=S2.ST01(+) AND CP13=S3.ST01(+) and CP01=CPM01(+) and CP10=CPM02(+) AND '1'=PTM01(+)  AND PA08=PTM02(+) And EP06 Is Not Null And EP13 Is Null And EP10 IS NULL And S1.ST06 In ('3', '4') And CP57 Is Null And CP10 In ('101','102','103','104','105','109','110','112','113','114','115') AND NOT (CP01='CFP' AND CP10='105')"
'    StrSql = "SELECT SUBSTR(CP09,1,1)," & SQLDate("CP05") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(PA57,'Y','＊',''),NVL(PA05,NVL(PA06,PA07)), Decode(pa09,'000',cpm03,cpm04),s1.st02, ROUND(cp18,2), EP13, DECODE(EP06,NULL,' ',0,' '," & SQLDate("EP06") & "), EP20, DECODE(EP14,NULL,'',0,''," & SQLDate("eP14") & "),'' As 草期限, " & SQLDate("eP15") & ",0, EP29, " & SQLDate("eP17") & ", '' As 墨期限, " & SQLDate("EP18") & ",0," & SQLDate("cP06") & "," & SQLDate("cP27") & ",ep26,s3.st02,CP09,DECODE(PA09,'000',PTM03,PTM04)," & SQLDate("CP07") & "," & SQLDate("CP57") & ",ep21,ep22,ep23,ep24,ep25,ep16,ep19, CP10, S2.ST02, PA08 " & _
'                " FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,PATENT,CASEPROPERTYMAP,PATENTTRADEMARKMAP " & _
'                " WHERE EP02=CP09(+) And CP01 In ('P','CFP','FCP') AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP14=S1.ST01(+) AND EP13=S2.ST01(+) AND CP13=S3.ST01(+) and CP01=CPM01(+) and CP10=CPM02(+) AND '1'=PTM01(+)  AND PA08=PTM02(+) And EP06 Is Not Null And EP13 Is Null And EP10 IS NULL And S1.ST06 In ('3', '4') And CP57 Is Null And CP10 In ('101','102','103','104','105','109','110','112','113','114','115')"
'edit by nickc 2005/04/07 不鎖案件性質，只要計件，都要出來
'    StrSql = "SELECT '',SUBSTR(CP09,1,1)," & SQLDate("CP05") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(PA57,'Y','＊',''),NVL(PA05,NVL(PA06,PA07)), Decode(pa09,'000',cpm03,cpm04),s1.st02, ROUND(cp18,2), S2.ST02, DECODE(EP06,NULL,' ',0,' '," & SQLDate("EP06") & "), EP20, DECODE(EP14,NULL,'',0,''," & SQLDate("eP14") & "),'' As 草期限, " & SQLDate("eP15") & ",0, EP29, " & SQLDate("eP17") & ", '' As 墨期限, " & SQLDate("EP18") & ",0," & SQLDate("cP06") & "," & SQLDate("cP27") & ",ep26,s3.st02,CP09,DECODE(PA09,'000',PTM03,PTM04)," & SQLDate("CP07") & "," & SQLDate("CP57") & ",ep21,ep22,ep23,ep24,ep25,ep16,ep19, CP10, ep13, PA08 " & _
'                " FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,PATENT,CASEPROPERTYMAP,PATENTTRADEMARKMAP " & _
'                " WHERE EP02=CP09(+) And CP01 In ('P','CFP','FCP') AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP14=S1.ST01(+) AND EP13=S2.ST01(+) AND CP13=S3.ST01(+) and CP01=CPM01(+) and CP10=CPM02(+) AND '1'=PTM01(+)  AND PA08=PTM02(+) And EP06 Is Not Null  And EP10 IS NULL And S1.ST06 In ('3', '4') And CP57 Is Null And CP10 In ('101','102','103','104','105','109','110','112','113','114','115') and cp107 is null and (ep20 is null or ep29 is null)  "
'edit by nickc 2005/04/12 改共用一句
'    StrSql = "SELECT '',SUBSTR(CP09,1,1)," & SQLDate("CP05") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(PA57,'Y','＊',''),NVL(PA05,NVL(PA06,PA07)), Decode(pa09,'000',cpm03,cpm04),s1.st02, ROUND(cp18,2), S2.ST02, DECODE(EP06,NULL,' ',0,' '," & SQLDate("EP06") & "), EP20, DECODE(EP14,NULL,'',0,''," & SQLDate("eP14") & "),'' As 草期限, " & SQLDate("eP15") & ",0, EP29, " & SQLDate("eP17") & ", '' As 墨期限, " & SQLDate("EP18") & ",0," & SQLDate("cP06") & "," & SQLDate("cP27") & ",ep26,s3.st02,CP09,DECODE(PA09,'000',PTM03,PTM04)," & SQLDate("CP07") & "," & SQLDate("CP57") & ",ep21,ep22,ep23,ep24,ep25,ep16,ep19, CP10, ep13, PA08 " & _
'                " FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,PATENT,CASEPROPERTYMAP,PATENTTRADEMARKMAP " & _
'                " WHERE EP02=CP09(+) And CP01 In ('P','CFP','FCP') AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP14=S1.ST01(+) AND EP13=S2.ST01(+) AND CP13=S3.ST01(+) and CP01=CPM01(+) and CP10=CPM02(+) AND '1'=PTM01(+)  AND PA08=PTM02(+) And EP06 Is Not Null  And EP10 IS NULL And S1.ST06 In ('3', '4') And CP57 Is Null  and cp107 is null and (ep20 is null or ep29 is null)  and ep13 is not null "
strSql = strSql & " And S1.ST06 In ('3', '4') "
'若使用者非南所人員, 只可看到自所所別的資料
'edit by nickc 2005/03/31  北所的繪圖沒上不出來
'Else
'edit by nickc 2005/04/12 改可查其他所
'ElseIf PUB_GetST06(strUserNum) = "1" Then
ElseIf cmdData(0).Enabled = False Then
   'edit by nickc 2005/03/01 加入未經主管確認的才出來，且草或墨有計件才出現
   'Modify by Morgan 2004/5/19
   '加專利種類
'    strSQL = "SELECT SUBSTR(CP09,1,1)," & SQLDate("CP05") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(PA57,'Y','＊',''),NVL(PA05,NVL(PA06,PA07)), Decode(pa09,'000',cpm03,cpm04),s1.st02, ROUND(cp18,2), EP13, DECODE(EP06,NULL,' ',0,' '," & SQLDate("EP06") & "), EP20, DECODE(EP14,NULL,'',0,''," & SQLDate("eP14") & "),'' As 草期限, " & SQLDate("eP15") & ",0, EP29, " & SQLDate("eP17") & ", '' As 墨期限, " & SQLDate("EP18") & ",0," & SQLDate("cP06") & "," & SQLDate("cP27") & ",ep26,s3.st02,CP09,DECODE(PA09,'000',PTM03,PTM04)," & SQLDate("CP07") & "," & SQLDate("CP57") & ",ep21,ep22,ep23,ep24,ep25,ep16,ep19, CP10, S2.ST02 " & _
'                " FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,PATENT,CASEPROPERTYMAP,PATENTTRADEMARKMAP, Staff S4 " & _
'                " WHERE EP02=CP09(+) And CP01 In ('P','CFP','FCP') AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP14=S1.ST01(+) AND EP13=S2.ST01(+) AND CP13=S3.ST01(+) and CP01=CPM01(+) and CP10=CPM02(+) AND '1'=PTM01(+)  AND PA08=PTM02(+) And EP06 Is Not Null And EP13 Is Null And EP10 IS NULL And S4.ST06=S1.ST06 And '" & strUserNum & "'=S4.ST01 And CP57 Is Null And CP10 In ('101','102','103','104','105','109','110','112','113','114','115') AND NOT (CP01='CFP' AND CP10='105')"
'    StrSql = "SELECT SUBSTR(CP09,1,1)," & SQLDate("CP05") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(PA57,'Y','＊',''),NVL(PA05,NVL(PA06,PA07)), Decode(pa09,'000',cpm03,cpm04),s1.st02, ROUND(cp18,2), EP13, DECODE(EP06,NULL,' ',0,' '," & SQLDate("EP06") & "), EP20, DECODE(EP14,NULL,'',0,''," & SQLDate("eP14") & "),'' As 草期限, " & SQLDate("eP15") & ",0, EP29, " & SQLDate("eP17") & ", '' As 墨期限, " & SQLDate("EP18") & ",0," & SQLDate("cP06") & "," & SQLDate("cP27") & ",ep26,s3.st02,CP09,DECODE(PA09,'000',PTM03,PTM04)," & SQLDate("CP07") & "," & SQLDate("CP57") & ",ep21,ep22,ep23,ep24,ep25,ep16,ep19, CP10, S2.ST02, PA08 " & _
'                " FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,PATENT,CASEPROPERTYMAP,PATENTTRADEMARKMAP, Staff S4 " & _
'                " WHERE EP02=CP09(+) And CP01 In ('P','CFP','FCP') AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP14=S1.ST01(+) AND EP13=S2.ST01(+) AND CP13=S3.ST01(+) and CP01=CPM01(+) and CP10=CPM02(+) AND '1'=PTM01(+)  AND PA08=PTM02(+) And EP06 Is Not Null And EP13 Is Null And EP10 IS NULL And S4.ST06=S1.ST06 And '" & strUserNum & "'=S4.ST01 And CP57 Is Null And CP10 In ('101','102','103','104','105','109','110','112','113','114','115')"
'edit by nickc 2005/04/07 不鎖案件性質，只要計件，都要出來
'    StrSql = "SELECT '',SUBSTR(CP09,1,1)," & SQLDate("CP05") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(PA57,'Y','＊',''),NVL(PA05,NVL(PA06,PA07)), Decode(pa09,'000',cpm03,cpm04),s1.st02, ROUND(cp18,2), S2.ST02, DECODE(EP06,NULL,' ',0,' '," & SQLDate("EP06") & "), EP20, DECODE(EP14,NULL,'',0,''," & SQLDate("eP14") & "),'' As 草期限, " & SQLDate("eP15") & ",0, EP29, " & SQLDate("eP17") & ", '' As 墨期限, " & SQLDate("EP18") & ",0," & SQLDate("cP06") & "," & SQLDate("cP27") & ",ep26,s3.st02,CP09,DECODE(PA09,'000',PTM03,PTM04)," & SQLDate("CP07") & "," & SQLDate("CP57") & ",ep21,ep22,ep23,ep24,ep25,ep16,ep19, CP10, ep13, PA08 " & _
'                " FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,PATENT,CASEPROPERTYMAP,PATENTTRADEMARKMAP, Staff S4 " & _
'                " WHERE EP02=CP09(+) And CP01 In ('P','CFP','FCP') AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP14=S1.ST01(+) AND EP13=S2.ST01(+) AND CP13=S3.ST01(+) and CP01=CPM01(+) and CP10=CPM02(+) AND '1'=PTM01(+)  AND PA08=PTM02(+) And EP06 Is Not Null  And EP10 IS NULL And S4.ST06=S1.ST06 And '" & strUserNum & "'=S4.ST01 And CP57 Is Null And CP10 In ('101','102','103','104','105','109','110','112','113','114','115') and cp107 is null  and (ep20 is null or ep29 is null) and ep13 is not null "
'edit by nickc 2005/04/12 改共用一句
'    StrSql = "SELECT '',SUBSTR(CP09,1,1)," & SQLDate("CP05") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(PA57,'Y','＊',''),NVL(PA05,NVL(PA06,PA07)), Decode(pa09,'000',cpm03,cpm04),s1.st02, ROUND(cp18,2), S2.ST02, DECODE(EP06,NULL,' ',0,' '," & SQLDate("EP06") & "), EP20, DECODE(EP14,NULL,'',0,''," & SQLDate("eP14") & "),'' As 草期限, " & SQLDate("eP15") & ",0, EP29, " & SQLDate("eP17") & ", '' As 墨期限, " & SQLDate("EP18") & ",0," & SQLDate("cP06") & "," & SQLDate("cP27") & ",ep26,s3.st02,CP09,DECODE(PA09,'000',PTM03,PTM04)," & SQLDate("CP07") & "," & SQLDate("CP57") & ",ep21,ep22,ep23,ep24,ep25,ep16,ep19, CP10, ep13, PA08 " & _
'                " FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,PATENT,CASEPROPERTYMAP,PATENTTRADEMARKMAP, Staff S4 " & _
'                " WHERE EP02=CP09(+) And CP01 In ('P','CFP','FCP') AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP14=S1.ST01(+) AND EP13=S2.ST01(+) AND CP13=S3.ST01(+) and CP01=CPM01(+) and CP10=CPM02(+) AND '1'=PTM01(+)  AND PA08=PTM02(+) And EP06 Is Not Null  And EP10 IS NULL And S4.ST06=S1.ST06 And '" & strUserNum & "'=S4.ST01 And CP57 Is Null  and cp107 is null  and (ep20 is null or ep29 is null) and ep13 is not null "
strSql = strSql & " And S1.ST06 In ('1', '5') "
'add by nickc 2005/03/31 除南高之外，中所也不能鎖繪圖
Else
'edit by nickc 2005/04/07 不鎖案件性質，只要計件，都要出來
'    StrSql = "SELECT '',SUBSTR(CP09,1,1)," & SQLDate("CP05") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(PA57,'Y','＊',''),NVL(PA05,NVL(PA06,PA07)), Decode(pa09,'000',cpm03,cpm04),s1.st02, ROUND(cp18,2), S2.ST02, DECODE(EP06,NULL,' ',0,' '," & SQLDate("EP06") & "), EP20, DECODE(EP14,NULL,'',0,''," & SQLDate("eP14") & "),'' As 草期限, " & SQLDate("eP15") & ",0, EP29, " & SQLDate("eP17") & ", '' As 墨期限, " & SQLDate("EP18") & ",0," & SQLDate("cP06") & "," & SQLDate("cP27") & ",ep26,s3.st02,CP09,DECODE(PA09,'000',PTM03,PTM04)," & SQLDate("CP07") & "," & SQLDate("CP57") & ",ep21,ep22,ep23,ep24,ep25,ep16,ep19, CP10, ep13, PA08 " & _
'                " FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,PATENT,CASEPROPERTYMAP,PATENTTRADEMARKMAP, Staff S4 " & _
'                " WHERE EP02=CP09(+) And CP01 In ('P','CFP','FCP') AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP14=S1.ST01(+) AND EP13=S2.ST01(+) AND CP13=S3.ST01(+) and CP01=CPM01(+) and CP10=CPM02(+) AND '1'=PTM01(+)  AND PA08=PTM02(+) And EP06 Is Not Null  And EP10 IS NULL And S4.ST06=S1.ST06 And '" & strUserNum & "'=S4.ST01 And CP57 Is Null And CP10 In ('101','102','103','104','105','109','110','112','113','114','115') and cp107 is null  and (ep20 is null or ep29 is null) "
'edit by nickc 2005/04/12 改共用一句
'    StrSql = "SELECT '',SUBSTR(CP09,1,1)," & SQLDate("CP05") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(PA57,'Y','＊',''),NVL(PA05,NVL(PA06,PA07)), Decode(pa09,'000',cpm03,cpm04),s1.st02, ROUND(cp18,2), S2.ST02, DECODE(EP06,NULL,' ',0,' '," & SQLDate("EP06") & "), EP20, DECODE(EP14,NULL,'',0,''," & SQLDate("eP14") & "),'' As 草期限, " & SQLDate("eP15") & ",0, EP29, " & SQLDate("eP17") & ", '' As 墨期限, " & SQLDate("EP18") & ",0," & SQLDate("cP06") & "," & SQLDate("cP27") & ",ep26,s3.st02,CP09,DECODE(PA09,'000',PTM03,PTM04)," & SQLDate("CP07") & "," & SQLDate("CP57") & ",ep21,ep22,ep23,ep24,ep25,ep16,ep19, CP10, ep13, PA08 " & _
'                " FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,PATENT,CASEPROPERTYMAP,PATENTTRADEMARKMAP, Staff S4 " & _
'                " WHERE EP02=CP09(+) And CP01 In ('P','CFP','FCP') AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP14=S1.ST01(+) AND EP13=S2.ST01(+) AND CP13=S3.ST01(+) and CP01=CPM01(+) and CP10=CPM02(+) AND '1'=PTM01(+)  AND PA08=PTM02(+) And EP06 Is Not Null  And EP10 IS NULL And S4.ST06=S1.ST06 And '" & strUserNum & "'=S4.ST01 And CP57 Is Null  and cp107 is null  and (ep20 is null or ep29 is null) and ep13 is not null "
strSql = strSql & " And S1.ST06 In ('2') "
End If

'Modify end

strSql = strSql + " ORDER BY CP14, 3 Desc "
CheckOC
With adoRecordset
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        Set grd1.Recordset = adoRecordset
        grd1.Visible = False
        Screen.MousePointer = vbHourglass
        Me.Enabled = False
        For i = 1 To grd1.Rows - 1
            grd1.row = i
            '文齊日(將空白消除)
'            Me.grd1.TextMatrix(i, 7) = Trim(Me.grd1.TextMatrix(i, 7))
            Me.grd1.TextMatrix(i, 9) = Trim(Me.grd1.TextMatrix(i, 9))
            '草完日
'            grd1.col = 11
            grd1.col = 13
            strDate1 = grd1.Text
            '草齊日
'            grd1.col = 9
            grd1.col = 11
            strDate2 = grd1.Text
            If Trim(strDate1) <> "" And Trim(strDate2) <> "" Then
                '草天
'                grd1.col = 12
                grd1.col = 14
                grd1.Text = GetWorkDay(ChangeTStringToWString(ChangeTDateStringToTString(strDate1)), ChangeTStringToWString(ChangeTDateStringToTString(strDate2)))
            Else
                '草天
'                grd1.col = 12
                grd1.col = 14
                grd1.Text = ""
            End If
            '墨完日
'            grd1.col = 16
            grd1.col = 18
            strDate1 = grd1.Text
            '墨齊日
'            grd1.col = 14
            grd1.col = 16
            strDate2 = grd1.Text
            If Trim(strDate1) <> "" And Trim(strDate2) <> "" Then
                '墨天
'                grd1.col = 17
                grd1.col = 19
                grd1.Text = GetWorkDay(ChangeTStringToWString(ChangeTDateStringToTString(strDate1)), ChangeTStringToWString(ChangeTDateStringToTString(strDate2)))
            Else
                '墨天
'                grd1.col = 17
                grd1.col = 19
                grd1.Text = ""
            End If
            '若草圖不計件
'            If grd1.TextMatrix(i, 8) = "N" Then
'                grd1.TextMatrix(i, 9) = "******"
'                grd1.TextMatrix(i, 11) = "******"
'                grd1.TextMatrix(i, 12) = ""
'            End If
            If grd1.TextMatrix(i, 10) = "N" Then
                grd1.TextMatrix(i, 11) = "******"
                grd1.TextMatrix(i, 13) = "******"
                grd1.TextMatrix(i, 14) = ""
            End If
            '若墨圖不計件
'            If grd1.TextMatrix(i, 13) = "N" Then
'                grd1.TextMatrix(i, 14) = "******"
'                grd1.TextMatrix(i, 16) = "******"
'                grd1.TextMatrix(i, 17) = ""
'            End If
            If grd1.TextMatrix(i, 15) = "N" Then
                grd1.TextMatrix(i, 16) = "******"
                grd1.TextMatrix(i, 18) = "******"
                grd1.TextMatrix(i, 19) = ""
            End If
            '草圖承辦期限
'            If Me.grd1.TextMatrix(i, 9) <> "" And Me.grd1.TextMatrix(i, 9) <> "******" Then
'                '設計申請
'                If Me.grd1.TextMatrix(i, 33) = "103" Or Me.grd1.TextMatrix(i, 33) = "105" Then
'                    Me.grd1.TextMatrix(i, 10) = ChangeTStringToTDateString(CompWorkDay(5, Replace(Me.grd1.TextMatrix(i, 9), "/", "") + 19110000) - 19110000)
'                '非設計申請
'                Else
'                    Me.grd1.TextMatrix(i, 10) = ChangeTStringToTDateString(CompWorkDay(4, Replace(Me.grd1.TextMatrix(i, 9), "/", "") + 19110000) - 19110000)
'                End If
'            End If
            If Me.grd1.TextMatrix(i, 11) <> "" And Me.grd1.TextMatrix(i, 11) <> "******" Then
                '設計申請
                'Modify by Morgan 2004/5/19
                '設計改用專利種類判斷
                'If Me.grd1.TextMatrix(i, 34) = "103" Or Me.grd1.TextMatrix(i, 34) = "105" Then
                If Me.grd1.TextMatrix(i, 37) = "3" Then
                    Me.grd1.TextMatrix(i, 12) = ChangeTStringToTDateString(CompWorkDay(5, Replace(Me.grd1.TextMatrix(i, 11), "/", "") + 19110000) - 19110000)
                '非設計申請
                Else
                    Me.grd1.TextMatrix(i, 12) = ChangeTStringToTDateString(CompWorkDay(4, Replace(Me.grd1.TextMatrix(i, 11), "/", "") + 19110000) - 19110000)
                End If
            End If
            '墨圖承辦期限
'            If Me.grd1.TextMatrix(i, 14) <> "" And Me.grd1.TextMatrix(i, 14) <> "******" Then
'                Me.grd1.TextMatrix(i, 15) = ChangeTStringToTDateString(CompWorkDay(3, Replace(Me.grd1.TextMatrix(i, 14), "/", "") + 19110000) - 19110000)
'            End If
            If Me.grd1.TextMatrix(i, 16) <> "" And Me.grd1.TextMatrix(i, 16) <> "******" Then
                Me.grd1.TextMatrix(i, 17) = ChangeTStringToTDateString(CompWorkDay(3, Replace(Me.grd1.TextMatrix(i, 16), "/", "") + 19110000) - 19110000)
            End If
        Next i
        Me.Enabled = True
        SetGrd1
        Screen.MousePointer = vbDefault
        grd1.MousePointer = flexDefault
        grd1.Visible = True
        ChkNoData = False
    Else
        Me.grd1.Clear
        Me.grd1.Rows = 2
        SetGrd1
        ChkNoData = True
         Screen.MousePointer = vbDefault
         grd1.MousePointer = flexDefault
    End If
End With
CheckOC
         Screen.MousePointer = vbDefault
         grd1.MousePointer = flexDefault
End Sub

Sub ChgGrdColor()
Dim tmpcolor1 As Integer
Dim tmpcolor2 As Integer
'Add by Morgan 2004/5/19
'專利種類
Dim stPA08 As String

With grd1
    .Visible = False
    For i = 1 To grd1.Rows - 1
        .row = i
        '法定期限
'        .col = 24
        .col = 26
'        '若有法定期限
        If .Text <> "" Then
            '若法定期限 = 系統日
            If DBDATE(.Text) = strSrvDate(1) Then
                For j = 2 To .Cols - 1
                    .col = j
                    .CellBackColor = &H8080FF '淺紅色
                Next j
            End If
        End If
            '取消收文日
'            .col = 25
            .col = 27
            '若有取消收文日
            If .Text <> "" Then
                For j = 2 To .Cols - 1
                    .col = j
                    .CellBackColor = &HC0C0C0 '灰色
                Next j
            End If
            '若無取消收文日
                '專利種類
                'Add by Morgan 2004/5/19
                stPA08 = .TextMatrix(.row, 37)
                
                '設計？
'                .col = 33
                .col = 35
                strCP10 = Trim(.Text)
                '發文日
'                .col = 19
                .col = 21
                StrColor1 = .Text
                '取消收文日
'                .col = 25
                .col = 27
                StrColor2 = .Text
                '草圖完稿
'                .col = 11
                .col = 13
                StrColor3 = Replace(.Text, "******", "")
                '墨圖完稿
'                .col = 16
                .col = 18
                StrColor4 = Replace(.Text, "******", "")
                '草圖齊備
'                .col = 9
                .col = 11
                StrColor5 = Replace(.Text, "******", "")
                '墨圖齊備
'                .col = 14
                .col = 16
                StrColor6 = Replace(.Text, "******", "")
                '草圖作業天數
'                .col = 12
                .col = 14
                tmpcolor1 = Val(.Text)
                '墨圖作業天數
'                .col = 17
                .col = 19
                tmpcolor2 = Val(.Text)
                
                'Modify by Morgan 2004/5/19
                '改依專利種類判斷
'                Select Case StrCp10
'                Case "103", "105" '設計申請
                Select Case stPA08
                  Case "3"
                    'nick  91/04/10
                    If tmpcolor1 > 5 Or tmpcolor2 > 3 Then
                        For j = 2 To .Cols - 1
                            .col = j
                            .CellBackColor = &H80FFFF '黃色
                        Next j
                    Else
                        '無發文日, 無取消收文日, 無草完日
                        If Len(Trim(StrColor1)) = 0 And Len(Trim(StrColor2)) = 0 And Len(Trim(StrColor3)) = 0 Then
                            '若有草齊日
                            If Len(Trim(StrColor5)) <> 0 Then
                                '若系統日超過草齊日5個工作天
                                If GetWorkDay(strSrvDate(1), ChangeTStringToWString(ChangeTDateStringToTString(StrColor5))) > 5 Then
                                    For j = 2 To .Cols - 1
                                        .col = j
                                        .CellBackColor = &H80FFFF '黃色
                                    Next j
                                End If
                            End If
                        Else
                            '無發文日, 無取消收文日, 無墨完日
                            If Len(Trim(StrColor1)) = 0 And Len(Trim(StrColor2)) = 0 And Len(Trim(StrColor4)) = 0 Then
                                '若有墨齊日
                                If Len(Trim(StrColor6)) <> 0 Then
                                    '若系統日大於墨齊日3個工作天
                                    If GetWorkDay(strSrvDate(1), ChangeTStringToWString(ChangeTDateStringToTString(StrColor6))) > 3 Then
                                        For j = 2 To .Cols - 1
                                        .col = j
                                        .CellBackColor = &H80FFFF '黃色
                                        Next j
                                    End If
                                End If
                            End If
                        End If
                    End If
                Case Else '非設計申請
                    '91/04/10    ncik
                    If tmpcolor1 > 4 Or tmpcolor2 > 3 Then
                        For j = 2 To .Cols - 1
                           .col = j
                           .CellBackColor = &H80FFFF '黃色
                        Next j
                    Else
                        '無發文日, 無取消收文日, 無草完日
                        If Len(Trim(StrColor1)) = 0 And Len(Trim(StrColor2)) = 0 And Len(Trim(StrColor3)) = 0 Then
                           '若有草齊日
                           If Len(Trim(StrColor5)) <> 0 Then
                              '若系統日超過草齊日4個工作天
                              If GetWorkDay(strSrvDate(1), ChangeTStringToWString(ChangeTDateStringToTString(StrColor5))) > 4 Then
                                 For j = 2 To .Cols - 1
                                    .col = j
                                    .CellBackColor = &H80FFFF '黃色
                                 Next j
                              End If
                           End If
                        Else
                            '無發文日, 無取消收文日, 無墨完日
                            If Len(Trim(StrColor1)) = 0 And Len(Trim(StrColor2)) = 0 And Len(Trim(StrColor4)) = 0 Then
                               '若有墨齊日
                               If Len(Trim(StrColor6)) <> 0 Then
                                  '若系統日大於墨齊日3個工作天
                                  If GetWorkDay(strSrvDate(1), ChangeTStringToWString(ChangeTDateStringToTString(StrColor6))) > 3 Then
                                     For j = 2 To .Cols - 1
                                        .col = j
                                        .CellBackColor = &H80FFFF '黃色
                                     Next j
                                  End If
                               End If
                            End If
                        End If
                    End If
                End Select
                '若草圖計件, 有文齊日但無草齊日
'                If .TextMatrix(i, 8) = "" And .TextMatrix(i, 7) <> "" And .TextMatrix(i, 9) = "" Then
'                    .col = 9
'                    .CellBackColor = &HFF8080 '淺藍色
'                End If
                If .TextMatrix(i, 10) = "" And .TextMatrix(i, 9) <> "" And .TextMatrix(i, 11) = "" Then
                    .col = 11
                    .CellBackColor = &HFF8080 '淺藍色
                End If
                '若草圖計件, 有草齊日但無草完日
'                If .TextMatrix(i, 8) = "" And .TextMatrix(i, 9) <> "" And .TextMatrix(i, 11) = "" Then
'                    .col = 11
'                    .CellBackColor = &HFF80FF '粉紅色
'                End If
                If .TextMatrix(i, 10) = "" And .TextMatrix(i, 11) <> "" And .TextMatrix(i, 13) = "" Then
                    .col = 13
                    .CellBackColor = &HFF80FF '粉紅色
                End If
                '若墨圖計件, 有墨齊日但無墨完日
'                If .TextMatrix(i, 13) = "" And .TextMatrix(i, 14) <> "" And .TextMatrix(i, 16) = "" Then
'                    .col = 16
'                    .CellBackColor = &HFF80FF '粉紅色
'                End If
                If .TextMatrix(i, 15) = "" And .TextMatrix(i, 16) <> "" And .TextMatrix(i, 18) = "" Then
                    .col = 18
                    .CellBackColor = &HFF80FF '粉紅色
                End If
    Next i
    .Visible = True
End With
End Sub

Private Sub SetGrd1()

With grd1
    .Visible = False
'    .Cols = 34
    'Modify by Morgan 2004/5/19
    '加專利種類
    '.Cols = 36
'edit by nickc 2005/03/07  加入繪圖主管確認
'    .Cols = 37
'    .Row = 0
'    .RowHeight(0) = 400
'    .col = 0:   .Text = "類別"
'    .ColWidth(0) = 300
'    .CellAlignment = flexAlignCenterCenter
'    .col = 1:   .Text = "收文日"
'    .ColWidth(1) = 700
'    .CellAlignment = flexAlignCenterCenter
'    .col = 2:   .Text = "本所案號"
'    .ColWidth(2) = 1400
'    .CellAlignment = flexAlignCenterCenter
'    .col = 3:   .Text = "案件名稱"
'    .ColWidth(3) = 1500
'    .CellAlignment = flexAlignCenterCenter
'    .col = 4:   .Text = "案件性質"
'    .ColWidth(4) = 800
'    .CellAlignment = flexAlignCenterCenter
'    .col = 5:   .Text = "承辦人"
'    .ColWidth(5) = 700
'    .CellAlignment = flexAlignCenterCenter
'    .col = 6:   .Text = "點數"
'    .ColWidth(6) = 400
'    .CellAlignment = flexAlignCenterCenter
'    .col = 7:   .Text = "繪圖人員"
'    .ColWidth(7) = 800
'    .CellAlignment = flexAlignCenterCenter
'    .col = 8:   .Text = "文齊日"
'    .ColWidth(8) = 700
'    .CellAlignment = flexAlignCenterCenter
'    .col = 9:   .Text = "草計"
'    .ColWidth(9) = 400
'    .CellAlignment = flexAlignCenterCenter
'    .col = 10:   .Text = "草齊日"
'    .ColWidth(10) = 700
'    .CellAlignment = flexAlignCenterCenter
'    .col = 11:   .Text = "草期限"
''    .ColWidth(10) = 700
'    .ColWidth(11) = 0
'    .CellAlignment = flexAlignCenterCenter
'    .col = 12:  .Text = "草完日"
'    .ColWidth(12) = 700
'    .CellAlignment = flexAlignCenterCenter
'    .col = 13:  .Text = "草天"
'    .ColWidth(13) = 400
'    .CellAlignment = flexAlignCenterCenter
'    .col = 14:   .Text = "墨計"
'    .ColWidth(14) = 400
'    .CellAlignment = flexAlignCenterCenter
'    .col = 15:  .Text = "墨齊日"
'    .ColWidth(15) = 700
'    .CellAlignment = flexAlignCenterCenter
'    .col = 16:  .Text = "墨期限"
''    .ColWidth(15) = 700
'    .ColWidth(16) = 0
'    .CellAlignment = flexAlignCenterCenter
'    .col = 17:  .Text = "墨完日"
'    .ColWidth(17) = 700
'    .CellAlignment = flexAlignCenterCenter
'    .col = 18:  .Text = "墨天"
'    .ColWidth(18) = 400
'    .CellAlignment = flexAlignCenterCenter
'    .col = 19:  .Text = "本所期限"
'    .ColWidth(19) = 700
'    .CellAlignment = flexAlignCenterCenter
'    .col = 20:  .Text = "發文日"
'    .ColWidth(20) = 700
'    .CellAlignment = flexAlignCenterCenter
'    .col = 21:  .Text = "備註"
'    .ColWidth(21) = 800
'    .CellAlignment = flexAlignCenterCenter
'    .col = 22:  .Text = "智權人員"
'    .ColWidth(22) = 700
'    .CellAlignment = flexAlignCenterCenter
'    .col = 23:  .Text = "" '收文號
'    .ColWidth(23) = 0
'    .CellAlignment = flexAlignCenterCenter
'    .col = 24:  .Text = "" '案件性質名稱
'    .ColWidth(24) = 0
'    .CellAlignment = flexAlignCenterCenter
'    .col = 25:  .Text = "" '法定期限
'    .ColWidth(25) = 0
'    .CellAlignment = flexAlignCenterCenter
'    .col = 26:  .Text = "" '取消收文日
'    .ColWidth(26) = 0
'    .CellAlignment = flexAlignCenterCenter
'    .col = 27:  .Text = "" '草圖承辦時數
'    .ColWidth(27) = 0
'    .CellAlignment = flexAlignCenterCenter
'    .col = 28:  .Text = "" '墨圖承辦時數
'    .ColWidth(28) = 0
'    .CellAlignment = flexAlignCenterCenter
'    .col = 29:  .Text = "" '修改時數1
'    .ColWidth(29) = 0
'    .CellAlignment = flexAlignCenterCenter
'    .col = 30:  .Text = "" '修改時數2
'    .ColWidth(30) = 0
'    .CellAlignment = flexAlignCenterCenter
'    .col = 31:  .Text = "" '修改時數3
'    .ColWidth(31) = 0
'    .CellAlignment = flexAlignCenterCenter
'    .col = 32:  .Text = "" '草圖張數
'    .ColWidth(32) = 0
'    .CellAlignment = flexAlignCenterCenter
'    .col = 33:  .Text = "" '墨圖張數
'    .ColWidth(33) = 0
'    .CellAlignment = flexAlignCenterCenter
'    .col = 34:  .Text = "" '案件性質代號
'    .ColWidth(34) = 0
'    .CellAlignment = flexAlignCenterCenter
'    .col = 35:  .Text = "" '繪圖人員代號
'    .ColWidth(35) = 0
'    .CellAlignment = flexAlignCenterCenter
'
'    'Add by Morgan 2004/5/19
'    .col = 36:  .Text = "" '專利種類代號
'    .ColWidth(36) = 0
'    .CellAlignment = flexAlignCenterCenter
    '.Cols = 38
    .Cols = 39
    .row = 0
    .RowHeight(0) = 400
    .col = 0:   .Text = "確認分案"
    .ColWidth(0) = 700
    .CellAlignment = flexAlignCenterCenter
    .col = 1:   .Text = "類別"
    .ColWidth(1) = 300
    .CellAlignment = flexAlignCenterCenter
    .col = 2:   .Text = "收文日"
    .ColWidth(2) = 800
    .CellAlignment = flexAlignCenterCenter
    .col = 3:   .Text = "本所案號"
    .ColWidth(3) = 1400
    .CellAlignment = flexAlignCenterCenter
    .col = 4:   .Text = "案件名稱"
    .ColWidth(4) = 1500
    .CellAlignment = flexAlignCenterCenter
    .col = 5:   .Text = "案件性質"
    .ColWidth(5) = 800
    .CellAlignment = flexAlignCenterCenter
    .col = 6:   .Text = "承辦人"
    .ColWidth(6) = 700
    .CellAlignment = flexAlignCenterCenter
    .col = 7:   .Text = "點數"
    .ColWidth(7) = 400
    .CellAlignment = flexAlignCenterCenter
    .col = 8:   .Text = "繪圖人員"
    .ColWidth(8) = 800
    .CellAlignment = flexAlignCenterCenter
    .col = 9:   .Text = "文齊日"
    .ColWidth(9) = 800
    .CellAlignment = flexAlignCenterCenter
    .col = 10:   .Text = "草計"
    .ColWidth(10) = 400
    .CellAlignment = flexAlignCenterCenter
    .col = 11:   .Text = "草齊日"
    .ColWidth(11) = 800
    .CellAlignment = flexAlignCenterCenter
    .col = 12:   .Text = "草期限"
'    .ColWidth(10) = 700
    .ColWidth(12) = 0
    .CellAlignment = flexAlignCenterCenter
    .col = 13:  .Text = "草完日"
    .ColWidth(13) = 800
    .CellAlignment = flexAlignCenterCenter
    .col = 14:  .Text = "草天"
    .ColWidth(14) = 400
    .CellAlignment = flexAlignCenterCenter
    .col = 15:   .Text = "墨計"
    .ColWidth(15) = 400
    .CellAlignment = flexAlignCenterCenter
    .col = 16:  .Text = "墨齊日"
    .ColWidth(16) = 800
    .CellAlignment = flexAlignCenterCenter
    .col = 17:  .Text = "墨期限"
'    .ColWidth(15) = 700
    .ColWidth(17) = 0
    .CellAlignment = flexAlignCenterCenter
    .col = 18:  .Text = "墨完日"
    .ColWidth(18) = 800
    .CellAlignment = flexAlignCenterCenter
    .col = 19:  .Text = "墨天"
    .ColWidth(19) = 400
    .CellAlignment = flexAlignCenterCenter
    .col = 20:  .Text = "本所期限"
    .ColWidth(20) = 800
    .CellAlignment = flexAlignCenterCenter
    .col = 21:  .Text = "發文日"
    .ColWidth(21) = 800
    .CellAlignment = flexAlignCenterCenter
    .col = 22:  .Text = "備註"
    .ColWidth(22) = 800
    .CellAlignment = flexAlignCenterCenter
    .col = 23:  .Text = "智權人員"
    .ColWidth(23) = 700
    .CellAlignment = flexAlignCenterCenter
    .col = 24:  .Text = "" '收文號
    .ColWidth(24) = 0
    .CellAlignment = flexAlignCenterCenter
    .col = 25:  .Text = "" '案件性質名稱
    .ColWidth(25) = 0
    .CellAlignment = flexAlignCenterCenter
    .col = 26:  .Text = "" '法定期限
    .ColWidth(26) = 0
    .CellAlignment = flexAlignCenterCenter
    .col = 27:  .Text = "" '取消收文日
    .ColWidth(27) = 0
    .CellAlignment = flexAlignCenterCenter
    .col = 28:  .Text = "" '草圖承辦時數
    .ColWidth(28) = 0
    .CellAlignment = flexAlignCenterCenter
    .col = 29:  .Text = "" '墨圖承辦時數
    .ColWidth(29) = 0
    .CellAlignment = flexAlignCenterCenter
    .col = 30:  .Text = "" '修改時數1
    .ColWidth(30) = 0
    .CellAlignment = flexAlignCenterCenter
    .col = 31:  .Text = "" '修改時數2
    .ColWidth(31) = 0
    .CellAlignment = flexAlignCenterCenter
    .col = 32:  .Text = "" '修改時數3
    .ColWidth(32) = 0
    .CellAlignment = flexAlignCenterCenter
    .col = 33:  .Text = "" '草圖張數
    .ColWidth(33) = 0
    .CellAlignment = flexAlignCenterCenter
    .col = 34:  .Text = "" '墨圖張數
    .ColWidth(34) = 0
    .CellAlignment = flexAlignCenterCenter
    .col = 35:  .Text = "" '案件性質代號
    .ColWidth(35) = 0
    .CellAlignment = flexAlignCenterCenter
    .col = 36:  .Text = "" '繪圖人員代號
    .ColWidth(36) = 0
    .CellAlignment = flexAlignCenterCenter
    
    'Add by Morgan 2004/5/19
    .col = 36:  .Text = "" '專利種類代號
    .ColWidth(36) = 0
    .CellAlignment = flexAlignCenterCenter
    'add by nickc 2005/03/21
    .col = 37:  .Text = ""
    .ColWidth(37) = 0
    .CellAlignment = flexAlignCenterCenter
    'add by nickc 2008/01/22
    .col = 38:  .Text = ""
    .ColWidth(38) = 0
    .CellAlignment = flexAlignCenterCenter
    .Visible = True
End With
ChgGrdColor
End Sub

Private Sub GRD1_DblClick()
If Me.grd1.MouseRow > 0 Then
    SSTab1.Tab = 1
    'Modify By Cheng 2004/04/19
'    cmdOK(2).Caption = "存檔(&O)"
    cmdok(2).Caption = "存檔"
    'End
End If
End Sub

Private Sub grd1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Me.grd1.MouseRow > 0 Then
    If Button = 1 Then
        SWPRow = str(grd1.MouseRow)
        MouseClick Val(SWPRow)
        Me.txt1(0).SetFocus
    End If
End If
End Sub

Sub MouseClick(Optional Strindex As Integer)
With grd1
    .Visible = False
    For i = 0 To .Rows - 1
        .col = 0
        .row = i
        '若為點選的資料
        If .CellBackColor = &HFFC0C0 Then
            For k = 0 To 1
                .col = k
                .CellBackColor = QBColor(15) '白色
            Next k
            '草計
'            .col = 8
            .col = 9
            .CellBackColor = QBColor(15) '白色
            '墨計
'            .col = 13
            .col = 14
            .CellBackColor = QBColor(15) '白色
            Exit For
        End If
    Next i
    .col = 0
    If Strindex <> 0 Then
        .row = Strindex
    Else
        .row = .MouseRow
    End If
    If .row = 0 Then
        .row = 1
    End If
    '收文號
'    .col = 22
    .col = 23
    Process (.Text)
    For i = 0 To 1
        .col = i
        .CellBackColor = &HFFC0C0
    Next i
    '草計
'    .col = 8
    .col = 9
    .CellBackColor = &HFFC0C0
    '墨計
'    .col = 13
    .col = 14
    .CellBackColor = &HFFC0C0
    .Visible = True
End With
Combo2.Enabled = False
End Sub

Private Sub grd1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    'Add By Cheng 2004/03/23
    If Me.grd1.MouseRow < 1 Then
        If m_blnColOrderAsc = True Then
            Me.grd1.Sort = 5 '昇冪
            m_blnColOrderAsc = False
        Else
            Me.grd1.Sort = 6 '降冪
            m_blnColOrderAsc = True
        End If
    End If
    'End
End Sub

Private Sub grd1_SelChange()
grd1.Visible = False
grd1.row = grd1.MouseRow
grd1.col = 0
If grd1.row <> 0 Then
If grd1.Text = "V" Then
     grd1.Text = ""
      For i = 0 To 1
          grd1.col = i
          grd1.CellBackColor = QBColor(15) '白色
      Next i
      grd1.col = 9
      grd1.CellBackColor = QBColor(15) '白色
      grd1.col = 14
      grd1.CellBackColor = QBColor(15) '白色
Else
     grd1.Text = "V"
      For i = 0 To 1
          grd1.col = i
          grd1.CellBackColor = &HFFC0C0 '白色
      Next i
      grd1.col = 9
      grd1.CellBackColor = &HFFC0C0 '白色
      grd1.col = 14
      grd1.CellBackColor = &HFFC0C0 '白色

End If
End If
grd1.Visible = True
End Sub


Private Sub Option1_Click(Index As Integer)
If Option1(0).Value = True Then
    Combo2.Enabled = False
    txt1(13).Enabled = True
Else
    Combo2.Enabled = True
    txt1(13).Enabled = False
End If
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
If SSTab1.Tab = 0 Then
    'Modify By Cheng 2004/04/19
'    cmdOK(2).Caption = "確定(&O)"
    cmdok(2).Caption = "確定"
    'End
Else
    'Modify By Cheng 2004/04/19
'    cmdOK(2).Caption = "存檔(&O)"
    cmdok(2).Caption = "存檔"
    'End
    Me.txt1(0).SetFocus
End If
End Sub

Private Sub txt1_Change(Index As Integer)
    Select Case Index
    Case 7 '草圖是否計件
        If Me.txt1(Index).Text = "N" Then
            Me.txt1(1).Enabled = False
            Me.txt1(2).Enabled = False
            Me.txt1(1).Text = "******"
            Me.txt1(2).Text = "******"
        Else
            Me.txt1(1).Enabled = True
            Me.txt1(2).Enabled = True
            Me.txt1(1).Text = ""
            Me.txt1(2).Text = ""
        End If
    Case 14 '墨圖否計件
        If Me.txt1(Index).Text = "N" Then
            Me.txt1(4).Enabled = False
            Me.txt1(5).Enabled = False
            Me.txt1(4).Text = "******"
            Me.txt1(5).Text = "******"
        Else
            Me.txt1(4).Enabled = True
            Me.txt1(5).Enabled = True
            Me.txt1(4).Text = ""
            Me.txt1(5).Text = ""
        End If
    End Select
End Sub

Private Sub txt1_GotFocus(Index As Integer)
txt1(Index).SelStart = 0
txt1(Index).SelLength = Len(txt1(Index))
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As MSForms.ReturnInteger)
    KeyAscii = UpperCase(KeyAscii)
    Select Case Index
    Case 7, 14
        If KeyAscii <> 8 And KeyAscii <> 78 Then
            KeyAscii = 0
        End If
    Case Else
    End Select
End Sub

Private Sub txt1_LostFocus(Index As Integer)
ChkData2 = True
Select Case Index
Case 0
'edit by nickc 2005/06/13 取消
    If Me.txt1(0).Text = "" Then
'        Me.lbl1(0).Caption = ""
'        '2005/6/2 ADD BY SONIA
'        s = MsgBox("繪圖人員不可空白!!", , "USER 輸入錯誤")
'        txt1(Index).SetFocus
'        txt1(Index).SelStart = 0
'        txt1(Index).SelLength = Len(txt1(Index))
'        ChkData2 = False
'        Exit Sub
'        '2005/6/2 END
    Else
      'add by nickc 2006/07/04 剔除電腦中心，不然每次都出問題
      If Pub_StrUserSt03 <> "M51" Then
        CheckOC2
        strSql = "SELECT S1.ST02 FROM STAFF S1,STAFF S2 WHERE S2.ST01='" & strUserNum & "' AND SUBSTR(S1.ST03,1,1) = SUBSTR(S2.ST03,1,1) AND S1.ST05 in ('79','81','82','AC') AND S1.ST04='1' AND S2.ST04='1' AND S1.ST01='" & Trim(txt1(0)) & "' "
        With adoRecordset1
            .CursorLocation = adUseClient
            .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
            If .RecordCount <> 0 And .RecordCount > 0 Then
                lbl1(0).Caption = CheckStr(.Fields(0))
            Else
                s = MsgBox("此 " & txt1(0) & " 不屬於繪圖部門,並與 USER '" & strUserNum & "' 部門第一碼不同!!", , "USER 輸入錯誤")
                txt1(0).SetFocus
                txt1(0).SelStart = 0
                txt1(0).SelLength = Len(txt1(0))
                ChkData2 = False
                Exit Sub
            End If
        End With
      End If
    End If
Case 2, 3, 4, 5, 6, 8, 9, 10, 11, 12, 1
    Select Case Index
    Case 1, 2, 4, 5
        If Me.txt1(Index).Text = "******" Then Exit Sub
    Case Else
    End Select
     If Len(txt1(Index)) <> 0 Then
        If IsNumeric(txt1(Index)) = False Then
            s = MsgBox("請輸入數字!!", , "USER 輸入錯誤")
            txt1(Index).SetFocus
            txt1(Index).SelStart = 0
            txt1(Index).SelLength = Len(txt1(Index))
            ChkData2 = False
            Exit Sub
        End If
     End If
     If Index >= 8 And Index <= 12 Then
        If Val(txt1(Index)) >= 100 Or Val(txt1(Index)) < 0 Then
            s = MsgBox("時數輸入錯誤 0-99.9 !!", , "USER 輸入錯誤")
            txt1(Index).SetFocus
            txt1(Index).SelStart = 0
            txt1(Index).SelLength = Len(txt1(Index))
            ChkData2 = False
            Exit Sub
        End If
     End If
     If Index = 2 Or Index = 5 Or Index = 4 Or Index = 1 Then
         If Len(txt1(Index)) <> 0 Then
            If Not ChkWorkDay(ChangeTStringToWString(txt1(Index))) Then
               ShowDateErr
               txt1(Index).SetFocus
               txt1_GotFocus (Index)
               ChkData2 = False
               Exit Sub
            End If
         End If
      End If
     If Index = 2 Then
         If Len(txt1(1)) <> 0 And Len(txt1(2)) <> 0 Then
            If RunNick2(txt1(1), txt1(2)) Then
               txt1(1).SetFocus
               txt1_GotFocus (1)
               ChkData2 = False
               Exit Sub
            End If
         End If
     End If
     If Index = 5 Then
         If Len(txt1(4)) <> 0 And Len(txt1(5)) <> 0 Then
            If RunNick2(txt1(4), txt1(5)) Then
               txt1(4).SetFocus
               txt1_GotFocus (4)
               ChkData2 = False
               Exit Sub
            End If
         End If
    End If
Case 7, 14
     Select Case Trim(txt1(7))
     Case "", "N"
     Case Else
          s = MsgBox("只能輸入 N 或空白!!", , "USER 輸入錯誤")
          txt1(Index).SetFocus
          txt1(Index).SelStart = 0
          txt1(Index).SelLength = Len(txt1(Index))
          ChkData2 = False
          Exit Sub
     End Select
Case Else
End Select
End Sub
