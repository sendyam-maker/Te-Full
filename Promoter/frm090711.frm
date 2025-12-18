VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm090711 
   BorderStyle     =   1  '單線固定
   Caption         =   "個人工作進度資料維護"
   ClientHeight    =   8880
   ClientLeft      =   -1790
   ClientTop       =   960
   ClientWidth     =   15010
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8880
   ScaleWidth      =   15010
   Begin VB.CommandButton cmdok 
      Caption         =   "確定"
      Height          =   400
      Index           =   2
      Left            =   12690
      TabIndex        =   28
      Top             =   45
      Width           =   756
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   270
      Left            =   6105
      MaxLength       =   5
      TabIndex        =   1
      Top             =   480
      Width           =   1215
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "本月統計(&A)"
      Height          =   400
      Index           =   0
      Left            =   11460
      TabIndex        =   27
      Top             =   45
      Width           =   1200
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "回前畫面(&U)"
      Height          =   400
      Index           =   1
      Left            =   13455
      TabIndex        =   29
      Top             =   45
      Width           =   1200
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7950
      Left            =   120
      TabIndex        =   30
      Top             =   840
      Width           =   14685
      _ExtentX        =   25912
      _ExtentY        =   14023
      _Version        =   393216
      Style           =   1
      Tab             =   2
      TabHeight       =   520
      TabCaption(0)   =   "瀏覽資料"
      TabPicture(0)   =   "frm090711.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "grd1"
      Tab(0).Control(1)=   "Combo3"
      Tab(0).Control(2)=   "cmd(1)"
      Tab(0).Control(3)=   "cmd(0)"
      Tab(0).Control(4)=   "Label5"
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "詳細資料"
      TabPicture(1)   =   "frm090711.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label1(30)"
      Tab(1).Control(1)=   "lbl1(30)"
      Tab(1).Control(2)=   "lbl1(25)"
      Tab(1).Control(3)=   "lbl1(20)"
      Tab(1).Control(4)=   "lbl1(23)"
      Tab(1).Control(5)=   "lbl1(22)"
      Tab(1).Control(6)=   "lbl1(24)"
      Tab(1).Control(7)=   "lbl1(19)"
      Tab(1).Control(8)=   "lbl1(18)"
      Tab(1).Control(9)=   "Label1(31)"
      Tab(1).Control(10)=   "Label1(22)"
      Tab(1).Control(11)=   "Label1(6)"
      Tab(1).Control(12)=   "Label1(3)"
      Tab(1).Control(13)=   "Label1(26)"
      Tab(1).Control(14)=   "Label1(23)"
      Tab(1).Control(15)=   "Label1(17)"
      Tab(1).Control(16)=   "Label1(2)"
      Tab(1).Control(17)=   "Label1(28)"
      Tab(1).Control(18)=   "Label1(34)"
      Tab(1).Control(19)=   "lbl1(21)"
      Tab(1).Control(20)=   "Label1(5)"
      Tab(1).Control(21)=   "Label1(24)"
      Tab(1).Control(22)=   "lbl1(26)"
      Tab(1).Control(23)=   "Label1(27)"
      Tab(1).Control(24)=   "lbl1(28)"
      Tab(1).Control(25)=   "lbl1(29)"
      Tab(1).Control(26)=   "Label1(35)"
      Tab(1).Control(27)=   "Label1(1)"
      Tab(1).Control(28)=   "Label1(33)"
      Tab(1).Control(29)=   "Label1(29)"
      Tab(1).Control(30)=   "Label1(21)"
      Tab(1).Control(31)=   "Label1(20)"
      Tab(1).Control(32)=   "Label1(19)"
      Tab(1).Control(33)=   "Label1(18)"
      Tab(1).Control(34)=   "Label1(16)"
      Tab(1).Control(35)=   "Label1(15)"
      Tab(1).Control(36)=   "Label1(14)"
      Tab(1).Control(37)=   "Label1(13)"
      Tab(1).Control(38)=   "Label1(12)"
      Tab(1).Control(39)=   "Label1(11)"
      Tab(1).Control(40)=   "Label1(10)"
      Tab(1).Control(41)=   "Label1(9)"
      Tab(1).Control(42)=   "Label1(8)"
      Tab(1).Control(43)=   "Label1(4)"
      Tab(1).Control(44)=   "Label1(25)"
      Tab(1).Control(45)=   "lbl1(27)"
      Tab(1).Control(46)=   "lbl1(10)"
      Tab(1).Control(47)=   "lbl1(1)"
      Tab(1).Control(48)=   "lbl1(3)"
      Tab(1).Control(49)=   "lbl1(4)"
      Tab(1).Control(50)=   "lbl1(5)"
      Tab(1).Control(51)=   "lbl1(8)"
      Tab(1).Control(52)=   "lbl1(9)"
      Tab(1).Control(53)=   "lbl1(11)"
      Tab(1).Control(54)=   "lbl1(6)"
      Tab(1).Control(55)=   "lbl1(16)"
      Tab(1).Control(56)=   "lbl1(12)"
      Tab(1).Control(57)=   "lbl1(14)"
      Tab(1).Control(58)=   "lbl1(15)"
      Tab(1).Control(59)=   "lbl1(13)"
      Tab(1).Control(60)=   "lbl1(7)"
      Tab(1).Control(61)=   "lbl1(2)"
      Tab(1).Control(62)=   "lbl1(0)"
      Tab(1).Control(63)=   "lblClose"
      Tab(1).Control(64)=   "Label1(7)"
      Tab(1).Control(65)=   "Label1(36)"
      Tab(1).Control(66)=   "lbl1(17)"
      Tab(1).Control(67)=   "Label1(37)"
      Tab(1).Control(68)=   "lbl1(31)"
      Tab(1).Control(69)=   "Label1(38)"
      Tab(1).Control(70)=   "lbl1(32)"
      Tab(1).Control(71)=   "lbl1(33)"
      Tab(1).Control(72)=   "Label7"
      Tab(1).Control(73)=   "Label6"
      Tab(1).Control(74)=   "lbl1(34)"
      Tab(1).Control(75)=   "Label2"
      Tab(1).Control(76)=   "Label3"
      Tab(1).Control(77)=   "Label1(39)"
      Tab(1).Control(78)=   "Label1(40)"
      Tab(1).Control(79)=   "lblEApp"
      Tab(1).Control(80)=   "lblCM10"
      Tab(1).Control(81)=   "lblCMboth"
      Tab(1).Control(82)=   "txt1(3)"
      Tab(1).Control(83)=   "txt1(14)"
      Tab(1).Control(84)=   "txt1(4)"
      Tab(1).Control(85)=   "txt1(5)"
      Tab(1).Control(86)=   "txt1(6)"
      Tab(1).Control(87)=   "txt1(8)"
      Tab(1).Control(88)=   "txt1(9)"
      Tab(1).Control(89)=   "txt1(10)"
      Tab(1).Control(90)=   "txt1(11)"
      Tab(1).Control(91)=   "txt1(12)"
      Tab(1).Control(92)=   "txt1(13)"
      Tab(1).Control(93)=   "txt1(0)"
      Tab(1).Control(94)=   "txt1(7)"
      Tab(1).Control(95)=   "txt1(1)"
      Tab(1).Control(96)=   "txt1(2)"
      Tab(1).Control(97)=   "txt1(16)"
      Tab(1).Control(98)=   "txt1(15)"
      Tab(1).Control(99)=   "txt1(18)"
      Tab(1).Control(100)=   "txt1(17)"
      Tab(1).Control(101)=   "txt1(19)"
      Tab(1).Control(102)=   "Option1(0)"
      Tab(1).Control(103)=   "Option1(1)"
      Tab(1).Control(104)=   "Combo2"
      Tab(1).Control(105)=   "cmd1"
      Tab(1).Control(106)=   "cmd2"
      Tab(1).Control(107)=   "cmd3"
      Tab(1).Control(108)=   "cmdPic"
      Tab(1).Control(109)=   "cmdok(3)"
      Tab(1).Control(110)=   "cmd(2)"
      Tab(1).ControlCount=   111
      TabCaption(2)   =   "待辦歷程"
      TabPicture(2)   =   "frm090711.frx":0038
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Label16"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Label1(48)"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "grd2"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "cmdQuery"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "cmdDetail"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "Combo5"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).ControlCount=   6
      Begin VB.ComboBox Combo5 
         Height          =   300
         ItemData        =   "frm090711.frx":0054
         Left            =   4800
         List            =   "frm090711.frx":0064
         Style           =   2  '單純下拉式
         TabIndex        =   127
         Top             =   420
         Width           =   1350
      End
      Begin VB.CommandButton cmdDetail 
         Caption         =   "明細資料(&D)"
         Height          =   360
         Left            =   6450
         TabIndex        =   125
         Top             =   360
         Width           =   1305
      End
      Begin VB.CommandButton cmdQuery 
         Caption         =   "畫面更新(&Q)"
         Height          =   360
         Left            =   7980
         TabIndex        =   122
         Top             =   360
         Width           =   1155
      End
      Begin VB.CommandButton cmd 
         Caption         =   "承辦歷程(&E)"
         Height          =   375
         Index           =   2
         Left            =   -65520
         TabIndex        =   121
         Top             =   870
         Width           =   1665
      End
      Begin VB.CommandButton cmdok 
         Caption         =   "繪圖超時內部收文(&B)"
         Height          =   555
         Index           =   3
         Left            =   -67665
         TabIndex        =   120
         Top             =   3150
         Visible         =   0   'False
         Width           =   1965
      End
      Begin VB.CommandButton cmdPic 
         BackColor       =   &H00C0C0C0&
         Caption         =   "代表圖(&I)"
         Height          =   375
         Left            =   -65520
         Style           =   1  '圖片外觀
         TabIndex        =   119
         Top             =   390
         Width           =   1665
      End
      Begin VB.CommandButton cmd3 
         Caption         =   "相關國內案件"
         Height          =   330
         Left            =   -74850
         TabIndex        =   110
         Top             =   5580
         Width           =   1485
      End
      Begin VB.CommandButton cmd2 
         Caption         =   "相關多國案件"
         Height          =   330
         Left            =   -74850
         TabIndex        =   109
         Top             =   6000
         Width           =   1485
      End
      Begin VB.CommandButton cmd1 
         Caption         =   "相關國外案件"
         Height          =   330
         Left            =   -74850
         TabIndex        =   108
         Top             =   5130
         Width           =   1485
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grd1 
         Height          =   6990
         Left            =   -74925
         TabIndex        =   31
         Top             =   795
         Width           =   14505
         _ExtentX        =   25576
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
      Begin VB.ComboBox Combo2 
         Height          =   300
         Left            =   -69180
         TabIndex        =   21
         Text            =   "Combo2"
         Top             =   4845
         Width           =   1560
      End
      Begin VB.ComboBox Combo3 
         Height          =   300
         ItemData        =   "frm090711.frx":0083
         Left            =   -73980
         List            =   "frm090711.frx":009C
         TabIndex        =   106
         Top             =   390
         Width           =   3465
      End
      Begin VB.CommandButton cmd 
         Caption         =   "未發文(&L)"
         Height          =   400
         Index           =   1
         Left            =   -61605
         TabIndex        =   3
         Top             =   336
         Width           =   1200
      End
      Begin VB.CommandButton cmd 
         Caption         =   "當月資料(&M)"
         Height          =   400
         Index           =   0
         Left            =   -62835
         TabIndex        =   2
         Top             =   336
         Width           =   1200
      End
      Begin VB.OptionButton Option1 
         Height          =   300
         Index           =   1
         Left            =   -69840
         TabIndex        =   20
         Top             =   4785
         Width           =   285
      End
      Begin VB.OptionButton Option1 
         Height          =   300
         Index           =   0
         Left            =   -69840
         TabIndex        =   18
         Top             =   4500
         Value           =   -1  'True
         Width           =   300
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grd2 
         Height          =   4395
         Left            =   90
         TabIndex        =   123
         Top             =   750
         Width           =   9645
         _ExtentX        =   17022
         _ExtentY        =   7743
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
      Begin MSForms.TextBox txt1 
         Height          =   300
         Index           =   19
         Left            =   -68925
         TabIndex        =   26
         Top             =   6405
         Width           =   585
         VariousPropertyBits=   671107099
         MaxLength       =   1
         Size            =   "1032;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   300
         Index           =   17
         Left            =   -69345
         TabIndex        =   24
         Top             =   6060
         Width           =   525
         VariousPropertyBits=   671107099
         MaxLength       =   3
         Size            =   "926;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   300
         Index           =   18
         Left            =   -68745
         TabIndex        =   25
         Top             =   6060
         Width           =   7365
         VariousPropertyBits=   671107099
         Size            =   "12991;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   300
         Index           =   15
         Left            =   -69345
         TabIndex        =   22
         Top             =   5460
         Width           =   525
         VariousPropertyBits=   671107099
         MaxLength       =   3
         Size            =   "926;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   300
         Index           =   16
         Left            =   -68745
         TabIndex        =   23
         Top             =   5460
         Width           =   7365
         VariousPropertyBits=   671107099
         Size            =   "12991;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   300
         Index           =   2
         Left            =   -69180
         TabIndex        =   7
         Top             =   975
         Width           =   1200
         VariousPropertyBits=   671107099
         MaxLength       =   7
         Size            =   "2117;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   300
         Index           =   1
         Left            =   -69180
         TabIndex        =   6
         Top             =   690
         Width           =   1200
         VariousPropertyBits=   671107099
         MaxLength       =   7
         Size            =   "2117;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   300
         Index           =   7
         Left            =   -69180
         TabIndex        =   5
         Top             =   420
         Width           =   585
         VariousPropertyBits=   671107099
         MaxLength       =   1
         Size            =   "1032;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   300
         Index           =   0
         Left            =   -74025
         TabIndex        =   4
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
         Index           =   13
         Left            =   -69180
         TabIndex        =   19
         Top             =   4560
         Width           =   3810
         VariousPropertyBits=   -1467989989
         MaxLength       =   2000
         ScrollBars      =   2
         Size            =   "6720;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   300
         Index           =   12
         Left            =   -69180
         TabIndex        =   17
         Top             =   4275
         Width           =   1200
         VariousPropertyBits=   671107099
         MaxLength       =   4
         Size            =   "2117;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   300
         Index           =   11
         Left            =   -69180
         TabIndex        =   16
         Top             =   4005
         Width           =   1200
         VariousPropertyBits=   671107099
         MaxLength       =   4
         Size            =   "2117;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   300
         Index           =   10
         Left            =   -69180
         TabIndex        =   15
         Top             =   3720
         Width           =   1200
         VariousPropertyBits=   671107099
         MaxLength       =   4
         Size            =   "2117;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   300
         Index           =   9
         Left            =   -68970
         TabIndex        =   14
         Top             =   3450
         Width           =   1200
         VariousPropertyBits=   671107099
         MaxLength       =   4
         Size            =   "2117;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   300
         Index           =   8
         Left            =   -68970
         TabIndex        =   13
         Top             =   3165
         Width           =   1200
         VariousPropertyBits=   671107099
         MaxLength       =   4
         Size            =   "2117;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   300
         Index           =   6
         Left            =   -69180
         TabIndex        =   12
         Top             =   2625
         Width           =   1200
         VariousPropertyBits=   671107099
         MaxLength       =   2
         Size            =   "2117;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   300
         Index           =   5
         Left            =   -69180
         TabIndex        =   11
         Top             =   2340
         Width           =   1200
         VariousPropertyBits=   671107099
         MaxLength       =   7
         Size            =   "2117;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   300
         Index           =   4
         Left            =   -69180
         TabIndex        =   10
         Top             =   2070
         Width           =   1200
         VariousPropertyBits=   671107099
         MaxLength       =   7
         Size            =   "2117;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   300
         Index           =   14
         Left            =   -69180
         TabIndex        =   9
         Top             =   1800
         Width           =   585
         VariousPropertyBits=   671107099
         MaxLength       =   1
         Size            =   "1032;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   300
         Index           =   3
         Left            =   -69180
         TabIndex        =   8
         Top             =   1245
         Width           =   1200
         VariousPropertyBits=   671107099
         MaxLength       =   2
         Size            =   "2117;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label lblCMboth 
         Caption         =   "lblCMboth"
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
         Left            =   -71940
         TabIndex        =   130
         Top             =   690
         Width           =   945
      End
      Begin VB.Label lblCM10 
         Caption         =   "一案兩請"
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
         Left            =   -71940
         TabIndex        =   129
         Top             =   1245
         Visible         =   0   'False
         Width           =   825
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "最近聯絡："
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   48
         Left            =   3870
         TabIndex        =   128
         Top             =   480
         Width           =   900
      End
      Begin VB.Label lblEApp 
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
         Left            =   -71940
         TabIndex        =   126
         Top             =   975
         Visible         =   0   'False
         Width           =   945
      End
      Begin VB.Label Label16 
         Caption         =   "註：雙擊選取時，開啟承辦歷程"
         ForeColor       =   &H000000C0&
         Height          =   225
         Left            =   180
         TabIndex        =   124
         Top             =   510
         Width           =   2895
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "(Y：提供；空白為不提供)"
         Height          =   180
         Index           =   40
         Left            =   -68220
         TabIndex        =   118
         Top             =   6450
         Width           =   2040
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "客戶是否提供圖檔："
         Height          =   180
         Index           =   39
         Left            =   -70575
         TabIndex        =   117
         Top             =   6465
         Width           =   1620
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "墨圖計件值："
         Height          =   180
         Left            =   -70590
         TabIndex        =   116
         Top             =   5850
         Width           =   1080
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "墨圖加乘註記："
         Height          =   180
         Left            =   -70590
         TabIndex        =   115
         Top             =   6075
         Width           =   1260
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   34
         Left            =   -69345
         TabIndex        =   114
         Top             =   5865
         Width           =   780
         VariousPropertyBits=   27
         Caption         =   "123"
         Size            =   "1376;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "草圖計件值："
         Height          =   180
         Left            =   -70590
         TabIndex        =   113
         Top             =   5250
         Width           =   1080
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "草圖加乘註記："
         Height          =   180
         Left            =   -70590
         TabIndex        =   112
         Top             =   5460
         Width           =   1260
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   33
         Left            =   -69375
         TabIndex        =   111
         Top             =   5250
         Width           =   780
         VariousPropertyBits=   27
         Caption         =   "123"
         Size            =   "1376;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label5 
         Caption         =   "顏色說明："
         Height          =   225
         Left            =   -74880
         TabIndex        =   107
         Top             =   435
         Width           =   915
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   32
         Left            =   -66510
         TabIndex        =   105
         Top             =   2070
         Visible         =   0   'False
         Width           =   900
         VariousPropertyBits=   27
         Caption         =   "123"
         Size            =   "1587;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         Caption         =   "墨圖承辦期限："
         Height          =   195
         Index           =   38
         Left            =   -67860
         TabIndex        =   104
         Top             =   2070
         Visible         =   0   'False
         Width           =   1305
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   31
         Left            =   -66480
         TabIndex        =   103
         Top             =   690
         Visible         =   0   'False
         Width           =   900
         VariousPropertyBits=   27
         Caption         =   "123"
         Size            =   "1587;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         Caption         =   "草圖承辦期限："
         Height          =   195
         Index           =   37
         Left            =   -67830
         TabIndex        =   102
         Top             =   690
         Visible         =   0   'False
         Width           =   1305
      End
      Begin MSForms.Label lbl1 
         Height          =   285
         Index           =   17
         Left            =   -63795
         TabIndex        =   101
         Top             =   1800
         Visible         =   0   'False
         Width           =   600
         VariousPropertyBits=   27
         Caption         =   "123"
         Size            =   "1058;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         Caption         =   "墨圖是否計件："
         Height          =   195
         Index           =   36
         Left            =   -70560
         TabIndex        =   100
         Top             =   1800
         Width           =   1275
      End
      Begin VB.Label Label1 
         Caption         =   "(N：不計)"
         Height          =   180
         Index           =   7
         Left            =   -68520
         TabIndex        =   99
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
         Left            =   -71940
         TabIndex        =   79
         Top             =   1800
         Width           =   930
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   0
         Left            =   -72750
         TabIndex        =   58
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
         Left            =   -74130
         TabIndex        =   59
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
         Left            =   -74310
         TabIndex        =   63
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
         Left            =   -73590
         TabIndex        =   72
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
         Left            =   -73560
         TabIndex        =   71
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
         Left            =   -73770
         TabIndex        =   70
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
         Left            =   -73560
         TabIndex        =   69
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
         Left            =   -73320
         TabIndex        =   68
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
         Left            =   -73950
         TabIndex        =   67
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
         Left            =   -73950
         TabIndex        =   66
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
         Left            =   -73950
         TabIndex        =   65
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
         Left            =   -73920
         TabIndex        =   64
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
         Left            =   -73530
         TabIndex        =   62
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
         Left            =   -73980
         TabIndex        =   61
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
         Left            =   -73980
         TabIndex        =   60
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
         Left            =   -73950
         TabIndex        =   57
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
         Left            =   -74130
         TabIndex        =   56
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
         Left            =   -63795
         TabIndex        =   98
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
         Left            =   -69600
         TabIndex        =   97
         Top             =   4005
         Width           =   345
      End
      Begin VB.Label Label1 
         Caption         =   "繪圖人員："
         Height          =   195
         Index           =   4
         Left            =   -74880
         TabIndex        =   96
         Top             =   420
         Width           =   885
      End
      Begin VB.Label Label1 
         Caption         =   "承辦人："
         Height          =   195
         Index           =   8
         Left            =   -74880
         TabIndex        =   95
         Top             =   3195
         Width           =   765
      End
      Begin VB.Label Label1 
         Caption         =   "國外案承辦人："
         Height          =   195
         Index           =   9
         Left            =   -74880
         TabIndex        =   94
         Top             =   4575
         Width           =   1305
      End
      Begin VB.Label Label1 
         Caption         =   "國外案本所案號："
         Height          =   195
         Index           =   10
         Left            =   -74880
         TabIndex        =   93
         Top             =   4860
         Width           =   1485
      End
      Begin VB.Label Label1 
         Caption         =   "點數："
         Height          =   195
         Index           =   11
         Left            =   -74880
         TabIndex        =   92
         Top             =   2355
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "智權人員："
         Height          =   195
         Index           =   12
         Left            =   -74880
         TabIndex        =   91
         Top             =   3465
         Width           =   900
      End
      Begin VB.Label Label1 
         Caption         =   "法定期限："
         Height          =   195
         Index           =   13
         Left            =   -74880
         TabIndex        =   90
         Top             =   2910
         Width           =   945
      End
      Begin VB.Label Label1 
         Caption         =   "本所期限："
         Height          =   195
         Index           =   14
         Left            =   -74880
         TabIndex        =   89
         Top             =   2640
         Width           =   945
      End
      Begin VB.Label Label1 
         Caption         =   "案件性質："
         Height          =   195
         Index           =   15
         Left            =   -74880
         TabIndex        =   88
         Top             =   2085
         Width           =   945
      End
      Begin VB.Label Label1 
         Caption         =   "專利/商標種類："
         Height          =   195
         Index           =   16
         Left            =   -74880
         TabIndex        =   87
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "案件名稱："
         Height          =   195
         Index           =   18
         Left            =   -74880
         TabIndex        =   86
         Top             =   1530
         Width           =   945
      End
      Begin VB.Label Label1 
         Caption         =   "本所案號："
         Height          =   195
         Index           =   19
         Left            =   -74880
         TabIndex        =   85
         Top             =   1245
         Width           =   945
      End
      Begin VB.Label Label1 
         Caption         =   "收文日："
         Height          =   195
         Index           =   20
         Left            =   -74880
         TabIndex        =   84
         Top             =   975
         Width           =   765
      End
      Begin VB.Label Label1 
         Caption         =   "總收文號："
         Height          =   195
         Index           =   21
         Left            =   -74880
         TabIndex        =   83
         Top             =   690
         Width           =   945
      End
      Begin VB.Label Label1 
         Caption         =   "取消收文日："
         Height          =   195
         Index           =   29
         Left            =   -74880
         TabIndex        =   82
         Top             =   4305
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "草圖作業天數："
         Height          =   195
         Index           =   33
         Left            =   -74880
         TabIndex        =   81
         Top             =   3750
         Width           =   1305
      End
      Begin VB.Label Label1 
         Caption         =   "墨圖作業天數："
         Height          =   195
         Index           =   1
         Left            =   -74880
         TabIndex        =   80
         Top             =   4020
         Width           =   1305
      End
      Begin VB.Label Label1 
         Caption         =   "(N：不計)"
         Height          =   180
         Index           =   35
         Left            =   -68520
         TabIndex        =   78
         Top             =   420
         Width           =   1065
      End
      Begin MSForms.Label lbl1 
         Height          =   285
         Index           =   29
         Left            =   -63795
         TabIndex        =   74
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
         Left            =   -63795
         TabIndex        =   73
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
         Left            =   -69600
         TabIndex        =   55
         Top             =   4275
         Width           =   345
      End
      Begin MSForms.Label lbl1 
         Height          =   285
         Index           =   26
         Left            =   -63795
         TabIndex        =   54
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
         Left            =   -69660
         TabIndex        =   53
         Top             =   3450
         Width           =   555
      End
      Begin VB.Label Label1 
         Caption         =   "草圖："
         Height          =   180
         Index           =   5
         Left            =   -69660
         TabIndex        =   52
         Top             =   3165
         Width           =   555
      End
      Begin MSForms.Label lbl1 
         Height          =   285
         Index           =   21
         Left            =   -63795
         TabIndex        =   51
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
         Left            =   -70560
         TabIndex        =   50
         Top             =   2625
         Width           =   945
      End
      Begin VB.Label Label1 
         Caption         =   "墨圖完稿日："
         Height          =   195
         Index           =   28
         Left            =   -70560
         TabIndex        =   49
         Top             =   2340
         Width           =   1125
      End
      Begin VB.Label Label1 
         Caption         =   "墨圖齊備日："
         Height          =   195
         Index           =   2
         Left            =   -70560
         TabIndex        =   48
         Top             =   2070
         Width           =   1125
      End
      Begin VB.Label Label1 
         Caption         =   "草圖是否計件："
         Height          =   195
         Index           =   17
         Left            =   -70560
         TabIndex        =   47
         Top             =   420
         Width           =   1275
      End
      Begin VB.Label Label1 
         Caption         =   "草圖齊備日："
         Height          =   195
         Index           =   23
         Left            =   -70560
         TabIndex        =   46
         Top             =   690
         Width           =   1125
      End
      Begin VB.Label Label1 
         Caption         =   "草圖完稿日："
         Height          =   195
         Index           =   26
         Left            =   -70560
         TabIndex        =   45
         Top             =   975
         Width           =   1125
      End
      Begin VB.Label Label1 
         Caption         =   "草圖張數："
         Height          =   195
         Index           =   3
         Left            =   -70560
         TabIndex        =   44
         Top             =   1245
         Width           =   945
      End
      Begin VB.Label Label1 
         Caption         =   "修改時數："
         Height          =   180
         Index           =   6
         Left            =   -70560
         TabIndex        =   43
         Top             =   3720
         Width           =   930
      End
      Begin VB.Label Label1 
         Caption         =   "複雜時數："
         Height          =   180
         Index           =   22
         Left            =   -70560
         TabIndex        =   42
         Top             =   3165
         Width           =   930
      End
      Begin VB.Label Label1 
         Caption         =   "備註："
         Height          =   180
         Index           =   31
         Left            =   -70560
         TabIndex        =   41
         Top             =   4560
         Width           =   690
      End
      Begin MSForms.Label lbl1 
         Height          =   285
         Index           =   18
         Left            =   -63795
         TabIndex        =   40
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
         Left            =   -63795
         TabIndex        =   39
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
         Left            =   -63795
         TabIndex        =   38
         Top             =   420
         Visible         =   0   'False
         Width           =   600
         VariousPropertyBits=   27
         Caption         =   "123"
         Size            =   "1058;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   285
         Index           =   22
         Left            =   -63795
         TabIndex        =   37
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
         Left            =   -63795
         TabIndex        =   36
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
         Left            =   -63795
         TabIndex        =   35
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
         Left            =   -63795
         TabIndex        =   34
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
         Left            =   -62655
         TabIndex        =   33
         Top             =   4560
         Visible         =   0   'False
         Width           =   510
         VariousPropertyBits=   27
         Caption         =   "123"
         Size            =   "900;476"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         Caption         =   "1."
         Height          =   180
         Index           =   30
         Left            =   -69600
         TabIndex        =   32
         Top             =   3720
         Width           =   345
      End
   End
   Begin MSForms.ComboBox Combo1 
      Height          =   300
      Left            =   1125
      TabIndex        =   0
      Top             =   480
      Width           =   2625
      VariousPropertyBits=   679479323
      DisplayStyle    =   7
      Size            =   "4630;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label lblCal1 
      Alignment       =   1  '靠右對齊
      Appearance      =   0  '平面
      BackColor       =   &H80000005&
      BorderStyle     =   1  '單線固定
      ForeColor       =   &H80000008&
      Height          =   210
      Index           =   2
      Left            =   5930
      TabIndex        =   139
      Top             =   0
      Width           =   1110
   End
   Begin VB.Label Label10 
      Alignment       =   2  '置中對齊
      Appearance      =   0  '平面
      BackColor       =   &H8000000A&
      BorderStyle     =   1  '單線固定
      Caption         =   "累計墨圖完成量"
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   2100
      TabIndex        =   138
      Top             =   0
      Width           =   1400
   End
   Begin VB.Label Label11 
      Alignment       =   2  '置中對齊
      Appearance      =   0  '平面
      BackColor       =   &H8000000A&
      BorderStyle     =   1  '單線固定
      Caption         =   "草圖目前進度"
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   4440
      TabIndex        =   137
      Top             =   0
      Width           =   1500
   End
   Begin VB.Label Label12 
      Alignment       =   2  '置中對齊
      Appearance      =   0  '平面
      BackColor       =   &H8000000A&
      BorderStyle     =   1  '單線固定
      Caption         =   "草圖累計達成比例"
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   4440
      TabIndex        =   136
      Top             =   210
      Width           =   1500
   End
   Begin VB.Label lblCal1 
      Alignment       =   1  '靠右對齊
      Appearance      =   0  '平面
      BackColor       =   &H80000005&
      BorderStyle     =   1  '單線固定
      ForeColor       =   &H80000008&
      Height          =   210
      Index           =   0
      Left            =   3330
      TabIndex        =   135
      Top             =   0
      Width           =   1110
   End
   Begin VB.Label lblCal1 
      Alignment       =   1  '靠右對齊
      Appearance      =   0  '平面
      BackColor       =   &H80000005&
      BorderStyle     =   1  '單線固定
      ForeColor       =   &H80000008&
      Height          =   210
      Index           =   3
      Left            =   5930
      TabIndex        =   134
      Top             =   210
      Width           =   1110
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      Caption         =   "此四項數據僅算到昨日"
      ForeColor       =   &H000000FF&
      Height          =   180
      Left            =   228
      TabIndex        =   133
      Top             =   30
      Width           =   1800
   End
   Begin VB.Label Label14 
      Alignment       =   2  '置中對齊
      Appearance      =   0  '平面
      BackColor       =   &H8000000A&
      BorderStyle     =   1  '單線固定
      Caption         =   "累計草圖完成量"
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   2100
      TabIndex        =   132
      Top             =   210
      Width           =   1400
   End
   Begin VB.Label lblCal1 
      Alignment       =   1  '靠右對齊
      Appearance      =   0  '平面
      BackColor       =   &H80000005&
      BorderStyle     =   1  '單線固定
      ForeColor       =   &H80000008&
      Height          =   210
      Index           =   1
      Left            =   3330
      TabIndex        =   131
      Top             =   210
      Width           =   1110
   End
   Begin MSForms.Label LBL2 
      Height          =   255
      Left            =   3435
      TabIndex        =   77
      Top             =   525
      Visible         =   0   'False
      Width           =   1410
      VariousPropertyBits=   27
      Size            =   "2487;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      Caption         =   "發文年月： "
      Height          =   180
      Index           =   32
      Left            =   5040
      TabIndex        =   76
      Top             =   516
      Width           =   912
   End
   Begin VB.Label Label1 
      Caption         =   "繪圖人員： "
      Height          =   180
      Index           =   0
      Left            =   228
      TabIndex        =   75
      Top             =   540
      Width           =   912
   End
End
Attribute VB_Name = "frm090711"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2022/05/30 (已檢查)整理frm880005改用寄信模組
'Memo by Morgan 2022/1/14 改成Form2.0 (grd1,grd2,txt1,Combo1,lbl1)
'Memo By Morgan 2012/12/10 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/17 日期欄已修改
Option Explicit

Public TextOk As Boolean, StrGrp090711 As String
Dim strSql As String, strSQL1 As String, i As Integer, j As Integer, s As Integer, SavDay1 As String, SavDay2 As String, SavDay3 As String, StrSQL7 As String, StrSQL4 As String, strSQL5 As String
Dim strSQL2 As String, iPrint As Integer, Page As Integer, strTemp(0 To 21) As String, strTemp3 As String, TestOk As Boolean, StrTemp99(0 To 21) As String
Dim PLeft(0 To 21) As Integer, strTemp1 As Variant, strTemp2 As Variant, StrSQL6 As String, Str020401SysKind As String, Seekok As Integer, SeekTemp As Integer, strDate1 As String, StrDate2 As String
Dim ChkData2 As Boolean, strCP10 As String, k As Integer, ChkNoData As Boolean, TXT090711 As Object
Public SWPRow As String
Dim NickRS As ADODB.Recordset, StrColor1 As String, StrColor2 As String, StrColor3 As String, StrColor4 As String, StrColor5 As String, StrColor6 As String
Dim ll As Integer
Dim m_blnColOrderAsc As Boolean '欄位資料由小到大排序
Dim m_CP09 As String
'add by nickc 2005/03/17
Dim m_CP101 As String
Dim m_CP102 As String
Dim m_CP104 As String
Dim m_CP105 As String
'add by nickc 2005/04/13 控制列印時的統計都會算
Public From090706 As Boolean
'add by nickc 2006/04/07
Dim StrSPa As String
'add by nickc 2006/12/29   紀錄 mail 資料，在 trans 後發
Dim skMail() As SeekMails
Dim m_CP21 As String 'Add by Morgan 2011/3/30
Dim m_PA09 As String 'Add by Morgan 2011/3/30
Dim m_bolExistsInCase As Boolean 'Add by Morgan 2011/3/30
'Add By Sindy 2013/6/7
Dim m_chkcmdok1 As Boolean '記錄確定鍵是否存檔成功
Dim dblPrevRow As Double
Public intBackTab As Integer
Dim ii As Integer
'2013/6/7 End
Dim m_CPM29 As String 'Add By Sindy 2013/9/30
Dim m_ProState As String 'Add By Sindy 2017/8/10 記錄目前權限


'Added by Morgan 2013/10/3
Private Sub UpdateOneRecord()
   Dim adoRst As ADODB.Recordset
   Dim stCP09 As String, stSQL As String, iR As Integer, iCol As Integer, iRow As Integer
   Dim strDate1 As String, StrDate2 As String
   
   iRow = grd1.row
   stCP09 = grd1.TextMatrix(iRow, 23)
   stSQL = "SELECT SUBSTR(CP09,1,1)||decode(ibf13,null,'','+'),SQLDateT2(CP05),substrb(na03,1,4),CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(PA57,'Y','＊',''),NVL(PA05,NVL(PA06,PA07)), Decode(pa09,'000',cpm03,cpm04),s1.st02, ROUND(cp18,2),nvl(SQLDateT2(EP06),' '), nvl(EP20,round(cp100 * cp101,2)),NVL(SQLDateT2(EP14),' '),'' As 草期限, SQLDateT2(EP15),0, nvl(EP29,round(cp103 * cp104,2)),NVL(SQLDateT2(EP17),' '), '' As 墨期限,SQLDateT2(EP18),0,SQLDateT2(CP06),SQLDateT2(CP27),ep26,s3.st02,CP09,decode(pa01,'CFP',ptm03,DECODE(PA09,'000',PTM03,PTM04)),SQLDateT2(CP07),SQLDateT2(NVL(CP57,PA58)),ep21,ep22,ep23,ep24,ep25,ep16,ep19, CP10, PA08,decode(pa09,'000','9999','020','8888','101','7777','0000') as PaSort FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,PATENT,CASEPROPERTYMAP,PATENTTRADEMARKMAP,nation,imgbytefile  " & _
            " WHERE cp09='" & stCP09 & "' and EP02(+)=CP09 AND cp01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND cp14=S1.ST01(+) AND eP13=S2.ST01(+) AND cP13=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND '1'=PTM01(+)  AND PA08=PTM02(+) and pa09=na01(+) and cp01=ibf01(+) and cp02=ibf02(+) and cp03=ibf03(+) and cp04=ibf04(+) "
   iR = 1
   Set adoRst = ClsLawReadRstMsg(iR, stSQL)
   If iR = 1 Then
      For iCol = 0 To adoRst.Fields.Count - 1
         grd1.TextMatrix(iRow, iCol) = "" & adoRst.Fields(iCol)
      Next
      grd1.col = 12
      strDate1 = LTrim(grd1.Text)
      grd1.col = 10
      StrDate2 = LTrim(grd1.Text)
      If Trim(strDate1) <> "" And Trim(StrDate2) <> "" Then
          '草天
          grd1.col = 13
          grd1.Text = GetWorkDay(ChangeTStringToWString(ChangeTDateStringToTString(strDate1)), ChangeTStringToWString(ChangeTDateStringToTString(StrDate2)))
      Else
          '草天
          grd1.col = 13
          grd1.Text = ""
      End If
      '墨完日
      grd1.col = 17
      strDate1 = LTrim(grd1.Text)
      '墨齊日
      grd1.col = 15
      StrDate2 = LTrim(grd1.Text)
      
      If Trim(strDate1) <> "" And Trim(StrDate2) <> "" Then
          '墨天
          grd1.col = 18
          grd1.Text = GetWorkDay(ChangeTStringToWString(ChangeTDateStringToTString(strDate1)), ChangeTStringToWString(ChangeTDateStringToTString(StrDate2)))
      Else
          '墨天
          grd1.col = 18
          grd1.Text = ""
      End If
      grd1.Text = grd1.Text & GetSign(grd1.TextMatrix(iRow, 23))
      If grd1.TextMatrix(iRow, 9) = "N" Then
          grd1.TextMatrix(iRow, 10) = " ******"
          grd1.TextMatrix(iRow, 12) = " ******"
          grd1.TextMatrix(iRow, 13) = ""
      End If
      If grd1.TextMatrix(iRow, 14) = "N" Then
          grd1.TextMatrix(iRow, 15) = " ******"
          grd1.TextMatrix(iRow, 17) = " ******"
          grd1.TextMatrix(iRow, 18) = ""
      End If
      ChgGrdColor True
   End If
   Set adoRst = Nothing
End Sub

'Modify By Sindy 2013/6/10
'Sub Process(strText As String)
Public Sub Process(strText As String)
'2013/6/10 End

'Added by Morgan 2013/10/3 重新抓資料庫更新到Grid否則會不即時(如工程師確認會稿完成日更新墨圖齊備日)
If Val(SWPRow) > 0 Then UpdateOneRecord
    
'代第2畫面資料
With grd1
    '收文號
'    .col = 20
    'edit by nick 2004/12/21 加了申請國家，要往後退
    '.col = 22
    .col = 23
    'edit by nickc 2005/06/28 修正，因為串錯了
    'strSQL = "SELECT CM01||'-'||CM02||'-'||CM03||'-'||CM04,ST02 FROM CASEMAP,CASEPROGRESS,STAFF WHERE CP01=CM01(+) AND CP02=CM02(+) AND CP03=CM03(+) AND CP04=CM04(+) AND CP14=ST01(+) AND CP31='Y' and cp09='" & .Text & "' order by CM01||'-'||CM02||'-'||CM03||'-'||CM04 "
    strSql = "SELECT CM01||'-'||CM02||'-'||CM03||'-'||CM04,ST02 FROM CASEMAP,CASEPROGRESS C1,caseprogress C2,STAFF WHERE c1.cp01=cm05(+) AND c1.cp02=CM06(+) AND c1.cp03=CM07(+) AND c1.cp04=CM08(+) AND C2.CP14=ST01(+) AND c2.CP31='Y' and c1.cp09='" & .Text & "' and cm01=c2.cp01(+) and cm02=c2.cp02(+) and cm03=c2.cp03(+) and cm04=c2.cp04(+) order by CM01||'-'||CM02||'-'||CM03||'-'||CM04 "
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
    'Modify By Cheng 2003/06/05
'    txt1(0) = Combo1.Text
    Txt1(0) = Trim(Left(Combo1.Text, 6))
    'Add By Cheng 2003/06/30
    '記錄原繪圖人員
    Me.Txt1(0).Tag = Me.Txt1(0).Text
    'Modify By Cheng 2003/09/22
    'Begin
'    lbl1(0).Caption = LBL2.Caption
    lbl1(0).Caption = GetPrjSalesNM(Me.Txt1(0).Text)
    'End
    
    'Modify by Morgan 2011/1/4 修正日期排序問題前面會補空白讀取資料時要去除
    
    '收文號
'    .col = 20
    'edit by nick 2004/12/21 加了申請國家，要往後退
    '.col = 22
    .col = 23
    lbl1(1).Caption = .Text
    '收文日
    .col = 1
    lbl1(2).Caption = .Text
    '本所案號
    'edit by nick 2004/12/21 加了申請國家，要往後退
    '.col = 2
    .col = 3
    lbl1(3).Caption = .Text
    'Add By Cheng 2002/04/29
    '是否閉卷
    If Right("" & .Text, 1) = "＊" Then
        Me.lblClose.Caption = "已閉卷"
    Else
        Me.lblClose.Caption = ""
    End If
    '案件名稱
    'edit by nick 2004/12/21 加了申請國家，要往後退
    '.col = 3
    .col = 4
    lbl1(4).Caption = .Text
    '案件性質名稱
'    .col = 21
    'edit by nick 2004/12/21 加了申請國家，要往後退
    '.col = 23
    .col = 24
    lbl1(5).Caption = .Text
    '案件性質
'    .col = 5
'    .col = 6
    'edit by nick 2004/12/21 加了申請國家，要往後退
    '.col = 4
    .col = 5
    lbl1(6).Caption = .Text
    '點數
'    .col = 8
    'edit by nick 2004/12/21 加了申請國家，要往後退
    '.col = 6
    .col = 7
    lbl1(7).Caption = .Text
    '本所期限
'    .col = 16
    'edit by nick 2004/12/21 加了申請國家，要往後退
    '.col = 18
    .col = 19
    lbl1(8).Caption = LTrim(.Text)
    '法定期限
'    .col = 22
    'edit by nick 2004/12/21 加了申請國家，要往後退
    '.col = 24
    .col = 25
    lbl1(9).Caption = LTrim(.Text)
    '承辦人
'    .col = 6
'    .col = 7
    'edit by nick 2004/12/21 加了申請國家，要往後退
    '.col = 5
    .col = 6
    lbl1(10).Caption = .Text
    '智權人員
'    .col = 19
    'edit by nick 2004/12/21 加了申請國家，要往後退
    .col = 22
    lbl1(11).Caption = .Text
    '取消收文日
'    .col = 23
    'edit by nick 2004/12/21 加了申請國家，要往後退
    '.col = 25
    .col = 26
    lbl1(14).Caption = LTrim(.Text)
    '草計
'    .col = 4
    'edit by nick 2004/12/21 加了申請國家，要往後退
    '.col = 8
    .col = 9
    'edit by nickc 2005/04/12
    'txt1(7).Text = .Text
    If .Text <> "N" Then
      Txt1(7).Text = ""
    Else
     Txt1(7).Text = "N"
   End If
    'Add By Cheng 2003/06/27
    '墨計
'    .col = 5
    'edit by nick 2004/12/21 加了申請國家，要往後退
    '.col = 13
    .col = 14
    'edit by nickc 2005/04/12
    'txt1(14).Text = .Text
    If .Text <> "N" Then
       Txt1(14).Text = ""
    Else
       Txt1(14).Text = "N"
    End If
    '草齊日
'    .col = 10
    'edit by nick 2004/12/21 加了申請國家，要往後退
    '.col = 9
    .col = 10
    
    If LTrim(.Text) = "******" Then
        Txt1(1).Text = LTrim(.Text)
    Else
        Txt1(1).Text = ChangeTDateStringToTString(LTrim(.Text))
    End If
    Txt1(1).Tag = Txt1(1) 'Added by Morgan 2012/8/13
    
    'Add By Cheng 2004/02/18
    '若從個人維護進入
    If ProState = "1" Then
        '草齊日若有值時設定不可改
        If Me.Txt1(1).Text <> "" Then
            Me.Txt1(1).Enabled = False
        Else
            Me.Txt1(1).Enabled = True
        End If
    End If
    'End
    'Add By Cheng 2003/06/30
    '草圖承辦期限
    'edit by nick 2004/12/21 加了申請國家，要往後退
    '.col = 10
    .col = 11
    Me.lbl1(31).Caption = LTrim(.Text)
    '草完日
    'edit by nick 2004/12/21 加了申請國家，要往後退
    '.col = 11
    .col = 12
    If LTrim(.Text) = "******" Then
        Txt1(2).Text = LTrim(.Text)
    Else
        Txt1(2).Text = ChangeTDateStringToTString(LTrim(.Text))
    End If
    'Add By Cheng 2004/02/18
    '若從個人維護進入
    If ProState = "1" Then
        '草完日若有值時設定不可改
        If Me.Txt1(2).Text <> "" Then
            Me.Txt1(2).Enabled = False
        Else
'            'Add By Sindy 2013/6/10
'            If Pub_StrUserSt03 = "M51" Or InStr(UCase(Pub_GetModuleFileName), "VB6.EXE") <> 0 Then
'               Me.txt1(2).Enabled = False
'            Else
'            '2013/6/10 End
               Me.Txt1(2).Enabled = True
'            End If
        End If
    End If
    'End
    '草圖張數
'    .col = 29
    'edit by nick 2004/12/21 加了申請國家，要往後退
    '.col = 31
    .col = 32
    Txt1(3) = .Text
    '墨齊日
'    .col = 13
    'edit by nick 2004/12/21 加了申請國家，要往後退
    '.col = 14
    .col = 15
    If LTrim(.Text) = "******" Then
        Txt1(4) = LTrim(.Text)
    Else
        Txt1(4) = ChangeTDateStringToTString(LTrim(.Text))
    End If
    Txt1(4).Tag = Txt1(4) 'Added by Morgan 2012/8/13
    
    'Add By Cheng 2004/02/18
    '若從個人維護進入
    If ProState = "1" Then
        '墨齊日若有值時設定不可改
        If Me.Txt1(4).Text <> "" Then
            Me.Txt1(4).Enabled = False
        Else
            Me.Txt1(4).Enabled = True
        End If
    End If
    'End
    'Add By Cheng 2003/06/30
    '墨圖承辦期限
    'edit by nick 2004/12/21 加了申請國家，要往後退
    '.col = 15
    .col = 16
    Me.lbl1(32).Caption = LTrim(.Text)
    '墨完日
'    .col = 14
    'edit by nick 2004/12/21 加了申請國家，要往後退
    '.col = 16
    .col = 17
    If LTrim(.Text) = "******" Then
        Txt1(5) = LTrim(.Text)
    Else
        Txt1(5) = ChangeTDateStringToTString(LTrim(.Text))
    End If
    'Add By Cheng 2004/02/18
    '若從個人維護進入
    If ProState = "1" Then
        '墨完日若有值時設定不可改
        If Me.Txt1(5).Text <> "" Then
            Me.Txt1(5).Enabled = False
        Else
'            'Add By Sindy 2013/6/10
'            If Pub_StrUserSt03 = "M51" Or InStr(UCase(Pub_GetModuleFileName), "VB6.EXE") <> 0 Then
'               Me.txt1(5).Enabled = False
'            Else
'            '2013/6/10 End
               Me.Txt1(5).Enabled = True
'            End If
        End If
    End If
    'End
    '草圖承辦天數
    If Len(Txt1(1)) <> 0 And Len(Txt1(2)) <> 0 Then
        'Modify By Cheng 2003/09/18
'        lbl1(12).Caption = GetWorkDay(ChangeTStringToWString(txt1(2)), ChangeTStringToWString(txt1(1)))
        'edit by nick 2004/12/21 加了申請國家，要往後退
        '.col = 12
        .col = 13
        lbl1(12).Caption = .Text
    Else
        lbl1(12).Caption = "0"
    End If
    '墨圖承辦天數
    If Len(Txt1(4)) <> 0 And Len(Txt1(5)) <> 0 Then
        'Modify By Cheng 2003/09/18
'        lbl1(13).Caption = GetWorkDay(ChangeTStringToWString(txt1(5)), ChangeTStringToWString(txt1(4)))
        'edit by nick 2004/12/21 加了申請國家，要往後退
        '.col = 17
        .col = 18
        lbl1(13).Caption = .Text
    Else
        lbl1(13).Caption = "0"
    End If
    '墨圖張數
'    .col = 30
    'edit by nick 2004/12/21 加了申請國家，要往後退
    '.col = 32
    .col = 33
    Txt1(6).Text = .Text
    '草圖承辦時數
'    .col = 24
    'edit by nick 2004/12/21 加了申請國家，要往後退
    '.col = 26
    .col = 27
    Txt1(8).Text = .Text
    '墨圖承辦時數
'    .col = 25
    'edit by nick 2004/12/21 加了申請國家，要往後退
    '.col = 27
    .col = 28
    Txt1(9).Text = .Text
    '修改時數1
'    .col = 26
    'edit by nick 2004/12/21 加了申請國家，要往後退
    '.col = 28
    .col = 29
    Txt1(10).Text = .Text
    '修改時數2
'    .col = 27
    'edit by nick 2004/12/21 加了申請國家，要往後退
    '.col = 29
    .col = 30
    Txt1(11).Text = .Text
    '修改時數3
'    .col = 28
    'edit by nick 2004/12/21 加了申請國家，要往後退
    '.col = 30
    .col = 31
    Txt1(12).Text = .Text
    '備註
'    .col = 18
    'edit by nick 2004/12/21 加了申請國家，要往後退
    '.col = 20
    .col = 21
    Txt1(13).Text = .Text
    'Add By Cheng 2004/02/18
    '若從個人維護進入
    If ProState = "1" Then
        '草齊日或草完日若有值時設定草圖是否計件不可改
        If Me.Txt1(1).Text <> "" Or Me.Txt1(2).Text <> "" Then
            Me.Txt1(7).Enabled = False
        Else
            Me.Txt1(7).Enabled = True
        End If
        '墨齊日或墨完日若有值時設定墨圖是否計件不可改
        If Me.Txt1(4).Text <> "" Or Me.Txt1(5).Text <> "" Then
            Me.Txt1(14).Enabled = False
        Else
            Me.Txt1(14).Enabled = True
        End If
        'add by nickc 2005/04/12         個人不能修改提不提供圖檔
        Txt1(19).Enabled = False
    Else
        'add by nickc 2005/04/12         個人不能修改提不提供圖檔
        Txt1(19).Enabled = True
    End If
    'End
    
   'Add by Morgan 2011/3/30
   m_CP21 = ""
   m_PA09 = ""
   m_bolExistsInCase = False
   'end 2011/3/30
   
   'Add By Sindy 2016/5/9 一案兩請
   strExc(1) = SystemNumber(lbl1(3).Caption, 1)
   strExc(2) = SystemNumber(lbl1(3).Caption, 2)
   strExc(3) = SystemNumber(lbl1(3).Caption, 3)
   strExc(4) = SystemNumber(lbl1(3).Caption, 4)
   strSql = "select * from casemap where cm01='" & strExc(1) & "' and cm02='" & strExc(2) & "' and cm03='" & strExc(3) & "' and cm04='" & strExc(4) & "' and cm10='3'" & _
            " Union select * from casemap where cm05='" & strExc(1) & "' and cm06='" & strExc(2) & "' and cm07='" & strExc(3) & "' and cm08='" & strExc(4) & "' and cm10='3'"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      lblCM10.Visible = True
   Else
      lblCM10.Visible = False
   End If
   '2016/5/9 END
    'Added by Lydia 2016/06/14 +台灣大陸案件提示
    lblCMboth.Caption = ""
    strExc(9) = GetPrjNation1(lbl1(3).Caption)
    If (strExc(1) = "P" Or strExc(1) = "FCP") And strExc(9) = 台灣國家代號 Then
       If PUB_GetRefCaseChk(strExc(1), strExc(2), strExc(3), strExc(4), "CASEMAP", "0", "A", 大陸國家代號) Then
          lblCMboth.Caption = "有大陸案"
       End If
    ElseIf strExc(1) = "P" And strExc(9) = 大陸國家代號 Then
       If PUB_GetRefCaseChk(strExc(1), strExc(2), strExc(3), strExc(4), "CASEMAP", "0", "A", 台灣國家代號) Then
          lblCMboth.Caption = "有台灣案"
       End If
    End If
    'end 2016/06/14
    
   'Add by Sindy 2013/9/17 電子送件
   'Modify by Sindy 2013/9/30 +casepropertymap
   strSql = "select CP118,CPM29 from caseprogress,casepropertymap" & _
            " where cp09='" & lbl1(1).Caption & "'" & _
            " and cp01=cpm01(+) and cp10=cpm02(+)"
   CheckOC3
   AdoRecordSet3.CursorLocation = adUseClient
   AdoRecordSet3.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   m_CPM29 = ""
   If AdoRecordSet3.RecordCount <> 0 Then
      If Not IsNull(AdoRecordSet3.Fields("CP118")) Then
         lblEApp.Visible = True
      Else
         lblEApp.Visible = False
      End If
      'Add By Sindy 2013/9/30
      m_CPM29 = "" & AdoRecordSet3.Fields("CPM29")
      If ProState = "1" Or Trim(Left("" & Combo1.Text, 6)) = strUserNum Then
         If m_CPM29 = "" Then '要電子簽核的案件性質
            '草圖完稿日
            Me.Txt1(2).Enabled = False
            '墨圖完稿日
            Me.Txt1(5).Enabled = False
         End If
      End If
      '2013/9/30 END
   End If
   CheckOC3
   '2013/9/17 END
   
    'add by nickc 2005/03/17 計件值加乘註記
    'Modify by Morgan 2011/3/30 +patent,casemap
    strSql = "select * from caseprogress,patent,casemap where cp09='" & lbl1(1).Caption & "'" & _
      " and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04" & _
      " and cm01(+)=cp01 and cm02(+)=cp02 and cm03(+)=cp03 and cm04(+)=cp04 and cm10(+)='0'"
    CheckOC3
    AdoRecordSet3.CursorLocation = adUseClient
    AdoRecordSet3.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If AdoRecordSet3.RecordCount <> 0 Then
         lbl1(33).Caption = "" & AdoRecordSet3.Fields("cp100").Value
         Txt1(15).Text = "" & AdoRecordSet3.Fields("cp101").Value
         m_CP101 = "" & AdoRecordSet3.Fields("cp101").Value
         Txt1(16).Text = "" & AdoRecordSet3.Fields("cp102").Value
         m_CP102 = "" & AdoRecordSet3.Fields("cp102").Value
         lbl1(34).Caption = "" & AdoRecordSet3.Fields("cp103").Value
         Txt1(17).Text = "" & AdoRecordSet3.Fields("cp104").Value
         m_CP104 = "" & AdoRecordSet3.Fields("cp104").Value
         Txt1(18).Text = "" & AdoRecordSet3.Fields("cp105").Value
         m_CP105 = "" & AdoRecordSet3.Fields("cp105").Value
         'add by nickc 2005/04/04
         Txt1(19) = "" & AdoRecordSet3.Fields("cp106").Value
         'Add by Morgan 2011/3/30
         m_CP21 = "" & AdoRecordSet3.Fields("cp21").Value
         m_PA09 = "" & AdoRecordSet3.Fields("pa09").Value
         If Not IsNull(AdoRecordSet3.Fields("cm05")) Then
            m_bolExistsInCase = True
         End If
         'end 2011/3/30
    Else
         lbl1(33).Caption = ""
         Txt1(15).Text = ""
         m_CP101 = ""
         Txt1(16).Text = ""
         m_CP102 = ""
         lbl1(34).Caption = ""
         Txt1(17).Text = ""
         m_CP104 = ""
         Txt1(18).Text = ""
         m_CP105 = ""
         'add by nickc 2005/04/04
         Txt1(19).Text = ""
    End If
    CheckOC3
        'add by nickc 2007/08/03 檢查有無代表圖
        'Modify by Amy 2018/07/19  改寫至function
'        strSql = "SELECT ibf01,ibf02 FROM imgbytefile WHERE ibf01='" & SystemNumber(lbl1(3), 1) & "' and ibf02='" & SystemNumber(lbl1(3), 2) & "' and ibf03='" & SystemNumber(lbl1(3), 3) & "' and ibf04='" & SystemNumber(lbl1(3), 4) & "' and ibf05='1' "
'        CheckOC2
'        adoRecordset1.CursorLocation = adUseClient
'        adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
'        If adoRecordset1.RecordCount <> 0 And adoRecordset1.RecordCount > 0 Then
        If ChkImgByteFile(SystemNumber(lbl1(3), 1), SystemNumber(lbl1(3), 2), SystemNumber(lbl1(3), 3), SystemNumber(lbl1(3), 4)) = True Then
            cmdPic.Caption = "已設定代表圖(&I)"
            cmdPic.BackColor = &HC0FFC0
        Else
            cmdPic.Caption = "未設定代表圖(&I)"
            cmdPic.BackColor = &HC0C0FF
        End If
'        CheckOC2
        'end 2018/07/19
    'add by nickc 2008/02/01 若是關聯案，草圖不計件，個人不可以改 墨圖張數
    Txt1(6).Enabled = True
    If Txt1(6) = "" Then Txt1(6) = "0"
    If Txt1(7) = "N" And Val(lbl1(34).Caption) = 0.4 And Txt1(14) = "" Then
         Txt1(6).Enabled = False
    ElseIf Txt1(7) = "" And ProState = "2" Then
        Txt1(6).Enabled = True
    End If
    
    'Add by Morgan 2010/5/10 個人才要--瓊玉
    Txt1(3).Enabled = True
    If ProState = "1" Then
      'Add by Morgan 2010/5/4
      '草圖要計件,加乘註記=0.2,張數0 不可改
      If Txt1(7) = "" And Val(Txt1(15)) = 0.2 And Val(Txt1(3)) = 0 Then
         Txt1(3).Enabled = False
      End If
      '墨圖要計件,加乘註記=0.2,張數0 不可改
      If Txt1(14) = "" And Val(Txt1(17)) = 0.2 And Val(Txt1(6)) = 0 Then
         Txt1(6).Enabled = False
      End If
      'end 2010/5/4
      
      'Add by Morgan 2011/3/30
      '多國案或有國內案的大陸案草墨不開放輸入
      If m_CP21 = "Y" Or (m_PA09 = "020" And m_bolExistsInCase = True) Then
         Txt1(3).Enabled = False
         Txt1(6).Enabled = False
      End If
   End If
End With
   
   Call SetColTag(True) 'Add By Sindy 2013/6/10
End Sub

'Add By Sindy 2013/6/10
'bolSetTag=true : 將輸入欄位值記錄至.tag裡面
'bolSetTag=false : 比較輸入欄位值.Tag與畫面上資料是否一致
Private Function SetColTag(bolSetTag As Boolean) As Boolean
   If bolSetTag = True Then
      Txt1(0).Tag = Txt1(0)
      Txt1(7).Tag = Txt1(7)
      Txt1(1).Tag = Txt1(1)
      Txt1(2).Tag = Txt1(2)
      Txt1(3).Tag = Txt1(3)
      Txt1(14).Tag = Txt1(14)
      Txt1(4).Tag = Txt1(4)
      Txt1(5).Tag = Txt1(5)
      Txt1(6).Tag = Txt1(6)
      Txt1(8).Tag = Txt1(8)
      Txt1(9).Tag = Txt1(9)
      Txt1(10).Tag = Txt1(10)
      Txt1(11).Tag = Txt1(11)
      Txt1(12).Tag = Txt1(12)
      Txt1(13).Tag = Txt1(13)
      Combo2.Tag = Combo2.Text
      Txt1(15).Tag = Txt1(15)
      Txt1(16).Tag = Txt1(16)
      Txt1(17).Tag = Txt1(17)
      Txt1(18).Tag = Txt1(18)
      Txt1(19).Tag = Txt1(19)
   Else
      SetColTag = True
      If Txt1(0).Tag <> Txt1(0) Then SetColTag = False: Exit Function
      If Txt1(7).Tag <> Txt1(7) Then SetColTag = False: Exit Function
      If Txt1(1).Tag <> Txt1(1) Then SetColTag = False: Exit Function
      If Txt1(2).Tag <> Txt1(2) Then SetColTag = False: Exit Function
      If Txt1(3).Tag <> Txt1(3) Then SetColTag = False: Exit Function
      If Txt1(14).Tag <> Txt1(14) Then SetColTag = False: Exit Function
      If Txt1(4).Tag <> Txt1(4) Then SetColTag = False: Exit Function
      If Txt1(5).Tag <> Txt1(5) Then SetColTag = False: Exit Function
      If Txt1(6).Tag <> Txt1(6) Then SetColTag = False: Exit Function
      If Txt1(8).Tag <> Txt1(8) Then SetColTag = False: Exit Function
      If Txt1(9).Tag <> Txt1(9) Then SetColTag = False: Exit Function
      If Txt1(10).Tag <> Txt1(10) Then SetColTag = False: Exit Function
      If Txt1(11).Tag <> Txt1(11) Then SetColTag = False: Exit Function
      If Txt1(12).Tag <> Txt1(12) Then SetColTag = False: Exit Function
      If Txt1(13).Tag <> Txt1(13) Then SetColTag = False: Exit Function
      If Combo2.Tag <> Combo2.Text Then SetColTag = False: Exit Function
      If Txt1(15).Tag <> Txt1(15) Then SetColTag = False: Exit Function
      If Txt1(16).Tag <> Txt1(16) Then SetColTag = False: Exit Function
      If Txt1(17).Tag <> Txt1(17) Then SetColTag = False: Exit Function
      If Txt1(18).Tag <> Txt1(18) Then SetColTag = False: Exit Function
      If Txt1(19).Tag <> Txt1(19) Then SetColTag = False: Exit Function
   End If
End Function

Sub StrMenu2()       '作本月統計資料
Dim strMonthLastDate As String '某月份最後一天
Dim strBeginDate As String
Dim strEndDate As String
    
    Screen.MousePointer = vbHourglass
    Me.Enabled = False
    DoEvents
    cnnConnection.Execute "DELETE FROM R090711_1 WHERE ID='" & strUserNum & "' "
    If ProState = "2" Or ProState = "3" Then
        strMonthLastDate = (Val(frm090706.Txt1(3).Text) + 1911) & Format(frm090706.Txt1(4).Text, "00") & PUB_GetMonthDays(Val(frm090706.Txt1(3).Text) + 1911, Val(frm090706.Txt1(4).Text))
    Else
        strMonthLastDate = strSrvDate(1)
    End If
    '統計其他項目
    '可辦草圖
    'Modify By Cheng 2003/07/07
'    strSQL = "INSERT INTO R090711_1 (R111001,R111002,R111003,ID) select EP13,1,count(*),'" & strUserNum & "' from engineerprogress,caseprogress where EP02=CP09(+) AND  cp01 in ('CFP','P')  AND EP13='" & Trim(Left(Combo1.Text, 6)) & "' AND CP27 is null and cp57 is null and ep14 is not null AND EP15 IS NULL GROUP BY EP13 "
    strSql = ""
    If ProState = "2" Or ProState = "3" Then '管理
        '若發文年月<系統年月
        If Val(frm090706.Txt1(3).Text) + 1911 & Format(frm090706.Txt1(4).Text, "00") < Left(strSrvDate(1), 6) Then
            strSql = strSql & " And CP05<=" & Val(strMonthLastDate) & " "
            'edit by nickc 2005/05/13
            'strSQL = strSQL & " And ((CP27 Is Null And CP57 Is Null) Or CP27>" & Val(strMonthLastDate) & " Or CP57>" & Val(strMonthLastDate) & " )"
            'edit by nickc 2006/01/11
            'strSQL = strSQL & " And ((CP27 Is Null And CP57 Is Null) Or (CP27>" & Val(strMonthLastDate) & ") Or (CP57>" & Val(strMonthLastDate) & " and cp27 is null) )"
            strSql = strSql & " And ((cp57 is null and cp27 is null) Or (CP27>" & Val(strMonthLastDate) & ") Or (CP57>" & Val(strMonthLastDate) & " and cp27 is null) or (CP05>" & Val(strMonthLastDate) & ") )"
            strSql = strSql & " And (EP14 Is Not Null And EP14<=" & Val(strMonthLastDate) & " ) "
            strSql = strSql & " And (EP15 Is Null Or EP15>" & Val(strMonthLastDate) & " ) "
        '若發文年月>=系統年月
        Else
            'edit by nick 2004/12/03 程弘改錯了
            'StrSql = StrSql & " AND CP27 is null and cp57 is null and ep14 is not null AND EP15 IS NULL And EP29 Is Null "
            strSql = strSql & " and cp57 is null and cp27 is null and ep14 is not null AND EP15 IS NULL And EP20 Is Null "
        End If
    Else '個人
        'edit by nick 2004/12/03 程弘改錯了
        'StrSql = StrSql & " AND CP27 is null and cp57 is null and ep14 is not null AND EP15 IS NULL And EP29 Is Null"
        strSql = strSql & " and cp57 is null and cp27 is null and ep14 is not null AND EP15 IS NULL And EP20 Is Null"
    End If
      'add by nick 2004/12/20   加多國案且草圖不計件不秀
      'edit by nickc 2005/03/01 墨圖也要判斷
      'StrSql = StrSql & " and ((cp21='Y' and ep20 is null) or cp21 is null) "
      strSql = strSql & " and ((cp21='Y' and (ep20 is null or ep29 is null)) or cp21 is null) "
      'add by nickc 2005/03/21
      strSql = strSql & " and  cp107='Y' "
      
'    strSQL = "INSERT INTO R090711_1 (R111001,R111002,R111003,ID) select EP13,1,count(*),'" & strUserNum & "' from engineerprogress,caseprogress where EP02=CP09(+) AND  cp01 in ('CFP','P','FCP')  AND EP13='" & Trim(Left(Combo1.Text, 6)) & "' AND CP27 is null and cp57 is null and ep14 is not null AND EP15 IS NULL And EP29 Is Null GROUP BY EP13 "
    'edit by nickc 2005/04/13
    'strSQL = "INSERT INTO R090711_1 (R111001,R111002,R111003,ID) select EP13,1,count(*),'" & strUserNum & "' from engineerprogress,caseprogress where EP02=CP09(+) AND  cp01 in ('CFP','P','FCP')  AND EP13='" & Trim(Left(Combo1.Text, 6)) & "' " & strSQL & " GROUP BY EP13 "
    'edit by nickc 2005/05/04
    'strSQL = "INSERT INTO R090711_1 (R111001,R111002,R111003,ID) select EP13,1,count(*),'" & strUserNum & "' from engineerprogress,caseprogress where EP02=CP09(+) AND  cp01 in ('CFP','P','FCP')  " & IIf(From090706 = True, "", " AND EP13='" & Trim(Left(Combo1.Text, 6)) & "' ") & " " & strSQL & " GROUP BY EP13 "
    strSql = "INSERT INTO R090711_1 (R111001,R111002,R111003,ID,r111006) select EP13,1,count(*),'" & strUserNum & "',sum(nvl(cp100,0) * nvl(cp101,0)) from engineerprogress,caseprogress where EP02=CP09(+) AND  cp01 in ('CFP','P','FCP')  " & IIf(From090706 = True, "", " AND EP13='" & Trim(Left(Combo1.Text, 6)) & "' ") & " " & strSql & " GROUP BY EP13 "
    cnnConnection.Execute strSql
    DoEvents
    '可辦墨圖
    'Modify By Cheng 2003/07/07
'    strSQL = "INSERT INTO R090711_1 (R111001,R111002,R111003,ID) select EP13,2,count(*),'" & strUserNum & "' from engineerprogress,caseprogress where EP02=CP09(+) AND  cp01 in ('CFP','P')   AND EP13='" & Trim(Left(Combo1.Text, 6)) & "' AND CP27 is null and cp57 is  null  and ep17 is not null and ep18 IS NULL GROUP BY EP13 "
    strSql = ""
    If ProState = "2" Or ProState = "3" Then '管理
        '若發文年月<系統年月
        If Val(frm090706.Txt1(3).Text) + 1911 & Format(frm090706.Txt1(4).Text, "00") < Left(strSrvDate(1), 6) Then
            strSql = strSql & " And CP05<=" & Val(strMonthLastDate) & " "
            'edit by nickc 2005/05/13
            'strSQL = strSQL & " And ((CP27 Is Null And CP57 Is Null) Or CP27>" & Val(strMonthLastDate) & " Or CP57>" & Val(strMonthLastDate) & " )"
            'edit by nickc 2006/01/11
            'strSQL = strSQL & " And ((CP27 Is Null And CP57 Is Null) Or (CP27>" & Val(strMonthLastDate) & ") Or (CP57>" & Val(strMonthLastDate) & " and cp27 is null) )"
            strSql = strSql & " And ((cp57 is null and cp27 is null) Or (CP27>" & Val(strMonthLastDate) & ") Or (CP57>" & Val(strMonthLastDate) & " and cp27 is null) or (CP05>" & Val(strMonthLastDate) & ") )"
            strSql = strSql & " And (EP17 Is Not Null And EP17<=" & Val(strMonthLastDate) & " ) "
            strSql = strSql & " And (EP18 Is Null Or EP18>" & Val(strMonthLastDate) & " ) "
        '若發文年月>=系統年月
        Else
            strSql = strSql & " and cp57 is null and cp27 is null and ep17 is not null and ep18 IS NULL And EP29 Is Null "
        End If
    Else '個人
        strSql = strSql & " and cp57 is null and cp27 is null and ep17 is not null and ep18 IS NULL And EP29 Is Null "
    End If
      'add by nick 2004/12/20   加多國案且草圖不計件不秀
      'StrSql = StrSql & " and ((cp21='Y' and ep20 is null) or cp21 is null) "
      'edit by nickc 2005/03/01
      strSql = strSql & " and ((cp21='Y' and (ep20 is null or ep29 is null)) or cp21 is null) "
      'add by nickc 2005/03/21
      strSql = strSql & " and  cp107='Y' "
      
'    strSQL = "INSERT INTO R090711_1 (R111001,R111002,R111003,ID) select EP13,2,count(*),'" & strUserNum & "' from engineerprogress,caseprogress where EP02=CP09(+) AND  cp01 in ('CFP','P','FCP')   AND EP13='" & Trim(Left(Combo1.Text, 6)) & "' AND CP27 is null and cp57 is  null  and ep17 is not null and ep18 IS NULL And EP29 Is Null GROUP BY EP13 "
    'edit by nickc 2005/04/13
    'strSQL = "INSERT INTO R090711_1 (R111001,R111002,R111003,ID) select EP13,2,count(*),'" & strUserNum & "' from engineerprogress,caseprogress where EP02=CP09(+) AND  cp01 in ('CFP','P','FCP')   AND EP13='" & Trim(Left(Combo1.Text, 6)) & "' " & strSQL & " GROUP BY EP13 "
    'edit by nickc 2005/05/04
    'strSQL = "INSERT INTO R090711_1 (R111001,R111002,R111003,ID) select EP13,2,count(*),'" & strUserNum & "' from engineerprogress,caseprogress where EP02=CP09(+) AND  cp01 in ('CFP','P','FCP')   " & IIf(From090706 = True, "", " AND EP13='" & Trim(Left(Combo1.Text, 6)) & "' ") & " " & strSQL & " GROUP BY EP13 "
    strSql = "INSERT INTO R090711_1 (R111001,R111002,R111003,ID,r111006) select EP13,2,count(*),'" & strUserNum & "',sum(nvl(cp103,0) * nvl(cp104,0)) from engineerprogress,caseprogress where EP02=CP09(+) AND  cp01 in ('CFP','P','FCP')   " & IIf(From090706 = True, "", " AND EP13='" & Trim(Left(Combo1.Text, 6)) & "' ") & " " & strSql & " GROUP BY EP13 "
    cnnConnection.Execute strSql
    DoEvents
    '達成草圖
    'Modify By Cheng 2003/07/07
'    strSQL = "INSERT INTO R090711_1 (R111001,R111002,r111003,R111004,ID) select EP13,3,count(*),SUM(eP16),'" & strUserNum & "' from engineerprogress,caseprogress where EP02=CP09(+) AND EP13='" & Trim(Left(Combo1.Text, 6)) & "' AND  CP01 IN ('P','CFP')  and cp57 is  null  and SUBSTR(EP15,1,6)=" & Val(Text1) + 191100 & " AND EP16 IS NOT NULL AND EP16>0 GROUP BY EP13 "
'    strSQL = "INSERT INTO R090711_1 (R111001,R111002,r111003,R111004,ID) select EP13,3,Sum(Decode(EP29, Null, 1, 0)),SUM(eP16),'" & strUserNum & "' from engineerprogress,caseprogress where EP02=CP09(+) AND EP13='" & Trim(Left(Combo1.Text, 6)) & "' AND  CP01 IN ('P','CFP','FCP')  and cp57 is  null  and SUBSTR(EP15,1,6)=" & Val(Text1) + 191100 & " AND EP16 IS NOT NULL AND EP16>0 GROUP BY EP13 "
    'edit by nick 2005/01/04
    'StrSql = "INSERT INTO R090711_1 (R111001,R111002,r111003,R111004,ID) select EP13,3,Sum(Decode(EP29, Null, 1, 0)),SUM(eP16),'" & strUserNum & "' from engineerprogress,caseprogress where EP02=CP09(+) AND EP13='" & Trim(Left(Combo1.Text, 6)) & "' AND  CP01 IN ('P','CFP','FCP') And SUBSTR(EP15,1,6)=" & Val(Text1) + 191100 & " AND EP16 IS NOT NULL AND EP16>0  GROUP BY EP13 "
    'edit by nickc 2005/03/01 墨圖也要判斷
    'StrSql = "INSERT INTO R090711_1 (R111001,R111002,r111003,R111004,ID) select EP13,3,Sum(Decode(EP29, Null, 1, 0)),SUM(eP16),'" & strUserNum & "' from engineerprogress,caseprogress where EP02=CP09(+) AND EP13='" & Trim(Left(Combo1.Text, 6)) & "' AND  CP01 IN ('P','CFP','FCP') And SUBSTR(EP15,1,6)=" & Val(Text1) + 191100 & " AND EP16 IS NOT NULL AND EP16>0 and ((cp21='Y' and ep20 is null) or cp21 is null)  GROUP BY EP13 "
    'edit by nickc 2005/04/13
    'strSQL = "INSERT INTO R090711_1 (R111001,R111002,r111003,R111004,ID) select EP13,3,Sum(Decode(EP29, Null, 1, 0)),SUM(eP16),'" & strUserNum & "' from engineerprogress,caseprogress where EP02=CP09(+) AND EP13='" & Trim(Left(Combo1.Text, 6)) & "' AND  CP01 IN ('P','CFP','FCP') And SUBSTR(EP15,1,6)=" & Val(Text1) + 191100 & " AND EP16 IS NOT NULL AND EP16>0 and ((cp21='Y' and (ep20 is null or ep29 is null)) or cp21 is null) and cp107='Y'  GROUP BY EP13 "
    'edit by nickc 2005/05/04
    'strSQL = "INSERT INTO R090711_1 (R111001,R111002,r111003,R111004,ID) select EP13,3,Sum(Decode(EP29, Null, 1, 0)),SUM(eP16),'" & strUserNum & "' from engineerprogress,caseprogress where EP02=CP09(+) " & IIf(From090706 = True, "", " AND EP13='" & Trim(Left(Combo1.Text, 6)) & "' ") & " AND  CP01 IN ('P','CFP','FCP') And SUBSTR(EP15,1,6)=" & Val(Text1) + 191100 & " AND EP16 IS NOT NULL AND EP16>0 and ((cp21='Y' and (ep20 is null or ep29 is null)) or cp21 is null) and cp107='Y'  GROUP BY EP13 "
    strSql = "INSERT INTO R090711_1 (R111001,R111002,r111003,R111004,ID,r111006,r111007) select EP13,3,Sum(Decode(EP29, Null, 1, 0)),sum(nvl(ep16,0)),'" & strUserNum & "',sum(nvl(cp100,0) * nvl(cp101,0)),sum(nvl(ep16,0)) from engineerprogress,caseprogress where EP02=CP09(+) " & IIf(From090706 = True, "", " AND EP13='" & Trim(Left(Combo1.Text, 6)) & "' ") & " AND  CP01 IN ('P','CFP','FCP') And (EP15>=" & Val(Text1) + 191100 & "01 and EP15<=" & Val(Text1) + 191100 & "31) and ((cp21='Y' and (ep20 is null or ep29 is null)) or cp21 is null) and cp107='Y'  GROUP BY EP13 "
    cnnConnection.Execute strSql
    'Add By Cheng 2004/03/30
    '將支援記錄加入達成草圖
    'edit by nickc 2005/04/13
    'strSQL = "INSERT INTO R090711_1 (R111001,R111002,r111003,R111004,ID) select SH02,3,Sum(Decode(SH06, 'CFP', Nvl(SH05, 0)/8, Nvl(SH05, 0)/4)), 0,'" & strUserNum & "' from SupportHour where SH02='" & Trim(Left(Combo1.Text, 6)) & "' AND  SH06 IN ('P','CFP','FCP') And SUBSTR(SH01,1,6)=" & Val(Text1) + 191100 & " And SH11 ='V'  GROUP BY SH02 "
    'edit by nickc 2005/05/04
    'strSQL = "INSERT INTO R090711_1 (R111001,R111002,r111003,R111004,ID) select SH02,3,Sum(Decode(SH06, 'CFP', Nvl(SH05, 0)/4, Nvl(SH05, 0)/4)), 0,'" & strUserNum & "' from SupportHour where  SH06 IN ('P','CFP','FCP') " & IIf(From090706 = True, "", " AND sh02='" & Trim(Left(Combo1.Text, 6)) & "' ") & " And SUBSTR(SH01,1,6)=" & Val(Text1) + 191100 & " And SH11 ='V'  GROUP BY SH02 "
    'Modified by Morgan 2012/1/2 不限制系統別(與每週速度考核一致)
    'strSql = "INSERT INTO R090711_1 (R111001,R111002,r111003,R111004,ID,r111006,r111007) select SH02,3,Sum(Decode(SH06, 'CFP', Nvl(SH05, 0)/4, Nvl(SH05, 0)/4)), 0,'" & strUserNum & "',Sum(Decode(SH06, 'CFP', Nvl(SH05, 0)/4, Nvl(SH05, 0)/4)), 0 from SupportHour where  SH06 IN ('P','CFP','FCP') " & IIf(From090706 = True, "", " AND sh02='" & Trim(Left(Combo1.Text, 6)) & "' ") & " And SUBSTR(SH01,1,6)=" & Val(Text1) + 191100 & " And SH11 ='V'  GROUP BY SH02 "
    'Modified by Morgan 2014/3/20 --2014/4/1起支援改每小時折計0.2基數
    'strSql = "INSERT INTO R090711_1 (R111001,R111002,r111003,R111004,ID,r111006,r111007) select SH02,3,Sum(Decode(SH06, 'CFP', Nvl(SH05, 0)/4, Nvl(SH05, 0)/4)), 0,'" & strUserNum & "',Sum(Decode(SH06, 'CFP', Nvl(SH05, 0)/3, Nvl(SH05, 0)/4))*0.65*2, 0 from SupportHour where SH11 ='V' " & IIf(From090706 = True, "", " AND sh02='" & Trim(Left(Combo1.Text, 6)) & "' ") & " And SUBSTR(SH01,1,6)=" & Val(Text1) + 191100 & " GROUP BY SH02 "
    strSql = "INSERT INTO R090711_1 (R111001,R111002,r111003,R111004,ID,r111006,r111007) select SH02,3,Sum(Nvl(SH05, 0)/4), 0,'" & strUserNum & "',Sum(" & Sh2EPtCode & ")*0.65*2, 0 from SupportHour where SH11 ='V' " & IIf(From090706 = True, "", " AND sh02='" & Trim(Left(Combo1.Text, 6)) & "' ") & " And (SH01>=" & Val(Text1) + 191100 & "01 and SH01<=" & Val(Text1) + 191100 & "31) GROUP BY SH02 "
    'end 2014/3/20
    cnnConnection.Execute strSql
    'End
    DoEvents
    '達成墨圖
    'Modify By Cheng 2003/07/07
'    strSQL = "INSERT INTO R090711_1 (R111001,R111002,r111003,R111004,r111005,ID) select EP13,4,count(*),SUM(eP19),sum(cp18),'" & strUserNum & "' from engineerprogress,caseprogress where EP02=CP09(+) AND EP13='" & Trim(Left(Combo1.Text, 6)) & "' AND  cp01 IN ('P','CFP')  and cp57 is  null and SUBSTR(EP18,1,6)=" & Val(Text1) + 191100 & " GROUP BY EP13 "
'    strSQL = "INSERT INTO R090711_1 (R111001,R111002,r111003,R111004,r111005,ID) select EP13,4,Sum(Decode(EP29, Null, 1, 0)),SUM(eP19),sum(cp18),'" & strUserNum & "' from engineerprogress,caseprogress where EP02=CP09(+) AND EP13='" & Trim(Left(Combo1.Text, 6)) & "' AND  cp01 IN ('P','CFP','FCP')  and cp57 is  null and SUBSTR(EP18,1,6)=" & Val(Text1) + 191100 & " GROUP BY EP13 "
    'edit by nick 2005/01/04
    'StrSql = "INSERT INTO R090711_1 (R111001,R111002,r111003,R111004,r111005,ID) select EP13,4,Sum(Decode(EP29, Null, 1, 0)),SUM(eP19),sum(cp18),'" & strUserNum & "' from engineerprogress,caseprogress where EP02=CP09(+) AND EP13='" & Trim(Left(Combo1.Text, 6)) & "' AND  cp01 IN ('P','CFP','FCP') And SUBSTR(EP18,1,6)=" & Val(Text1) + 191100 & " GROUP BY EP13 "
    'StrSql = "INSERT INTO R090711_1 (R111001,R111002,r111003,R111004,r111005,ID) select EP13,4,Sum(Decode(EP29, Null, 1, 0)),SUM(eP19),sum(cp18),'" & strUserNum & "' from engineerprogress,caseprogress where EP02=CP09(+) AND EP13='" & Trim(Left(Combo1.Text, 6)) & "' AND  cp01 IN ('P','CFP','FCP') And SUBSTR(EP18,1,6)=" & Val(Text1) + 191100 & " and ((cp21='Y' and ep20 is null) or cp21 is null)  GROUP BY EP13 "
    'edit by nickc 2005/03/01
    'edit by nickc 2005/04/13
    'strSQL = "INSERT INTO R090711_1 (R111001,R111002,r111003,R111004,r111005,ID) select EP13,4,Sum(Decode(EP29, Null, 1, 0)),SUM(eP19),sum(cp18),'" & strUserNum & "' from engineerprogress,caseprogress where EP02=CP09(+) AND EP13='" & Trim(Left(Combo1.Text, 6)) & "' AND  cp01 IN ('P','CFP','FCP') And SUBSTR(EP18,1,6)=" & Val(Text1) + 191100 & " and ((cp21='Y' and (ep20 is null or ep29 is null)) or cp21 is null) and cp107='Y'  GROUP BY EP13 "
    'edit by nickc 2005/05/04
    'strSQL = "INSERT INTO R090711_1 (R111001,R111002,r111003,R111004,r111005,ID) select EP13,4,Sum(Decode(EP29, Null, 1, 0)),SUM(eP19),sum(cp18),'" & strUserNum & "' from engineerprogress,caseprogress where EP02=CP09(+) " & IIf(From090706 = True, "", " AND EP13='" & Trim(Left(Combo1.Text, 6)) & "' ") & " AND  cp01 IN ('P','CFP','FCP') And SUBSTR(EP18,1,6)=" & Val(Text1) + 191100 & " and ((cp21='Y' and (ep20 is null or ep29 is null)) or cp21 is null) and cp107='Y'  GROUP BY EP13 "
    strSql = "INSERT INTO R090711_1 (R111001,R111002,r111003,R111004,r111005,ID,r111006,r111007,r111008) select EP13,4,Sum(ep291),sum(ep292),sum(ep293),'" & strUserNum & "',sum(ep294),sum(ep295),sum(ep296) from (select distinct cp09,ep13,Decode(EP29, Null, 1, 0) as ep291,nvl(ep19,0) as ep292,cp18 as ep293,nvl(cp103,0) * nvl(cp104,0) as ep294,nvl(ep19,0) as ep295,nvl(cp18,0) - nvl(a1u07/1000,0) as ep296 from engineerprogress,caseprogress,(select a1u03,sum(nvl(a1u07,0)) as a1u07 from acc1u0 group by a1u03) ABCDE where EP02=CP09(+) and ep02=a1u03(+) " & IIf(From090706 = True, "", " AND EP13='" & Trim(Left(Combo1.Text, 6)) & "' ") & " AND  cp01 IN ('P','CFP','FCP') And (EP18>=" & Val(Text1) + 191100 & "01 and EP18<=" & Val(Text1) + 191100 & "31) and ((cp21='Y' and (ep20 is null or ep29 is null)) or cp21 is null) and cp107='Y' ) AA  GROUP BY AA.EP13 "
    cnnConnection.Execute strSql
    'Add By Cheng 2004/03/30
    '將支援記錄加入達成墨圖
    'edit by nickc 2005/04/13
    'strSQL = "INSERT INTO R090711_1 (R111001,R111002,r111003,R111004,r111005,ID) select SH02,4,Sum(Decode(SH06,'CFP', Nvl(SH05,0)/8, Nvl(SH05, 0)/4)), 0, 0,'" & strUserNum & "' from SupportHour where SH02='" & Trim(Left(Combo1.Text, 6)) & "' AND  SH06 IN ('P','CFP','FCP') And SUBSTR(SH01,1,6)=" & Val(Text1) + 191100 & " And SH11 ='V' GROUP BY SH02 "
    'edit by nickc 2005/05/04
    'strSQL = "INSERT INTO R090711_1 (R111001,R111002,r111003,R111004,r111005,ID) select SH02,4,Sum(Decode(SH06,'CFP', Nvl(SH05,0)/8, Nvl(SH05, 0)/4)), 0, 0,'" & strUserNum & "' from SupportHour where  SH06 IN ('P','CFP','FCP') " & IIf(From090706 = True, "", " AND sh02='" & Trim(Left(Combo1.Text, 6)) & "' ") & " And SUBSTR(SH01,1,6)=" & Val(Text1) + 191100 & " And SH11 ='V' GROUP BY SH02 "
    'Modified by Morgan 2012/1/2 不限制系統別且比例有改(與每週速度考核一致)
    'strSql = "INSERT INTO R090711_1 (R111001,R111002,r111003,R111004,r111005,ID,r111006,r111007,r111008) select SH02,4,Sum(Decode(SH06,'CFP', Nvl(SH05,0)/4, Nvl(SH05, 0)/4)), 0, 0,'" & strUserNum & "',Sum(Decode(SH06,'CFP', Nvl(SH05,0)/4, Nvl(SH05, 0)/4)), 0, 0 from SupportHour where  SH06 IN ('P','CFP','FCP') " & IIf(From090706 = True, "", " AND sh02='" & Trim(Left(Combo1.Text, 6)) & "' ") & " And SUBSTR(SH01,1,6)=" & Val(Text1) + 191100 & " And SH11 ='V' GROUP BY SH02 "
    'Modified by Morgan 2014/3/20 --2014/4/1起支援改每小時折計0.2基數
    'strSql = "INSERT INTO R090711_1 (R111001,R111002,r111003,R111004,r111005,ID,r111006,r111007,r111008) select SH02,4,Sum(Decode(SH06,'CFP', Nvl(SH05,0)/4, Nvl(SH05, 0)/4)), 0, 0,'" & strUserNum & "',Sum(Decode(SH06,'CFP', Nvl(SH05,0)/3, Nvl(SH05, 0)/4))*0.35*2, 0, 0 from SupportHour where SH11 ='V' " & IIf(From090706 = True, "", " AND sh02='" & Trim(Left(Combo1.Text, 6)) & "' ") & " And SUBSTR(SH01,1,6)=" & Val(Text1) + 191100 & " GROUP BY SH02 "
    strSql = "INSERT INTO R090711_1 (R111001,R111002,r111003,R111004,r111005,ID,r111006,r111007,r111008) select SH02,4,Sum(Nvl(SH05,0)/4), 0, 0,'" & strUserNum & "',Sum(" & Sh2EPtCode & ")*0.35*2, 0, 0 from SupportHour where SH11 ='V' " & IIf(From090706 = True, "", " AND sh02='" & Trim(Left(Combo1.Text, 6)) & "' ") & " And (SH01>=" & Val(Text1) + 191100 & "01 and SH01<=" & Val(Text1) + 191100 & "31) GROUP BY SH02 "
    'end 2014/3/20
    cnnConnection.Execute strSql
    'End
    DoEvents
    '其他新案(抓CFP案草圖張數>0且草完日為當月的資料)
    'Modify By Cheng 2003/07/07
'    strSQL = "INSERT INTO R090711_1 (R111001,R111002,r111003,R111004,r111005,ID) select EP13,5,count(*),SUM(eP19),sum(cp18),'" & strUserNum & "' from engineerprogress,caseprogress where EP02=CP09(+) AND EP13='" & Trim(Left(Combo1.Text, 6)) & "' AND  CP01 NOT IN ('P','CFP')  AND SUBSTR(CP27,1,6)=" & Val(Text1) + 191100 & "  and cp57 is  null and EP18 IS NOT NULL AND cp31='Y'  GROUP BY EP13 "
    'Modify By Cheng 2004/02/16
'    strSQL = "INSERT INTO R090711_1 (R111001,R111002,r111003,R111004,r111005,ID) select EP13,5,Sum(Decode(EP29, Null, 1, 0)),SUM(eP19),sum(cp18),'" & strUserNum & "' from engineerprogress,caseprogress where EP02=CP09(+) AND EP13='" & Trim(Left(Combo1.Text, 6)) & "' AND  CP01 NOT IN ('P','CFP','FCP')  AND SUBSTR(CP27,1,6)=" & Val(Text1) + 191100 & "  and cp57 is  null and EP18 IS NOT NULL AND cp31='Y'  GROUP BY EP13 "
'    strSQL = "INSERT INTO R090711_1 (R111001,R111002,r111003,R111004,r111005,ID) select EP13,5,Sum(Decode(EP20, Null, 1, 0)),SUM(EP16),sum(cp18),'" & strUserNum & "' from engineerprogress,caseprogress where EP02=CP09(+) AND EP13='" & Trim(Left(Combo1.Text, 6)) & "' AND  CP01 IN ('CFP')  AND ((SUBSTR(CP27,1,6)=" & Val(Text1) + 191100 & "  and cp57 is  null ) Or (SUBSTR(CP57,1,6)=" & Val(Text1) + 191100 & "  and cp27 is  null) Or (CP27 Is Null  and cp57 is  null)) And substr(EP15,1,6)=" & Val(Text1) + 191100 & " and EP16>0  GROUP BY EP13 "
    'edit by nick 2005/01/04
    'StrSql = "INSERT INTO R090711_1 (R111001,R111002,r111003,R111004,r111005,ID) select EP13,5,Sum(Decode(EP20, Null, 1, 0)),SUM(EP16),sum(cp18),'" & strUserNum & "' from engineerprogress,caseprogress where EP02=CP09(+) AND EP13='" & Trim(Left(Combo1.Text, 6)) & "' AND  CP01 IN ('CFP') And substr(EP15,1,6)=" & Val(Text1) + 191100 & " and EP16>0  GROUP BY EP13 "
    'edit by nick 2005/03/01
    'StrSql = "INSERT INTO R090711_1 (R111001,R111002,r111003,R111004,r111005,ID) select EP13,5,Sum(Decode(EP20, Null, 1, 0)),SUM(EP16),sum(cp18),'" & strUserNum & "' from engineerprogress,caseprogress where EP02=CP09(+) AND EP13='" & Trim(Left(Combo1.Text, 6)) & "' AND  CP01 IN ('CFP') And substr(EP15,1,6)=" & Val(Text1) + 191100 & " and EP16>0  and ((cp21='Y' and ep20 is null) or cp21 is null) GROUP BY EP13 "
    'edit by nickc 2005/04/13
    'strSQL = "INSERT INTO R090711_1 (R111001,R111002,r111003,R111004,r111005,ID) select EP13,5,Sum(Decode(EP20, Null, 1, 0)),SUM(EP16),sum(cp18),'" & strUserNum & "' from engineerprogress,caseprogress where EP02=CP09(+) AND EP13='" & Trim(Left(Combo1.Text, 6)) & "' AND  CP01 IN ('CFP') And substr(EP15,1,6)=" & Val(Text1) + 191100 & " and EP16>0  and ((cp21='Y' and (ep20 is null or ep29 is null) ) or cp21 is null) GROUP BY EP13 "
    strSql = "INSERT INTO R090711_1 (R111001,R111002,r111003,R111004,r111005,ID) select EP13,5,Sum(Decode(EP20, Null, 1, 0)),sum(nvl(ep16,0)),sum(cp18),'" & strUserNum & "' from engineerprogress,caseprogress where EP02=CP09(+) " & IIf(From090706 = True, "", " AND EP13='" & Trim(Left(Combo1.Text, 6)) & "' ") & " AND  CP01 IN ('CFP') And (EP15>=" & Val(Text1) + 191100 & "01 and EP15<=" & Val(Text1) + 191100 & "31) and ((cp21='Y' and (ep20 is null or ep29 is null) ) or cp21 is null) GROUP BY EP13 "
    'End
    cnnConnection.Execute strSql
    DoEvents
    '其他舊案(抓CFP案草圖張數<=0, 墨圖張數<=0且墨完日為當月的資料)
    'Modify By Cheng 2003/07/07
'    strSQL = "INSERT INTO R090711_1 (R111001,R111002,r111003,R111004,r111005,ID) select EP13,6,count(*),SUM(eP19),sum(cp18),'" & strUserNum & "' from engineerprogress,caseprogress where EP02=CP09(+) AND EP13='" & Trim(Left(Combo1.Text, 6)) & "' AND  CP01 NOT IN ('P','CFP')  AND SUBSTR(CP27,1,6)=" & Val(Text1) + 191100 & "  and cp57 is  null and EP18 IS NOT NULL AND (CP31<>'Y' OR CP31 IS NULL) GROUP BY EP13 "
    'Modify By Cheng 2004/02/16
'    strSQL = "INSERT INTO R090711_1 (R111001,R111002,r111003,R111004,r111005,ID) select EP13,6,Sum(Decode(EP29, Null, 1, 0)),SUM(eP19),sum(cp18),'" & strUserNum & "' from engineerprogress,caseprogress where EP02=CP09(+) AND EP13='" & Trim(Left(Combo1.Text, 6)) & "' AND  CP01 NOT IN ('P','CFP','FCP')  AND SUBSTR(CP27,1,6)=" & Val(Text1) + 191100 & "  and cp57 is  null and EP18 IS NOT NULL AND (CP31<>'Y' OR CP31 IS NULL) GROUP BY EP13 "
'    strSQL = "INSERT INTO R090711_1 (R111001,R111002,r111003,R111004,r111005,ID) select EP13,6,Sum(Decode(EP29, Null, 1, 0)),SUM(EP19),sum(cp18),'" & strUserNum & "' from engineerprogress,caseprogress where EP02=CP09(+) AND EP13='" & Trim(Left(Combo1.Text, 6)) & "' AND  CP01 IN ('CFP')  AND ((SUBSTR(CP27,1,6)=" & Val(Text1) + 191100 & "  and cp57 is  null ) Or (SUBSTR(CP57,1,6)=" & Val(Text1) + 191100 & "  and cp27 is  null) Or (CP27 Is Null  and cp57 is  null)) and substr(EP18,1,6)=" & Val(Text1) + 191100 & " and ((EP16<=0 Or EP16 Is Null) And (EP19<=0 Or EP19 Is Null)) GROUP BY EP13 "
    '不必判斷墨圖=0
    'strSQL = "INSERT INTO R090711_1 (R111001,R111002,r111003,R111004,r111005,ID) select EP13,6,Sum(Decode(EP29, Null, 1, 0)),SUM(EP19),sum(cp18),'" & strUserNum & "' from engineerprogress,caseprogress where EP02=CP09(+) AND EP13='" & Trim(Left(Combo1.Text, 6)) & "' AND  CP01 IN ('CFP')  and substr(EP18,1,6)=" & Val(Text1) + 191100 & " and ((EP16<=0 Or EP16 Is Null) And (EP19<=0 Or EP19 Is Null)) GROUP BY EP13 "
    'edit by nick 2005/01/04
    'StrSql = "INSERT INTO R090711_1 (R111001,R111002,r111003,R111004,r111005,ID) select EP13,6,Sum(Decode(EP29, Null, 1, 0)),SUM(EP19),sum(cp18),'" & strUserNum & "' from engineerprogress,caseprogress where EP02=CP09(+) AND EP13='" & Trim(Left(Combo1.Text, 6)) & "' AND  CP01 IN ('CFP')  and substr(EP18,1,6)=" & Val(Text1) + 191100 & " and (EP16<=0 Or EP16 Is Null) GROUP BY EP13 "
    'edit by nickc 2005/03/01
    'StrSql = "INSERT INTO R090711_1 (R111001,R111002,r111003,R111004,r111005,ID) select EP13,6,Sum(Decode(EP29, Null, 1, 0)),SUM(EP19),sum(cp18),'" & strUserNum & "' from engineerprogress,caseprogress where EP02=CP09(+) AND EP13='" & Trim(Left(Combo1.Text, 6)) & "' AND  CP01 IN ('CFP')  and substr(EP18,1,6)=" & Val(Text1) + 191100 & " and (EP16<=0 Or EP16 Is Null) and ((cp21='Y' and ep20 is null) or cp21 is null) GROUP BY EP13 "
    'edit by nickc 2005/04/13
    'strSQL = "INSERT INTO R090711_1 (R111001,R111002,r111003,R111004,r111005,ID) select EP13,6,Sum(Decode(EP29, Null, 1, 0)),SUM(EP19),sum(cp18),'" & strUserNum & "' from engineerprogress,caseprogress where EP02=CP09(+) AND EP13='" & Trim(Left(Combo1.Text, 6)) & "' AND  CP01 IN ('CFP')  and substr(EP18,1,6)=" & Val(Text1) + 191100 & " and (EP16<=0 Or EP16 Is Null) and ((cp21='Y' and (ep20 is null or ep29 is null)) or cp21 is null) and cp107='Y' GROUP BY EP13 "
    strSql = "INSERT INTO R090711_1 (R111001,R111002,r111003,R111004,r111005,ID) select EP13,6,Sum(Decode(EP29, Null, 1, 0)),sum(nvl(ep19,0)),sum(cp18),'" & strUserNum & "' from engineerprogress,caseprogress where EP02=CP09(+) " & IIf(From090706 = True, "", " AND EP13='" & Trim(Left(Combo1.Text, 6)) & "' ") & " AND  CP01 IN ('CFP')  and (EP18>=" & Val(Text1) + 191100 & "01 and EP18<=" & Val(Text1) + 191100 & "31) and (EP16<=0 Or EP16 Is Null) and ((cp21='Y' and (ep20 is null or ep29 is null)) or cp21 is null) and cp107='Y' GROUP BY EP13 "
    'End
    cnnConnection.Execute strSql
    DoEvents
    'Add By Cheng 2003/07/01
    '本月發文
    'Modify By Cheng 2003/07/07
'    strSQL = "INSERT INTO R090711_1 (R111001,R111002,r111003,R111004,r111005,ID) select EP13,9,count(*),SUM(eP19),sum(cp18),'" & strUserNum & "' from engineerprogress,caseprogress where EP02=CP09(+) AND EP13='" & Trim(Left(Combo1.Text, 6)) & "' AND  CP01 IN ('P','CFP')  AND SUBSTR(CP27,1,6)=" & Val(Text1) + 191100 & "  and cp57 is  null and EP18 IS NOT NULL GROUP BY EP13 "
    'Modify By Cheng 2003/07/16
'    strSQL = "INSERT INTO R090711_1 (R111001,R111002,r111003,R111004,r111005,ID) select EP13,9,Sum(Decode(EP29, Null, 1, 0)),SUM(eP19),sum(cp18),'" & strUserNum & "' from engineerprogress,caseprogress where EP02=CP09(+) AND EP13='" & Trim(Left(Combo1.Text, 6)) & "' AND  CP01 IN ('P','CFP','FCP')  AND SUBSTR(CP27,1,6)=" & Val(Text1) + 191100 & "  and cp57 is  null and EP18 IS NOT NULL GROUP BY EP13 "
    'Modify By Cheng 2004/02/17
'    strSQL = "INSERT INTO R090711_1 (R111001,R111002,r111003,R111004,r111005,ID) select EP13,9,Sum(Decode(EP29, Null, 1, 0)),SUM(eP19),sum(cp18),'" & strUserNum & "' from engineerprogress,caseprogress where EP02=CP09(+) AND EP13='" & Trim(Left(Combo1.Text, 6)) & "' AND  CP01 IN ('P','CFP','FCP')  AND SUBSTR(CP27,1,6)=" & Val(Text1) + 191100 & "  and cp57 is  null GROUP BY EP13 "
    '本月發文計件與點數
    'Modify By Cheng 2004/03/30
    '抓墨圖計件的資料
'    strSQL = "INSERT INTO R090711_1 (R111001,R111002,r111003,R111004,r111005,ID) select EP13,9,Sum(Decode(EP29, Null, 1, 0)), 0, sum(cp18),'" & strUserNum & "' from engineerprogress,caseprogress where EP02=CP09(+) AND EP13='" & Trim(Left(Combo1.Text, 6)) & "' AND  CP01 IN ('P','CFP','FCP')  AND SUBSTR(CP27,1,6)=" & Val(Text1) + 191100 & "  and cp57 is  null GROUP BY EP13 "
    'edit by nick 2005/01/04
    'StrSql = "INSERT INTO R090711_1 (R111001,R111002,r111003,R111004,r111005,ID) select EP13,9,Sum(Decode(EP29, Null, 1, 0)), 0, sum(cp18),'" & strUserNum & "' from engineerprogress,caseprogress where EP02=CP09(+) AND EP13='" & Trim(Left(Combo1.Text, 6)) & "' AND  CP01 IN ('P','CFP','FCP')  AND SUBSTR(CP27,1,6)=" & Val(Text1) + 191100 & "  and cp57 is  null And EP29 Is Null GROUP BY EP13 "
    'edit by nickc 2005/03/01
    'StrSql = "INSERT INTO R090711_1 (R111001,R111002,r111003,R111004,r111005,ID) select EP13,9,Sum(Decode(EP29, Null, 1, 0)), 0, sum(cp18),'" & strUserNum & "' from engineerprogress,caseprogress where EP02=CP09(+) AND EP13='" & Trim(Left(Combo1.Text, 6)) & "' AND  CP01 IN ('P','CFP','FCP')  AND SUBSTR(CP27,1,6)=" & Val(Text1) + 191100 & "  and cp57 is  null And EP29 Is Null and ((cp21='Y' and ep20 is null) or cp21 is null) GROUP BY EP13 "
    'edit by nickc 2005/04/13
    'strSQL = "INSERT INTO R090711_1 (R111001,R111002,r111003,R111004,r111005,ID) select EP13,9,Sum(Decode(EP29, Null, 1, 0)), 0, sum(cp18),'" & strUserNum & "' from engineerprogress,caseprogress where EP02=CP09(+) AND EP13='" & Trim(Left(Combo1.Text, 6)) & "' AND  CP01 IN ('P','CFP','FCP')  AND SUBSTR(CP27,1,6)=" & Val(Text1) + 191100 & "  and cp57 is  null And EP29 Is Null and ((cp21='Y' and (ep20 is null or ep29 is null)) or cp21 is null) and cp107='Y' GROUP BY EP13 "
    'edit by nickc 2005/05/04
    'strSQL = "INSERT INTO R090711_1 (R111001,R111002,r111003,R111004,r111005,ID) select EP13,9,Sum(Decode(EP29, Null, 1, 0)), 0, sum(cp18),'" & strUserNum & "' from engineerprogress,caseprogress where EP02=CP09(+) " & IIf(From090706 = True, "", " AND EP13='" & Trim(Left(Combo1.Text, 6)) & "' ") & " AND  CP01 IN ('P','CFP','FCP')  AND SUBSTR(CP27,1,6)=" & Val(Text1) + 191100 & "  and cp57 is  null And EP29 Is Null and ((cp21='Y' and (ep20 is null or ep29 is null)) or cp21 is null) and cp107='Y' GROUP BY EP13 "
    strSql = "INSERT INTO R090711_1 (R111001,R111002,r111003,R111004,r111005,ID,r111006,r111007,r111008) select EP13,9,Sum(Decode(EP29, Null, 1, 0)), 0, sum(cp18),'" & strUserNum & "',Sum(nvl(cp103,0) * nvl(cp104,0)), 0, sum(nvl(cp18,0)-nvl(a1u07/1000,0)) from engineerprogress,caseprogress,(select a1u03,sum(nvl(a1u07,0)) as a1u07 from acc1u0 group by a1u03) ABCDE where CP09=ep02(+) and CP09=a1u03(+) " & IIf(From090706 = True, "", " AND EP13='" & Trim(Left(Combo1.Text, 6)) & "' ") & " AND  CP01 IN ('P','CFP','FCP')  AND CP27>=" & Val(Text1) + 191100 & "01 AND CP27<=" & Val(Text1) + 191100 & "31  and cp57 is  null And EP29 Is Null and ((cp21='Y' and (ep20 is null or ep29 is null)) or cp21 is null) and cp107='Y' GROUP BY EP13 "
    'End
    cnnConnection.Execute strSql
    'Add By Cheng 2004/03/30
    '將支援記錄加入本月發文件數
    'edit by nickc 2005/04/13
    'strSQL = "INSERT INTO R090711_1 (R111001,R111002,r111003,R111004,r111005,ID) select SH02,9,Sum(Decode(SH06, 'CFP', Nvl(SH05, 0)/8, Nvl(SH05, 0)/4)), 0, 0,'" & strUserNum & "' from SupportHour where SH02='" & Trim(Left(Combo1.Text, 6)) & "' AND  SH06 IN ('P','CFP','FCP') And SUBSTR(SH01,1,6)=" & Val(Text1) + 191100 & " And SH11 ='V' GROUP BY SH02 "
    'edit by nickc 2005/05/04
    'strSQL = "INSERT INTO R090711_1 (R111001,R111002,r111003,R111004,r111005,ID) select SH02,9,Sum(Decode(SH06, 'CFP', Nvl(SH05, 0)/8, Nvl(SH05, 0)/4)), 0, 0,'" & strUserNum & "' from SupportHour where  SH06 IN ('P','CFP','FCP') " & IIf(From090706 = True, "", " AND sh02='" & Trim(Left(Combo1.Text, 6)) & "' ") & " And SUBSTR(SH01,1,6)=" & Val(Text1) + 191100 & " And SH11 ='V' GROUP BY SH02 "
    'Modified by Morgan 2012/1/2 不限制系統別且比例有改(與每週速度考核一致)
    'strSql = "INSERT INTO R090711_1 (R111001,R111002,r111003,R111004,r111005,ID,r111006,r111007,r111008) select SH02,9,Sum(Decode(SH06, 'CFP', Nvl(SH05, 0)/4, Nvl(SH05, 0)/4)), 0, 0,'" & strUserNum & "',Sum(Decode(SH06, 'CFP', Nvl(SH05, 0)/4, Nvl(SH05, 0)/4)), 0, 0 from SupportHour where  SH06 IN ('P','CFP','FCP') " & IIf(From090706 = True, "", " AND sh02='" & Trim(Left(Combo1.Text, 6)) & "' ") & " And SUBSTR(SH01,1,6)=" & Val(Text1) + 191100 & " And SH11 ='V' GROUP BY SH02 "
    'Modified by Morgan 2014/3/20 --2014/4/1起支援改每小時折計0.2基數
    'strSql = "INSERT INTO R090711_1 (R111001,R111002,r111003,R111004,r111005,ID,r111006,r111007,r111008) select SH02,9,Sum(Decode(SH06, 'CFP', Nvl(SH05, 0)/4, Nvl(SH05, 0)/4)), 0, 0,'" & strUserNum & "',Sum(Decode(SH06, 'CFP', Nvl(SH05, 0)/3, Nvl(SH05, 0)/4)), 0, 0 from SupportHour where SH11 ='V' " & IIf(From090706 = True, "", " AND sh02='" & Trim(Left(Combo1.Text, 6)) & "' ") & " And SUBSTR(SH01,1,6)=" & Val(Text1) + 191100 & " GROUP BY SH02 "
    strSql = "INSERT INTO R090711_1 (R111001,R111002,r111003,R111004,r111005,ID,r111006,r111007,r111008) select SH02,9,Sum(Nvl(SH05, 0)/4), 0, 0,'" & strUserNum & "',Sum(" & Sh2EPtCode & "), 0, 0 from SupportHour where SH11 ='V' " & IIf(From090706 = True, "", " AND sh02='" & Trim(Left(Combo1.Text, 6)) & "' ") & " And (SH01>=" & Val(Text1) + 191100 & "01 and SH01<=" & Val(Text1) + 191100 & "31) GROUP BY SH02 "
    'end 2014/3/20
    cnnConnection.Execute strSql
    'End
    '本月草圖張數/2(無墨完日或墨完日不在當月)
    'edit by nick 2005/01/04
    'StrSql = "INSERT INTO R090711_1 (R111001,R111002,r111003,R111004,r111005,ID) select EP13,9, 0, Sum(Nvl(EP16,0)/2), 0,'" & strUserNum & "' from engineerprogress,caseprogress where EP02=CP09(+) AND EP13='" & Trim(Left(Combo1.Text, 6)) & "' AND  CP01 IN ('P','CFP','FCP')  AND SUBSTR(EP15,1,6)=" & Val(Text1) + 191100 & " And (EP18 Is Null Or substr(EP18,1,6)<>" & Val(Me.Text1.Text) + 191100 & ") GROUP BY EP13 "
    'edit by nickc 2005/03/01 墨圖也要判斷
    'StrSql = "INSERT INTO R090711_1 (R111001,R111002,r111003,R111004,r111005,ID) select EP13,9, 0, Sum(Nvl(EP16,0)/2), 0,'" & strUserNum & "' from engineerprogress,caseprogress where EP02=CP09(+) AND EP13='" & Trim(Left(Combo1.Text, 6)) & "' AND  CP01 IN ('P','CFP','FCP')  AND SUBSTR(EP15,1,6)=" & Val(Text1) + 191100 & " And (EP18 Is Null Or substr(EP18,1,6)<>" & Val(Me.Text1.Text) + 191100 & ") and ((cp21='Y' and ep20 is null) or cp21 is null)  GROUP BY EP13 "
    'edit by nickc 2005/04/13
    'strSQL = "INSERT INTO R090711_1 (R111001,R111002,r111003,R111004,r111005,ID) select EP13,9, 0, Sum(Nvl(EP16,0)/2), 0,'" & strUserNum & "' from engineerprogress,caseprogress where EP02=CP09(+) AND EP13='" & Trim(Left(Combo1.Text, 6)) & "' AND  CP01 IN ('P','CFP','FCP')  AND SUBSTR(EP15,1,6)=" & Val(Text1) + 191100 & " And (EP18 Is Null Or substr(EP18,1,6)<>" & Val(Me.Text1.Text) + 191100 & ") and ((cp21='Y' and (ep20 is null or ep29 is null)) or cp21 is null) and cp107='Y'  GROUP BY EP13 "
    'edit by nickc 2005/05/04
    'strSQL = "INSERT INTO R090711_1 (R111001,R111002,r111003,R111004,r111005,ID) select EP13,9, 0, Sum(Nvl(EP16,0)/2), 0,'" & strUserNum & "' from engineerprogress,caseprogress where EP02=CP09(+) " & IIf(From090706 = True, "", " AND EP13='" & Trim(Left(Combo1.Text, 6)) & "' ") & " AND  CP01 IN ('P','CFP','FCP')  AND SUBSTR(EP15,1,6)=" & Val(Text1) + 191100 & " And (EP18 Is Null Or substr(EP18,1,6)<>" & Val(Me.Text1.Text) + 191100 & ") and ((cp21='Y' and (ep20 is null or ep29 is null)) or cp21 is null) and cp107='Y'  GROUP BY EP13 "
    strSql = "INSERT INTO R090711_1 (R111001,R111002,r111003,R111004,r111005,ID,r111006,r111007,r111008) select EP13,9, 0, Sum(Nvl(EP16,0)/2), 0,'" & strUserNum & "',0, Sum(Nvl(EP16,0)/2), 0 from engineerprogress,caseprogress where EP02=CP09(+) " & IIf(From090706 = True, "", " AND EP13='" & Trim(Left(Combo1.Text, 6)) & "' ") & " AND  CP01 IN ('P','CFP','FCP')  AND SUBSTR(EP15,1,6)=" & Val(Text1) + 191100 & " And (EP18 Is Null Or substr(EP18,1,6)<>" & Val(Me.Text1.Text) + 191100 & ") and ((cp21='Y' and (ep20 is null or ep29 is null)) or cp21 is null) and cp107='Y'  GROUP BY EP13 "
    cnnConnection.Execute strSql
    '本月墨圖張數/2(無草完日或草完日不在當月)
    'edit by nick 2005/01/04
    'StrSql = "INSERT INTO R090711_1 (R111001,R111002,r111003,R111004,r111005,ID) select EP13,9, 0, Sum(Nvl(EP19,0)/2), 0,'" & strUserNum & "' from engineerprogress,caseprogress where EP02=CP09(+) AND EP13='" & Trim(Left(Combo1.Text, 6)) & "' AND  CP01 IN ('P','CFP','FCP')  AND SUBSTR(EP18,1,6)=" & Val(Text1) + 191100 & " AND (EP15 Is Null Or substr(EP15,1,6)<>" & Val(Me.Text1.Text) + 191100 & ")  GROUP BY EP13 "
    'edit by nickc 2005/03/01 墨圖也要判斷
    'StrSql = "INSERT INTO R090711_1 (R111001,R111002,r111003,R111004,r111005,ID) select EP13,9, 0, Sum(Nvl(EP19,0)/2), 0,'" & strUserNum & "' from engineerprogress,caseprogress where EP02=CP09(+) AND EP13='" & Trim(Left(Combo1.Text, 6)) & "' AND  CP01 IN ('P','CFP','FCP')  AND SUBSTR(EP18,1,6)=" & Val(Text1) + 191100 & " AND (EP15 Is Null Or substr(EP15,1,6)<>" & Val(Me.Text1.Text) + 191100 & ") and ((cp21='Y' and ep20 is null) or cp21 is null) GROUP BY EP13 "
    'edit by nickc 2005/04/13
    'strSQL = "INSERT INTO R090711_1 (R111001,R111002,r111003,R111004,r111005,ID) select EP13,9, 0, Sum(Nvl(EP19,0)/2), 0,'" & strUserNum & "' from engineerprogress,caseprogress where EP02=CP09(+) AND EP13='" & Trim(Left(Combo1.Text, 6)) & "' AND  CP01 IN ('P','CFP','FCP')  AND SUBSTR(EP18,1,6)=" & Val(Text1) + 191100 & " AND (EP15 Is Null Or substr(EP15,1,6)<>" & Val(Me.Text1.Text) + 191100 & ") and ((cp21='Y' and (ep20 is null or ep29 is null) ) or cp21 is null) GROUP BY EP13 "
    'edit by nickc 2005/05/04
    'strSQL = "INSERT INTO R090711_1 (R111001,R111002,r111003,R111004,r111005,ID) select EP13,9, 0, Sum(Nvl(EP19,0)/2), 0,'" & strUserNum & "' from engineerprogress,caseprogress where EP02=CP09(+) " & IIf(From090706 = True, "", " AND EP13='" & Trim(Left(Combo1.Text, 6)) & "' ") & " AND  CP01 IN ('P','CFP','FCP')  AND SUBSTR(EP18,1,6)=" & Val(Text1) + 191100 & " AND (EP15 Is Null Or substr(EP15,1,6)<>" & Val(Me.Text1.Text) + 191100 & ") and ((cp21='Y' and (ep20 is null or ep29 is null) ) or cp21 is null) GROUP BY EP13 "
    strSql = "INSERT INTO R090711_1 (R111001,R111002,r111003,R111004,r111005,ID,r111006,r111007,r111008) select EP13,9, 0, Sum(Nvl(EP19,0)/2), 0,'" & strUserNum & "',0, Sum(Nvl(EP19,0)/2), 0 from engineerprogress,caseprogress where EP02=CP09(+) " & IIf(From090706 = True, "", " AND EP13='" & Trim(Left(Combo1.Text, 6)) & "' ") & " AND  CP01 IN ('P','CFP','FCP')  AND SUBSTR(EP18,1,6)=" & Val(Text1) + 191100 & " AND (EP15 Is Null Or substr(EP15,1,6)<>" & Val(Me.Text1.Text) + 191100 & ") and ((cp21='Y' and (ep20 is null or ep29 is null) ) or cp21 is null) GROUP BY EP13 "
    cnnConnection.Execute strSql
    '本月(草圖+墨圖張數)/2(草完日及墨完日皆在當月)
    'edit by nick 2005/01/04
    'StrSql = "INSERT INTO R090711_1 (R111001,R111002,r111003,R111004,r111005,ID) select EP13,9, 0, Sum((Nvl(EP16,0)+Nvl(EP19,0))/2), 0,'" & strUserNum & "' from engineerprogress,caseprogress where EP02=CP09(+) AND EP13='" & Trim(Left(Combo1.Text, 6)) & "' AND  CP01 IN ('P','CFP','FCP') AND SUBSTR(EP15,1,6)=" & Val(Text1) + 191100 & " AND SUBSTR(EP18,1,6)=" & Val(Text1) + 191100 & " GROUP BY EP13 "
    'edit by nickc 2005/03/01  墨圖也要判斷
    'StrSql = "INSERT INTO R090711_1 (R111001,R111002,r111003,R111004,r111005,ID) select EP13,9, 0, Sum((Nvl(EP16,0)+Nvl(EP19,0))/2), 0,'" & strUserNum & "' from engineerprogress,caseprogress where EP02=CP09(+) AND EP13='" & Trim(Left(Combo1.Text, 6)) & "' AND  CP01 IN ('P','CFP','FCP') AND SUBSTR(EP15,1,6)=" & Val(Text1) + 191100 & " AND SUBSTR(EP18,1,6)=" & Val(Text1) + 191100 & " and ((cp21='Y' and ep20 is null) or cp21 is null) GROUP BY EP13 "
    'strSQL = "INSERT INTO R090711_1 (R111001,R111002,r111003,R111004,r111005,ID) select EP13,9, 0, Sum((Nvl(EP16,0)+Nvl(EP19,0))/2), 0,'" & strUserNum & "' from engineerprogress,caseprogress where EP02=CP09(+) AND EP13='" & Trim(Left(Combo1.Text, 6)) & "' AND  CP01 IN ('P','CFP','FCP') AND SUBSTR(EP15,1,6)=" & Val(Text1) + 191100 & " AND SUBSTR(EP18,1,6)=" & Val(Text1) + 191100 & " and ((cp21='Y' and (ep20 is null or ep29 is null)) or cp21 is null) and cp107='Y' GROUP BY EP13 "
    'edit by nickc 2005/05/04
    'strSQL = "INSERT INTO R090711_1 (R111001,R111002,r111003,R111004,r111005,ID) select EP13,9, 0, Sum((Nvl(EP16,0)+Nvl(EP19,0))/2), 0,'" & strUserNum & "' from engineerprogress,caseprogress where EP02=CP09(+) " & IIf(From090706 = True, "", " AND EP13='" & Trim(Left(Combo1.Text, 6)) & "' ") & " AND  CP01 IN ('P','CFP','FCP') AND SUBSTR(EP15,1,6)=" & Val(Text1) + 191100 & " AND SUBSTR(EP18,1,6)=" & Val(Text1) + 191100 & " and ((cp21='Y' and (ep20 is null or ep29 is null)) or cp21 is null) and cp107='Y' GROUP BY EP13 "
    strSql = "INSERT INTO R090711_1 (R111001,R111002,r111003,R111004,r111005,ID,r111006,r111007,r111008) select EP13,9, 0, Sum((Nvl(EP16,0)+Nvl(EP19,0))/2), 0,'" & strUserNum & "',0, Sum((Nvl(EP16,0)+Nvl(EP19,0))/2), 0 from engineerprogress,caseprogress where EP02=CP09(+) " & IIf(From090706 = True, "", " AND EP13='" & Trim(Left(Combo1.Text, 6)) & "' ") & " AND  CP01 IN ('P','CFP','FCP') AND SUBSTR(EP15,1,6)=" & Val(Text1) + 191100 & " AND SUBSTR(EP18,1,6)=" & Val(Text1) + 191100 & " and ((cp21='Y' and (ep20 is null or ep29 is null)) or cp21 is null) and cp107='Y' GROUP BY EP13 "
    cnnConnection.Execute strSql
    
    'add by nickc 2005/04/13 增加提供圖檔及關聯
    strSql = "INSERT INTO R090711_1 (R111001,R111002,r111003,R111004,r111005,ID) select EP13,10,count(*), 0, 0,'" & strUserNum & "' from engineerprogress,caseprogress where EP02=CP09(+) " & IIf(From090706 = True, "", " AND EP13='" & Trim(Left(Combo1.Text, 6)) & "' ") & " AND SUBSTR(ep15,1,6)=" & Val(Text1) + 191100 & "   and ep20 is null and ((cp21='Y' and (ep20 is null or ep29 is null)) or cp21 is null) and cp107='Y' and cp106='Y'  GROUP BY EP13 "
    cnnConnection.Execute strSql
    strSql = "INSERT INTO R090711_1 (R111001,R111002,r111003,R111004,r111005,ID) select EP13,11,count(*), 0, 0,'" & strUserNum & "' from engineerprogress,caseprogress where EP02=CP09(+) and ep20='N' and cp103=0.4 and cp100=0 " & IIf(From090706 = True, "", " AND EP13='" & Trim(Left(Combo1.Text, 6)) & "' ") & " AND SUBSTR(ep18,1,6)=" & Val(Text1) + 191100 & "   And EP29 Is Null and ((cp21='Y' and (ep20 is null or ep29 is null)) or cp21 is null) and cp107='Y' GROUP BY EP13 "
    cnnConnection.Execute strSql
    
    'End
    DoEvents
    Set NickRS = New ADODB.Recordset
    Select Case ProState
    Case "1"
        strBeginDate = Left(ChangeWDateStringToWString(DateAdd("m", -1, ChangeWStringToWDateString(strSrvDate(1)))), 6) & "01"
        strEndDate = strSrvDate(1)
          StrSQL6 = "  AND EP13='" & Trim(Left(Combo1.Text, 6)) & "' and cp01 in ('FCP','P','CFP')  "
'          StrSQL6 = StrSQL6 + " and CP27 IS NULL  and CP57 IS NULL and cp05>=19980101 "
          StrSQL6 = StrSQL6 + " And cp05>=19980101 "
          strSQL1 = " AND EP13='" & Trim(Left(Combo1.Text, 6)) & "' and cp01 in ('FCP','P','CFP')  "
'          strSQL1 = strSQL1 & " and (SUBSTR(CP27,1,6)=" & Mid(strSrvDate(1), 1, 6) & " or SUBSTR(CP57,1,6)=" & Mid(strSrvDate(1), 1, 6) & ") and cp05>=19980101 "
          strSQL1 = strSQL1 & " And cp05>=19980101 "
    Case "4" '繪圖人員個人工作進度資料查詢
        strBeginDate = Left(ChangeWDateStringToWString(DateAdd("m", -1, ChangeWStringToWDateString(strSrvDate(1)))), 6) & "01"
        strEndDate = strSrvDate(1)
          StrSQL6 = " AND EP13='" & Trim(Left(Combo1.Text, 6)) & "' and cp01 in ('FCP','P','CFP')  "
          'edit by nickc 2006/01/11
          'StrSQL6 = StrSQL6 & " AND SUBSTR(CP27,1,6)=" & Val(frm090303_1.Text1.Text) + 191100 & " AND cp05>=19980101 "
          'Modify By Sindy 2016/5/10
          'StrSQL6 = StrSQL6 & " AND (SUBSTR(CP27,1,6)=" & Val(frm090303_1.Text1.Text) + 191100 & " or SUBSTR(CP05,1,6)=" & Val(frm090303_1.Text1.Text) + 191100 & ") AND cp05>=19980101 "
          StrSQL6 = StrSQL6 & " AND ((CP27>=" & Val(frm090303_1.Text1.Text) + 191100 & "01 and CP27<=" & Val(frm090303_1.Text1.Text) + 191100 & "31) or (CP05>=" & Val(frm090303_1.Text1.Text) + 191100 & "01 and CP05<=" & Val(frm090303_1.Text1.Text) + 191100 & "31)) AND cp05>=19980101 "
          '2016/5/10 END
          strSQL1 = " AND EP13='" & Trim(Left(Combo1.Text, 6)) & "' and cp01 in ('FCP','P','CFP')  "
          'edit by nickc 2006/01/11
          'strSQL1 = strSQL1 & " AND SUBSTR(CP27,1,6)=" & Val(frm090303_1.Text1.Text) + 191100 & " AND cp05>=19980101 "
          'Modify By Sindy 2016/5/10
          'strSQL1 = strSQL1 & " AND (SUBSTR(CP27,1,6)=" & Val(frm090303_1.Text1.Text) + 191100 & " or SUBSTR(CP05,1,6)=" & Val(frm090303_1.Text1.Text) + 191100 & ") AND cp05>=19980101 "
          strSQL1 = strSQL1 & " AND ((CP27>=" & Val(frm090303_1.Text1.Text) + 191100 & "01 and CP27<=" & Val(frm090303_1.Text1.Text) + 191100 & "31) or (CP05>=" & Val(frm090303_1.Text1.Text) + 191100 & "01 and CP05<=" & Val(frm090303_1.Text1.Text) + 191100 & "31)) AND cp05>=19980101 "
          '2016/5/10 END
    Case "2" '繪圖人員管理查詢作業
        strBeginDate = IIf(Val(Me.Text1.Text) + 191100 < Val(Left(strSrvDate(1), 6)), _
                            Left(ChangeWDateStringToWString(DateAdd("m", -1, ChangeWStringToWDateString((Val(Me.Text1.Text) + 191100) & "01"))), 6) & "01", _
                            Left(ChangeWDateStringToWString(DateAdd("m", -1, ChangeWStringToWDateString(strSrvDate(1)))), 6) & "01")
        strEndDate = IIf(Val(Me.Text1.Text) + 191100 < Val(Left(strSrvDate(1), 6)), _
                            Left((Val(Me.Text1.Text) + 191100), 6) & PUB_GetMonthDays(Left((Val(Me.Text1.Text) + 191100), 4), Mid((Val(Me.Text1.Text) + 191100), 5, 2)), _
                            strSrvDate(1))
          strSQL1 = " AND CP05<=" & Val(Me.Text1.Text) + 191100 & "31 "
          StrSQL6 = " AND CP05<=" & Val(Me.Text1.Text) + 191100 & "31 "
          StrGrp090711 = ""
          strSQL1 = " AND EP13='" & Trim(Left(Combo1.Text, 6)) & "' AND CP01 IN (" & SQLGrpStr(StrGrp090711, 1) & ") " & strSQL1
'          strSQL1 = strSQL1 & " and ((SUBSTR(CP27,1,6)=" & (Val(Text1.Text) + 191100) & " AND CP57 IS NULL ) OR ( CP27 IS NULL AND SUBSTR(CP57,1,6)=" & (Val(Text1.Text) + 191100) & ")) and cp05>=19980101 "
          strSQL1 = strSQL1 & " And cp05>=19980101 "
          StrSQL6 = " AND EP13='" & Trim(Left(Combo1.Text, 6)) & "' AND CP01 IN (" & SQLGrpStr(StrGrp090711, 1) & ") " & StrSQL6
'          StrSQL6 = StrSQL6 & " and CP27 IS NULL  and CP57 IS NULL and cp05>=19980101 "
          StrSQL6 = StrSQL6 & " And cp05>=19980101 "
    Case "3"
        strBeginDate = IIf(Val(Me.Text1.Text) + 191100 < Val(Left(strSrvDate(1), 6)), _
                            Left(ChangeWDateStringToWString(DateAdd("m", -1, ChangeWStringToWDateString((Val(Me.Text1.Text) + 191100) & "01"))), 6) & "01", _
                            Left(ChangeWDateStringToWString(DateAdd("m", -1, ChangeWStringToWDateString(strSrvDate(1)))), 6) & "01")
        strEndDate = IIf(Val(Me.Text1.Text) + 191100 < Val(Left(strSrvDate(1), 6)), _
                            Left((Val(Me.Text1.Text) + 191100), 6) & PUB_GetMonthDays(Left((Val(Me.Text1.Text) + 191100), 4), Mid((Val(Me.Text1.Text) + 191100), 5, 2)), _
                            strSrvDate(1))
          strSQL1 = ""
          StrSQL6 = ""
          StrGrp090711 = ""
          strSQL1 = " AND CP05<=" & Val(Me.Text1.Text) + 191100 & "31 "
          StrSQL6 = " AND CP05<=" & Val(Me.Text1.Text) + 191100 & "31 "
          strSQL1 = strSQL1 & " AND EP13='" & Trim(Left(Combo1.Text, 6)) & "' AND CP01 IN (" & SQLGrpStr(StrGrp090711, 1) & ") " & StrSQL6
'          strSQL1 = strSQL1 & " and (SUBSTR(CP27,1,6)=" & (Val(Text1.Text) + 191100) & " or SUBSTR(CP57,1,6)=" & (Val(Text1.Text) + 191100) & ") and cp05>=19980101 "
          strSQL1 = strSQL1 & " And cp05>=19980101 "
          StrSQL6 = StrSQL6 & " AND EP13='" & Trim(Left(Combo1.Text, 6)) & "' AND CP01 IN (" & SQLGrpStr(StrGrp090711, 1) & ") " & StrSQL6
'          StrSQL6 = StrSQL6 & " and CP27 IS NULL  and CP57 IS NULL and cp05>=19980101 "
          StrSQL6 = StrSQL6 & " And cp05>=19980101 "
    Case Else
    End Select
      'add by nick 2004/12/20   加多國案且草圖不計件不秀
'edit by nick 2005/03/01 墨圖也要判斷
'      strSQL1 = strSQL1 & " and ((cp21='Y' and ep20 is null) or cp21 is null) "
'      StrSQL6 = StrSQL6 & " and ((cp21='Y' and ep20 is null) or cp21 is null) "
      strSQL1 = strSQL1 & " and ((cp21='Y' and (ep20 is null or ep29 is null)) or cp21 is null) "
      StrSQL6 = StrSQL6 & " and ((cp21='Y' and (ep20 is null or ep29 is null)) or cp21 is null) "
      'add by nickc 2005/03/21
      strSQL1 = strSQL1 & " and  cp107='Y' "
      StrSQL6 = StrSQL6 & " and cp107='Y' "
    'Modify By Cheng 2003/06/30
    '草圖或墨圖逾期
'    strSQL = "SELECT EP13,PA08,EP14,'1' FROM ENGINEERPROGRESS,CASEPROGRESS,PATENT WHERE CP01 IN ('P','CFP') AND EP02=CP09(+) AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND EP13='" & Trim(Left(Combo1.Text, 6)) & "' AND CP57 IS NULL AND CP27 is null and ep14 is not null AND EP15 IS NULL "
'    strSQL = strSQL & " UNION all  SELECT EP13,PA08,EP17,'2' FROM ENGINEERPROGRESS,CASEPROGRESS,PATENT WHERE CP01 IN ('P','CFP') AND EP02=CP09(+) AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND EP13='" & Trim(Left(Combo1.Text, 6)) & "' AND CP57 IS NULL AND CP27 is null and ep17 is not null and ep18 IS NULL "
'    strSQL = "SELECT EP13, CP10, EP14, '1', EP15, CP07, CP27, CP09, CP57 FROM ENGINEERPROGRESS,CASEPROGRESS,PATENT WHERE CP01 IN ('P','CFP','FCP') AND EP02=CP09(+) AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND EP13='" & Trim(Left(Combo1.Text, 6)) & "' AND CP57 IS NULL and ep14 is not null And EP20 Is Null " & StrSQL6
'    strSQL = strSQL & " Union SELECT EP13, CP10, Nvl(EP17, EP08), '2', EP18, CP07, CP27, CP09, CP57 FROM ENGINEERPROGRESS,CASEPROGRESS,PATENT WHERE CP01 IN ('P','CFP','FCP') AND EP02=CP09(+) AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND EP13='" & Trim(Left(Combo1.Text, 6)) & "' AND CP57 IS NULL and (EP17 is not null Or EP08 Is Not Null) And EP29 Is Null " & StrSQL6
    
    'Modify by Morgan 2004/5/19
    '加專利種類
'    strSQL = "SELECT EP13, CP10, EP14, '1', EP15, CP07, CP27, CP09, CP57 FROM ENGINEERPROGRESS,CASEPROGRESS,PATENT WHERE CP01 IN ('P','CFP','FCP') AND EP02=CP09(+) AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND EP13='" & Trim(Left(Combo1.Text, 6)) & "' And EP20 Is Null And (EP14>=" & strBeginDate & " And EP14<=" & strEndDate & " ) " & StrSQL6
'    strSQL = strSQL & " Union SELECT EP13, CP10, Nvl(EP17, EP08), '2', EP18, CP07, CP27, CP09, CP57 FROM ENGINEERPROGRESS,CASEPROGRESS,PATENT WHERE CP01 IN ('P','CFP','FCP') AND EP02=CP09(+) AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND EP13='" & Trim(Left(Combo1.Text, 6)) & "' And EP29 Is Null And (EP17>=" & strBeginDate & " And EP17<=" & strEndDate & " ) " & StrSQL6
'edit by nickc 2005/04/13
'    strSQL = "SELECT EP13, CP10, EP14, '1', EP15, CP07, CP27, CP09, CP57, PA08 FROM ENGINEERPROGRESS,CASEPROGRESS,PATENT WHERE CP01 IN ('P','CFP','FCP') AND EP02=CP09(+) AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND EP13='" & Trim(Left(Combo1.Text, 6)) & "' And EP20 Is Null And (EP14>=" & strBeginDate & " And EP14<=" & strEndDate & " ) " & StrSQL6
'    strSQL = strSQL & " Union SELECT EP13, CP10, Nvl(EP17, EP08), '2', EP18, CP07, CP27, CP09, CP57, PA08 FROM ENGINEERPROGRESS,CASEPROGRESS,PATENT WHERE CP01 IN ('P','CFP','FCP') AND EP02=CP09(+) AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND EP13='" & Trim(Left(Combo1.Text, 6)) & "' And EP29 Is Null And (EP17>=" & strBeginDate & " And EP17<=" & strEndDate & " ) " & StrSQL6
    strSql = "SELECT EP13, CP10, EP14, '1', EP15, CP07, CP27, CP09, CP57, PA08 FROM ENGINEERPROGRESS,CASEPROGRESS,PATENT WHERE CP01 IN ('P','CFP','FCP') AND EP02=CP09(+) AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) " & IIf(From090706 = True, "", " AND EP13='" & Trim(Left(Combo1.Text, 6)) & "' ") & " And EP20 Is Null And (EP14>=" & strBeginDate & " And EP14<=" & strEndDate & " ) " & StrSQL6
    strSql = strSql & " Union SELECT EP13, CP10, Nvl(EP17, EP08), '2', EP18, CP07, CP27, CP09, CP57, PA08 FROM ENGINEERPROGRESS,CASEPROGRESS,PATENT WHERE CP01 IN ('P','CFP','FCP') AND EP02=CP09(+) AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) " & IIf(From090706 = True, "", " AND EP13='" & Trim(Left(Combo1.Text, 6)) & "' ") & " And EP29 Is Null And (EP17>=" & strBeginDate & " And EP17<=" & strEndDate & " ) " & StrSQL6
    
'    If ProState <> 4 Then
'        strSQL = strSQL & " Union SELECT EP13, CP10, EP14, '1', EP15, CP07, CP27, CP09, CP57 FROM ENGINEERPROGRESS,CASEPROGRESS,PATENT WHERE CP01 IN ('P','CFP','FCP') AND EP02=CP09(+) AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND EP13='" & Trim(Left(Combo1.Text, 6)) & "' AND CP57 IS NULL and ep14 is not null And EP20 Is Null " & strSQL1
'        strSQL = strSQL & " Union SELECT EP13, CP10, Nvl(EP17, EP08), '2', EP18, CP07, CP27, CP09, CP57 FROM ENGINEERPROGRESS,CASEPROGRESS,PATENT WHERE CP01 IN ('P','CFP','FCP') AND EP02=CP09(+) AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND EP13='" & Trim(Left(Combo1.Text, 6)) & "' AND CP57 IS NULL and (EP17 is not null Or EP08 Is Not Null) And EP29 Is Null " & strSQL1
'    End If
    NickRS.CursorLocation = adUseClient
    NickRS.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If NickRS.RecordCount <> 0 Then
        NickRS.MoveFirst
        Do While NickRS.EOF = False
            'Modify by Morgan 2004/5/19
            '改依專利種類判斷
'            Select Case CheckStr(NickRS.Fields(1))
'            Case "103", "105" '設計申請
            Select Case CheckStr(NickRS.Fields("PA08").Value)
               Case "3"
               
                If CheckStr(NickRS.Fields(3)) = "1" Then  '草圖
                    '若有草齊日及草完日
                    If "" & NickRS.Fields(2).Value <> "" And "" & NickRS.Fields(4).Value <> "" Then
                        'Modify By Cheng 2004/02/16
                        '草完日必須為當月
'                        If Left("" & NickRS.Fields(4).Value, 6) = Left(strSrvDate(1), 6) Then
                        If Left("" & NickRS.Fields(4).Value, 6) = IIf(ProState = "2" Or ProState = "3", "" & (Val(Me.Text1.Text) + 191100), Left(strSrvDate(1), 6)) Then
                        'End
                            If GetWorkDay(CheckStr(NickRS.Fields(4)), "" & NickRS.Fields(2).Value) > 5 Then
                                cnnConnection.Execute "INSERT INTO R090711_1 (R111001,R111002,R111003,ID) VALUES ('" & CheckStr(NickRS.Fields(0)) & "',7,1,'" & strUserNum & "') "
                            End If
                        End If
                        'End
'                    '若無發文日有草齊日無草完日無取消收文日
'                    ElseIf "" & NickRS.Fields(6).Value = "" And "" & NickRS.Fields(2).Value <> "" And "" & NickRS.Fields(4).Value = "" And "" & NickRS.Fields(8).Value = "" Then
                    '若有草齊日無草完日
'edit by nickc 2005/04/18 瓊玉說無完稿日不算逾時
'                    ElseIf "" & NickRS.Fields(2).Value <> "" And "" & NickRS.Fields(4).Value = "" Then
''                        If GetWorkDay(strSrvDate(1), CheckStr(NickRS.Fields(2))) > 5 Then
'                        If GetWorkDay(strEndDate, CheckStr(NickRS.Fields(2))) > 5 Then
'                           cnnConnection.Execute "INSERT INTO R090711_1 (R111001,R111002,R111003,ID) VALUES ('" & CheckStr(NickRS.Fields(0)) & "',7,1,'" & strUserNum & "') "
'                        End If
                    End If
                Else '墨圖
                    '若有墨齊日及墨完日
                    If "" & NickRS.Fields(2).Value <> "" And "" & NickRS.Fields(4).Value <> "" Then
                        'Modify By Cheng 2004/02/16
                        '墨完日必須為當月
                        'Modify By Cheng 2004/03/12
'                        If Left("" & NickRS.Fields(4).Value, 6) = Left(strSrvDate(1), 6) Then
                        If Left("" & NickRS.Fields(4).Value, 6) = IIf(ProState = "2" Or ProState = "3", "" & (Val(Me.Text1.Text) + 191100), Left(strSrvDate(1), 6)) Then
                        'End
                            If GetWorkDay(CheckStr(NickRS.Fields(4)), "" & NickRS.Fields(2).Value) > 3 Then
                                cnnConnection.Execute "INSERT INTO R090711_1 (R111001,R111002,R111003,ID) VALUES ('" & CheckStr(NickRS.Fields(0)) & "',8,1,'" & strUserNum & "') "
                            End If
                        End If
                        'End
'                    '若無發文日有墨齊日無墨完日無取消收文日
'                    ElseIf "" & NickRS.Fields(6).Value = "" And "" & NickRS.Fields(2).Value <> "" And "" & NickRS.Fields(4).Value = "" And "" & NickRS.Fields(8).Value = "" Then
                    '若有墨齊日無墨完日
'edit by nickc 2005/04/18 瓊玉說無完稿日不算逾時
'                    ElseIf "" & NickRS.Fields(2).Value <> "" And "" & NickRS.Fields(4).Value = "" Then
''                        If GetWorkDay(strSrvDate(1), CheckStr(NickRS.Fields(2))) > 3 Then
'                        If GetWorkDay(strEndDate, CheckStr(NickRS.Fields(2))) > 3 Then
'                            cnnConnection.Execute "INSERT INTO R090711_1 (R111001,R111002,R111003,ID) VALUES ('" & CheckStr(NickRS.Fields(0)) & "',8,1,'" & strUserNum & "') "
'                        End If
                    End If
                End If
            Case Else
                If CheckStr(NickRS.Fields(3)) = "1" Then  '草圖
                    '若有草齊日及草完日
                    If "" & NickRS.Fields(2).Value <> "" And "" & NickRS.Fields(4).Value <> "" Then
                        'Modify By Cheng 2004/02/16
                        '草完日必須為當月
                        'Modify By Cheng 2004/03/12
'                        If Left("" & NickRS.Fields(4).Value, 6) = Left(strSrvDate(1), 6) Then
                        If Left("" & NickRS.Fields(4).Value, 6) = IIf(ProState = "2" Or ProState = "3", "" & (Val(Me.Text1.Text) + 191100), Left(strSrvDate(1), 6)) Then
                        'End
                            If GetWorkDay(CheckStr(NickRS.Fields(4)), "" & NickRS.Fields(2).Value) > 4 Then
                               cnnConnection.Execute "INSERT INTO R090711_1 (R111001,R111002,R111003,ID) VALUES ('" & CheckStr(NickRS.Fields(0)) & "',7,1,'" & strUserNum & "') "
                            End If
                        End If
                        'End
'                    '若無發文日有草齊日無草完日無取消收文日
'                    ElseIf "" & NickRS.Fields(6).Value = "" And "" & NickRS.Fields(2).Value <> "" And "" & NickRS.Fields(4).Value = "" And "" & NickRS.Fields(8).Value = "" Then
                    '若有草齊日無草完日
'edit by nickc 2005/04/18 瓊玉說無完稿日不算逾時
'                    ElseIf "" & NickRS.Fields(2).Value <> "" And "" & NickRS.Fields(4).Value = "" Then
''                        If GetWorkDay(strSrvDate(1), CheckStr(NickRS.Fields(2))) > 4 Then
'                        If GetWorkDay(strEndDate, CheckStr(NickRS.Fields(2))) > 4 Then
'                           cnnConnection.Execute "INSERT INTO R090711_1 (R111001,R111002,R111003,ID) VALUES ('" & CheckStr(NickRS.Fields(0)) & "',7,1,'" & strUserNum & "') "
'                        End If
                    End If
                Else '墨圖
                    '若有墨齊日及墨完日
                    If "" & NickRS.Fields(2).Value <> "" And "" & NickRS.Fields(4).Value <> "" Then
                        'Modify By Cheng 2004/02/16
                        '墨完日必須為當月
                        'Modify By Cheng 2004/03/12
'                        If Left("" & NickRS.Fields(4).Value, 6) = Left(strSrvDate(1), 6) Then
                        If Left("" & NickRS.Fields(4).Value, 6) = IIf(ProState = "2" Or ProState = "3", "" & (Val(Me.Text1.Text) + 191100), Left(strSrvDate(1), 6)) Then
                        'End
                            If GetWorkDay((CheckStr(NickRS.Fields(4))), "" & NickRS.Fields(2).Value) > 3 Then
                                cnnConnection.Execute "INSERT INTO R090711_1 (R111001,R111002,R111003,ID) VALUES ('" & CheckStr(NickRS.Fields(0)) & "',8,1,'" & strUserNum & "') "
                            End If
                        End If
                        'End
'                    '若無發文日有墨齊日無墨完日無取消收文日
'                    ElseIf "" & NickRS.Fields(6).Value = "" And "" & NickRS.Fields(2).Value <> "" And "" & NickRS.Fields(4).Value = "" And "" & NickRS.Fields(8).Value = "" Then
                    '若有墨齊日無墨完日
'edit by nickc 2005/04/18 瓊玉說無完稿日不算逾時
'                    ElseIf "" & NickRS.Fields(2).Value <> "" And "" & NickRS.Fields(4).Value = "" Then
''                        If GetWorkDay(strSrvDate(1), CheckStr(NickRS.Fields(2))) > 3 Then
'                        If GetWorkDay(strEndDate, CheckStr(NickRS.Fields(2))) > 3 Then
'                           cnnConnection.Execute "INSERT INTO R090711_1 (R111001,R111002,R111003,ID) VALUES ('" & CheckStr(NickRS.Fields(0)) & "',8,1,'" & strUserNum & "') "
'                        End If
                    End If
                End If
            End Select
            NickRS.MoveNext
        Loop
    End If
    If NickRS.State = 1 Then NickRS.Close
    
    'strSQL = "INSERT INTO R090711_1 (R111001,R111002,R111003,ID) select EP13,7,count(*),'" & strUserNum & "' from engineerprogress,caseprogress where EP02=CP09(+) AND  cp01 in ('CFP','P') AND SUBSTR(CP27,1,6)=" & Val(Text1) + 191100 & " and cp57 is  null AND cp26 is null AND EP15>CP48 GROUP BY ST02 "
    'cnnConnection.Execute strSQL
    'DoEvents
    'strSQL = "INSERT INTO R090711_1 (R111001,R111002,R111003,ID) select EP13,8,count(*),'" & strUserNum & "' from engineerprogress,caseprogress where EP02=CP09(+) AND  cp01 in ('CFP','P') AND SUBSTR(CP27,1,6)=" & Val(Text1) + 191100 & " and cp57 is  null  AND cp26 is null AND EP18>CP48 GROUP BY ST02 "
    'cnnConnection.Execute strSQL
    'DoEvents
    '寫入沒有資料的繪圖人員
With adoRecordset
    CheckOC
    strSql = "SELECT DISTINCT R111001 FROM R090711_1 WHERE ID='" & strUserNum & "' "
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        .MoveFirst
        Do While .EOF = False
            'edit by nickc 2005/04/18
            'For i = 1 To 8
            For i = 1 To 11
                strSql = "SELECT * FROM R090711_1 WHERE ID='" & strUserNum & "' AND R111002='" & Trim(str(i)) & "' AND R111001='" & CheckStr(.Fields(0)) & "' "
                CheckOC2
                adoRecordset1.CursorLocation = adUseClient
                adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
                If adoRecordset1.RecordCount <> 0 And adoRecordset1.RecordCount > 0 Then
                Else
                    'edit by nickc 2005/05/04
                    'strSQL = "INSERT INTO R090711_1 VALUES('" & CheckStr(.Fields(0)) & "','" & Trim(str(i)) & "',0,0,0,'" & strUserNum & "') "
                    strSql = "INSERT INTO R090711_1 (r111001,r111002,r111003,r111004,r111005,id,r111006,r111007,r111008) VALUES('" & CheckStr(.Fields(0)) & "','" & Trim(str(i)) & "',0,0,0,'" & strUserNum & "',0,0,0) "
                    cnnConnection.Execute strSql
                End If
            Next i
            DoEvents
            .MoveNext
            CheckOC2
        Loop
    End If
End With
CheckOC
Me.Enabled = True
Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Click(Index As Integer)
Dim nFrm As Form
   
Select Case Index
Case 0 '當月資料
     StrMenu
Case 1 '未發文
     StrMenu5
'Modify By Sindy 2013/5/16
Case 2 '承辦歷程
      'NowPrint lbl1(3).Caption, "99", "00", True, strUserNum '申請書
      'Add By Sindy 2013/6/10
      If SetColTag(False) = False Then
         Call cmdOK_Click(2)
         If m_chkcmdok1 = False Then Exit Sub
      End If
      
'      'Add By Sindy 2017/9/19
'      '檢查表單是否已開啟，若是，則關閉
'      For Each nFrm In Forms
'         If StrComp(nFrm.Name, "frm090202_2", vbTextCompare) = 0 Then
'            Unload frm090202_2
'         End If
'      Next
'      '2017/9/19 END
      If PUB_ChkFormIsClose("frm090202_2") = False Then Exit Sub 'Add By Sindy 2020/1/17
      intBackTab = 1
      '2013/6/10 End
      frm090202_2.Hide
      frm090202_2.m_EEP01 = lbl1(1) '總收文號
      frm090202_2.m_FlowUserNum = Trim(Left("" & Combo1.Text, 6)) 'Add By Sindy 2013/9/12 案件流程所屬人員
      frm090202_2.intReceiveKind = 3 '繪圖人員工作進度
      frm090202_2.SetParent Me
      If frm090202_2.QueryData = True Then
         frm090202_2.Show
         Me.Hide
      End If
      Exit Sub
'2013/5/16 End
Case Else
End Select
If grd1.col = 0 Then
    If Len(grd1.Text) <> 0 Then
        TextOk = True
        MouseClick (1)
        TextOk = False
    End If
End If
End Sub

'add by nick 2004/12/20
Private Sub Cmd1_Click()
Me.Hide
Screen.MousePointer = vbHourglass
frm090711_3.Show
frm090711_3.StrMenu (lbl1(3).Caption)
Screen.MousePointer = vbDefault
End Sub

Private Sub cmd2_Click()
Me.Hide
Screen.MousePointer = vbHourglass
frm090711_4.Show
frm090711_4.StrMenu (lbl1(3).Caption)
Screen.MousePointer = vbDefault
End Sub

Private Sub cmd3_Click()
Me.Hide
Screen.MousePointer = vbHourglass
frm090711_5.Show
frm090711_5.StrMenu (lbl1(3).Caption)
Screen.MousePointer = vbDefault
End Sub

'Add By Sindy 2013/8/19
Private Sub cmdDetail_Click()
   Call grd2_DblClick
End Sub

'Modify By Sindy 2013/5/21
'Private Sub cmdOK_Click(index As Integer)
Public Sub cmdOK_Click(Index As Integer)
'2013/5/21 End
Dim ii As Integer '序號
Dim strOfficeKind As String '所別
Dim bolInTrans As Boolean 'Added by Morgan 2022/6/15

Select Case Index
Case 0 '本月統計
     Me.Hide
     StrMenu2
     frm090711_1.Show
Case 1 '回前畫面
     Select Case ProState
     Case "1" '個人
         Unload Me
     Case "2"
         Me.Hide
         frm090706.Show
         Unload Me
     Case "3"
         Unload Me
     Case "4"
         Me.Hide
         frm090303_1.Show
         Unload Me
     Case Else
     End Select
Case 2 '存檔
     If SSTab1.Tab = 1 Then
         If ChkNoData = False Then
            ChkData2 = True
            'Modify By Cheng 2003/09/17
            '因為取消按確定時存檔, 所以可不檢查欄位有效性
            'Begin
'            For Each TXT090711 In txt1
'               If txt1(TXT090711.Index).Visible = True And txt1(TXT090711.Index).Enabled = True Then
'                  txt1_LostFocus (TXT090711.Index)
'                  If ChkData2 = False Then Exit Sub
'               End If
'            Next
            'End
            m_chkcmdok1 = False 'Add By Sindy 2013/6/7 進入承辦歷程時會先執行一次確定鍵,因有可能已在此畫面先修改資料,且有些日期檢查條件須先執行
            'add by nickc 2005/03/17 重新檢查欄位有效性
            If TxtValidate = False Then
               Screen.MousePointer = vbDefault
               Exit Sub
            End If
            
            Screen.MousePointer = vbHourglass
            Me.Enabled = False
            'Modify by Morgan 2011/1/4 配合日期排序前面會加空白要去除
            Txt1(1).Text = LTrim(Replace(Txt1(1).Text, "******", ""))
            Txt1(2).Text = LTrim(Replace(Txt1(2).Text, "******", ""))
            Txt1(4).Text = LTrim(Replace(Txt1(4).Text, "******", ""))
            Txt1(5).Text = LTrim(Replace(Txt1(5).Text, "******", ""))
            
            On Error GoTo ErrorHandler
            cnnConnection.BeginTrans
            bolInTrans = True 'Added by Morgan 2022/6/15
            'Modify By Cheng 2003/06/27
            '加更新墨是否計件欄(EP29)
'            strSQL = "UPDATE ENGINEERPROGRESS SET EP13='" & txt1(0) & "',EP14=" & IIf(Val(ChangeTStringToWString(txt1(1))) <> 0, ChangeTStringToWString(txt1(1)), "NULL") & ",EP15=" & IIf(Val(ChangeTStringToWString(txt1(2))) <> 0, Val(ChangeTStringToWString(txt1(2))), "NULL") & ",EP16=" & Val(txt1(3)) & ",EP17=" & IIf(Val(ChangeTStringToWString(txt1(4))) <> 0, Val(ChangeTStringToWString(txt1(4))), "NULL") & ",EP18=" & IIf(Val(ChangeTStringToWString(txt1(5))) <> 0, Val(ChangeTStringToWString(txt1(5))), "NULL") & ",EP19=" & Val(txt1(6)) & ",EP20='" & Trim(txt1(7)) & "',EP21=" & Val(txt1(8)) & ",EP22=" & Val(txt1(9)) & ",EP23=" & Val(txt1(10)) & ",EP24=" & Val(txt1(11)) & ",EP25=" & Val(txt1(12))
            'edit by nickc 2005/04/14 修正是否計件欄位
            'StrSql = "UPDATE ENGINEERPROGRESS SET EP13='" & txt1(0) & "',EP14=" & IIf(Val(ChangeTStringToWString(txt1(1))) <> 0, ChangeTStringToWString(txt1(1)), "NULL") & ",EP15=" & IIf(Val(ChangeTStringToWString(txt1(2))) <> 0, Val(ChangeTStringToWString(txt1(2))), "NULL") & ",EP16=" & Val(txt1(3)) & ",EP17=" & IIf(Val(ChangeTStringToWString(txt1(4))) <> 0, Val(ChangeTStringToWString(txt1(4))), "NULL") & ",EP18=" & IIf(Val(ChangeTStringToWString(txt1(5))) <> 0, Val(ChangeTStringToWString(txt1(5))), "NULL") & ",EP19=" & Val(txt1(6)) & ",EP20='" & Trim(txt1(7)) & "',EP21=" & Val(txt1(8)) & ",EP22=" & Val(txt1(9)) & ",EP23=" & Val(txt1(10)) & ",EP24=" & Val(txt1(11)) & ",EP25=" & Val(txt1(12)) & ", EP29='" & Me.txt1(14).Text & "' "
            '2009/9/11 modify by sonia EP13由trigger更新
            'strSQL = "UPDATE ENGINEERPROGRESS SET EP13='" & txt1(0) & "',EP14=" & IIf(Val(ChangeTStringToWString(txt1(1))) <> 0, ChangeTStringToWString(txt1(1)), "NULL") & ",EP15=" & IIf(Val(ChangeTStringToWString(txt1(2))) <> 0, Val(ChangeTStringToWString(txt1(2))), "NULL") & ",EP16=" & Val(txt1(3)) & ",EP17=" & IIf(Val(ChangeTStringToWString(txt1(4))) <> 0, Val(ChangeTStringToWString(txt1(4))), "NULL") & ",EP18=" & IIf(Val(ChangeTStringToWString(txt1(5))) <> 0, Val(ChangeTStringToWString(txt1(5))), "NULL") & ",EP19=" & Val(txt1(6)) & ",EP20=" & IIf(Trim(txt1(7)) = "", "null", "'" & Trim(txt1(7)) & "'") & ",EP21=" & Val(txt1(8)) & ",EP22=" & Val(txt1(9)) & ",EP23=" & Val(txt1(10)) & ",EP24=" & Val(txt1(11)) & ",EP25=" & Val(txt1(12)) & ", EP29=" & IIf(Trim(txt1(14)) = "", "null", "'" & Me.txt1(14).Text & "'") & " "
            'Modified by Morgan 2012/8/13 EP14,EP17 判斷有修改才更新,否則若恰好有工程師會稿完成更新草墨齊時會被覆蓋
            'strSql = "UPDATE ENGINEERPROGRESS SET EP14=" & IIf(Val(ChangeTStringToWString(txt1(1))) <> 0, ChangeTStringToWString(txt1(1)), "NULL") & ",EP15=" & IIf(Val(ChangeTStringToWString(txt1(2))) <> 0, Val(ChangeTStringToWString(txt1(2))), "NULL") & ",EP16=" & Val(txt1(3)) & ",EP17=" & IIf(Val(ChangeTStringToWString(txt1(4))) <> 0, Val(ChangeTStringToWString(txt1(4))), "NULL") & ",EP18=" & IIf(Val(ChangeTStringToWString(txt1(5))) <> 0, Val(ChangeTStringToWString(txt1(5))), "NULL") & ",EP19=" & Val(txt1(6)) & ",EP20=" & IIf(Trim(txt1(7)) = "", "null", "'" & Trim(txt1(7)) & "'") & ",EP21=" & Val(txt1(8)) & ",EP22=" & Val(txt1(9)) & ",EP23=" & Val(txt1(10)) & ",EP24=" & Val(txt1(11)) & ",EP25=" & Val(txt1(12)) & ", EP29=" & IIf(Trim(txt1(14)) = "", "null", "'" & Me.txt1(14).Text & "'") & " "
            strSql = "UPDATE ENGINEERPROGRESS SET EP15=" & IIf(Val(ChangeTStringToWString(Txt1(2))) <> 0, Val(ChangeTStringToWString(Txt1(2))), "NULL") & ",EP16=" & Val(Txt1(3)) & ",EP18=" & IIf(Val(ChangeTStringToWString(Txt1(5))) <> 0, Val(ChangeTStringToWString(Txt1(5))), "NULL") & ",EP19=" & Val(Txt1(6)) & ",EP20=" & IIf(Trim(Txt1(7)) = "", "null", "'" & Trim(Txt1(7)) & "'") & ",EP21=" & Val(Txt1(8)) & ",EP22=" & Val(Txt1(9)) & ",EP23=" & Val(Txt1(10)) & ",EP24=" & Val(Txt1(11)) & ",EP25=" & Val(Txt1(12)) & ", EP29=" & IIf(Trim(Txt1(14)) = "", "null", "'" & Me.Txt1(14).Text & "'") & " "
            If Txt1(1).Text <> LTrim(Replace(Txt1(1).Tag, "******", "")) Then
               strSql = strSql + ",EP14=" & IIf(Val(ChangeTStringToWString(Txt1(1))) <> 0, ChangeTStringToWString(Txt1(1)), "NULL") & " "
            End If
            If Txt1(4).Text <> LTrim(Replace(Txt1(4).Tag, "******", "")) Then
               strSql = strSql + ",EP17=" & IIf(Val(ChangeTStringToWString(Txt1(4))) <> 0, Val(ChangeTStringToWString(Txt1(4))), "NULL") & " "
            End If
            'end 2012/8/13
            '2009/9/11 end
            If Option1(0).Value = True Then
               strSql = strSql + ",EP26='" & Txt1(13) & "' "
            Else
               strSql = strSql + ",EP26='" & Combo2.Text & "' "
            End If
            strSql = strSql + " WHERE EP02='" & lbl1(1).Caption & "' "
            cnnConnection.Execute strSql
            'Add By Cheng 2003/11/18
            '更新進度檔的繪圖人員
            strSql = "Update CaseProgress Set CP29='" & Me.Txt1(0).Text & "' Where CP09='" & Me.lbl1(1).Caption & "' "
            cnnConnection.Execute strSql
            
            'add by nickc 2006/05/29 若是有更改繪圖人員時，一併更改其他關聯案
            If Me.Txt1(0).Tag <> Me.Txt1(0).Text Then
                'edit by nickc 2006/07/04 因為之前語法太慢，所以修正
                'strSQL = "UPDATE caseprogress set cp29='" & Me.txt1(0).Text & "' where cp01||cp02||cp03||cp04 in (select cm01||cm02||cm03||cm04 from casemap where cm05||'-'||cm06||'-'||cm07||'-'||cm08='" & lbl1(3).Caption & "' and cm10='0'  union select cm05||cm06||cm07||cm08 from casemap where cm01||'-'||cm02||'-'||cm03||'-'||cm04='" & lbl1(3).Caption & "' and cm10='0' )"
                '2009/9/11 modify by sonia casemap的國內案改繪圖人員,其他關聯案才可一併修改且僅限於新申請案,國外案修改時不可改回國內案CFP-021836
                'strSQL = "UPDATE caseprogress set cp29='" & Me.txt1(0).Text & "' where cp09 in (select cp09 from casemap,caseprogress where cm05||'-'||cm06||'-'||cm07||'-'||cm08='" & lbl1(3).Caption & "' and cm10='0' and cm01=cp01(+) and cm02=cp02(+) and cm03=cp03(+) and cm04=cp04(+)  union select cp09 from casemap,caseprogress where cm01||'-'||cm02||'-'||cm03||'-'||cm04='" & lbl1(3).Caption & "' and cm10='0'  and cm05=cp01(+) and cm06=cp02(+) and cm07=cp03(+) and cm08=cp04(+) ) "
                strSql = "UPDATE caseprogress set cp29='" & Me.Txt1(0).Text & "' where cp09 in (select cp09 from casemap,caseprogress where cm05||'-'||cm06||'-'||cm07||'-'||cm08='" & lbl1(3).Caption & "' and cm10='0' and cm01=cp01(+) and cm02=cp02(+) and cm03=cp03(+) and cm04=cp04(+) and instr('" & NewCasePtyList & "',cp10)>0) and cp27 is null "
                '2009/9/11 END
                cnnConnection.Execute strSql
                
               'Added by Morgan 2013/7/3
               '更新一案兩請繪圖人員
               strSql = "update CaseProgress set cp29='" & Me.Txt1(0).Text & "' Where cp09 in " & _
                  "(select cp09 from caseprogress,casemap where cm05||'-'||cm06||'-'||cm07||'-'||cm08='" & lbl1(3).Caption & "' " & _
                  " and cm10='3' and cm01=cp01(+) and cm02=cp02(+) and cm03=cp03(+) and cm04=cp04(+) AND CP27 IS NULL and instr('" & NewCasePtyList & "',cp10)>0" & _
                  " union all select cp09 from caseprogress,casemap where  cm01||'-'||cm02||'-'||cm03||'-'||cm04='" & lbl1(3).Caption & "' " & _
                  " and cm10='3' and cm05=cp01(+) and cm06=cp02(+) and cm07=cp03(+) and cm08=cp04(+) AND CP27 IS NULL and instr('" & NewCasePtyList & "',cp10)>0) "
               cnnConnection.Execute strSql, intI
               'end 2013/7/3
            End If
            'add by nickc 2005/04/12 繪圖主管可修改是否提供圖檔
            strSql = "Update CaseProgress Set cp106=" & IIf(Trim(Txt1(19)) = "", "null", "'" & Me.Txt1(19).Text & "'") & " Where CP09='" & Me.lbl1(1).Caption & "' "
            cnnConnection.Execute strSql
            'add by nickc 2005/03/17  加重新計算寄件值
            m_CP09 = Me.lbl1(1).Caption
            'PUB_UpdateCaseValue m_CP09 'Remove by Morgan 2005/4/13 改由 trigger 更新
            
            
'Modified by Morgan 2013/5/9 草墨合併通知 --瓊玉
'            'add by nickc 2005/03/17  加入儲存加乘註記及理由
'            If Val(txt1(15)) <> Val(m_CP101) Then
'                  strSql = "Update CaseProgress Set cp101='" & txt1(15).Text & "',cp102='" & ChgSQL(txt1(16).Text) & "'  where CP09='" & Me.lbl1(1).Caption & "' "
'                  cnnConnection.Execute strSql
'                  '增加紀錄
'                  strSql = "insert into flagstory (fs01,fs02,fs03,fs04,fs05,fs06,fs07,fs08) select '" & Me.lbl1(1).Caption & "',to_number(to_char(sysdate,'YYYYMMDD')),to_number(to_char(sysdate,'HH24MIss')),'2','" & m_CP101 & "','" & Trim(txt1(15)) & "','" & ChgSQL(Trim(txt1(16))) & "','" & strUserNum & "' from dual  "
'                  cnnConnection.Execute strSql
'                  'edit by nickc 2006/12/29 改在 trans 後發
'                  'add by nickc 2005/04/13  發 mail
'                  'PUB_SendMail strUserNum, "72006", Me.lbl1(1).Caption, "更改草加乘註記", "本所案號：" & lbl1(3).Caption & vbCrLf & "收文號：" & lbl1(1).Caption & vbCrLf & "原加乘註記：" & m_CP101 & vbCrLf & "更改後加乘註記：" & txt1(15) & vbCrLf & "理由：" & txt1(16)
'                    ReDim Preserve skMail(UBound(skMail) + 1) As SeekMails
'                    skMail(UBound(skMail)).fiSender = strUserNum
'                    skMail(UBound(skMail)).fiReceiver = "72006"
'                    skMail(UBound(skMail)).fiContent = "本所案號：" & lbl1(3).Caption & vbCrLf & "收文號：" & lbl1(1).Caption & vbCrLf & "原加乘註記：" & m_CP101 & vbCrLf & "更改後加乘註記：" & txt1(15) & vbCrLf & "理由：" & txt1(16)
'                    skMail(UBound(skMail)).fiSubject = "更改草加乘註記"
'                    skMail(UBound(skMail)).fiRecriverNo = Me.lbl1(1).Caption
'                  'edit by nickc 2006/12/29 改在 trans 後發
'                  'add by nickc 2005/04/14  加發 mail 給協理
'                  'PUB_SendMail strUserNum, "71011", Me.lbl1(1).Caption, "更改草加乘註記", "本所案號：" & lbl1(3).Caption & vbCrLf & "收文號：" & lbl1(1).Caption & vbCrLf & "原加乘註記：" & m_CP101 & vbCrLf & "更改後加乘註記：" & txt1(15) & vbCrLf & "理由：" & txt1(16)
'                    ReDim Preserve skMail(UBound(skMail) + 1) As SeekMails
'                    skMail(UBound(skMail)).fiSender = strUserNum
'                    skMail(UBound(skMail)).fiReceiver = "71011"
'                    skMail(UBound(skMail)).fiContent = "本所案號：" & lbl1(3).Caption & vbCrLf & "收文號：" & lbl1(1).Caption & vbCrLf & "原加乘註記：" & m_CP101 & vbCrLf & "更改後加乘註記：" & txt1(15) & vbCrLf & "理由：" & txt1(16)
'                    skMail(UBound(skMail)).fiSubject = "更改草加乘註記"
'                    skMail(UBound(skMail)).fiRecriverNo = Me.lbl1(1).Caption
'            End If
'            If Val(txt1(17)) <> Val(m_CP104) Then
'                  strSql = "Update CaseProgress Set cp104='" & txt1(17).Text & "',cp105='" & ChgSQL(txt1(18).Text) & "'  where CP09='" & Me.lbl1(1).Caption & "' "
'                  cnnConnection.Execute strSql
'                  '增加紀錄
'                  strSql = "insert into flagstory (fs01,fs02,fs03,fs04,fs05,fs06,fs07,fs08) select '" & Me.lbl1(1).Caption & "',to_number(to_char(sysdate,'YYYYMMDD')),to_number(to_char(sysdate,'HH24MIss')),'3','" & m_CP104 & "','" & Trim(txt1(17)) & "','" & ChgSQL(Trim(txt1(18))) & "','" & strUserNum & "' from dual  "
'                  cnnConnection.Execute strSql
'                  'edit by nickc 2006/12/29 改在 trans 後發
'                  'add by nickc 2005/04/13  發 mail
'                  'PUB_SendMail strUserNum, "72006", Me.lbl1(1).Caption, "更改墨加乘註記", "本所案號：" & lbl1(3).Caption & vbCrLf & "收文號：" & lbl1(1).Caption & vbCrLf & "原加乘註記：" & m_CP104 & vbCrLf & "更改後加乘註記：" & txt1(17) & vbCrLf & "理由：" & txt1(18)
'                    ReDim Preserve skMail(UBound(skMail) + 1) As SeekMails
'                    skMail(UBound(skMail)).fiSender = strUserNum
'                    skMail(UBound(skMail)).fiReceiver = "72006"
'                    skMail(UBound(skMail)).fiContent = "本所案號：" & lbl1(3).Caption & vbCrLf & "收文號：" & lbl1(1).Caption & vbCrLf & "原加乘註記：" & m_CP104 & vbCrLf & "更改後加乘註記：" & txt1(17) & vbCrLf & "理由：" & txt1(18)
'                    skMail(UBound(skMail)).fiSubject = "更改墨加乘註記"
'                    skMail(UBound(skMail)).fiRecriverNo = Me.lbl1(1).Caption
'                  'edit by nickc 2006/12/29 改在 trans 後發
'                  'add by nickc 2005/04/14  加發 mail 給協理
'                  'PUB_SendMail strUserNum, "71011", Me.lbl1(1).Caption, "更改墨加乘註記", "本所案號：" & lbl1(3).Caption & vbCrLf & "收文號：" & lbl1(1).Caption & vbCrLf & "原加乘註記：" & m_CP104 & vbCrLf & "更改後加乘註記：" & txt1(17) & vbCrLf & "理由：" & txt1(18)
'                    ReDim Preserve skMail(UBound(skMail) + 1) As SeekMails
'                    skMail(UBound(skMail)).fiSender = strUserNum
'                    skMail(UBound(skMail)).fiReceiver = "71011"
'                    skMail(UBound(skMail)).fiContent = "本所案號：" & lbl1(3).Caption & vbCrLf & "收文號：" & lbl1(1).Caption & vbCrLf & "原加乘註記：" & m_CP104 & vbCrLf & "更改後加乘註記：" & txt1(17) & vbCrLf & "理由：" & txt1(18)
'                    skMail(UBound(skMail)).fiSubject = "更改墨加乘註記"
'                    skMail(UBound(skMail)).fiRecriverNo = Me.lbl1(1).Caption
'            End If
            If Val(Txt1(15)) <> Val(m_CP101) Or Val(Txt1(17)) <> Val(m_CP104) Then
               If Val(Txt1(15)) <> Val(m_CP101) Then
                  strSql = "Update CaseProgress Set cp101='" & Txt1(15).Text & "',cp102='" & ChgSQL(Txt1(16).Text) & "'  where CP09='" & Me.lbl1(1).Caption & "' "
                  cnnConnection.Execute strSql
                  '增加紀錄
                  strSql = "insert into flagstory (fs01,fs02,fs03,fs04,fs05,fs06,fs07,fs08) select '" & Me.lbl1(1).Caption & "',to_number(to_char(sysdate,'YYYYMMDD')),to_number(to_char(sysdate,'HH24MIss')),'2','" & m_CP101 & "','" & Trim(Txt1(15)) & "','" & ChgSQL(Trim(Txt1(16))) & "','" & strUserNum & "' from dual  "
                  cnnConnection.Execute strSql
               End If
               
               If Val(Txt1(17)) <> Val(m_CP104) Then
                  strSql = "Update CaseProgress Set cp104='" & Txt1(17).Text & "',cp105='" & ChgSQL(Txt1(18).Text) & "'  where CP09='" & Me.lbl1(1).Caption & "' "
                  cnnConnection.Execute strSql
                  '增加紀錄
                  strSql = "insert into flagstory (fs01,fs02,fs03,fs04,fs05,fs06,fs07,fs08) select '" & Me.lbl1(1).Caption & "',to_number(to_char(sysdate,'YYYYMMDD')),to_number(to_char(sysdate,'HH24MIss')),'3','" & m_CP104 & "','" & Trim(Txt1(17)) & "','" & ChgSQL(Trim(Txt1(18))) & "','" & strUserNum & "' from dual  "
                  cnnConnection.Execute strSql
               End If
               
               ReDim Preserve skMail(UBound(skMail) + 1) As SeekMails
               skMail(UBound(skMail)).fiSender = strUserNum
               'modify by sonia 2016/3/3 改72006為73022
               'Modified by Lydia 2022/12/27 改成系統特殊設定
               'skMail(UBound(skMail)).fiReceiver = "71011;73022"
               skMail(UBound(skMail)).fiReceiver = Pub_GetSpecMan("更改草墨加乘註記收受者")
               If Val(Txt1(15)) <> Val(m_CP101) And Val(Txt1(17)) <> Val(m_CP104) Then
                  skMail(UBound(skMail)).fiContent = "本所案號：" & lbl1(3).Caption & vbCrLf & "收文號：" & lbl1(1).Caption & _
                     vbCrLf & vbCrLf & "原草加乘註記：" & m_CP101 & vbCrLf & "更改後草加乘註記：" & Txt1(15) & vbCrLf & "理由：" & Txt1(16) & _
                     vbCrLf & vbCrLf & "原墨加乘註記：" & m_CP104 & vbCrLf & "更改後墨加乘註記：" & Txt1(17) & vbCrLf & "理由：" & Txt1(18)
                  skMail(UBound(skMail)).fiSubject = "更改草墨加乘註記"
               ElseIf Val(Txt1(15)) <> Val(m_CP101) Then
                  skMail(UBound(skMail)).fiContent = "本所案號：" & lbl1(3).Caption & vbCrLf & "收文號：" & lbl1(1).Caption & vbCrLf & "原加乘註記：" & m_CP101 & vbCrLf & "更改後加乘註記：" & Txt1(15) & vbCrLf & "理由：" & Txt1(16)
                  skMail(UBound(skMail)).fiSubject = "更改草加乘註記"
               Else
                  skMail(UBound(skMail)).fiContent = "本所案號：" & lbl1(3).Caption & vbCrLf & "收文號：" & lbl1(1).Caption & vbCrLf & "原加乘註記：" & m_CP104 & vbCrLf & "更改後加乘註記：" & Txt1(17) & vbCrLf & "理由：" & Txt1(18)
                  skMail(UBound(skMail)).fiSubject = "更改墨加乘註記"
               End If
               skMail(UBound(skMail)).fiRecriverNo = Me.lbl1(1).Caption
            End If
'end 2013/5/9
            
            cnnConnection.CommitTrans
            bolInTrans = False 'Added by Morgan 2022/6/15
            m_chkcmdok1 = True 'Add By Sindy 2013/6/7
            'Add By Cheng 2003/06/30
            '若更改繪圖人員, 則發E-Mail通知
            'edit by nickc 2006/06/28
            'If Me.txt1(0).Text <> Me.txt1(0).Tag Then
            If Me.Txt1(0).Text <> Me.Txt1(0).Tag And Txt1(0).Text <> "99999" Then
            
                'add by nickc 2007/07/12 加入詢問要不要問
                If MsgBox("要不要寄通知信給工程師??", vbYesNo + vbExclamation, "請回答!!") = vbYes Then
                        Screen.MousePointer = vbDefault
        '                MsgBox "若您未開啟OutLook，請先開啟!!!", vbExclamation + vbOKOnly
                        'edit by nickc 2006/12/29 改在 trans 後發
                        'Load frm880005
                         ReDim Preserve skMail(UBound(skMail) + 1) As SeekMails
                         skMail(UBound(skMail)).fiSender = strUserNum
        
                        'Modify By Cheng 2003/08/11
                        strOfficeKind = PUB_GetST06(strUserNum)
                        '若使用者為北所人員, 則E-Mail後面不加@taie.com.tw
                        If strOfficeKind = "1" Then
                            'edit by nickc 2006/12/29 改在 trans 後發
                            'frm880005.txtEmail(0).Text = IIf(Me.txt1(0).Tag <> "", Trim(Me.txt1(0).Tag) & ";", "") & IIf(Me.txt1(0).Text <> "", Trim(Me.txt1(0).Text) & ";", "") & _
                                                                         GetCP14EMail(Me.lbl1(1).Caption, strOfficeKind)
                            skMail(UBound(skMail)).fiReceiver = IIf(Me.Txt1(0).Tag <> "", Trim(Me.Txt1(0).Tag) & ";", "") & IIf(Me.Txt1(0).Text <> "", Trim(Me.Txt1(0).Text) & ";", "") & _
                                                                         GetCP14EMail(Me.lbl1(1).Caption, strOfficeKind)
                        '若使用者非北所人員, 則E-Mail後面加@taie.com.tw
                        Else
                            'edit by nickc 2006/12/29 改在 trans 後發
                            'frm880005.txtEmail(0).Text = IIf(Me.txt1(0).Tag <> "", Trim(Me.txt1(0).Tag) & "@taie.com.tw;", "") & IIf(Me.txt1(0).Text <> "", Trim(Me.txt1(0).Text) & "@taie.com.tw;", "") & _
                                                                         GetCP14EMail(Me.lbl1(1).Caption, strOfficeKind)
                            skMail(UBound(skMail)).fiReceiver = IIf(Me.Txt1(0).Tag <> "", Trim(Me.Txt1(0).Tag) & "@taie.com.tw;", "") & IIf(Me.Txt1(0).Text <> "", Trim(Me.Txt1(0).Text) & "@taie.com.tw;", "") & _
                                                                         GetCP14EMail(Me.lbl1(1).Caption, strOfficeKind)
                        End If
                        'edit by nickc 2006/12/29 改在 trans 後發
                        'frm880005.txtEmail(1).Text = "繪圖人員變更通知"
                        'frm880005.txtEmail(2).Text = "本所案號：" & Me.lbl1(3).Caption & vbCrLf & _
                                                                    "案件名稱：" & Me.lbl1(4).Caption & vbCrLf & _
                                                                    "收文日：" & DBYEAR(Me.lbl1(2).Caption) - 1911 & " 年 " & DBMONTH(Me.lbl1(2).Caption) & " 月 " & DBDAY(Me.lbl1(2).Caption) & " 日 " & vbCrLf & _
                                                                    "原繪圖人員：" & GetStaffName(Me.txt1(0).Tag) & vbCrLf & _
                                                                    "變更後繪圖人員：" & GetStaffName(Me.txt1(0).Text) & vbCrLf & _
                                                                    "承辦工程師：" & GetStaffName(Replace(Left(GetCP14EMail(Me.lbl1(1).Caption, strOfficeKind), 6), "@", ""))
                        'frm880005.Visible = False
                        'frm880005.Show vbModal
                        skMail(UBound(skMail)).fiSubject = "繪圖人員變更通知"
                        skMail(UBound(skMail)).fiContent = "本所案號：" & Me.lbl1(3).Caption & vbCrLf & _
                                                                    "案件名稱：" & Me.lbl1(4).Caption & vbCrLf & _
                                                                    "收文日：" & DBYEAR(Me.lbl1(2).Caption) - 1911 & " 年 " & DBMONTH(Me.lbl1(2).Caption) & " 月 " & DBDAY(Me.lbl1(2).Caption) & " 日 " & vbCrLf & _
                                                                    "原繪圖人員：" & GetStaffName(Me.Txt1(0).Tag) & vbCrLf & _
                                                                    "變更後繪圖人員：" & GetStaffName(Me.Txt1(0).Text) & vbCrLf & _
                                                                    "承辦工程師：" & GetStaffName(Replace(Left(GetCP14EMail(Me.lbl1(1).Caption, strOfficeKind), 6), "@", ""))
                        'Modified by Lydia 2022/05/30 傳入收文號
                        'skMail(UBound(skMail)).fiRecriverNo = ""
                        skMail(UBound(skMail)).fiRecriverNo = lbl1(1).Caption

                End If
                Screen.MousePointer = vbHourglass
                'Add By Cheng 2003/09/16
                '若有修改繪圖人員, 則移除此筆瀏覽資料
                If SWPRow <> "" And SWPRow <> "0" Then
                    'Modified by Morgan 2022/6/15
                    'For ii = Val(SWPRow) To Val(SWPRow)
                    '    Me.GRD1.RemoveItem ii
                    'Next ii
                    'SWPRow = "1"
                    If grd1.Rows = 2 Then
                        For ii = 0 To grd1.Cols - 1
                           grd1.TextMatrix(Val(SWPRow), ii) = ""
                        Next
                        SWPRow = "0"
                    Else
                        grd1.RemoveItem Val(SWPRow)
                        SWPRow = "1"
                    End If
                    'end 2022/6/15
                End If
            End If
            
            'add by nickc 2006/12/29 集中發信
            For i = 1 To UBound(skMail)
                 PUB_SendMail skMail(i).fiSender, skMail(i).fiReceiver, skMail(i).fiRecriverNo, skMail(i).fiSubject, skMail(i).fiContent
            Next i
            ReDim skMail(0) As SeekMails
            
            'Combo1.Clear
            'Modify By Cheng 2003/09/16
            '取消
            'Begin
'            '抓繪圖人員
'            StrMenu1
'            '抓繪圖人員的工作進度資料
'            grd1.Clear
'            grd1.Rows = 2
'            StrMenu
'            SetGrd1
            'End
            'Text1 = ""
            'Combo1.Text = Combo1.List(0)
            'Modify By Cheng 2003/09/16
            '取消
            'Begin
'            '設定某一繪圖人員的資料
'            Combo1_Click
            'End
            'Add By Cheng 2003/09/16
            '若未更改繪圖人員, 直接更新此筆瀏覽資料
            If Me.Txt1(0).Text = Me.Txt1(0).Tag Then
                'add by nickc 2005/04/13 重抓一次資料 計件值加乘註記
               strSql = "select * from caseprogress where cp09='" & lbl1(1).Caption & "' "
               CheckOC3
               AdoRecordSet3.CursorLocation = adUseClient
               AdoRecordSet3.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
               If AdoRecordSet3.RecordCount <> 0 Then
                    lbl1(33).Caption = "" & AdoRecordSet3.Fields("cp100").Value
                    Txt1(15).Text = "" & AdoRecordSet3.Fields("cp101").Value
                    m_CP101 = "" & AdoRecordSet3.Fields("cp101").Value
                    Txt1(16).Text = "" & AdoRecordSet3.Fields("cp102").Value
                    m_CP102 = "" & AdoRecordSet3.Fields("cp102").Value
                    lbl1(34).Caption = "" & AdoRecordSet3.Fields("cp103").Value
                    Txt1(17).Text = "" & AdoRecordSet3.Fields("cp104").Value
                    m_CP104 = "" & AdoRecordSet3.Fields("cp104").Value
                    Txt1(18).Text = "" & AdoRecordSet3.Fields("cp105").Value
                    m_CP105 = "" & AdoRecordSet3.Fields("cp105").Value
                    'add by nickc 2005/04/04
                    Txt1(19) = "" & AdoRecordSet3.Fields("cp106").Value
               Else
                    lbl1(33).Caption = ""
                    Txt1(15).Text = ""
                    m_CP101 = ""
                    Txt1(16).Text = ""
                    m_CP102 = ""
                    lbl1(34).Caption = ""
                    Txt1(17).Text = ""
                    m_CP104 = ""
                    Txt1(18).Text = ""
                    m_CP105 = ""
                    'add by nickc 2005/04/04
                    Txt1(19).Text = ""
               End If
               CheckOC3
                RefreshOneRecord
            End If
         End If
            'Add By Cheng 2002/04/17
'            MouseClick (1)
            MouseClick IIf(Val("" & SWPRow) < 1, 1, SWPRow)
            'Option1(0).Value = True
            'Txt1(13).Enabled = True
            'Combo2.AddItem "速件", 0
            'Combo2.AddItem "未齊備", 1
            'Combo2.AddItem "複雜", 2
            'Combo2.AddItem "其他新案", 3
            'Combo2.AddItem "ACAD", 4
            'Combo2.Text = "請選擇...."
            'Combo2.Enabled = False
            SSTab1.Tab = 0
            'Modify By Cheng 2004/04/19
'            cmdOK(2).Caption = "確定(&O)"
            cmdOK(2).Caption = "確定"
            'End
            Me.Enabled = True
            Screen.MousePointer = vbDefault
      Else
         SSTab1.Tab = 1
        'Modify By Cheng 2004/04/19
'         cmdOK(2).Caption = "存檔(&O)"
         cmdOK(2).Caption = "存檔"
        'End
      End If

Case 3 'Add by Morgan 2011/3/30
      frm090711_6.p_CP43 = lbl1(1).Caption
      frm090711_6.Show vbModal
      PUB_SendMailCache
Case Else
End Select
'Add By Cheng 2003/11/18
Exit Sub
ErrorHandler:
    If bolInTrans Then 'Added by Morgan 2022/6/15
      cnnConnection.RollbackTrans
    End If
    MsgBox Err.Description
    Me.Enabled = True
    Screen.MousePointer = vbDefault
End Sub

Private Sub Combo1_Click()
Dim txt As Object, Lbl As Object

'Add By Cheng 2003/05/22
Screen.MousePointer = vbHourglass
For Each txt In frm090711.Txt1
    txt.Text = ""
Next
For Each Lbl In frm090711.lbl1
    Lbl.Caption = ""
Next
grd1.Clear
grd1.Rows = 2
'Add By Sindy 2013/9/16 在切換繪圖人員時,會出現”陣列索引超出範圍”
If Combo1.Tag <> Combo1.Text Then
   SWPRow = 0
   dblPrevRow = 0
   Combo1.Tag = Combo1.Text
   StrMenu 'Modify by Sindy 2016/9/6
End If
'2013/9/16
'StrMenu
'Add By Cheng 2003/05/22
Screen.MousePointer = vbHourglass
SetGrd1
With grd1
    For i = 0 To .Rows - 1
        .col = 0
        .row = i
        If .CellBackColor = &HFFC0C0 Then
            For k = 0 To .Cols - 1
                .col = k
                .CellBackColor = QBColor(15)
            Next k
            Exit For
        End If
    Next i
End With
MouseClick (1)
'Modify By Cheng 2003/06/05
'LBL2.Caption = GetPrjSalesNM(Combo1.Text)
lbl2.Caption = GetPrjSalesNM(Trim(Left(Combo1.Text, 6)))
Option1(0).Value = True
If ProState <> "4" Then
    Txt1(13).Enabled = True
End If
Combo2.Clear
Combo2.AddItem "速件", 0
Combo2.AddItem "未齊備", 1
Combo2.AddItem "複雜", 2
Combo2.AddItem "其他新案", 3
Combo2.AddItem "ACAD", 4
Combo2.Text = "請選擇...."
Combo2.Enabled = False
If ChkNoData = True Then
   For s = 0 To 13
      Txt1(s).Enabled = False
   Next s
   Combo2.Enabled = False
   Option1(0).Enabled = False
   Option1(1).Enabled = False
Else
   If ProState <> "4" Then
      For s = 0 To 13
         Txt1(s).Enabled = True
      Next s
      Combo2.Enabled = True
      Option1(0).Enabled = True
      Option1(1).Enabled = True
   End If
End If
If Me.SSTab1.Tab <> 0 Then Me.SSTab1.Tab = 0: DoEvents
Call GetMonthAssess 'Add by Amy 2018/10/08
'Add By Cheng 2003/05/22
Screen.MousePointer = vbDefault
'StrMenu
'SetGrd1
'TextOk = True
'grd1_Click
'TextOk = False
End Sub

Private Sub Combo5_Click()
   If Me.Visible = True Then
      If QueryData(True) = False Then ShowNoData 'Add By Sindy 2023/4/12
   End If
End Sub

Private Sub Form_Activate()
'Dim nFrm As Form
   
   ProState = m_ProState 'Add By Sindy 2017/8/10 重新設定權限
   
'   'Add By Sindy 2017/8/30
'   '檢查表單是否已開啟，若是，則關閉
'   If Me.Visible = True Then
'      For Each nFrm In Forms
'         If StrComp(nFrm.Name, "frm090202_2", vbTextCompare) = 0 Then
'            If UCase(frm090202_2.m_PrevForm.Name) <> UCase(Me.Name) Then Exit For
'            'If frm090202_2.intReceiveKind = 3 Then '3.繪圖人員工作進度
'               Unload frm090202_2
'            'End If
'         End If
'      Next
'   End If
'   '2017/8/30 END
End Sub

Private Sub Form_Load()
   m_ProState = ProState 'Add By Sindy 2017/8/10 記錄目前權限
   'Add By Cheng 2003/05/22
   Screen.MousePointer = vbHourglass
   MoveFormToCenter Me
   
   Me.Combo3.ListIndex = 0
   'add by nickc 2005/04/13 控制列印時統計都會算
   From090706 = False
   'Add By Cheng 2002/04/29
   Me.lblClose.Caption = ""
   'add by nickc 2006/12/29   紀錄 mail 資料，在 trans 後發
   ReDim skMail(0) As SeekMails
   
   Combo5.Text = Combo5.List(3) 'Add By Sindy 2013/9/17
   
   Select Case ProState
   Case "1" '個人(繪圖人員個人工作進度資料維護)
       Label1(32).Visible = False
       Text1.Text = Trim(Val(Mid(strSrvDate(1), 1, 6)) - 191100)
       Text1.Visible = False
       Text1.Enabled = False
       Txt1(0).Visible = False
      'add by nickc 2005/03/17 個人只能看
      Txt1(15).Enabled = False
      Txt1(16).Enabled = False
      Txt1(17).Enabled = False
      Txt1(18).Enabled = False
       TextOk = True
       StrMenu1
       cmdOK(1).Caption = "結束(&X)"
       
   Case "2" '管理
       Label1(32).Visible = False
       Text1.Text = Trim(Val(frm090706.Txt1(3)) * 100 + Val(Right(ChgNumByNick(frm090706.Txt1(4)), 2)))
       Text1.Visible = False
       Text1.Enabled = False
       'Modify By Cheng 2003/09/22
       'Begin
   '    txt1(0).Visible = False
       Txt1(0).Visible = True
       'End
       frm090706.TextOk = True
       frm090706.Process2
       StrMenu1
       cmdOK(1).Caption = "回前畫面(&U)"
       cmdOK(3).Visible = True 'Add by Morgan 2011/3/30
       
   Case "3" '分所(繪圖人員管理工作進度資料維護)
       CheckOC
         With adoRecordset
                .CursorLocation = adUseClient
                   'Modify By Cheng 2003/06/05
   '             .Open "select DISTINCT st01 from staff where st04='1' AND ST05 in ('79','81','AC') ", cnnConnection, adOpenStatic, adLockReadOnly
                   'Modify By Cheng 2003/07/16
   '             .Open "select DISTINCT st01, ST02 from staff where st04='1' AND ST05 in ('79','81','82','AC') ", cnnConnection, adOpenStatic, adLockReadOnly
                .Open "select DISTINCT st01, ST02, ST06 from staff where st04='1' AND ST05 in ('79','81','82','AC') Order By 3, 1 ", cnnConnection, adOpenStatic, adLockReadOnly
                If .RecordCount <> 0 Then
                     Combo1.Clear
                     s = 0
                     Do While .EOF = False
                           'Modify By Cheng 2003/06/05
   '                      Combo1.AddItem CheckStr(.Fields(0)), s
                         Combo1.AddItem CheckStr(.Fields(0).Value) & " " & .Fields(1).Value, s
                         s = s + 1
                         .MoveNext
                     Loop
                     Combo1.Text = Combo1.List(0)
                     TextOk = True
                 Else
                     TextOk = False
                 End If
         End With
   
       Txt1(0).Visible = True
       Txt1(0).Enabled = True
       Text1.Enabled = True
       cmdOK(1).Caption = "結束(&X)"
       
   Case "4" '繪圖人員個人工作進度資料查詢
       Label1(0).Visible = False
       Combo1.Visible = False
       lbl2.Visible = False
       Label1(32).Visible = False
       Text1.Text = frm090303_1.Text1
      'add by nickc 2005/03/17 個人只能看
      Txt1(15).Enabled = False
      Txt1(16).Enabled = False
      Txt1(17).Enabled = False
      Txt1(18).Enabled = False
       
       Text1.Visible = False
       cmdOK(0).Visible = False
       cmd(0).Visible = False
       cmd(1).Visible = False
       For ll = 0 To 13
         Txt1(ll).Enabled = False
       Next ll
       frm090303_1.Process
       Combo2.Enabled = False
       Option1(0).Enabled = False
       Option1(1).Enabled = False
       cmdOK(2).Visible = False
   Case Else
   End Select
   Select Case ProState
   Case "1"
       If TextOk = False Then Screen.MousePointer = vbDefault: GoTo EXITSUB
   Case "2"
       If frm090706.TextOk = False Then Screen.MousePointer = vbDefault: TextOk = True: GoTo EXITSUB
   Case "3"
   Case "4"
   Case Else
   End Select
   SSTab1.Tab = 0
   
   'Add By Sindy 2013/5/16
   'If Pub_StrUserSt03 = "M51" Or InStr(UCase(Pub_GetModuleFileName), "VB6.EXE") <> 0 Then
      Me.cmd(2).Visible = True '承辦歷程
   '   'Add By Sindy 2013/6/7
   '   If ProState = "1" Then '個人
         Me.SSTab1.TabVisible(2) = True '待辦歷程
         SSTab1.Tab = 2
         If QueryData(True) = False Then
            SSTab1.Tab = 0
         End If
   '   Else
   '      Me.SSTab1.TabVisible(2) = False
   '   End If
      '2013/6/7 End
   'Else
   '   Me.cmd(2).Visible = False
   '   Me.SSTab1.TabVisible(2) = False
   'End If
   '2013/5/16 End
   
   TextOk = True
   Option1(0).Value = True
   If ProState <> "4" Then
      Txt1(13).Enabled = True
   End If
   Combo2.Clear
   Combo2.AddItem "速件", 0
   Combo2.AddItem "未齊備", 1
   Combo2.AddItem "複雜", 2
   Combo2.AddItem "其他新案", 3
   Combo2.AddItem "ACAD", 4
   Combo2.Text = "請選擇...."
   Combo2.Enabled = False
   
   'Modify by Morgan 2011/1/4 觸發過就不必再重複
   'Combo1.Text = Combo1.List(0)
   'Combo1_Click
   If Combo1.ListIndex < 0 Then
      Combo1.Text = Combo1.List(0)
   End If
   Call GetMonthAssess 'Add by Amy 2018/10/08
   MouseClick (1)
   'Add By Cheng 2003/05/22
   Screen.MousePointer = vbDefault
   Exit Sub
   
EXITSUB:
   Me.Hide
   Select Case ProState
   Case "1"
        'Me.Hide
   Case "2"
       frm090706.Show
       Me.Hide
   Case "3"
       'Me.Hide
   Case "4"
       frm090303_1.Show
       Me.Hide
   Case Else
   End Select
   'Add By Cheng 2003/05/22
   Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frm090711 = Nothing
End Sub

Sub StrMenu1()      '抓繪圖人員
'Dim strCon As String 'Add By Sindy 2015/5/12
'
Select Case ProState
Case "1" '繪圖人員個人工作進度資料維護(直接抓系統年月資料)
''      StrSQL6 = " and EP13='" & strUserNum & "' and cp01 in ('FCP','P','CFP') "
'      strCon = " and EP13='" & strUserNum & "' and cp01 in ('FCP','P','CFP') " 'Add By Sindy 2015/5/12
''      StrSQL6 = StrSQL6 + " and CP27 IS NULL  and CP57 IS NULL and cp05>=19980101 "
''      strSQL1 = " and EP13='" & strUserNum & "' "
'      strCon = strCon & " and (CP27||CP57 IS NULL or ((SUBSTR(CP27,1,6)=" & Mid(strSrvDate(1), 1, 6) & ") or (SUBSTR(CP57,1,6)=" & Mid(strSrvDate(1), 1, 6) & " and cp27 is null) or (SUBSTR(CP05,1,6)=" & Mid(strSrvDate(1), 1, 6) & "))) and cp05>=19980101 " 'Add By Sindy 2015/5/12
'      'edit by nickc 2005/05/13
'      'strSQL1 = strSQL1 & " and (SUBSTR(CP27,1,6)=" & Mid(strSrvDate(1), 1, 6) & " or SUBSTR(CP57,1,6)=" & Mid(strSrvDate(1), 1, 6) & ") and cp05>=19980101 "
'      'edit by nickc 2006/01/11
'      'strSQL1 = strSQL1 & " and ((SUBSTR(CP27,1,6)=" & Mid(strSrvDate(1), 1, 6) & ") or (SUBSTR(CP57,1,6)=" & Mid(strSrvDate(1), 1, 6) & " and cp27 is null)) and cp05>=19980101 "
''      strSQL1 = strSQL1 & " and ((SUBSTR(CP27,1,6)=" & Mid(strSrvDate(1), 1, 6) & ") or (SUBSTR(CP57,1,6)=" & Mid(strSrvDate(1), 1, 6) & " and cp27 is null) or (SUBSTR(CP05,1,6)=" & Mid(strSrvDate(1), 1, 6) & ")) and cp05>=19980101 "
'      'add by nick 2004/12/20   加多國案且草圖不計件不秀
''edit by nickc 2005/03/01 墨圖也要判斷
''      strSQL1 = strSQL1 & " and ((cp21='Y' and ep20 is null) or cp21 is null) "
''      StrSQL6 = StrSQL6 & " and ((cp21='Y' and ep20 is null) or cp21 is null) "
''      strSQL1 = strSQL1 & " and ((cp21='Y' and (ep20 is null or ep29 is null)) or cp21 is null) "
''      StrSQL6 = StrSQL6 & " and ((cp21='Y' and (ep20 is null or ep29 is null) ) or cp21 is null) "
'      strCon = strCon & " and ((cp21='Y' and (ep20 is null or ep29 is null) ) or cp21 is null) " 'Add By Sindy 2015/5/12
'      'add by nickc 2005/03/21
''      strSQL1 = strSQL1 & " and cp107='Y' "
''      StrSQL6 = StrSQL6 & " and cp107='Y' "
'      strCon = strCon & " and cp107='Y' " 'Add By Sindy 2015/5/12
Case "2", "3"
'      'StrGrp090711 = ""
'      'StrSQL6 = " and ep20 is null  AND EP13='" & Trim(Combo1.Text) & "' AND CP01 IN (" & SQLGrpStr(StrGrp090711, 1) & ") "
'      'StrSQL6 = StrSQL6 + " and CP26 IS NULL  and CP27 IS NULL  and CP57 IS NULL and cp05>=19980101 "
'      'StrSQL1 = " and ep20 is null  AND EP13='" & Trim(Combo1.Text) & "' "
'      'StrSQL1 = StrSQL1 & " and CP26 IS NULL  and (SUBSTR(CP27,1,6)=" & Mid(GetTodayDate, 1, 6) & " or SUBSTR(CP57,1,6)=" & Mid(GetTodayDate, 1, 6) & ") and cp05>=19980101 "
      Exit Sub
Case Else
End Select
''Modify By Cheng 2003/05/22
''strSQL = "SELECT DISTINCT EP13 FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,PATENT,CASEPROPERTYMAP,PATENTTRADEMARKMAP WHERE EP02=CP09(+) AND cp01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND cp14=S1.ST01(+) AND eP13=S2.ST01(+) AND cP13=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND '1'=PTM01(+)  AND PA08=PTM02(+) " & StrSQL6
''strSQL = strSQL & " UNION all  SELECT DISTINCT EP13 FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,PATENT,CASEPROPERTYMAP,PATENTTRADEMARKMAP WHERE EP02=CP09(+) AND cp01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND cp14=S1.ST01(+) AND eP13=S2.ST01(+) AND cP13=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND '1'=PTM01(+)  AND PA08=PTM02(+) " & strSQL1
'
''Modify By Sindy 2015/5/12 調整SQL增加查詢速度
''strSql = "SELECT DISTINCT EP13 FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,PATENT,CASEPROPERTYMAP,PATENTTRADEMARKMAP WHERE EP02=CP09(+) AND cp01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND cp14=S1.ST01(+) AND eP13=S2.ST01(+) AND cP13=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND '1'=PTM01(+)  AND PA08=PTM02(+) " & StrSQL6
''strSql = strSql & " UNION SELECT DISTINCT EP13 FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,PATENT,CASEPROPERTYMAP,PATENTTRADEMARKMAP WHERE EP02=CP09(+) AND cp01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND cp14=S1.ST01(+) AND eP13=S2.ST01(+) AND cP13=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND '1'=PTM01(+)  AND PA08=PTM02(+) " & strSQL1
''strSql = strSql + " ORDER BY 1 "
'
''strSql = "SELECT DISTINCT EP13 FROM CASEPROGRESS,ENGINEERPROGRESS WHERE EP02=CP09(+) " & StrSQL6
''strSql = strSql & " UNION SELECT DISTINCT EP13 FROM CASEPROGRESS,ENGINEERPROGRESS WHERE EP02=CP09(+) " & strSQL1
''strSql = strSql + " ORDER BY 1"
'
'strSql = "SELECT DISTINCT EP13 FROM CASEPROGRESS,ENGINEERPROGRESS WHERE EP02=CP09(+) " & strCon
'strSql = strSql + " ORDER BY 1"
''2015/5/12 END
'
''Select Case ProState
''Case "1"
''      strSQL = "SELECT DISTINCT EP13 FROM ENGINEERPROGRESS where ep13='" & strUserNum & "' "
''Case "2"
''      strSQL = "SELECT DISTINCT EP13 FROM ENGINEERPROGRESS "
''Case Else
''End Select
'CheckOC
'j = 0
'Combo1.Clear
'With adoRecordset
'    .CursorLocation = adUseClient
'    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
'    If .RecordCount <> 0 And .RecordCount > 0 Then
'        .MoveFirst
'        Do While .EOF = False
'            'Modify By Cheng 2003/06/05
''            Combo1.AddItem CheckStr(.Fields(0)), j
'            Combo1.AddItem CheckStr(.Fields(0).Value) & " " & GetStaffName("" & .Fields(0).Value, True), j
'            j = j + 1
'            .MoveNext
'        Loop
'    End If
'End With
'Modify By Sindy 2016/5/9 直接帶入操作人員
Combo1.Clear
Combo1.AddItem CheckStr(strUserNum) & " " & GetStaffName(strUserNum, True), 0
'2016/5/9 END
CheckOC
'Add By Sindy 2013/9/16 檢查當時是否需要為他人職代
If ProState = "1" Then
   Call Pub_SetForOthersEmpCombo(strUserNum, Combo1, False)
End If
'2013/9/16 END
Combo1.Text = Combo1.List(0)
End Sub

Sub StrMenu5()  '代資料  未發文
If Len(Text1) = 0 Then
    grd1.Clear
    grd1.Rows = 2
    SetGrd1
    Text1.Text = Mid(Val(strSrvDate(1)), 1, 6) - 191100
End If
Select Case ProState
Case "1"
      '  薛說不管是否計件都要出來   91/04/10
      'StrSQL6 = " and ep20 is null  AND EP13='" & strUserNum & "' and cp01 in ('FCP','P','CFP')  "
      'Modify By Sindy 2016/9/7
'      StrSQL6 = " AND EP13='" & Trim(Left(Combo1.Text, 6)) & "' and cp01 in ('FCP','P','CFP')  "
'      StrSQL6 = StrSQL6 + " and cp57 is null and cp27 is null and cp05>=19980101 "
      StrSQL6 = " AND cp29='" & Trim(Left(Combo1.Text, 6)) & "' and cp01 in ('FCP','P','CFP')  "
      StrSQL6 = StrSQL6 + " and cp158=0 and cp159=0 and cp05>=19980101 "
      '2016/9/7 END
      'add by nick 2004/12/20   加多國案且草圖不計件不秀
      'edit by nickc 2005/03/01 墨圖也要判斷
      'StrSQL6 = StrSQL6 & " and ((cp21='Y' and ep20 is null) or cp21 is null) "
      StrSQL6 = StrSQL6 & " and ((cp21='Y' and (ep20 is null or ep29 is null)) or cp21 is null) "
      'add by nickc 2005/03/21
      StrSQL6 = StrSQL6 & " and cp107='Y' "
Case "2", "3"
      StrSQL6 = ""
      StrGrp090711 = ""
      '  薛說不管是否計件都要出來   91/04/10
      'StrSQL6 = " and ep20 is null  AND EP13='" & Trim(Combo1.Text) & "' AND CP01 IN (" & SQLGrpStr(StrGrp090711, 1) & ") " & StrSQL6
      'Modify By Sindy 2016/9/7
'      StrSQL6 = " AND EP13='" & Trim(Left(Combo1.Text, 6)) & "' AND CP01 IN (" & SQLGrpStr(StrGrp090711, 1) & ") " & StrSQL6
'      StrSQL6 = StrSQL6 + " and cp57 is null and cp27 is null and cp05>=19980101 "
      StrSQL6 = " AND cp29='" & Trim(Left(Combo1.Text, 6)) & "' AND CP01 IN (" & SQLGrpStr(StrGrp090711, 1) & ") " & StrSQL6
      StrSQL6 = StrSQL6 + " and cp158=0 and cp159=0 and cp05>=19980101 "
      '2016/9/7 END
      'add by nick 2004/12/20   加多國案且草圖不計件不秀
      'edit by nickc 2005/03/01
      'StrSQL6 = StrSQL6 & " and ((cp21='Y' and ep20 is null) or cp21 is null) "
      StrSQL6 = StrSQL6 & " and ((cp21='Y' and (ep20 is null or ep29 is null)) or cp21 is null) "
      'add by nickc 2005/03/21
      StrSQL6 = StrSQL6 & " and cp107='Y' "
Case Else
End Select

'Modify By Cheng 2002/04/16
'strSQL = "SELECT SUBSTR(CP09,1,1)," & SQLDate("CP05") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04,NVL(PA05,NVL(PA06,PA07)),eP20,decode(pa09,'000',cpm03,cpm04),s1.st02," & SQLDate("CP48") & ",ROUND(cp18,2),DECODE(EP14,NULL,DECODE(EP06,NULL,'',0,''," & SQLDate("EP06") & "),0,DECODE(EP06,NULL,'',0,''," & SQLDate("EP06") & ")," & SQLDate("eP14") & ")," & SQLDate("eP15") & ",0,DECODE(EP17,NULL,DECODE(EP08,NULL,'',0,''," & SQLDate("EP08") & "),0,DECODE(EP08,NULL,'',0,''," & SQLDate("EP08") & ")," & SQLDate("eP17") & ")," & SQLDate("EP18") & ",0," & SQLDate("cP06") & "," & SQLDate("cP27") & ",ep26,s3.st02,CP09,DECODE(PA09,'000',PTM03,PTM04)," & SQLDate("CP07") & "," & SQLDate("CP57") & ",ep21,ep22,ep23,ep24,ep25,ep16,ep19,PA08 FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,PATENT,CASEPROPERTYMAP,PATENTTRADEMARKMAP " & _
'         " WHERE EP02=CP09(+) AND cp01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND cp14=S1.ST01(+) AND eP13=S2.ST01(+) AND cP13=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND '1'=PTM01(+)  AND PA08=PTM02(+) " & StrSQL6
'strSQL = strSQL + " ORDER BY 1,2 "

'Modify By Cheng 2002/04/26
'若已閉卷, 則在本所案號後加"*"號
'strSQL = "SELECT SUBSTR(CP09,1,1)," & SQLDate("CP05") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(PA57,'Y','＊',''),NVL(PA05,NVL(PA06,PA07)),eP20,decode(pa09,'000',cpm03,cpm04),s1.st02," & SQLDate("CP48") & ",ROUND(cp18,2)," & SQLDate("EP06") & ",DECODE(EP14,NULL,'',0,''," & SQLDate("eP14") & ")," & SQLDate("eP15") & ",0,DECODE(EP17,NULL,DECODE(EP08,NULL,'',0,''," & SQLDate("EP08") & "),0,DECODE(EP08,NULL,'',0,''," & SQLDate("EP08") & ")," & SQLDate("eP17") & ")," & SQLDate("EP18") & ",0," & SQLDate("cP06") & "," & SQLDate("cP27") & ",ep26,s3.st02,CP09,DECODE(PA09,'000',PTM03,PTM04)," & SQLDate("CP07") & "," & SQLDate("CP57") & ",ep21,ep22,ep23,ep24,ep25,ep16,ep19,PA08 FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,PATENT,CASEPROPERTYMAP,PATENTTRADEMARKMAP " & _
'         " WHERE EP02=CP09(+) AND cp01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND cp14=S1.ST01(+) AND eP13=S2.ST01(+) AND cP13=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND '1'=PTM01(+)  AND PA08=PTM02(+) " & StrSQL6
'92.5.13 modify by sonia
'Modify By Cheng 2004/02/18
'strSQL = "SELECT SUBSTR(CP09,1,1)," & SQLDate("CP05") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(PA57,'Y','＊',''),NVL(PA05,NVL(PA06,PA07)), decode(pa09,'000',cpm03,cpm04), s1.st02, ROUND(cp18,2), DECODE(EP06,NULL,' ',0,' '," & SQLDate("EP06") & "), EP20, DECODE(EP14,NULL,'',0,''," & SQLDate("eP14") & "), '' As 草期限, " & SQLDate("eP15") & ",0, EP29, DECODE(EP17,NULL,DECODE(EP08,NULL,'',0,''," & SQLDate("EP08") & "),0, DECODE(EP08,NULL,'',0,''," & SQLDate("EP08") & ")," & SQLDate("eP17") & "), '' As 墨期限, " & SQLDate("EP18") & ",0," & SQLDate("cP06") & "," & SQLDate("cP27") & ",ep26,s3.st02,CP09,DECODE(PA09,'000',PTM03,PTM04)," & SQLDate("CP07") & "," & SQLDate("CP57") & ",ep21,ep22,ep23,ep24,ep25,ep16,ep19, CP10 FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,PATENT,CASEPROPERTYMAP,PATENTTRADEMARKMAP " & _
'         " WHERE EP02=CP09(+) AND cp01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND cp14=S1.ST01(+) AND eP13=S2.ST01(+) AND cP13=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND '1'=PTM01(+)  AND PA08=PTM02(+) " & StrSQL6
'Modify By Cheng 2004/03/08
'顯示墨齊日欄位時直接顯示墨齊日欄位(EP17)
'strSQL = "SELECT SUBSTR(CP09,1,1)," & SQLDate("CP05") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(PA57,'Y','＊',''),NVL(PA05,NVL(PA06,PA07)), decode(pa09,'000',cpm03,cpm04), s1.st02, ROUND(cp18,2), DECODE(EP06,NULL,' ',0,' '," & SQLDate("EP06") & "), EP20, DECODE(EP14,NULL,'',0,''," & SQLDate("eP14") & "), '' As 草期限, " & SQLDate("eP15") & ",0, EP29, DECODE(EP17,NULL,DECODE(EP08,NULL,'',0,''," & SQLDate("EP08") & "),0, DECODE(EP08,NULL,'',0,''," & SQLDate("EP08") & ")," & SQLDate("eP17") & "), '' As 墨期限, " & SQLDate("EP18") & ",0," & SQLDate("cP06") & "," & SQLDate("cP27") & ",ep26,s3.st02,CP09,DECODE(PA09,'000',PTM03,PTM04)," & SQLDate("CP07") & "," & SQLDate("CP57") & ",ep21,ep22,ep23,ep24,ep25,ep16,ep19, CP10 FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,PATENT,CASEPROPERTYMAP,PATENTTRADEMARKMAP " & _
'         " WHERE EP02=CP09(+) AND cp01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND cp14=S1.ST01(+) AND eP13=S2.ST01(+) AND cP13=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND '1'=PTM01(+)  AND PA08=PTM02(+) " & StrSQL6
'edit by nickc 2005/03/17 修正
'StrSql = "SELECT SUBSTR(CP09,1,1)," & SQLDate("CP05") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(PA57,'Y','＊',''),NVL(PA05,NVL(PA06,PA07)), decode(pa09,'000',cpm03,cpm04), s1.st02, ROUND(cp18,2), DECODE(EP06,NULL,' ',0,' '," & SQLDate("EP06") & "), EP20, DECODE(EP14,NULL,'',0,''," & SQLDate("eP14") & "), '' As 草期限, " & SQLDate("eP15") & ",0, EP29, DECODE(EP17,NULL,'' ," & SQLDate("eP17") & "), '' As 墨期限, " & SQLDate("EP18") & ",0," & SQLDate("cP06") & "," & SQLDate("cP27") & ",ep26,s3.st02,CP09,DECODE(PA09,'000',PTM03,PTM04)," & SQLDate("CP07") & "," & SQLDate("CP57") & ",ep21,ep22,ep23,ep24,ep25,ep16,ep19, CP10 FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,PATENT,CASEPROPERTYMAP,PATENTTRADEMARKMAP " & _
         " WHERE EP02=CP09(+) AND cp01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND cp14=S1.ST01(+) AND eP13=S2.ST01(+) AND cP13=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND '1'=PTM01(+)  AND PA08=PTM02(+) " & StrSQL6
'edit by nickc 2006/01/16
'strSQL = "SELECT SUBSTR(CP09,1,1)," & SQLDate("CP05") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(PA57,'Y','＊',''),NVL(PA05,NVL(PA06,PA07)), decode(pa09,'000',cpm03,cpm04), s1.st02, ROUND(cp18,2), DECODE(EP06,NULL,' ',0,' '," & SQLDate("EP06") & "), EP20, DECODE(EP14,NULL,'',0,''," & SQLDate("eP14") & "), '' As 草期限, " & SQLDate("eP15") & ",0, EP29, DECODE(EP17,NULL,'' ," & SQLDate("eP17") & "), '' As 墨期限, " & SQLDate("EP18") & ",0," & SQLDate("cP06") & "," & SQLDate("cP27") & ",ep26,s3.st02,CP09,DECODE(PA09,'000',PTM03,PTM04)," & SQLDate("CP07") & "," & SQLDate("CP57") & ",ep21,ep22,ep23,ep24,ep25,ep16,ep19, CP10, PA08,decode(pa09,'000','9999','020','8888','101','7777','0000') as PaSort FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,PATENT,CASEPROPERTYMAP,PATENTTRADEMARKMAP " & _
         " WHERE EP02=CP09(+) AND cp01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND cp14=S1.ST01(+) AND eP13=S2.ST01(+) AND cP13=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND '1'=PTM01(+)  AND PA08=PTM02(+) " & StrSQL6
'Modify By Sindy 2016/3/1 少了申國欄位
strSql = "SELECT distinct SUBSTR(CP09,1,1)," & SQLDate("CP05") & ",substrb(na03,1,4),CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(PA57,'Y','＊',''),NVL(PA05,NVL(PA06,PA07)), decode(pa09,'000',cpm03,cpm04), s1.st02, ROUND(cp18,2), DECODE(EP06,NULL,' ',0,' '," & SQLDate("EP06") & "), EP20, DECODE(EP14,NULL,'',0,''," & SQLDate("eP14") & "), '' As 草期限, " & SQLDate("eP15") & ",0, EP29, DECODE(EP17,NULL,'' ," & SQLDate("eP17") & "), '' As 墨期限, " & SQLDate("EP18") & ",0," & SQLDate("cP06") & "," & SQLDate("cP27") & ",ep26,s3.st02,CP09,decode(pa01,'CFP',ptm03,DECODE(PA09,'000',PTM03,PTM04))," & SQLDate("CP07") & "," & SQLDate("CP57") & ",ep21,ep22,ep23,ep24,ep25,ep16,ep19, CP10, PA08,decode(pa09,'000','9999','020','8888','101','7777','0000') as PaSort FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,PATENT,CASEPROPERTYMAP,PATENTTRADEMARKMAP,nation " & _
         " WHERE EP02(+)=CP09 AND cp01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND cp14=S1.ST01(+) AND eP13=S2.ST01(+) AND cP13=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND '1'=PTM01(+) AND PA08=PTM02(+) and pa09=na01(+) " & StrSQL6
         
'End
'End
'strSQL = strSQL + " ORDER BY 1,2 "
'strSQL = strSQL + " ORDER BY 8 desc, 3 desc "
'排序方式文齊日(大至小且Null的排在最底下), 墨齊日(大至小), 草齊日(大至小), 本所案號(小至大)
'strSQL = strSQL + " ORDER BY 8 desc, 3 desc "
'edit by nickc 2005/03/17 修正
'StrSql = StrSql + " ORDER BY 8 Desc, 16 Desc, 10 Desc, 3 Asc "
strSql = strSql + " ORDER BY 9 Desc,pasort desc, 17 Desc, 11 Desc, 4 Asc "
'End
'92.5.13 end

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
            Me.grd1.TextMatrix(i, 8) = Trim(Me.grd1.TextMatrix(i, 8)) '7
            '草完日
            grd1.col = 12 '11
            strDate1 = grd1.Text
            '草齊日
            grd1.col = 10 '9
            StrDate2 = grd1.Text
            If Trim(strDate1) <> "" And Trim(StrDate2) <> "" Then
                '草天
                grd1.col = 13 '12
                grd1.Text = GetWorkDay(ChangeTStringToWString(ChangeTDateStringToTString(strDate1)), ChangeTStringToWString(ChangeTDateStringToTString(StrDate2)))
            Else
                '草天
                grd1.col = 13 '12
                grd1.Text = ""
            End If
            '墨完日
            grd1.col = 17 '16
            strDate1 = grd1.Text
            '墨齊日
            grd1.col = 15 '14
            StrDate2 = grd1.Text
            If Trim(strDate1) <> "" And Trim(StrDate2) <> "" Then
                '墨天
                grd1.col = 18 '17
                grd1.Text = GetWorkDay(ChangeTStringToWString(ChangeTDateStringToTString(strDate1)), ChangeTStringToWString(ChangeTDateStringToTString(StrDate2)))
            Else
                '墨天
                grd1.col = 18 '17
                grd1.Text = ""
            End If
            grd1.Text = grd1.Text & GetSign(grd1.TextMatrix(i, 23)) 'Add by Morgan 2009/10/5 22
            
            'Add By Cheng 2003/06/27
            '若草圖不計件
            If grd1.TextMatrix(i, 9) = "N" Then '8
                grd1.TextMatrix(i, 10) = "******" '9
                grd1.TextMatrix(i, 12) = "******" '11
                grd1.TextMatrix(i, 13) = "" '12
            End If
            '若墨圖不計件
            If grd1.TextMatrix(i, 14) = "N" Then '13
                grd1.TextMatrix(i, 15) = "******" '14
                grd1.TextMatrix(i, 17) = "******" '16
                grd1.TextMatrix(i, 18) = "" '17
            End If
            'Modify By Cheng 2003/09/18
            '不計算草圖及墨圖承辦期限
            'Begin
'            'Add By Cheng 2003/06/30
'            '草圖承辦期限
'            If Me.grd1.TextMatrix(i, 9) <> "" And Me.grd1.TextMatrix(i, 9) <> "******" Then
'                '設計申請
'                If Me.grd1.TextMatrix(i, 33) = "103" Or Me.grd1.TextMatrix(i, 33) = "105" Then
'                    Me.grd1.TextMatrix(i, 10) = ChangeTStringToTDateString(CompWorkDay(5, Replace(Me.grd1.TextMatrix(i, 9), "/", "") + 19110000) - 19110000)
'                '非設計申請
'                Else
'                    Me.grd1.TextMatrix(i, 10) = ChangeTStringToTDateString(CompWorkDay(4, Replace(Me.grd1.TextMatrix(i, 9), "/", "") + 19110000) - 19110000)
'                End If
'            End If
'            '墨圖承辦期限
'            If Me.grd1.TextMatrix(i, 14) <> "" And Me.grd1.TextMatrix(i, 14) <> "******" Then
'                Me.grd1.TextMatrix(i, 15) = ChangeTStringToTDateString(CompWorkDay(3, Replace(Me.grd1.TextMatrix(i, 14), "/", "") + 19110000) - 19110000)
'            End If
            'End
        Next i
        Me.Enabled = True
        Screen.MousePointer = vbDefault
        SetGrd1
        grd1.Visible = True
        ChkNoData = False
    Else
         ChkNoData = True
    End If
End With
CheckOC
End Sub

Sub StrMenu()       '代資料  當月資料
Dim strVBA As String 'Add By Sindy 2015/5/12

If Len(Text1) = 0 Then
    grd1.Clear
    grd1.Rows = 2
    SetGrd1
    Text1.Text = Mid(Val(strSrvDate(1)), 1, 6) - 191100
End If

'Modified by Morgan 2013/10/3 調整語法,改以CP為主
Select Case ProState
'Modify By Cheng 2002/04/18
'Case "1", "4"
Case "1" '繪圖人員作業--維護
      '  薛說不管是否計件都要出來   91/04/10
      'StrSQL6 = " and ep20 is null  AND EP13='" & strUserNum & "' and cp01 in ('FCP','P','CFP')  "
      'Modify By Sindy 2015/5/12 cp29==>ep13
      'StrSQL6 = "  AND cp29='" & Trim(Left(Combo1.Text, 6)) & "' and cp01 in ('FCP','P','CFP')  "
      'Modify By Sindy 2016/9/7
      'StrSQL6 = " AND ep13='" & Trim(Left(Combo1.Text, 6)) & "' and cp01 in ('FCP','P','CFP') "
      StrSQL6 = " AND cp29='" & Trim(Left(Combo1.Text, 6)) & "' and cp01 in ('FCP','P','CFP') "
      '2016/9/7 END
      '2015/5/12 END
      
      'Modify By Sindy 2016/9/7
      'StrSQL6 = StrSQL6 + " and cp57 is null and cp27 is null and cp05>=19980101 "
      StrSQL6 = StrSQL6 + " and cp158=0 and cp159=0 and cp05>=19980101 "
      '2016/9/7 END
      'strSQL1 = " and ep20 is null  AND EP13='" & strUserNum & "' and cp01 in ('FCP','P','CFP')  "
      'Modify By Sindy 2015/5/12 cp29==>ep13
      'strSQL1 = " AND cp29='" & Trim(Left(Combo1.Text, 6)) & "' and cp01 in ('FCP','P','CFP')  "
      'Modify By Sindy 2016/9/7
      'strSQL1 = " AND ep13='" & Trim(Left(Combo1.Text, 6)) & "' and cp01 in ('FCP','P','CFP') "
      strSQL1 = " AND cp29='" & Trim(Left(Combo1.Text, 6)) & "' and cp01 in ('FCP','P','CFP') "
      '2016/9/7 END
      '2015/5/12 END
      
      'edit by nickc 2005/05/13
      'strSQL1 = strSQL1 & " and (SUBSTR(CP27,1,6)=" & Mid(strSrvDate(1), 1, 6) & " or SUBSTR(CP57,1,6)=" & Mid(strSrvDate(1), 1, 6) & ") and cp05>=19980101 "
      'edit by nickc 2005/01/11
      'strSQL1 = strSQL1 & " and ((SUBSTR(CP27,1,6)=" & Mid(strSrvDate(1), 1, 6) & ") or (SUBSTR(CP57,1,6)=" & Mid(strSrvDate(1), 1, 6) & " and cp27 is null)) and cp05>=19980101 "
      'Modify By Sindy 2016/5/10
      'strSQL1 = strSQL1 & " and ((SUBSTR(CP27,1,6)=" & Mid(strSrvDate(1), 1, 6) & ") or (SUBSTR(CP57,1,6)=" & Mid(strSrvDate(1), 1, 6) & " and cp27 is null) or (SUBSTR(CP05,1,6)=" & Mid(strSrvDate(1), 1, 6) & ")) and cp05>=19980101 "
      strSQL1 = strSQL1 & " and ((CP27>=" & Mid(strSrvDate(1), 1, 6) & "01 and CP27<=" & Mid(strSrvDate(1), 1, 6) & "31) or (CP57>=" & Mid(strSrvDate(1), 1, 6) & "01 and CP57<=" & Mid(strSrvDate(1), 1, 6) & "31 and cp27 is null) or (CP05>=" & Mid(strSrvDate(1), 1, 6) & "01 and CP05<=" & Mid(strSrvDate(1), 1, 6) & "31)) and cp05>=19980101 "
      '2016/5/10 END
      'add by nick 2004/12/20   加多國案且草圖不計件不秀
'edit by nickc 2005/03/01  墨圖也要判斷
'      strSQL1 = strSQL1 & " and ((cp21='Y' and ep20 is null) or cp21 is null) "
'      StrSQL6 = StrSQL6 & " and ((cp21='Y' and ep20 is null) or cp21 is null) "
      strSQL1 = strSQL1 & " and ((cp21='Y' and (ep20 is null or ep29 is null)) or cp21 is null) "
      StrSQL6 = StrSQL6 & " and ((cp21='Y' and (ep20 is null or ep29 is null) ) or cp21 is null) "
      'add by nickc 2005/03/21
      strSQL1 = strSQL1 & " and  cp107='Y' "
      StrSQL6 = StrSQL6 & " and cp107='Y' "
      
'Add By Cheng 2002/04/18
'以發文年月查詢, 搜尋發文年月等於查詢之年月
Case "4" '繪圖人員作業--查詢
      'Modify By Sindy 2015/5/12 cp29==>ep13
      'StrSQL6 = " AND cp29='" & Trim(Left(Combo1.Text, 6)) & "' and cp01 in ('FCP','P','CFP')  "
      'Modify By Sindy 2016/9/7
      'StrSQL6 = " AND ep13='" & Trim(Left(Combo1.Text, 6)) & "' and cp01 in ('FCP','P','CFP') "
      StrSQL6 = " AND cp29='" & Trim(Left(Combo1.Text, 6)) & "' and cp01 in ('FCP','P','CFP') "
      '2016/9/7 END
      '2015/5/12 END
      
'      StrSQL6 = StrSQL6 + " AND CP26 IS NULL and CP27 IS NULL and CP57 IS NULL and cp05>=19980101 "
      'edit by nickc 2006/01/11
      'StrSQL6 = StrSQL6 & " AND SUBSTR(CP27,1,6)=" & Val(frm090303_1.Text1.Text) + 191100 & " AND cp05>=19980101 "
      'Modify By Sindy 2016/5/10
      'StrSQL6 = StrSQL6 & " AND (SUBSTR(CP27,1,6)=" & Val(frm090303_1.Text1.Text) + 191100 & " or SUBSTR(CP05,1,6)=" & Val(frm090303_1.Text1.Text) + 191100 & ") AND cp05>=19980101 "
      StrSQL6 = StrSQL6 & " AND ((CP27>=" & Val(frm090303_1.Text1.Text) + 191100 & "01 and CP27<=" & Val(frm090303_1.Text1.Text) + 191100 & "31) or (CP05>=" & Val(frm090303_1.Text1.Text) + 191100 & "01 and CP05<=" & Val(frm090303_1.Text1.Text) + 191100 & "31)) AND cp05>=19980101 "
      '2016/5/10 END
      'Modify By Sindy 2015/5/12 cp29==>ep13
      'strSQL1 = " AND cp29='" & Trim(Left(Combo1.Text, 6)) & "' and cp01 in ('FCP','P','CFP')  "
      'Modify By Sindy 2016/9/7
      strSQL1 = " AND cp29='" & Trim(Left(Combo1.Text, 6)) & "' and cp01 in ('FCP','P','CFP') "
      '2016/9/7 END
      '2015/5/12 END
      
      'edit by nickc 2006/01/11
      'strSQL1 = strSQL1 & " AND SUBSTR(CP27,1,6)=" & Val(frm090303_1.Text1.Text) + 191100 & " AND cp05>=19980101 "
      'Modify By Sindy 2016/5/10
      'strSQL1 = strSQL1 & " AND (SUBSTR(CP27,1,6)=" & Val(frm090303_1.Text1.Text) + 191100 & " or SUBSTR(CP05,1,6)=" & Val(frm090303_1.Text1.Text) + 191100 & ") AND cp05>=19980101 "
      strSQL1 = strSQL1 & " AND ((CP27>=" & Val(frm090303_1.Text1.Text) + 191100 & "01 and CP27<=" & Val(frm090303_1.Text1.Text) + 191100 & "31) or (CP05>=" & Val(frm090303_1.Text1.Text) + 191100 & "01 and CP05<=" & Val(frm090303_1.Text1.Text) + 191100 & "31)) AND cp05>=19980101 "
      '2016/5/10 END
      'add by nick 2004/12/20   加多國案且草圖不計件不秀
'edit by nickc 2005/03/01 墨圖也要判斷
'      strSQL1 = strSQL1 & " and ((cp21='Y' and ep20 is null) or cp21 is null) "
'      StrSQL6 = StrSQL6 & " and ((cp21='Y' and ep20 is null) or cp21 is null) "
      strSQL1 = strSQL1 & " and ((cp21='Y' and (ep20 is null or ep29 is null)) or cp21 is null) "
      StrSQL6 = StrSQL6 & " and ((cp21='Y' and (ep20 is null or ep29 is null)) or cp21 is null) "
      'add by nickc 2005/03/21
      strSQL1 = strSQL1 & " and  cp107='Y' "
      StrSQL6 = StrSQL6 & " and cp107='Y' "
'Add By Cheng 2002/04/22
Case "2" '繪圖人員工作管理--查詢
      strSQL1 = " AND CP05<=" & Val(Me.Text1.Text) + 191100 & "31 "
      StrSQL6 = " AND CP05<=" & Val(Me.Text1.Text) + 191100 & "31 "
      StrGrp090711 = ""
      'Modify By Sindy 2015/5/12 cp29==>ep13
      'strSQL1 = strSQL1 & " AND cp29='" & Trim(Left(Combo1.Text, 6)) & "' AND CP01 IN (" & SQLGrpStr(StrGrp090711, 1) & ") " & strSQL1
      'Modify By Sindy 2016/9/7
      'strSQL1 = strSQL1 & " AND ep13='" & Trim(Left(Combo1.Text, 6)) & "' AND CP01 IN (" & SQLGrpStr(StrGrp090711, 1) & ") " & strSQL1
      strSQL1 = strSQL1 & " AND cp29='" & Trim(Left(Combo1.Text, 6)) & "' AND CP01 IN (" & SQLGrpStr(StrGrp090711, 1) & ") " & strSQL1
      '2016/9/7 END
      '2015/5/12 END
      
      'edit by nickc 2005/05/13
      'strSQL1 = strSQL1 & " and ((SUBSTR(CP27,1,6)=" & (Val(Text1.Text) + 191100) & " AND CP57 IS NULL ) OR ( CP27 IS NULL AND SUBSTR(CP57,1,6)=" & (Val(Text1.Text) + 191100) & ")) and cp05>=19980101 "
      'edit by nickc 2006/01/11
      'strSQL1 = strSQL1 & " and ((SUBSTR(CP27,1,6)=" & (Val(Text1.Text) + 191100) & " ) OR ( CP27 IS NULL AND SUBSTR(CP57,1,6)=" & (Val(Text1.Text) + 191100) & ")) and cp05>=19980101 "
      'Modify By Sindy 2016/5/10
      'strSQL1 = strSQL1 & " and ((SUBSTR(CP27,1,6)=" & (Val(Text1.Text) + 191100) & " ) OR ( CP27 IS NULL AND SUBSTR(CP57,1,6)=" & (Val(Text1.Text) + 191100) & ") or (SUBSTR(CP05,1,6)=" & (Val(Text1.Text) + 191100) & " )) and cp05>=19980101 "
      strSQL1 = strSQL1 & " and ((CP27>=" & (Val(Text1.Text) + 191100) & "01 and CP27<=" & (Val(Text1.Text) + 191100) & "31) OR (cp27 is null AND (CP57>=" & (Val(Text1.Text) + 191100) & "01 and CP57<=" & (Val(Text1.Text) + 191100) & "31)) or (CP05>=" & (Val(Text1.Text) + 191100) & "01 and CP05<=" & (Val(Text1.Text) + 191100) & "31)) and cp05>=19980101 "
      '2016/5/10 END
      'Modify By Sindy 2015/5/12 cp29==>ep13
      'StrSQL6 = StrSQL6 & " AND cp29='" & Trim(Left(Combo1.Text, 6)) & "' AND CP01 IN (" & SQLGrpStr(StrGrp090711, 1) & ") " & StrSQL6
      'Modify By Sindy 2016/9/7
      'StrSQL6 = StrSQL6 & " AND ep13='" & Trim(Left(Combo1.Text, 6)) & "' AND CP01 IN (" & SQLGrpStr(StrGrp090711, 1) & ") " & StrSQL6
      StrSQL6 = StrSQL6 & " AND cp29='" & Trim(Left(Combo1.Text, 6)) & "' AND CP01 IN (" & SQLGrpStr(StrGrp090711, 1) & ") " & StrSQL6
      '2016/9/7 END
      '2015/5/12 END
      'Modify By Sindy 2016/9/7
      'StrSQL6 = StrSQL6 & " and cp57 is null and cp27 is null and cp05>=19980101 "
      StrSQL6 = StrSQL6 & " and cp158=0 and cp159=0 and cp05>=19980101 "
      '2016/9/7 END
      'add by nick 2004/12/20   加多國案且草圖不計件不秀
'edit by nickc 2005/03/01 墨圖也要判斷
'      strSQL1 = strSQL1 & " and ((cp21='Y' and ep20 is null) or cp21 is null) "
'      StrSQL6 = StrSQL6 & " and ((cp21='Y' and ep20 is null) or cp21 is null) "
      strSQL1 = strSQL1 & " and ((cp21='Y' and (ep20 is null or ep29 is null)) or cp21 is null) "
      StrSQL6 = StrSQL6 & " and ((cp21='Y' and (ep20 is null or ep29 is null)) or cp21 is null) "
      'add by nickc 2005/03/21
      strSQL1 = strSQL1 & " and  cp107='Y' "
      StrSQL6 = StrSQL6 & " and cp107='Y' "
'Case "2", "3"
Case "3" '繪圖人員工作管理--維護
      strSQL1 = ""
      StrSQL6 = ""
      StrGrp090711 = ""
      'Add By Cheng 2002/04/26
      '加判定--發文年月不能大於欲查詢的發文年月
      strSQL1 = strSQL1 & " AND CP05<=" & Val(Me.Text1.Text) + 191100 & "31 "
      StrSQL6 = StrSQL6 & " AND CP05<=" & Val(Me.Text1.Text) + 191100 & "31 "
      '  薛說不管是否計件都要出來   91/04/10
      'strSQL1 = " and ep20 is null  AND EP13='" & Trim(Combo1.Text) & "' AND CP01 IN (" & SQLGrpStr(StrGrp090711, 1) & ") " & StrSQL6
      'Modify By Sindy 2015/5/12 cp29==>ep13
      'strSQL1 = strSQL1 & " AND cp29='" & Trim(Left(Combo1.Text, 6)) & "' AND CP01 IN (" & SQLGrpStr(StrGrp090711, 1) & ") " & StrSQL6
      'Modify By Sindy 2016/9/7
      'strSQL1 = strSQL1 & " AND ep13='" & Trim(Left(Combo1.Text, 6)) & "' AND CP01 IN (" & SQLGrpStr(StrGrp090711, 1) & ") " & StrSQL6
      strSQL1 = strSQL1 & " AND cp29='" & Trim(Left(Combo1.Text, 6)) & "' AND CP01 IN (" & SQLGrpStr(StrGrp090711, 1) & ") " & StrSQL6
      '2016/9/7 END
      '2015/5/12 END
      
      'edit by nickc 2005/05/13
      'strSQL1 = strSQL1 & " and (SUBSTR(CP27,1,6)=" & (Val(Text1.Text) + 191100) & " or SUBSTR(CP57,1,6)=" & (Val(Text1.Text) + 191100) & ") and cp05>=19980101 "
      'edit by nickc 2006/01/11
      'strSQL1 = strSQL1 & " and ((SUBSTR(CP27,1,6)=" & (Val(Text1.Text) + 191100) & ") or (SUBSTR(CP57,1,6)=" & (Val(Text1.Text) + 191100) & " and cp27 is null)) and cp05>=19980101 "
      'Modify By Sindy 2016/5/10
      strSQL1 = strSQL1 & " and ((CP27>=" & (Val(Text1.Text) + 191100) & "01 and CP27<=" & (Val(Text1.Text) + 191100) & "31) or (CP57>=" & (Val(Text1.Text) + 191100) & "01 and CP57<=" & (Val(Text1.Text) + 191100) & "31 and cp27 is null) or (CP05>=" & (Val(Text1.Text) + 191100) & "01 and CP05<=" & (Val(Text1.Text) + 191100) & "31)) and cp05>=19980101 "
      '2016/5/10 END
      'StrSQL6 = " and ep20 is null  AND EP13='" & Trim(Combo1.Text) & "' AND CP01 IN (" & SQLGrpStr(StrGrp090711, 1) & ") " & StrSQL6
      'Modify By Sindy 2015/5/12 cp29==>ep13
      'StrSQL6 = StrSQL6 & " AND CP29='" & Trim(Left(Combo1.Text, 6)) & "' AND CP01 IN (" & SQLGrpStr(StrGrp090711, 1) & ") " & StrSQL6
      'Modify By Sindy 2016/9/7
      'StrSQL6 = StrSQL6 & " AND ep13='" & Trim(Left(Combo1.Text, 6)) & "' AND CP01 IN (" & SQLGrpStr(StrGrp090711, 1) & ") " & StrSQL6
      StrSQL6 = StrSQL6 & " AND cp29='" & Trim(Left(Combo1.Text, 6)) & "' AND CP01 IN (" & SQLGrpStr(StrGrp090711, 1) & ") " & StrSQL6
      '2016/9/7 END
      '2015/5/12 END
      'Modify By Sindy 2016/9/7
      'StrSQL6 = StrSQL6 & " and cp57 is null and cp27 is null and cp05>=19980101 "
      StrSQL6 = StrSQL6 & " and cp158=0 and cp159=0 and cp05>=19980101 "
      '2016/9/7 END
      'add by nick 2004/12/20   加多國案且草圖不計件不秀
'edit by nickc 2005/03/01 墨圖也要判斷
'      strSQL1 = strSQL1 & " and ((cp21='Y' and ep20 is null) or cp21 is null) "
'      StrSQL6 = StrSQL6 & " and ((cp21='Y' and ep20 is null) or cp21 is null) "
      strSQL1 = strSQL1 & " and ((cp21='Y' and (ep20 is null or ep29 is null)) or cp21 is null) "
      StrSQL6 = StrSQL6 & " and ((cp21='Y' and (ep20 is null or ep29 is null)) or cp21 is null) "
      'add by nickc 2005/03/21
      strSQL1 = strSQL1 & " and  cp107='Y' "
      StrSQL6 = StrSQL6 & " and cp107='Y' "
      
Case Else
End Select
'add by nickc 2006/04/07
StrSPa = " and ((pa58>=" & Val(Me.Text1.Text) + 191100 & "01 and pa58<=" & Val(Me.Text1.Text) + 191100 & "31) or pa58 is null) "

'Modify By Cheng 2002/04/16
'多加文件齊備日欄(EP06)
'strSQL = "SELECT SUBSTR(CP09,1,1)," & SQLDate("CP05") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04,NVL(PA05,NVL(PA06,PA07)),eP20,decode(pa09,'000',cpm03,cpm04),s1.st02," & SQLDate("CP48") & ",ROUND(cp18,2),DECODE(EP14,NULL,DECODE(EP06,NULL,'',0,''," & SQLDate("EP06") & "),0,DECODE(EP06,NULL,'',0,''," & SQLDate("EP06") & ")," & SQLDate("eP14") & ")," & SQLDate("eP15") & ",0,DECODE(EP17,NULL,DECODE(EP08,NULL,'',0,''," & SQLDate("EP08") & "),0,DECODE(EP08,NULL,'',0,''," & SQLDate("EP08") & ")," & SQLDate("eP17") & ")," & SQLDate("EP18") & ",0," & SQLDate("cP06") & "," & SQLDate("cP27") & ",ep26,s3.st02,CP09,DECODE(PA09,'000',PTM03,PTM04)," & SQLDate("CP07") & "," & SQLDate("CP57") & ",ep21,ep22,ep23,ep24,ep25,ep16,ep19,PA08 FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,PATENT,CASEPROPERTYMAP,PATENTTRADEMARKMAP " & _
'         " WHERE EP02=CP09(+) AND cp01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND cp14=S1.ST01(+) AND eP13=S2.ST01(+) AND cP13=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND '1'=PTM01(+)  AND PA08=PTM02(+) " & StrSQL6
'strSQL = strSQL & " UNION all  SELECT SUBSTR(CP09,1,1)," & SQLDate("CP05") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04,NVL(PA05,NVL(PA06,PA07)),eP20,decode(pa09,'000',cpm03,cpm04),s1.st02," & SQLDate("CP48") & ",ROUND(cp18,2),DECODE(EP14,NULL,DECODE(EP06,NULL,'',0,''," & SQLDate("EP06") & "),0,DECODE(EP06,NULL,'',0,''," & SQLDate("EP06") & ")," & SQLDate("eP14") & ")," & SQLDate("eP15") & ",0,DECODE(EP17,NULL,DECODE(EP08,NULL,'',0,''," & SQLDate("EP08") & "),0,DECODE(EP08,NULL,'',0,''," & SQLDate("EP08") & ")," & SQLDate("eP17") & ")," & SQLDate("EP18") & ",0," & SQLDate("cP06") & "," & SQLDate("cP27") & ",ep26,s3.st02,CP09,DECODE(PA09,'000',PTM03,PTM04)," & SQLDate("CP07") & "," & SQLDate("CP57") & ",ep21,ep22,ep23,ep24,ep25,ep16,ep19,PA08 FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,PATENT,CASEPROPERTYMAP,PATENTTRADEMARKMAP " & _
'         " WHERE EP02=CP09(+) AND cp01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND cp14=S1.ST01(+) AND eP13=S2.ST01(+) AND cP13=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND '1'=PTM01(+)  AND PA08=PTM02(+) " & strSQL1
'strSQL = strSQL + " ORDER BY 1,2 "
'Modify By Cheng 2002/04/26
'若已閉卷, 則在本所案號後加"*"號
'strSQL = "SELECT SUBSTR(CP09,1,1)," & SQLDate("CP05") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(PA57,'Y','＊',''),NVL(PA05,NVL(PA06,PA07)),eP20,decode(pa09,'000',cpm03,cpm04),s1.st02," & SQLDate("CP48") & ",ROUND(cp18,2),DECODE(EP06,NULL,'',0,''," & SQLDate("EP06") & "),DECODE(EP14,NULL,'',0,''," & SQLDate("eP14") & ")," & SQLDate("eP15") & ",0,DECODE(EP17,NULL,DECODE(EP08,NULL,'',0,''," & SQLDate("EP08") & "),0,DECODE(EP08,NULL,'',0,''," & SQLDate("EP08") & ")," & SQLDate("eP17") & ")," & SQLDate("EP18") & ",0," & SQLDate("cP06") & "," & SQLDate("cP27") & ",ep26,s3.st02,CP09,DECODE(PA09,'000',PTM03,PTM04)," & SQLDate("CP07") & "," & SQLDate("CP57") & ",ep21,ep22,ep23,ep24,ep25,ep16,ep19,PA08 FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,PATENT,CASEPROPERTYMAP,PATENTTRADEMARKMAP " & _
'            " WHERE EP02=CP09(+) AND cp01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND cp14=S1.ST01(+) AND eP13=S2.ST01(+) AND cP13=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND '1'=PTM01(+)  AND PA08=PTM02(+) " & StrSQL6
'Modify By Cheng 2003/06/30
'strSQL = "SELECT SUBSTR(CP09,1,1)," & SQLDate("CP05") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(PA57,'Y','＊',''),NVL(PA05,NVL(PA06,PA07)), EP20, EP29, Decode(pa09,'000',cpm03,cpm04),s1.st02, ROUND(cp18,2),DECODE(EP06,NULL,' ',0,' '," & SQLDate("EP06") & "),DECODE(EP14,NULL,'',0,''," & SQLDate("eP14") & ")," & SQLDate("eP15") & ",0,DECODE(EP17,NULL,DECODE(EP08,NULL,'',0,''," & SQLDate("EP08") & "),0,DECODE(EP08,NULL,'',0,''," & SQLDate("EP08") & ")," & SQLDate("eP17") & ")," & SQLDate("EP18") & ",0," & SQLDate("cP06") & "," & SQLDate("cP27") & ",ep26,s3.st02,CP09,DECODE(PA09,'000',PTM03,PTM04)," & SQLDate("CP07") & "," & SQLDate("CP57") & ",ep21,ep22,ep23,ep24,ep25,ep16,ep19,PA08 FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,PATENT,CASEPROPERTYMAP,PATENTTRADEMARKMAP " & _
'            " WHERE EP02=CP09(+) AND cp01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND cp14=S1.ST01(+) AND eP13=S2.ST01(+) AND cP13=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND '1'=PTM01(+)  AND PA08=PTM02(+) " & StrSQL6
'strSQL = "SELECT SUBSTR(CP09,1,1)," & SQLDate("CP05") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(PA57,'Y','＊',''),NVL(PA05,NVL(PA06,PA07)), Decode(pa09,'000',cpm03,cpm04),s1.st02, ROUND(cp18,2),DECODE(EP06,NULL,' ',0,' '," & SQLDate("EP06") & "), EP20, DECODE(EP14,NULL,'',0,''," & SQLDate("eP14") & "),'' As 草期限, " & SQLDate("eP15") & ",0, EP29, DECODE(EP17,NULL,DECODE(EP08,NULL,'',0,''," & SQLDate("EP08") & "),0,DECODE(EP08,NULL,'',0,''," & SQLDate("EP08") & ")," & SQLDate("eP17") & "), '' As 墨期限, " & SQLDate("EP18") & ",0," & SQLDate("cP06") & "," & SQLDate("cP27") & ",ep26,s3.st02,CP09,DECODE(PA09,'000',PTM03,PTM04)," & SQLDate("CP07") & "," & SQLDate("CP57") & ",ep21,ep22,ep23,ep24,ep25,ep16,ep19, CP10 FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,PATENT,CASEPROPERTYMAP,PATENTTRADEMARKMAP " & _
'            " WHERE EP02=CP09(+) AND cp01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND cp14=S1.ST01(+) AND eP13=S2.ST01(+) AND cP13=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND '1'=PTM01(+)  AND PA08=PTM02(+) " & StrSQL6
'Modify By Cheng 2004/02/18
'Modify By Cheng 2004/03/08
'顯示墨齊日欄位時直接顯示墨齊日欄位(EP17)
'strSQL = "SELECT SUBSTR(CP09,1,1)," & SQLDate("CP05") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(PA57,'Y','＊',''),NVL(PA05,NVL(PA06,PA07)), Decode(pa09,'000',cpm03,cpm04),s1.st02, ROUND(cp18,2),DECODE(EP06,NULL,' ',0,' '," & SQLDate("EP06") & "), EP20, DECODE(EP14,NULL,'',0,''," & SQLDate("eP14") & "),'' As 草期限, " & SQLDate("eP15") & ",0, EP29, DECODE(EP17,NULL,DECODE(EP08,NULL,'',0,''," & SQLDate("EP08") & "),0,DECODE(EP08,NULL,'',0,''," & SQLDate("EP08") & ")," & SQLDate("eP17") & "), '' As 墨期限, " & SQLDate("EP18") & ",0," & SQLDate("cP06") & "," & SQLDate("cP27") & ",ep26,s3.st02,CP09,DECODE(PA09,'000',PTM03,PTM04)," & SQLDate("CP07") & "," & SQLDate("CP57") & ",ep21,ep22,ep23,ep24,ep25,ep16,ep19, CP10 FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,PATENT,CASEPROPERTYMAP,PATENTTRADEMARKMAP " & _
'            " WHERE EP02=CP09(+) AND cp01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND cp14=S1.ST01(+) AND eP13=S2.ST01(+) AND cP13=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND '1'=PTM01(+)  AND PA08=PTM02(+) " & StrSQL6

'Modify by Morgan 2004/5/19
'加專利種類 PA08
'strSQL = "SELECT SUBSTR(CP09,1,1)," & SQLDate("CP05") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(PA57,'Y','＊',''),NVL(PA05,NVL(PA06,PA07)), Decode(pa09,'000',cpm03,cpm04),s1.st02, ROUND(cp18,2),DECODE(EP06,NULL,' ',0,' '," & SQLDate("EP06") & "), EP20, DECODE(EP14,NULL,'',0,''," & SQLDate("eP14") & "),'' As 草期限, " & SQLDate("eP15") & ",0, EP29, DECODE(EP17,NULL,'' ," & SQLDate("eP17") & "), '' As 墨期限, " & SQLDate("EP18") & ",0," & SQLDate("cP06") & "," & SQLDate("cP27") & ",ep26,s3.st02,CP09,DECODE(PA09,'000',PTM03,PTM04)," & SQLDate("CP07") & "," & SQLDate("CP57") & ",ep21,ep22,ep23,ep24,ep25,ep16,ep19, CP10 FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,PATENT,CASEPROPERTYMAP,PATENTTRADEMARKMAP " & _
'            " WHERE EP02=CP09(+) AND cp01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND cp14=S1.ST01(+) AND eP13=S2.ST01(+) AND cP13=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND '1'=PTM01(+)  AND PA08=PTM02(+) " & StrSQL6
'edit by nick 2004/12/21 加申請國家
'StrSql = "SELECT SUBSTR(CP09,1,1)," & SQLDate("CP05") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(PA57,'Y','＊',''),NVL(PA05,NVL(PA06,PA07)), Decode(pa09,'000',cpm03,cpm04),s1.st02, ROUND(cp18,2),DECODE(EP06,NULL,' ',0,' '," & SQLDate("EP06") & "), EP20, DECODE(EP14,NULL,'',0,''," & SQLDate("eP14") & "),'' As 草期限, " & SQLDate("eP15") & ",0, EP29, DECODE(EP17,NULL,'' ," & SQLDate("eP17") & "), '' As 墨期限, " & SQLDate("EP18") & ",0," & SQLDate("cP06") & "," & SQLDate("cP27") & ",ep26,s3.st02,CP09,DECODE(PA09,'000',PTM03,PTM04)," & SQLDate("CP07") & "," & SQLDate("CP57") & ",ep21,ep22,ep23,ep24,ep25,ep16,ep19, CP10, PA08 FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,PATENT,CASEPROPERTYMAP,PATENTTRADEMARKMAP " & _
            " WHERE EP02=CP09(+) AND cp01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND cp14=S1.ST01(+) AND eP13=S2.ST01(+) AND cP13=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND '1'=PTM01(+)  AND PA08=PTM02(+) " & StrSQL6
'edit by nick 2004/12/22 加排序條件
'StrSql = "SELECT SUBSTR(CP09,1,1)," & SQLDate("CP05") & ",na03,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(PA57,'Y','＊',''),NVL(PA05,NVL(PA06,PA07)), Decode(pa09,'000',cpm03,cpm04),s1.st02, ROUND(cp18,2),DECODE(EP06,NULL,' ',0,' '," & SQLDate("EP06") & "), EP20, DECODE(EP14,NULL,'',0,''," & SQLDate("eP14") & "),'' As 草期限, " & SQLDate("eP15") & ",0, EP29, DECODE(EP17,NULL,'' ," & SQLDate("eP17") & "), '' As 墨期限, " & SQLDate("EP18") & ",0," & SQLDate("cP06") & "," & SQLDate("cP27") & ",ep26,s3.st02,CP09,DECODE(PA09,'000',PTM03,PTM04)," & SQLDate("CP07") & "," & SQLDate("CP57") & ",ep21,ep22,ep23,ep24,ep25,ep16,ep19, CP10, PA08 FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,PATENT,CASEPROPERTYMAP,PATENTTRADEMARKMAP,nation " & _
            " WHERE EP02=CP09(+) AND cp01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND cp14=S1.ST01(+) AND eP13=S2.ST01(+) AND cP13=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND '1'=PTM01(+)  AND PA08=PTM02(+) and pa09=na01(+) " & StrSQL6
'edit by nickc 2005/04/12 計不計件值欄秀計件值
'StrSql = "SELECT SUBSTR(CP09,1,1)," & SQLDate("CP05") & ",substrb(na03,1,4),CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(PA57,'Y','＊',''),NVL(PA05,NVL(PA06,PA07)), Decode(pa09,'000',cpm03,cpm04),s1.st02, ROUND(cp18,2),DECODE(EP06,NULL,' ',0,' '," & SQLDate("EP06") & "), EP20, DECODE(EP14,NULL,'',0,''," & SQLDate("eP14") & "),'' As 草期限, " & SQLDate("eP15") & ",0, EP29, DECODE(EP17,NULL,'' ," & SQLDate("eP17") & "), '' As 墨期限, " & SQLDate("EP18") & ",0," & SQLDate("cP06") & "," & SQLDate("cP27") & ",ep26,s3.st02,CP09,DECODE(PA09,'000',PTM03,PTM04)," & SQLDate("CP07") & "," & SQLDate("CP57") & ",ep21,ep22,ep23,ep24,ep25,ep16,ep19, CP10, PA08,decode(pa09,'000','9999','020','8888','101','7777','0000') as PaSort FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,PATENT,CASEPROPERTYMAP,PATENTTRADEMARKMAP,nation " & _
            " WHERE EP02=CP09(+) AND cp01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND cp14=S1.ST01(+) AND eP13=S2.ST01(+) AND cP13=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND '1'=PTM01(+)  AND PA08=PTM02(+) and pa09=na01(+) " & StrSQL6
'edit by nickc 2006/01/16
'strSQL = "SELECT SUBSTR(CP09,1,1)," & SQLDate("CP05") & ",substrb(na03,1,4),CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(PA57,'Y','＊',''),NVL(PA05,NVL(PA06,PA07)), Decode(pa09,'000',cpm03,cpm04),s1.st02, ROUND(cp18,2),DECODE(EP06,NULL,' ',0,' '," & SQLDate("EP06") & "), nvl(EP20,round(cp100 * cp101,2)), DECODE(EP14,NULL,'',0,''," & SQLDate("eP14") & "),'' As 草期限, " & SQLDate("eP15") & ",0, nvl(EP29,round(cp103 * cp104,2)), DECODE(EP17,NULL,'' ," & SQLDate("eP17") & "), '' As 墨期限, " & SQLDate("EP18") & ",0," & SQLDate("cP06") & "," & SQLDate("cP27") & ",ep26,s3.st02,CP09,DECODE(PA09,'000',PTM03,PTM04)," & SQLDate("CP07") & "," & SQLDate("CP57") & ",ep21,ep22,ep23,ep24,ep25,ep16,ep19, CP10, PA08,decode(pa09,'000','9999','020','8888','101','7777','0000') as PaSort FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,PATENT,CASEPROPERTYMAP,PATENTTRADEMARKMAP,nation " & _
            " WHERE EP02=CP09(+) AND cp01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND cp14=S1.ST01(+) AND eP13=S2.ST01(+) AND cP13=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND '1'=PTM01(+)  AND PA08=PTM02(+) and pa09=na01(+) " & StrSQL6
'edit by nickc 2006/04/07
'strSQL = "SELECT distinct SUBSTR(CP09,1,1)," & SQLDate("CP05") & ",substrb(na03,1,4),CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(PA57,'Y','＊',''),NVL(PA05,NVL(PA06,PA07)), Decode(pa09,'000',cpm03,cpm04),s1.st02, ROUND(cp18,2),DECODE(EP06,NULL,' ',0,' '," & SQLDate("EP06") & "), nvl(EP20,round(cp100 * cp101,2)), DECODE(EP14,NULL,'',0,''," & SQLDate("eP14") & "),'' As 草期限, " & SQLDate("eP15") & ",0, nvl(EP29,round(cp103 * cp104,2)), DECODE(EP17,NULL,'' ," & SQLDate("eP17") & "), '' As 墨期限, " & SQLDate("EP18") & ",0," & SQLDate("cP06") & "," & SQLDate("cP27") & ",ep26,s3.st02,CP09,DECODE(PA09,'000',PTM03,PTM04)," & SQLDate("CP07") & "," & SQLDate("CP57") & ",ep21,ep22,ep23,ep24,ep25,ep16,ep19, CP10, PA08,decode(pa09,'000','9999','020','8888','101','7777','0000') as PaSort FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,PATENT,CASEPROPERTYMAP,PATENTTRADEMARKMAP,nation " & _
            " WHERE EP02=CP09(+) AND cp01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND cp14=S1.ST01(+) AND eP13=S2.ST01(+) AND cP13=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND '1'=PTM01(+)  AND PA08=PTM02(+) and pa09=na01(+) " & StrSQL6
'edit by nickc 2007/08/16 加入代表圖的判斷
'strSQL = "SELECT distinct SUBSTR(CP09,1,1)," & SQLDate("CP05") & ",substrb(na03,1,4),CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(PA57,'Y','＊',''),NVL(PA05,NVL(PA06,PA07)), Decode(pa09,'000',cpm03,cpm04),s1.st02, ROUND(cp18,2),DECODE(EP06,NULL,' ',0,' '," & SQLDate("EP06") & "), nvl(EP20,round(cp100 * cp101,2)), DECODE(EP14,NULL,'',0,''," & SQLDate("eP14") & "),'' As 草期限, " & SQLDate("eP15") & ",0, nvl(EP29,round(cp103 * cp104,2)), DECODE(EP17,NULL,'' ," & SQLDate("eP17") & "), '' As 墨期限, " & SQLDate("EP18") & ",0," & SQLDate("cP06") & "," & SQLDate("cP27") & ",ep26,s3.st02,CP09,DECODE(PA09,'000',PTM03,PTM04)," & SQLDate("CP07") & ",nvl(" & SQLDate("CP57") & "," & SQLDate("pa58") & "),ep21,ep22,ep23,ep24,ep25,ep16,ep19, CP10, PA08,decode(pa09,'000','9999','020','8888','101','7777','0000') as PaSort FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,PATENT,CASEPROPERTYMAP,PATENTTRADEMARKMAP,nation " & _
            " WHERE EP02=CP09(+) AND cp01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND cp14=S1.ST01(+) AND eP13=S2.ST01(+) AND cP13=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND '1'=PTM01(+)  AND PA08=PTM02(+) and pa09=na01(+) " & StrSQL6 & StrSPa

'Modify By Sindy 2015/5/12 調整SQL增加查詢速度
'Modify by Morgan 2011/1/4 修正日期欄位排序問題(小於100年的前面補空白)
'strSql = "SELECT distinct SUBSTR(CP09,1,1)||decode(ibf13,null,'','+'),SQLDateT2(CP05),substrb(na03,1,4),CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(PA57,'Y','＊',''),NVL(PA05,NVL(PA06,PA07)), Decode(pa09,'000',cpm03,cpm04),s1.st02, ROUND(cp18,2),nvl(SQLDateT2(EP06),' '), nvl(EP20,round(cp100 * cp101,2)),NVL(SQLDateT2(EP14),' '),'' As 草期限, SQLDateT2(EP15),0, nvl(EP29,round(cp103 * cp104,2)),NVL(SQLDateT2(EP17),' '), '' As 墨期限,SQLDateT2(EP18),0,SQLDateT2(CP06),SQLDateT2(CP27),ep26,s3.st02,CP09,decode(pa01,'CFP',ptm03,DECODE(PA09,'000',PTM03,PTM04)),SQLDateT2(CP07),SQLDateT2(NVL(CP57,PA58)),ep21,ep22,ep23,ep24,ep25,ep16,ep19, CP10, PA08,decode(pa09,'000','9999','020','8888','101','7777','0000') as PaSort FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,PATENT,CASEPROPERTYMAP,PATENTTRADEMARKMAP,nation,imgbytefile  " & _
'            " WHERE EP02(+)=CP09 AND cp01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND cp14=S1.ST01(+) AND eP13=S2.ST01(+) AND cP13=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND '1'=PTM01(+)  AND PA08=PTM02(+) and pa09=na01(+) and cp01=ibf01(+) and cp02=ibf02(+) and cp03=ibf03(+) and cp04=ibf04(+) " & StrSQL6 & StrSPa
strVBA = "SELECT cp01,cp02,cp03,cp04,cp05,cp06,cp07,cp09,cp10,cp13,cp14,cp18,cp27,cp57,cp103,cp104,cp100,cp101,ep02,ep06,ep13,ep14,ep15,ep16,ep17,ep18,ep19,ep20,ep21,ep22,ep23,ep24,ep25,ep26,ep29" & _
         " From CASEPROGRESS,ENGINEERPROGRESS" & _
         " WHERE EP02(+)=CP09 " & StrSQL6
'2015/5/12 END
'End
'End
'Modify By Cheng 2002/04/19
If ProState <> 4 Then
'   strSQL = strSQL & " UNION all  SELECT SUBSTR(CP09,1,1)," & SQLDate("CP05") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(PA57,'Y','＊',''),NVL(PA05,NVL(PA06,PA07)),eP20,decode(pa09,'000',cpm03,cpm04),s1.st02," & SQLDate("CP48") & ",ROUND(cp18,2),DECODE(EP06,NULL,'',0,''," & SQLDate("EP06") & "),DECODE(EP14,NULL,'',0,''," & SQLDate("eP14") & ")," & SQLDate("eP15") & ",0,DECODE(EP17,NULL,DECODE(EP08,NULL,'',0,''," & SQLDate("EP08") & "),0,DECODE(EP08,NULL,'',0,''," & SQLDate("EP08") & ")," & SQLDate("eP17") & ")," & SQLDate("EP18") & ",0," & SQLDate("cP06") & "," & SQLDate("cP27") & ",ep26,s3.st02,CP09,DECODE(PA09,'000',PTM03,PTM04)," & SQLDate("CP07") & "," & SQLDate("CP57") & ",ep21,ep22,ep23,ep24,ep25,ep16,ep19,PA08 FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,PATENT,CASEPROPERTYMAP,PATENTTRADEMARKMAP " & _
'                " WHERE EP02=CP09(+) AND cp01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND cp14=S1.ST01(+) AND eP13=S2.ST01(+) AND cP13=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND '1'=PTM01(+)  AND PA08=PTM02(+) " & strSQL1
'Modify By Cheng 2003/06/30
'   strSQL = strSQL & " UNION all  SELECT SUBSTR(CP09,1,1)," & SQLDate("CP05") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(PA57,'Y','＊',''),NVL(PA05,NVL(PA06,PA07)), EP20, EP29, Decode(pa09,'000',cpm03,cpm04),s1.st02, ROUND(cp18,2),DECODE(EP06,NULL,' ',0,' '," & SQLDate("EP06") & "),DECODE(EP14,NULL,'',0,''," & SQLDate("eP14") & ")," & SQLDate("eP15") & ",0,DECODE(EP17,NULL,DECODE(EP08,NULL,'',0,''," & SQLDate("EP08") & "),0,DECODE(EP08,NULL,'',0,''," & SQLDate("EP08") & ")," & SQLDate("eP17") & ")," & SQLDate("EP18") & ",0," & SQLDate("cP06") & "," & SQLDate("cP27") & ",ep26,s3.st02,CP09,DECODE(PA09,'000',PTM03,PTM04)," & SQLDate("CP07") & "," & SQLDate("CP57") & ",ep21,ep22,ep23,ep24,ep25,ep16,ep19,PA08 FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,PATENT,CASEPROPERTYMAP,PATENTTRADEMARKMAP " & _
'                " WHERE EP02=CP09(+) AND cp01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND cp14=S1.ST01(+) AND eP13=S2.ST01(+) AND cP13=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND '1'=PTM01(+)  AND PA08=PTM02(+) " & strSQL1
    'Modify By Cheng 2004/02/18
'   strSQL = strSQL & " UNION all  SELECT SUBSTR(CP09,1,1)," & SQLDate("CP05") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(PA57,'Y','＊',''),NVL(PA05,NVL(PA06,PA07)), Decode(pa09,'000',cpm03,cpm04),s1.st02, ROUND(cp18,2),DECODE(EP06,NULL,' ',0,' '," & SQLDate("EP06") & "), EP20, DECODE(EP14,NULL,'',0,''," & SQLDate("eP14") & "), '' As 草期限, " & SQLDate("eP15") & ",0, EP29, DECODE(EP17,NULL,DECODE(EP08,NULL,'',0,''," & SQLDate("EP08") & "),0,DECODE(EP08,NULL,'',0,''," & SQLDate("EP08") & ")," & SQLDate("eP17") & "), '' As 墨期限, " & SQLDate("EP18") & ",0," & SQLDate("cP06") & "," & SQLDate("cP27") & ",ep26,s3.st02,CP09,DECODE(PA09,'000',PTM03,PTM04)," & SQLDate("CP07") & "," & SQLDate("CP57") & ",ep21,ep22,ep23,ep24,ep25,ep16,ep19, CP10 FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,PATENT,CASEPROPERTYMAP,PATENTTRADEMARKMAP " & _
'                " WHERE EP02=CP09(+) AND cp01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND cp14=S1.ST01(+) AND eP13=S2.ST01(+) AND cP13=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND '1'=PTM01(+)  AND PA08=PTM02(+) " & strSQL1
    'Modify By Cheng 2004/03/08
    '顯示墨齊日欄位時直接顯示墨齊日欄位(EP17)
'   strSQL = strSQL & " UNION all  SELECT SUBSTR(CP09,1,1)," & SQLDate("CP05") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(PA57,'Y','＊',''),NVL(PA05,NVL(PA06,PA07)), Decode(pa09,'000',cpm03,cpm04),s1.st02, ROUND(cp18,2),DECODE(EP06,NULL,' ',0,' '," & SQLDate("EP06") & "), EP20, DECODE(EP14,NULL,'',0,''," & SQLDate("eP14") & "), '' As 草期限, " & SQLDate("eP15") & ",0, EP29, DECODE(EP17,NULL,DECODE(EP08,NULL,'',0,''," & SQLDate("EP08") & "),0,DECODE(EP08,NULL,'',0,''," & SQLDate("EP08") & ")," & SQLDate("eP17") & "), '' As 墨期限, " & SQLDate("EP18") & ",0," & SQLDate("cP06") & "," & SQLDate("cP27") & ",ep26,s3.st02,CP09,DECODE(PA09,'000',PTM03,PTM04)," & SQLDate("CP07") & "," & SQLDate("CP57") & ",ep21,ep22,ep23,ep24,ep25,ep16,ep19, CP10 FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,PATENT,CASEPROPERTYMAP,PATENTTRADEMARKMAP " & _
'                " WHERE EP02=CP09(+) AND cp01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND cp14=S1.ST01(+) AND eP13=S2.ST01(+) AND cP13=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND '1'=PTM01(+)  AND PA08=PTM02(+) " & strSQL1
   
   'Modify by Morgan 2004/5/19
   '加專利種類 PA08
'   strSQL = strSQL & " UNION all  SELECT SUBSTR(CP09,1,1)," & SQLDate("CP05") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(PA57,'Y','＊',''),NVL(PA05,NVL(PA06,PA07)), Decode(pa09,'000',cpm03,cpm04),s1.st02, ROUND(cp18,2),DECODE(EP06,NULL,' ',0,' '," & SQLDate("EP06") & "), EP20, DECODE(EP14,NULL,'',0,''," & SQLDate("eP14") & "), '' As 草期限, " & SQLDate("eP15") & ",0, EP29, DECODE(EP17,NULL,'' ," & SQLDate("eP17") & "), '' As 墨期限, " & SQLDate("EP18") & ",0," & SQLDate("cP06") & "," & SQLDate("cP27") & ",ep26,s3.st02,CP09,DECODE(PA09,'000',PTM03,PTM04)," & SQLDate("CP07") & "," & SQLDate("CP57") & ",ep21,ep22,ep23,ep24,ep25,ep16,ep19, CP10 FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,PATENT,CASEPROPERTYMAP,PATENTTRADEMARKMAP " & _
'                " WHERE EP02=CP09(+) AND cp01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND cp14=S1.ST01(+) AND eP13=S2.ST01(+) AND cP13=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND '1'=PTM01(+)  AND PA08=PTM02(+) " & strSQL1
'edit by nick 2004/12/21 加申請國家
'   StrSql = StrSql & " UNION all  SELECT SUBSTR(CP09,1,1)," & SQLDate("CP05") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(PA57,'Y','＊',''),NVL(PA05,NVL(PA06,PA07)), Decode(pa09,'000',cpm03,cpm04),s1.st02, ROUND(cp18,2),DECODE(EP06,NULL,' ',0,' '," & SQLDate("EP06") & "), EP20, DECODE(EP14,NULL,'',0,''," & SQLDate("eP14") & "), '' As 草期限, " & SQLDate("eP15") & ",0, EP29, DECODE(EP17,NULL,'' ," & SQLDate("eP17") & "), '' As 墨期限, " & SQLDate("EP18") & ",0," & SQLDate("cP06") & "," & SQLDate("cP27") & ",ep26,s3.st02,CP09,DECODE(PA09,'000',PTM03,PTM04)," & SQLDate("CP07") & "," & SQLDate("CP57") & ",ep21,ep22,ep23,ep24,ep25,ep16,ep19, CP10, PA08 FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,PATENT,CASEPROPERTYMAP,PATENTTRADEMARKMAP " & _
                " WHERE EP02=CP09(+) AND cp01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND cp14=S1.ST01(+) AND eP13=S2.ST01(+) AND cP13=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND '1'=PTM01(+)  AND PA08=PTM02(+) " & strSQL1
'edit by nick 2004/12/22 加排序條件
'   StrSql = StrSql & " UNION all  SELECT SUBSTR(CP09,1,1)," & SQLDate("CP05") & ",na03,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(PA57,'Y','＊',''),NVL(PA05,NVL(PA06,PA07)), Decode(pa09,'000',cpm03,cpm04),s1.st02, ROUND(cp18,2),DECODE(EP06,NULL,' ',0,' '," & SQLDate("EP06") & "), EP20, DECODE(EP14,NULL,'',0,''," & SQLDate("eP14") & "), '' As 草期限, " & SQLDate("eP15") & ",0, EP29, DECODE(EP17,NULL,'' ," & SQLDate("eP17") & "), '' As 墨期限, " & SQLDate("EP18") & ",0," & SQLDate("cP06") & "," & SQLDate("cP27") & ",ep26,s3.st02,CP09,DECODE(PA09,'000',PTM03,PTM04)," & SQLDate("CP07") & "," & SQLDate("CP57") & ",ep21,ep22,ep23,ep24,ep25,ep16,ep19, CP10, PA08 FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,PATENT,CASEPROPERTYMAP,PATENTTRADEMARKMAP,nation " & _
                " WHERE EP02=CP09(+) AND cp01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND cp14=S1.ST01(+) AND eP13=S2.ST01(+) AND cP13=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND '1'=PTM01(+)  AND PA08=PTM02(+) and pa09=na01(+) " & strSQL1
'edit by nickc 2005/04/12 計不計件值欄秀計件值
'   StrSql = StrSql & " UNION all  SELECT SUBSTR(CP09,1,1)," & SQLDate("CP05") & ",substrb(na03,1,4),CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(PA57,'Y','＊',''),NVL(PA05,NVL(PA06,PA07)), Decode(pa09,'000',cpm03,cpm04),s1.st02, ROUND(cp18,2),DECODE(EP06,NULL,' ',0,' '," & SQLDate("EP06") & "), EP20, DECODE(EP14,NULL,'',0,''," & SQLDate("eP14") & "), '' As 草期限, " & SQLDate("eP15") & ",0, EP29, DECODE(EP17,NULL,'' ," & SQLDate("eP17") & "), '' As 墨期限, " & SQLDate("EP18") & ",0," & SQLDate("cP06") & "," & SQLDate("cP27") & ",ep26,s3.st02,CP09,DECODE(PA09,'000',PTM03,PTM04)," & SQLDate("CP07") & "," & SQLDate("CP57") & ",ep21,ep22,ep23,ep24,ep25,ep16,ep19, CP10, PA08,decode(pa09,'000','9999','020','8888','101','7777','0000') as PaSort FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,PATENT,CASEPROPERTYMAP,PATENTTRADEMARKMAP,nation " & _
                " WHERE EP02=CP09(+) AND cp01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND cp14=S1.ST01(+) AND eP13=S2.ST01(+) AND cP13=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND '1'=PTM01(+)  AND PA08=PTM02(+) and pa09=na01(+) " & strSQL1
'edit by nickc 2006/01/16
'   strSQL = strSQL & " UNION all  SELECT SUBSTR(CP09,1,1)," & SQLDate("CP05") & ",substrb(na03,1,4),CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(PA57,'Y','＊',''),NVL(PA05,NVL(PA06,PA07)), Decode(pa09,'000',cpm03,cpm04),s1.st02, ROUND(cp18,2),DECODE(EP06,NULL,' ',0,' '," & SQLDate("EP06") & "), nvl(EP20,round(cp100 * cp101,2)), DECODE(EP14,NULL,'',0,''," & SQLDate("eP14") & "), '' As 草期限, " & SQLDate("eP15") & ",0, nvl(EP29,round(cp103 * cp104,2)), DECODE(EP17,NULL,'' ," & SQLDate("eP17") & "), '' As 墨期限, " & SQLDate("EP18") & ",0," & SQLDate("cP06") & "," & SQLDate("cP27") & ",ep26,s3.st02,CP09,DECODE(PA09,'000',PTM03,PTM04)," & SQLDate("CP07") & "," & SQLDate("CP57") & ",ep21,ep22,ep23,ep24,ep25,ep16,ep19, CP10, PA08,decode(pa09,'000','9999','020','8888','101','7777','0000') as PaSort FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,PATENT,CASEPROPERTYMAP,PATENTTRADEMARKMAP,nation " & _
                " WHERE EP02=CP09(+) AND cp01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND cp14=S1.ST01(+) AND eP13=S2.ST01(+) AND cP13=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND '1'=PTM01(+)  AND PA08=PTM02(+) and pa09=na01(+) " & strSQL1
'edit by nickc 2006/04/07
'   strSQL = strSQL & " UNION SELECT SUBSTR(CP09,1,1)," & SQLDate("CP05") & ",substrb(na03,1,4),CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(PA57,'Y','＊',''),NVL(PA05,NVL(PA06,PA07)), Decode(pa09,'000',cpm03,cpm04),s1.st02, ROUND(cp18,2),DECODE(EP06,NULL,' ',0,' '," & SQLDate("EP06") & "), nvl(EP20,round(cp100 * cp101,2)), DECODE(EP14,NULL,'',0,''," & SQLDate("eP14") & "), '' As 草期限, " & SQLDate("eP15") & ",0, nvl(EP29,round(cp103 * cp104,2)), DECODE(EP17,NULL,'' ," & SQLDate("eP17") & "), '' As 墨期限, " & SQLDate("EP18") & ",0," & SQLDate("cP06") & "," & SQLDate("cP27") & ",ep26,s3.st02,CP09,DECODE(PA09,'000',PTM03,PTM04)," & SQLDate("CP07") & "," & SQLDate("CP57") & ",ep21,ep22,ep23,ep24,ep25,ep16,ep19, CP10, PA08,decode(pa09,'000','9999','020','8888','101','7777','0000') as PaSort FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,PATENT,CASEPROPERTYMAP,PATENTTRADEMARKMAP,nation " & _
                " WHERE EP02=CP09(+) AND cp01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND cp14=S1.ST01(+) AND eP13=S2.ST01(+) AND cP13=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND '1'=PTM01(+)  AND PA08=PTM02(+) and pa09=na01(+) " & strSQL1
'edit by nickc 2007/08/16 加入代表圖的判斷
'   strSQL = strSQL & " UNION SELECT SUBSTR(CP09,1,1)," & SQLDate("CP05") & ",substrb(na03,1,4),CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(PA57,'Y','＊',''),NVL(PA05,NVL(PA06,PA07)), Decode(pa09,'000',cpm03,cpm04),s1.st02, ROUND(cp18,2),DECODE(EP06,NULL,' ',0,' '," & SQLDate("EP06") & "), nvl(EP20,round(cp100 * cp101,2)), DECODE(EP14,NULL,'',0,''," & SQLDate("eP14") & "), '' As 草期限, " & SQLDate("eP15") & ",0, nvl(EP29,round(cp103 * cp104,2)), DECODE(EP17,NULL,'' ," & SQLDate("eP17") & "), '' As 墨期限, " & SQLDate("EP18") & ",0," & SQLDate("cP06") & "," & SQLDate("cP27") & ",ep26,s3.st02,CP09,DECODE(PA09,'000',PTM03,PTM04)," & SQLDate("CP07") & ",nvl(" & SQLDate("CP57") & "," & SQLDate("pa58") & "),ep21,ep22,ep23,ep24,ep25,ep16,ep19, CP10, PA08,decode(pa09,'000','9999','020','8888','101','7777','0000') as PaSort FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,PATENT,CASEPROPERTYMAP,PATENTTRADEMARKMAP,nation " & _
                " WHERE EP02=CP09(+) AND cp01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND cp14=S1.ST01(+) AND eP13=S2.ST01(+) AND cP13=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND '1'=PTM01(+)  AND PA08=PTM02(+) and pa09=na01(+) " & strSQL1 & StrSPa
   
   'Modify By Sindy 2015/5/12 調整SQL增加查詢速度
   'Modify by Morgan 2011/1/4 修正日期欄位排序問題(小於100年的前面補空白)
'   strSql = strSql & " UNION SELECT SUBSTR(CP09,1,1)||decode(ibf13,null,'','+'),SQLDateT2(CP05),substrb(na03,1,4),CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(PA57,'Y','＊',''),NVL(PA05,NVL(PA06,PA07)), Decode(pa09,'000',cpm03,cpm04),s1.st02, ROUND(cp18,2),NVL(SQLDateT2(EP06),' '), nvl(EP20,round(cp100 * cp101,2)),NVL(SQLDateT2(EP14),' '), '' As 草期限,SQLDateT2(EP15),0, nvl(EP29,round(cp103 * cp104,2)),NVL(SQLDateT2(EP17),' '), '' As 墨期限,SQLDateT2(EP18),0,SQLDateT2(CP06),SQLDateT2(CP27),ep26,s3.st02,CP09,decode(pa01,'CFP',ptm03,DECODE(PA09,'000',PTM03,PTM04)),SQLDateT2(CP07),SQLDateT2(NVL(CP57,PA58)),ep21,ep22,ep23,ep24,ep25,ep16,ep19, CP10, PA08,decode(pa09,'000','9999','020','8888','101','7777','0000') as PaSort FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,PATENT,CASEPROPERTYMAP,PATENTTRADEMARKMAP,nation,imgbytefile " & _
'                " WHERE EP02(+)=CP09 AND cp01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND cp14=S1.ST01(+) AND eP13=S2.ST01(+) AND cP13=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND '1'=PTM01(+)  AND PA08=PTM02(+) and pa09=na01(+)  and cp01=ibf01(+) and cp02=ibf02(+) and cp03=ibf03(+) and cp04=ibf04(+) " & strSQL1 & StrSPa
   strVBA = strVBA & " UNION SELECT cp01,cp02,cp03,cp04,cp05,cp06,cp07,cp09,cp10,cp13,cp14,cp18,cp27,cp57,cp103,cp104,cp100,cp101,ep02,ep06,ep13,ep14,ep15,ep16,ep17,ep18,ep19,ep20,ep21,ep22,ep23,ep24,ep25,ep26,ep29" & _
                     " From CASEPROGRESS,ENGINEERPROGRESS" & _
                     " WHERE EP02(+)=CP09 " & strSQL1
   '2015/5/12 END
    'End
    'End
End If
'92.5.13 modify by sonia
'strSQL = strSQL + " ORDER BY 1,2 "
'strSQL = strSQL + " ORDER BY 2 desc,3 desc,1 "
'Modify By Cheng 2004/02/26
'排序方式文齊日(大至小且Null的排在最底下), 墨齊日(大至小), 草齊日(大至小), 本所案號(小至大)
'strSQL = strSQL + " ORDER BY 8 desc, 3 desc "
'StrSql = StrSql + " ORDER BY 8 Desc, 16 Desc, 10 Desc, 3 Asc "
'edit by nick 2004/12/21
'StrSql = StrSql + " ORDER BY 9 Desc, 17 Desc, 11 Desc, 4 Asc "
'edit by nick 2004/12/22 文齊後面加國家 台灣排第一，再來大陸，再來日本，其他不管
'Modify By Sindy 2015/5/12 調整SQL增加查詢速度
strSql = "SELECT SUBSTR(CP09,1,1)||decode(ibf13,null,'','+'),SQLDateT2(CP05),substrb(na03,1,4),CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(PA57,'Y','＊',''),NVL(PA05,NVL(PA06,PA07)), Decode(pa09,'000',cpm03,cpm04),s1.st02, ROUND(cp18,2),NVL(SQLDateT2(EP06),' '), nvl(EP20,round(cp100 * cp101,2)),NVL(SQLDateT2(EP14),' '), '' As 草期限,SQLDateT2(EP15),0, nvl(EP29,round(cp103 * cp104,2)),NVL(SQLDateT2(EP17),' '), '' As 墨期限,SQLDateT2(EP18),0,SQLDateT2(CP06),SQLDateT2(CP27),ep26,s3.st02,CP09,decode(pa01,'CFP',ptm03,DECODE(PA09,'000',PTM03,PTM04)),SQLDateT2(CP07),SQLDateT2(NVL(CP57,PA58)),ep21,ep22,ep23,ep24,ep25,ep16,ep19, CP10, PA08,decode(pa09,'000','9999','020','8888','101','7777','0000') as PaSort" & _
         " FROM (" & strVBA & ") A,STAFF S1,STAFF S2,STAFF S3,PATENT,CASEPROPERTYMAP,PATENTTRADEMARKMAP,nation,imgbytefile " & _
         " WHERE cp01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND cp14=S1.ST01(+) AND eP13=S2.ST01(+) AND cP13=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND '1'=PTM01(+) AND PA08=PTM02(+) and pa09=na01(+) and cp01=ibf01(+) and cp02=ibf02(+) and cp03=ibf03(+) and cp04=ibf04(+) " & StrSPa
'2015/5/12 END
strSql = strSql + " ORDER BY 9 Desc,pasort desc, 17 Desc, 11 Desc, 4 Asc "

'add by nickc 2005/03/31
Screen.MousePointer = vbHourglass
grd1.MousePointer = flexHourglass
'End
'92.5.13 end
CheckOC
With adoRecordset
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        If (ProState = "2" Or ProState = "4") And pub_QL04 <> "" Then
            InsertQueryLog (.RecordCount) 'Add By Sindy 2010/12/16
            pub_QL04 = ""
        End If
        Set grd1.Recordset = adoRecordset
        grd1.Visible = False
        Screen.MousePointer = vbHourglass
        grd1.MousePointer = flexHourglass
        Me.Enabled = False
        For i = 1 To grd1.Rows - 1
            grd1.row = i
            '文齊日(將空白消除)
            'edit by nick 2004/12/21 加了申請國家，要往後退
            'Me.grd1.TextMatrix(i, 7) = Trim(Me.grd1.TextMatrix(i, 7))
            
            'Remove by Morgan 2011/1/4 修正日期欄位排序問題(小於100年的前面會補空白)
            'Me.grd1.TextMatrix(i, 8) = Trim(Me.grd1.TextMatrix(i, 8))
            
            '草完日
            'edit by nick 2004/12/21 加了申請國家，要往後退
            'grd1.col = 11
            grd1.col = 12
            'Modify by Morgan 2011/1/4 修正日期欄位排序問題(小於100年的前面會補空白)
            'StrDate1 = grd1.Text
            strDate1 = LTrim(grd1.Text)
            
            '草齊日
            'edit by nick 2004/12/21 加了申請國家，要往後退
            'grd1.col = 9
            grd1.col = 10
            'Modify by Morgan 2011/1/4 修正日期欄位排序問題(小於100年的前面會補空白)
            'StrDate2 = grd1.Text
            StrDate2 = LTrim(grd1.Text)
            
            If Trim(strDate1) <> "" And Trim(StrDate2) <> "" Then
                '草天
                'edit by nick 2004/12/21 加了申請國家，要往後退
                'grd1.col = 12
                grd1.col = 13
                grd1.Text = GetWorkDay(ChangeTStringToWString(ChangeTDateStringToTString(strDate1)), ChangeTStringToWString(ChangeTDateStringToTString(StrDate2)))
            Else
                '草天
                'edit by nick 2004/12/21 加了申請國家，要往後退
                'grd1.col = 12
                grd1.col = 13
                grd1.Text = ""
            End If
            '墨完日
            'edit by nick 2004/12/21 加了申請國家，要往後退
            'grd1.col = 16
            grd1.col = 17
            'Modify by Morgan 2011/1/4 修正日期欄位排序問題(小於100年的前面會補空白)
            'StrDate1 = grd1.Text
            strDate1 = LTrim(grd1.Text)
            
            '墨齊日
            'edit by nick 2004/12/21 加了申請國家，要往後退
            'grd1.col = 14
            grd1.col = 15
            'Modify by Morgan 2011/1/4 修正日期欄位排序問題(小於100年的前面會補空白)
            'StrDate2 = grd1.Text
            StrDate2 = LTrim(grd1.Text)
            
            If Trim(strDate1) <> "" And Trim(StrDate2) <> "" Then
                '墨天
                'edit by nick 2004/12/21 加了申請國家，要往後退
                'grd1.col = 17
                grd1.col = 18
                grd1.Text = GetWorkDay(ChangeTStringToWString(ChangeTDateStringToTString(strDate1)), ChangeTStringToWString(ChangeTDateStringToTString(StrDate2)))
            Else
                '墨天
                'edit by nick 2004/12/21 加了申請國家，要往後退
                'grd1.col = 17
                grd1.col = 18
                grd1.Text = ""
            End If
            grd1.Text = grd1.Text & GetSign(grd1.TextMatrix(i, 23)) 'Add by Morgan 2009/10/5
            
            'Add By Cheng 2003/06/27
            '若草圖不計件
            'edit by nick 2004/12/21 加了申請國家，要往後退
'            If grd1.TextMatrix(i, 8) = "N" Then
'                grd1.TextMatrix(i, 9) = "******"
'                grd1.TextMatrix(i, 11) = "******"
'                grd1.TextMatrix(i, 12) = ""
            If grd1.TextMatrix(i, 9) = "N" Then
                'Modify by Morgan 2011/1/4 修正日期排序問題前面補空白
                grd1.TextMatrix(i, 10) = " ******"
                grd1.TextMatrix(i, 12) = " ******"
                grd1.TextMatrix(i, 13) = ""
            End If
            '若墨圖不計件
            'edit by nick 2004/12/21 加了申請國家，要往後退
'            If grd1.TextMatrix(i, 13) = "N" Then
'                grd1.TextMatrix(i, 14) = "******"
'                grd1.TextMatrix(i, 16) = "******"
'                grd1.TextMatrix(i, 17) = ""
            If grd1.TextMatrix(i, 14) = "N" Then
                'Modify by Morgan 2011/1/4 修正日期排序問題前面補空白
                grd1.TextMatrix(i, 15) = " ******"
                grd1.TextMatrix(i, 17) = " ******"
                grd1.TextMatrix(i, 18) = ""
            End If
        Next i
        Me.Enabled = True
        Screen.MousePointer = vbDefault
        grd1.MousePointer = flexDefault
        SetGrd1
        grd1.Visible = True
        ChkNoData = False
    Else
        If (ProState = "2" Or ProState = "4") And pub_QL04 <> "" Then
            InsertQueryLog (0) 'Add By Sindy 2010/12/16
            pub_QL04 = ""
        End If
        'Add By Cheng 2002/04/26
        Me.grd1.Clear
        Me.grd1.Rows = 2
        Screen.MousePointer = vbDefault
        grd1.MousePointer = flexDefault
        SetGrd1
        ChkNoData = True
    End If
End With
CheckOC
Screen.MousePointer = vbDefault
grd1.MousePointer = flexDefault
End Sub

'Modify By Cheng 2003/09/17
'Sub ChgGrdColor()
Sub ChgGrdColor(blnOneRow As Boolean)
'blnOneRow 是否只改變一列的顏色
Dim tmpcolor1 As Integer
Dim tmpcolor2 As Integer
Dim jj As Integer
'Add By Cheng 2004/02/13
Dim blnRoughOver As Boolean '草圖是否逾時
Dim blnInkOver As Boolean '墨圖是否逾時
'End

'Add by Morgan 2004/5/19
'專利種類
Dim stPA08 As String

With grd1
    .Visible = False
'    For i = 1 To grd1.Rows - 1
    For i = IIf(blnOneRow = True, SWPRow, 1) To IIf(blnOneRow = True, SWPRow, grd1.Rows - 1)
        .row = i
        If blnOneRow = True Then
            For jj = 0 To .Cols - 1
                .col = jj
                .CellBackColor = QBColor(15)
            Next jj
        End If
        '法定期限
        'edit by nick 2004/12/21 加了申請國家，要往後退
        '.col = 24
        .col = 25
'        '若有法定期限
        'Modify by Morgan +判斷未發文
        'If .Text <> "" Then
        If .Text <> "" And .TextMatrix(.row, 20) = "" Then
            '若法定期限 = 系統日
            If DBDATE(.Text) = strSrvDate(1) Then
                For j = 2 To .Cols - 1
                    .col = j
                    .CellBackColor = &H8080FF '淺紅色
                Next j
            End If
        End If
'        '若無法定期限
'        Else
            '取消收文日
            'edit by nick 2004/12/21 加了申請國家，要往後退
            '.col = 25
            .col = 26
            '若有取消收文日
            If .Text <> "" Then
                For j = 2 To .Cols - 1
                    .col = j
                    .CellBackColor = &HC0C0C0 '灰色
                Next j
            End If
'            '若無取消收文日
'            Else
                
                '設計？
                'edit by nick 2004/12/21 加了申請國家，要往後退
                '.col = 33
                .col = 34
                strCP10 = Trim(.Text)
                
                '專利種類
                'Add by Morgan 2004/5/19
                'edit by nick 2004/12/21 加了申請國家，要往後退
                'stPA08 = .TextMatrix(.Row, 34)
                stPA08 = .TextMatrix(.row, 35)
                
                '發文日
                'edit by nick 2004/12/21 加了申請國家，要往後退
                '.col = 19
                .col = 20
                StrColor1 = .Text
                '取消收文日
                'edit by nick 2004/12/21 加了申請國家，要往後退
                '.col = 25
                .col = 26
                StrColor2 = .Text
                '草圖完稿
                'edit by nick 2004/12/21 加了申請國家，要往後退
                '.col = 11
                .col = 12
                StrColor3 = Replace(.Text, "******", "")
                '墨圖完稿
                'edit by nick 2004/12/21 加了申請國家，要往後退
                '.col = 16
                .col = 17
                StrColor4 = Replace(.Text, "******", "")
                '草圖齊備
                'edit by nick 2004/12/21 加了申請國家，要往後退
                '.col = 9
                .col = 10
                StrColor5 = Replace(.Text, "******", "")
                '墨圖齊備
                'edit by nick 2004/12/21 加了申請國家，要往後退
                '.col = 14
                .col = 15
                StrColor6 = Replace(.Text, "******", "")
                '草圖作業天數
                'edit by nick 2004/12/21 加了申請國家，要往後退
                '.col = 12
                .col = 13
                tmpcolor1 = Val(.Text)
                '墨圖作業天數
                'edit by nick 2004/12/21 加了申請國家，要往後退
                '.col = 17
                .col = 18
                tmpcolor2 = Val(.Text)
                
                'Modify by Morgan 2004/5/19
                '改依專利種類判斷
'                Select Case StrCp10
'                Case "103", "105" '設計申請
                Select Case stPA08
                  Case "3"
                
                    'nick  91/04/10
                    If tmpcolor1 > 5 Or tmpcolor2 > 3 Then
                        '判斷草圖或墨圖是否逾時
                        blnRoughOver = False: blnInkOver = False
                        '草圖
                        If tmpcolor1 > 5 Then
                            blnRoughOver = True
                        End If
                        '墨圖
                        If tmpcolor2 > 3 Then
                            blnInkOver = True
                        End If
                    Else
                        '判斷草圖或墨圖是否逾時
                        blnRoughOver = False: blnInkOver = False
'                        '無發文日, 無取消收文日, 無草完日
'                        If Len(Trim(StrColor1)) = 0 And Len(Trim(StrColor2)) = 0 And Len(Trim(StrColor3)) = 0 Then
                        '無草完日
                        If Len(Trim(StrColor3)) = 0 Then
                            '若有草齊日
                            If Len(Trim(StrColor5)) <> 0 Then
                                '若系統日超過草齊日5個工作天
                                If GetWorkDay(strSrvDate(1), ChangeTStringToWString(ChangeTDateStringToTString(StrColor5))) > 5 Then
                                    blnRoughOver = True
                                End If
                            End If
                        End If
'                        '無發文日, 無取消收文日, 無墨完日
'                        If Len(Trim(StrColor1)) = 0 And Len(Trim(StrColor2)) = 0 And Len(Trim(StrColor4)) = 0 Then
                        '無墨完日
                        If Len(Trim(StrColor4)) = 0 Then
                            '若有墨齊日
                            If Len(Trim(StrColor6)) <> 0 Then
                                '若系統日大於墨齊日3個工作天
                                If GetWorkDay(strSrvDate(1), ChangeTStringToWString(ChangeTDateStringToTString(StrColor6))) > 3 Then
                                    blnInkOver = True
                                End If
                            End If
                        End If
                    End If
                    If blnRoughOver = True Or blnInkOver = True Then
                        For j = 2 To .Cols - 1
                            'edit by nick 2004/12/21 加了申請國家，要往後退
                            'If j >= 8 And j <= 12 Then
                            If j >= 9 And j <= 13 Then
                                If blnRoughOver = True Then
                                    .col = j
                                    .CellBackColor = &H80FFFF '黃色
                                End If
                            'edit by nick 2004/12/21 加了申請國家，要往後退
                            'ElseIf j >= 13 And j <= 17 Then
                            ElseIf j >= 14 And j <= 18 Then
                                If blnInkOver = True Then
                                    .col = j
                                    .CellBackColor = &H80FFFF '黃色
                                End If
                            Else
'                                .Col = j
'                                .CellBackColor = &H80FFFF '黃色
                            End If
                        Next j
                    End If
                    
                Case Else '非設計申請
                
                    '91/04/10    ncik
                    If tmpcolor1 > 4 Or tmpcolor2 > 3 Then
                        '判斷草圖或墨圖是否逾時
                        blnRoughOver = False: blnInkOver = False
                        '草圖
                        If tmpcolor1 > 4 Then
                            blnRoughOver = True
                        End If
                        '墨圖
                        If tmpcolor2 > 3 Then
                            blnInkOver = True
                        End If
                    Else
                        '判斷草圖或墨圖是否逾時
                        blnRoughOver = False: blnInkOver = False
'                        '無發文日, 無取消收文日, 無草完日
'                        If Len(Trim(StrColor1)) = 0 And Len(Trim(StrColor2)) = 0 And Len(Trim(StrColor3)) = 0 Then
                        '無草完日
                        If Len(Trim(StrColor3)) = 0 Then
                            '若有草齊日
                            If Len(Trim(StrColor5)) <> 0 Then
                                '若系統日超過草齊日4個工作天
                                If GetWorkDay(strSrvDate(1), ChangeTStringToWString(ChangeTDateStringToTString(StrColor5))) > 4 Then
                                    blnRoughOver = True
                                End If
                            End If
                        End If
'                        '無發文日, 無取消收文日, 無墨完日
'                        If Len(Trim(StrColor1)) = 0 And Len(Trim(StrColor2)) = 0 And Len(Trim(StrColor4)) = 0 Then
                        '無墨完日
                        If Len(Trim(StrColor4)) = 0 Then
                            '若有墨齊日
                            If Len(Trim(StrColor6)) <> 0 Then
                                '若系統日大於墨齊日3個工作天
                                If GetWorkDay(strSrvDate(1), ChangeTStringToWString(ChangeTDateStringToTString(StrColor6))) > 3 Then
                                    blnInkOver = True
                                End If
                            End If
                        End If
                    End If
                    If blnRoughOver = True Or blnInkOver = True Then
                        For j = 2 To .Cols - 1
                            'edit by nick 2004/12/21 加了申請國家，要往後退
                            'If j >= 8 And j <= 12 Then
                            If j >= 9 And j <= 13 Then
                                If blnRoughOver = True Then
                                    .col = j
                                    .CellBackColor = &H80FFFF '黃色
                                End If
                            'edit by nick 2004/12/21 加了申請國家，要往後退
                            'ElseIf j >= 13 And j <= 17 Then
                            ElseIf j >= 14 And j <= 18 Then
                                If blnInkOver = True Then
                                    .col = j
                                    .CellBackColor = &H80FFFF '黃色
                                End If
                            Else
'                                .Col = j
'                                .CellBackColor = &H80FFFF '黃色
                            End If
                        Next j
                    End If
                    
                End Select
                'Add By Cheng 2003/06/30
'                '若草圖計件, 有文齊日但無草齊日
                '若草圖計件, 有文齊日但無草齊日, 且未取消收文
'                If .TextMatrix(i, 8) = "" And .TextMatrix(i, 7) <> "" And .TextMatrix(i, 9) = "" Then
                'edit by nick 2004/12/21 加了申請國家，要往後退
                'If .TextMatrix(i, 8) = "" And .TextMatrix(i, 7) <> "" And .TextMatrix(i, 9) = "" And StrColor2 = "" Then
                If Trim(.TextMatrix(i, 9)) = "" And Trim(.TextMatrix(i, 8)) <> "" And Trim(.TextMatrix(i, 10)) = "" And StrColor2 = "" Then
                    'edit by nick 2004/12/21 加了申請國家，要往後退
                    '.col = 9
                    .col = 10
                    .CellBackColor = &HFF8080 '淺藍色
                End If
'                '若草圖計件, 有草齊日但無草完日
                '若草圖計件, 有草齊日但無草完日, 且未取消收文
'                If .TextMatrix(i, 8) = "" And .TextMatrix(i, 9) <> "" And .TextMatrix(i, 11) = "" Then
                 'edit by nick 2004/12/21 加了申請國家，要往後退
                'If .TextMatrix(i, 8) = "" And .TextMatrix(i, 9) <> "" And .TextMatrix(i, 11) = "" And StrColor2 = "" Then
                'edit by nickc 2005/04/13 計件欄位加計件值
                'If .TextMatrix(i, 9) = "" And .TextMatrix(i, 10) <> "" And .TextMatrix(i, 12) = "" And StrColor2 = "" Then
                If .TextMatrix(i, 9) <> "N" And Trim(.TextMatrix(i, 10)) <> "" And Trim(.TextMatrix(i, 12)) = "" And StrColor2 = "" Then
                    'edit by nick 2004/12/21 加了申請國家，要往後退
                    '.col = 11
                    .col = 12
                    'Modify by Morgan 2011/3/29 計件值小於 0.6 的改用正藍色表示
                     If Val(.TextMatrix(i, 9)) <= 0.6 Then
                        .CellBackColor = &HFFBF00 '天藍
                     Else
                        .CellBackColor = &HFF80FF '粉紅色
                     End If
                End If
'                '若墨圖計件, 有墨齊日但無墨完日
                '若墨圖計件, 有墨齊日但無墨完日, 且未取消收文
'                If .TextMatrix(i, 13) = "" And .TextMatrix(i, 14) <> "" And .TextMatrix(i, 16) = ""  Then
                'edit by nick 2004/12/21 加了申請國家，要往後退
                'If .TextMatrix(i, 13) = "" And .TextMatrix(i, 14) <> "" And .TextMatrix(i, 16) = "" And StrColor2 = "" Then
                'edit by nickc 2005/04/13 計件欄位加計件值
                'If .TextMatrix(i, 14) = "" And .TextMatrix(i, 15) <> "" And .TextMatrix(i, 17) = "" And StrColor2 = "" Then
                If .TextMatrix(i, 14) <> "N" And Trim(.TextMatrix(i, 15)) <> "" And Trim(.TextMatrix(i, 17)) = "" And StrColor2 = "" Then
                    'edit by nick 2004/12/21 加了申請國家，要往後退
                    '.col = 16
                    .col = 17
                     'Modify by Morgan 2011/3/29 計件值小於 0.6 的改用正藍色表示
                     If Val(.TextMatrix(i, 14)) <= 0.6 Then
                        .CellBackColor = &HFFBF00 '天藍
                     Else
                        .CellBackColor = &HFF80FF '粉紅色
                     End If
                End If
'            End If
'        End If
    Next i
    .Visible = True
End With
End Sub

Private Sub SetGrd1()

With grd1
    .Visible = False
    'Modify by Morgan 2004/5/19
    '加專利種類 PA08
    '.Cols = 34
    'edit by nick 2004/12/21
    '.Cols = 35
    .Cols = 37
    
    .row = 0
    .RowHeight(0) = 400
    .col = 0:   .Text = "類別"
    .ColWidth(0) = 300
    .CellAlignment = flexAlignCenterCenter
    .col = 1:   .Text = "收文日"
    .ColWidth(1) = 795
    .ColAlignment(1) = flexAlignRightCenter
    .CellAlignment = flexAlignCenterCenter
'edit by nick 2004/12/21  加申請國家

    .col = 2:   .Text = "申國"
    .ColWidth(2) = 400
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
    .col = 8:   .Text = "文齊日"
    .ColWidth(8) = 795
    .ColAlignment(8) = flexAlignRightCenter
    .CellAlignment = flexAlignCenterCenter
    .col = 9:   .Text = "草計"
    .ColWidth(9) = 400
    .CellAlignment = flexAlignCenterCenter
    .col = 10:   .Text = "草齊日"
    .ColWidth(10) = 795
    .ColAlignment(10) = flexAlignRightCenter
    .CellAlignment = flexAlignCenterCenter
    .col = 11:   .Text = "草期限"
'    .ColWidth(10) = 700
    .ColWidth(11) = 0
    .CellAlignment = flexAlignCenterCenter
    .col = 12:  .Text = "草完日"
    .ColWidth(12) = 795
    .ColAlignment(12) = flexAlignRightCenter
    .CellAlignment = flexAlignCenterCenter
    .col = 13:  .Text = "草天"
    .ColWidth(13) = 400
    .CellAlignment = flexAlignCenterCenter
    .col = 14:   .Text = "墨計"
    .ColWidth(14) = 400
    .CellAlignment = flexAlignCenterCenter
    .col = 15:  .Text = "墨齊日"
    .ColWidth(15) = 795
    .ColAlignment(15) = flexAlignRightCenter
    .CellAlignment = flexAlignCenterCenter
    .col = 16:  .Text = "墨期限"
'    .ColWidth(15) = 700
    .ColWidth(16) = 0
    .CellAlignment = flexAlignCenterCenter
    .col = 17:  .Text = "墨完日"
    .ColWidth(17) = 795
    .ColAlignment(17) = flexAlignRightCenter
    .CellAlignment = flexAlignCenterCenter
    .col = 18:  .Text = "墨天"
    .ColWidth(18) = 400
    .CellAlignment = flexAlignCenterCenter
    .col = 19:  .Text = "本所期限"
    .ColWidth(19) = 795
    .ColAlignment(19) = flexAlignRightCenter
    .CellAlignment = flexAlignCenterCenter
    .col = 20:  .Text = "發文日"
    .ColWidth(20) = 795
    .ColAlignment(20) = flexAlignRightCenter
    .CellAlignment = flexAlignCenterCenter
    .col = 21:  .Text = "備註"
    .ColWidth(21) = 800
    .CellAlignment = flexAlignCenterCenter
    .col = 22:  .Text = "智權人員"
    .ColWidth(22) = 700
    .CellAlignment = flexAlignCenterCenter
    .col = 23:  .Text = "" '收文號
    .ColWidth(23) = 0
    .CellAlignment = flexAlignCenterCenter
    .col = 24:  .Text = "" '案件性質名稱
    .ColWidth(24) = 0
    .CellAlignment = flexAlignCenterCenter
    .col = 25:  .Text = "" '法定期限
    .ColWidth(25) = 0
    .CellAlignment = flexAlignCenterCenter
    .col = 26:  .Text = "" '取消收文日
    .ColWidth(26) = 0
    .CellAlignment = flexAlignCenterCenter
    .col = 27:  .Text = "" '草圖承辦時數
    .ColWidth(27) = 0
    .CellAlignment = flexAlignCenterCenter
    .col = 28:  .Text = "" '墨圖承辦時數
    .ColWidth(28) = 0
    .CellAlignment = flexAlignCenterCenter
    .col = 29:  .Text = "" '修改時數1
    .ColWidth(29) = 0
    .CellAlignment = flexAlignCenterCenter
    .col = 30:  .Text = "" '修改時數2
    .ColWidth(30) = 0
    .CellAlignment = flexAlignCenterCenter
    .col = 31:  .Text = "" '修改時數3
    .ColWidth(31) = 0
    .CellAlignment = flexAlignCenterCenter
    .col = 32:  .Text = "" '草圖張數
    .ColWidth(32) = 0
    .CellAlignment = flexAlignCenterCenter
    .col = 33:  .Text = "" '墨圖張數
    .ColWidth(33) = 0
    .CellAlignment = flexAlignCenterCenter
    .col = 34:  .Text = "" '案件性質代號
    .ColWidth(34) = 0
    .CellAlignment = flexAlignCenterCenter
    
    'Add by Morgan 2004/5/19
    .col = 35:  .Text = "" '專利種類代號
    .ColWidth(35) = 0
    .CellAlignment = flexAlignCenterCenter
    
    'add by nick 2004/12/22
    .col = 36:  .Text = "" '排序用
    .ColWidth(36) = 0
    .CellAlignment = flexAlignCenterCenter
    .Visible = True
End With
'ChgGrdColor
ChgGrdColor False
End Sub

Private Sub GRD1_DblClick()
If Me.grd1.MouseRow > 0 Then
    SSTab1.Tab = 1
    'Modify By Cheng 2004/04/19
'    cmdOK(2).Caption = "存檔(&O)"
    cmdOK(2).Caption = "存檔"
    'End
End If
End Sub

Private Sub grd1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Me.grd1.MouseRow > 0 Then
    If Button = 1 Then
       SWPRow = str(grd1.MouseRow)
       MouseClick Val(SWPRow)
    End If
End If
End Sub

'Modify By Sindy 2013/6/10
'Sub MouseClick(Optional Strindex As Integer)
Public Sub MouseClick(Optional Strindex As Integer)
'2013/6/10 End
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
'            '草齊日
'            .col = 10
            'edit by nick 2004/12/21 加了申請國家，要往後退
            '.col = 8
            .col = 9
            .CellBackColor = QBColor(15) '白色
            '墨計
'            '草齊日
'            .col = 10
            'edit by nick 2004/12/21 加了申請國家，要往後退
            '.col = 13
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
'    .col = 20
    'edit by nick 2004/12/21 加了申請國家，要往後退
    '.col = 22
    .col = 23
    Process (.Text)
    For i = 0 To 1
        .col = i
        .CellBackColor = &HFFC0C0
    Next i
    '草計
'    '草齊日
'    .col = 10
    'edit by nick 2004/12/21 加了申請國家，要往後退
    '.col = 8
    .col = 9
    .CellBackColor = &HFFC0C0
    '墨計
    'edit by nick 2004/12/21 加了申請國家，要往後退
    '.col = 13
    .col = 14
    .CellBackColor = &HFFC0C0
    .Visible = True
End With
Combo2.Enabled = False
End Sub

Private Sub grd1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    'Add By Cheng 2004/02/26
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

Private Sub Option1_Click(Index As Integer)
If Option1(0).Value = True Then
    Combo2.Enabled = False
    Txt1(13).Enabled = True
Else
    Combo2.Enabled = True
    Txt1(13).Enabled = False
End If
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
   If SSTab1.Tab = 0 Then
       'Modify By Cheng 2004/04/19
   '    cmdOK(2).Caption = "確定(&O)"
      cmdOK(2).Caption = "確定"
       'End
   ElseIf SSTab1.Tab = 1 Then
       'Modify By Cheng 2004/04/19
   '   cmdOK(2).Caption = "存檔(&O)"
      cmdOK(2).Caption = "存檔"
       'End
       If Me.Txt1(7).Enabled = True Then Me.Txt1(7).SetFocus
       'Add By Cheng 2004/03/17
       Me.Option1(0).Value = True
       Me.Option1(1).Value = False
       'End
   End If
   'Add By Sindy 2013/6/10
   If Me.SSTab1.TabVisible(2) = True Then
   If SSTab1.Tab = 2 Then
      Call QueryData(False)
   End If
   If PreviousTab = 0 Or PreviousTab = 1 Then
      '若有資料
      If Me.grd1.Rows > 1 Then
         '若點選的那筆無資料, 則退出函式
         If Me.grd1.TextMatrix(Val(SWPRow), 1) = "" Then SSTab1.Tab = 0: Exit Sub
         If Val(SWPRow) > 0 Then
            '上一筆資料列清除反白
            'Modify By Sindy 2016/5/9
            'If dblPrevRow > 0 Then
            If dblPrevRow > 0 And dblPrevRow <= (grd2.Rows - 1) Then
            '2016/5/9 END
               grd2.col = 0
               grd2.row = dblPrevRow
               grd2.Text = ""
               For ii = 0 To grd2.Cols - 1
                  grd2.col = ii
                  grd2.CellBackColor = QBColor(15)
               Next ii
               dblPrevRow = 0
            End If
            For i = 1 To Me.grd2.Rows - 1
               If Me.grd2.TextMatrix(i, 13) = Me.grd1.TextMatrix(Val(SWPRow), 23) Then
                  '目前資料列反白
                  dblPrevRow = i
                  grd2.col = 0
                  grd2.row = dblPrevRow
                  If grd2.TextMatrix(grd2.row, 1) <> "" Then
                     grd2.Text = "V"
                     For ii = 0 To grd2.Cols - 1
                        grd2.col = ii
                        grd2.CellBackColor = &HFFC0C0
                     Next ii
                  End If
                  Exit For
               End If
            Next i
         Else
            '目前資料列反白
            'Modify By Sindy 2016/5/9
            'If dblPrevRow > 0 Then
            If dblPrevRow > 0 And dblPrevRow <= (grd2.Rows - 1) Then
            '2016/5/9 END
               grd2.col = 0
               grd2.row = dblPrevRow
               If grd2.TextMatrix(grd2.row, 1) <> "" Then
                  grd2.Text = "V"
                  For ii = 0 To grd2.Cols - 1
                     grd2.col = ii
                     grd2.CellBackColor = &HFFC0C0
                  Next ii
               End If
            End If
         End If
      End If
   ElseIf PreviousTab = 2 Then
      '若有資料
      If (Me.grd2.Rows - 1) < dblPrevRow Then dblPrevRow = 0 'Add By Sindy 2024/7/10
      If Me.grd2.Rows > 1 And dblPrevRow > 0 Then
         If Me.grd2.TextMatrix(dblPrevRow, 1) <> "" Then
            For i = 1 To Me.grd1.Rows - 1
               If Me.grd2.TextMatrix(dblPrevRow, 13) = Me.grd1.TextMatrix(i, 23) Then
                  SWPRow = i
                  Exit For
               End If
            Next i
            MouseClick Val(SWPRow)
         End If
      End If
   End If
   End If
   '2013/6/10 End
End Sub

Private Sub Text1_GotFocus()
Text1.SelStart = 0
Text1.SelLength = Len(Text1)
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text1_LostFocus()
If IsNumeric(Text1) = False Then
    s = MsgBox("發文年月只能輸入數字!!", , "USER 輸入錯誤")
    Text1.SetFocus
    Text1.SelStart = 0
    Text1.SelLength = Len(Text1)
    Exit Sub
End If
StrMenu
End Sub

Private Sub txt1_Change(Index As Integer)
    'Add By Cheng 2003/06/27
    Select Case Index
    Case 7 '草圖是否計件
        If Me.Txt1(Index).Text = "N" Then
            Me.Txt1(1).Enabled = False
            Me.Txt1(2).Enabled = False
            Me.Txt1(1).Text = "******"
            Me.Txt1(2).Text = "******"
        Else
            Me.Txt1(1).Enabled = True
            Me.Txt1(2).Enabled = True
            Me.Txt1(1).Text = ""
            Me.Txt1(2).Text = ""
        End If
    Case 14 '墨圖否計件
        If Me.Txt1(Index).Text = "N" Then
            Me.Txt1(4).Enabled = False
            Me.Txt1(5).Enabled = False
            Me.Txt1(4).Text = "******"
            Me.Txt1(5).Text = "******"
        Else
            Me.Txt1(4).Enabled = True
            Me.Txt1(5).Enabled = True
            Me.Txt1(4).Text = ""
            Me.Txt1(5).Text = ""
        End If
    End Select
End Sub

Private Sub txt1_GotFocus(Index As Integer)
Txt1(Index).SelStart = 0
Txt1(Index).SelLength = Len(Txt1(Index))
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As MSForms.ReturnInteger)
    KeyAscii = UpperCase(KeyAscii)
    Select Case Index
    Case 7, 14
        If KeyAscii <> 8 And KeyAscii <> 78 Then
            KeyAscii = 0
        End If
    'add by nickc 2005/04/12
    Case 19
        If KeyAscii <> 8 And KeyAscii <> 89 Then
               KeyAscii = 0
        End If
    Case Else
    End Select
End Sub

Private Sub txt1_LostFocus(Index As Integer)
ChkData2 = True
Select Case Index
Case 0 '繪圖人員
     CheckOC2
     Txt1(0) = Trim(Txt1(0)) 'Add by Morgan 2005/1/14
     '92.04.03 nick add left join
     'strSQL = "SELECT S1.ST02 FROM STAFF S1,STAFF S2 WHERE S2.ST01='" & strUserNum & "' AND SUBSTR(S1.ST03,1,1) = SUBSTR(S2.ST03,1,1) AND S1.ST03='P13' AND S1.ST04='1' AND S2.ST04='1' AND S1.ST01='" & Trim(txt1(0)) & "' "
        'Modify By Cheng 2003/04/28
'     strSQL = "SELECT S1.ST02 FROM STAFF S1,STAFF S2 WHERE S2.ST01='" & strUserNum & "' AND SUBSTR(S1.ST03,1,1) = SUBSTR(S2.ST03,1,1)(+) AND S1.ST03='P13' AND S1.ST04='1' AND S2.ST04='1' AND S1.ST01='" & Trim(txt1(0)) & "' "
     strSql = "SELECT S1.ST02 FROM STAFF S1,STAFF S2 WHERE S2.ST01='" & strUserNum & "' AND SUBSTR(S1.ST03,1,1) = SUBSTR(S2.ST03,1,1) AND S1.ST05 in ('79','81','82','AC') AND S1.ST04='1' AND S2.ST04='1' AND S1.ST01='" & Trim(Txt1(0)) & "' "
     With adoRecordset1
        .CursorLocation = adUseClient
        .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
        If .RecordCount <> 0 And .RecordCount > 0 Then
            lbl1(0).Caption = CheckStr(.Fields(0))
        Else
            'add by nickc 2006/03/16 電腦中心不管
            If Pub_StrUserSt03 <> "M51" Then
                s = MsgBox("此 " & Txt1(0) & " 不屬於繪圖部門,並與 USER '" & strUserNum & "' 部門第一碼不同!!", , "USER 輸入錯誤")
                Txt1(0).SetFocus
                Txt1(0).SelStart = 0
                Txt1(0).SelLength = Len(Txt1(0))
                ChkData2 = False
                Exit Sub
            End If
        End If
     End With
Case 2, 3, 4, 5, 6, 8, 9, 10, 11, 12, 1
    'Add By Cheng 2003/06/27
    Select Case Index
    Case 1, 2, 4, 5
        If Me.Txt1(Index).Text = "******" Then Exit Sub
    Case Else
    End Select
     If Len(Txt1(Index)) <> 0 Then
        If IsNumeric(Txt1(Index)) = False Then
            s = MsgBox("請輸入數字!!", , "USER 輸入錯誤")
            Txt1(Index).SetFocus
            Txt1(Index).SelStart = 0
            Txt1(Index).SelLength = Len(Txt1(Index))
            ChkData2 = False
            Exit Sub
        End If
     End If
     If Index >= 8 And Index <= 12 Then
        If Val(Txt1(Index)) >= 100 Or Val(Txt1(Index)) < 0 Then
            s = MsgBox("時數輸入錯誤 0-99.9 !!", , "USER 輸入錯誤")
            Txt1(Index).SetFocus
            Txt1(Index).SelStart = 0
            Txt1(Index).SelLength = Len(Txt1(Index))
            ChkData2 = False
            Exit Sub
        End If
     End If
     If Index = 2 Or Index = 5 Or Index = 4 Or Index = 1 Then
         If Len(Txt1(Index)) <> 0 Then
            If Not ChkWorkDay(ChangeTStringToWString(Txt1(Index))) Then
               ShowDateErr
               Txt1(Index).SetFocus
               txt1_GotFocus (Index)
               ChkData2 = False
               Exit Sub
            End If
         End If
      End If
     If Index = 2 Then
         If Len(Txt1(1)) <> 0 And Len(Txt1(2)) <> 0 Then
            If RunNick2(Txt1(1), Txt1(2)) Then
               Txt1(1).SetFocus
               txt1_GotFocus (1)
               ChkData2 = False
               Exit Sub
            End If
         End If
     End If
     If Index = 5 Then
         If Len(Txt1(4)) <> 0 And Len(Txt1(5)) <> 0 Then
            If RunNick2(Txt1(4), Txt1(5)) Then
               Txt1(4).SetFocus
               txt1_GotFocus (4)
               ChkData2 = False
               Exit Sub
            End If
         End If
    End If
Case 7, 14
     Select Case Trim(Txt1(Index))
     Case "", "N"
     Case Else
          s = MsgBox("只能輸入 N 或空白!!", , "USER 輸入錯誤")
          Txt1(Index).SetFocus
          Txt1(Index).SelStart = 0
          Txt1(Index).SelLength = Len(Txt1(Index))
          ChkData2 = False
          Exit Sub
     End Select
'add by nickc 2005/04/12
Case 19
     Select Case Trim(Txt1(19))
     Case "", "Y"
     Case Else
          s = MsgBox("只能輸入 Y 或空白!!", , "USER 輸入錯誤")
          Txt1(Index).SetFocus
          Txt1(Index).SelStart = 0
          Txt1(Index).SelLength = Len(Txt1(Index))
          ChkData2 = False
          Exit Sub
     End Select
Case Else
End Select
End Sub

Private Function GetCP14EMail(strCP09 As String, strOfficeKind As String) As String
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset

GetCP14EMail = ""
StrSQLa = "Select CP14 From Caseprogress Where CP09='" & strCP09 & "' And CP14 Is Not Null "
rsA.CursorLocation = adUseClient
rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
If rsA.RecordCount > 0 Then
    'Modify By Cheng 2003/08/11
    '若使用者為北所人員, 則E-Mail後面不加@taie.com.tw
    If strOfficeKind = "1" Then
        GetCP14EMail = rsA.Fields(0).Value
    '若使用者非北所人員, 則E-Mail後面加@taie.com.tw
    Else
        GetCP14EMail = rsA.Fields(0).Value & "@taie.com.tw"
    End If
End If
If rsA.State <> adStateClosed Then rsA.Close
Set rsA = Nothing

End Function

'Add By Cheng 2003/09/17
'將修改後的資料更新回瀏覽畫面
Public Sub RefreshOneRecord()
Dim txt As Object
Dim Lbl As Object
Dim ii As Integer
    
    'Add By Cheng 2004/03/17
    '用收文號尋找瀏覽資料
    For ii = 1 To Me.grd1.Rows - 1
        'edit by nick 2004/12/21 加了申請國家，要往後退
        'If Me.lbl1(1).Caption = Me.grd1.TextMatrix(ii, 22) Then
        If Me.lbl1(1).Caption = Me.grd1.TextMatrix(ii, 23) Then
            SWPRow = ii
            Exit For
        End If
    Next ii
    'End
    If SWPRow <> "" And SWPRow <> "0" Then
        '草圖是否計件
        'edit by nick 2004/12/21 加了申請國家，要往後退
        'Me.grd1.TextMatrix(SWPRow, 8) = Me.txt1(7).Text
        'edit by nickc 2005/04/13  加入計件值
        'Me.grd1.TextMatrix(SWPRow, 9) = Me.txt1(7).Text
        Me.grd1.TextMatrix(SWPRow, 9) = IIf(Trim(Me.Txt1(7).Text) <> "N", Trim(Val(lbl1(33)) * Val(Txt1(15))), Txt1(7).Text)
        '草圖齊備日
        If Me.Txt1(7).Text = "N" Then
            'edit by nick 2004/12/21 加了申請國家，要往後退
            'Me.grd1.TextMatrix(SWPRow, 9) = "******"
            Me.grd1.TextMatrix(SWPRow, 10) = "******"
        Else
            'edit by nick 2004/12/21 加了申請國家，要往後退
            'Me.grd1.TextMatrix(SWPRow, 9) = ChangeTStringToTDateString(Me.txt1(1).Text)
            Me.grd1.TextMatrix(SWPRow, 10) = ChangeTStringToTDateString(Me.Txt1(1).Text)
        End If
        '草圖承辦期限
        'edit by nick 2004/12/21 加了申請國家，要往後退
        'Me.grd1.TextMatrix(SWPRow, 10) = Me.lbl1(31).Caption
        Me.grd1.TextMatrix(SWPRow, 11) = Me.lbl1(31).Caption
        '草圖完稿日
        If Me.Txt1(7).Text = "N" Then
            'edit by nick 2004/12/21 加了申請國家，要往後退
            'Me.grd1.TextMatrix(SWPRow, 11) = "******"
            Me.grd1.TextMatrix(SWPRow, 12) = "******"
        Else
            'edit by nick 2004/12/21 加了申請國家，要往後退
            'Me.grd1.TextMatrix(SWPRow, 11) = ChangeTStringToTDateString(Me.txt1(2).Text)
            Me.grd1.TextMatrix(SWPRow, 12) = ChangeTStringToTDateString(Me.Txt1(2).Text)
        End If
        '草圖張數
        'edit by nick 2004/12/21 加了申請國家，要往後退
        'Me.grd1.TextMatrix(SWPRow, 31) = Me.txt1(3).Text
        Me.grd1.TextMatrix(SWPRow, 32) = Me.Txt1(3).Text
        '墨圖是否計件
        'edit by nick 2004/12/21 加了申請國家，要往後退
        'Me.grd1.TextMatrix(SWPRow, 13) = Me.txt1(14).Text
        'edit by nickc 2005/04/13 加入計件值
        'Me.grd1.TextMatrix(SWPRow, 14) = Me.txt1(14).Text
        Me.grd1.TextMatrix(SWPRow, 14) = IIf(Trim(Me.Txt1(14).Text) <> "N", Trim(Val(lbl1(34)) * Val(Txt1(17))), Txt1(14).Text)
        '墨圖齊備日
        If Me.Txt1(14).Text = "N" Then
            'edit by nick 2004/12/21 加了申請國家，要往後退
            'Me.grd1.TextMatrix(SWPRow, 14) = "******"
            Me.grd1.TextMatrix(SWPRow, 15) = "******"
        Else
            'edit by nick 2004/12/21 加了申請國家，要往後退
            'Me.grd1.TextMatrix(SWPRow, 14) = ChangeTStringToTDateString(Me.txt1(4).Text)
            Me.grd1.TextMatrix(SWPRow, 15) = ChangeTStringToTDateString(Me.Txt1(4).Text)
        End If
        '墨圖承辦期限
        'edit by nick 2004/12/21 加了申請國家，要往後退
        'Me.grd1.TextMatrix(SWPRow, 15) = Me.lbl1(32).Caption
        Me.grd1.TextMatrix(SWPRow, 16) = Me.lbl1(32).Caption
        '墨圖完稿日
        If Me.Txt1(14).Text = "N" Then
            'edit by nick 2004/12/21 加了申請國家，要往後退
            'Me.grd1.TextMatrix(SWPRow, 16) = "******"
            Me.grd1.TextMatrix(SWPRow, 17) = "******"
        Else
            'edit by nick 2004/12/21 加了申請國家，要往後退
            'Me.grd1.TextMatrix(SWPRow, 16) = ChangeTStringToTDateString(Me.txt1(5).Text)
            Me.grd1.TextMatrix(SWPRow, 17) = ChangeTStringToTDateString(Me.Txt1(5).Text)
        End If
        '墨圖張數
        'edit by nick 2004/12/21 加了申請國家，要往後退
        'Me.grd1.TextMatrix(SWPRow, 32) = Me.txt1(6).Text
        Me.grd1.TextMatrix(SWPRow, 33) = Me.Txt1(6).Text
    
        '草圖承辦天數
        If Len(Txt1(1)) <> 0 And Len(Txt1(2)) <> 0 Then
            'edit by nick 2004/12/21 加了申請國家，要往後退
            'Me.grd1.TextMatrix(SWPRow, 12) = GetWorkDay(ChangeTStringToWString(txt1(2)), ChangeTStringToWString(txt1(1)))
            Me.grd1.TextMatrix(SWPRow, 13) = GetWorkDay(ChangeTStringToWString(Txt1(2)), ChangeTStringToWString(Txt1(1)))
        Else
            'edit by nick 2004/12/21 加了申請國家，要往後退
            'Me.grd1.TextMatrix(SWPRow, 12) = ""
            Me.grd1.TextMatrix(SWPRow, 13) = ""
        End If
        '墨圖承辦天數
        If Len(Txt1(4)) <> 0 And Len(Txt1(5)) <> 0 Then
            'edit by nick 2004/12/21 加了申請國家，要往後退
            'Me.grd1.TextMatrix(SWPRow, 17) = GetWorkDay(ChangeTStringToWString(txt1(5)), ChangeTStringToWString(txt1(4)))
            Me.grd1.TextMatrix(SWPRow, 18) = GetWorkDay(ChangeTStringToWString(Txt1(5)), ChangeTStringToWString(Txt1(4)))
        Else
            'edit by nick 2004/12/21 加了申請國家，要往後退
            'Me.grd1.TextMatrix(SWPRow, 17) = ""
            Me.grd1.TextMatrix(SWPRow, 18) = ""
        End If
        Me.grd1.TextMatrix(SWPRow, 18) = Me.grd1.TextMatrix(SWPRow, 18) & GetSign(grd1.TextMatrix(SWPRow, 23)) 'Add by Morgan 2009/10/7
    
        '承辦時數--草圖
        'edit by nick 2004/12/21 加了申請國家，要往後退
        'Me.grd1.TextMatrix(SWPRow, 26) = Me.txt1(8).Text
        Me.grd1.TextMatrix(SWPRow, 27) = Me.Txt1(8).Text
        '承辦時數--墨圖
        'edit by nick 2004/12/21 加了申請國家，要往後退
        'Me.grd1.TextMatrix(SWPRow, 27) = Me.txt1(9).Text
        Me.grd1.TextMatrix(SWPRow, 28) = Me.Txt1(9).Text
        '修改時數1
        'edit by nick 2004/12/21 加了申請國家，要往後退
        'Me.grd1.TextMatrix(SWPRow, 28) = Me.txt1(10).Text
        Me.grd1.TextMatrix(SWPRow, 29) = Me.Txt1(10).Text
        '修改時數2
        'edit by nick 2004/12/21 加了申請國家，要往後退
        'Me.grd1.TextMatrix(SWPRow, 29) = Me.txt1(11).Text
        Me.grd1.TextMatrix(SWPRow, 30) = Me.Txt1(11).Text
        '修改時數3
        'edit by nick 2004/12/21 加了申請國家，要往後退
        'Me.grd1.TextMatrix(SWPRow, 30) = Me.txt1(12).Text
        Me.grd1.TextMatrix(SWPRow, 31) = Me.Txt1(12).Text
    
        '備註
        If Me.Option1(0).Value = True Then
            'edit by nick 2004/12/21 加了申請國家，要往後退
            'Me.grd1.TextMatrix(SWPRow, 20) = Me.txt1(13).Text
            Me.grd1.TextMatrix(SWPRow, 21) = Me.Txt1(13).Text
        ElseIf Me.Option1(1).Value = True Then
            'edit by nick 2004/12/21 加了申請國家，要往後退
            'Me.grd1.TextMatrix(SWPRow, 20) = Me.Combo2.Text
            Me.grd1.TextMatrix(SWPRow, 21) = Me.Combo2.Text
        End If
    End If
    For Each txt In frm090711.Txt1
        txt.Text = ""
    Next
    For Each Lbl In frm090711.lbl1
        Lbl.Caption = ""
    Next
    ChgGrdColor True
End Sub

Private Sub txt1_Validate(Index As Integer, Cancel As Boolean)
Select Case Index
'add by nickc 2005/03/04
Case 15
       If Txt1(Index).Enabled = True Then
           If IsNumeric(Txt1(Index)) = False Then
               MsgBox "請輸入數字！", , "錯誤！"
               Cancel = True
               Exit Sub
           Else
               If Val(Txt1(Index)) > 3 Then
                  MsgBox "上限為 3 ，請重新輸入！", , "輸入錯誤！"
                  Cancel = True
                  Exit Sub
               End If
               If Val(Txt1(Index)) < 0 Then
                  MsgBox "下限為 0 ，請重新輸入！", , "輸入錯誤！"
                  Cancel = True
                  Exit Sub
               End If
           End If
           If Val(Txt1(Index)) <> Val(m_CP101) Then
               If Trim(Txt1(16)) = Trim(m_CP102) Then
                  Txt1(16).Text = ""
               End If
           Else
               Txt1(16).Text = m_CP102
           End If
       End If
Case 16
      If Txt1(Index).Enabled = True Then
         If CheckLengthIsOK(Txt1(Index), 100) = False Then
             MsgBox "最長為 50 個中文字！", , " 輸入錯誤！"
             Cancel = True
             Exit Sub
         End If
      End If
Case 17
       If Txt1(Index).Enabled = True Then
           If IsNumeric(Txt1(Index)) = False Then
               MsgBox "請輸入數字！", , "錯誤！"
               Cancel = True
               Exit Sub
           Else
               If Val(Txt1(Index)) > 3 Then
                  MsgBox "上限為 3 ，請重新輸入！", , "輸入錯誤！"
                  Cancel = True
                  Exit Sub
               End If
               If Val(Txt1(Index)) < 0 Then
                  MsgBox "下限為 0 ，請重新輸入！", , "輸入錯誤！"
                  Cancel = True
                  Exit Sub
               End If
           End If
           If Val(Txt1(Index)) <> Val(m_CP104) Then
               If Trim(Txt1(16)) = Trim(m_CP105) Then
                  Txt1(16).Text = ""
               End If
           Else
               Txt1(16).Text = m_CP105
           End If
       End If
Case 18
      If Txt1(Index).Enabled = True Then
         If CheckLengthIsOK(Txt1(Index), 100) = False Then
             MsgBox "最長為 50 個中文字！", , " 輸入錯誤！"
             Cancel = True
             Exit Sub
         End If
      End If
Case Else
End Select
End Sub

'add by nickc 2005/03/17
Private Function TxtValidate() As Boolean

TxtValidate = False
If Txt1(15).Enabled = True Then
   If Val(Me.Txt1(15).Text) <> Val(m_CP101) Then
         If Trim(Txt1(16).Text) = "" Then
            MsgBox "修改過草圖加乘註記，請輸入理由!!!", vbExclamation + vbOKOnly
            Exit Function
         End If
   End If
End If
If Txt1(17).Enabled = True Then
   If Val(Me.Txt1(17).Text) <> Val(m_CP104) Then
         If Trim(Txt1(18).Text) = "" Then
            MsgBox "修改過墨圖加乘註記，請輸入理由!!!", vbExclamation + vbOKOnly
            Exit Function
         End If
   End If
End If
'add by nickc 2005/06/13
If Txt1(0).Enabled = True Then
   If Trim(Txt1(0)) = "" Then
      If MsgBox("沒有繪圖人員，請確定！", vbYesNo, "警告！") = vbNo Then
         Exit Function
      End If
   End If
End If
TxtValidate = True
End Function
'add by nickc 2007/08/03
Private Sub CmdPic_Click()
frmPic001.oCP01 = SystemNumber(lbl1(3), 1)
frmPic001.oCP02 = SystemNumber(lbl1(3), 2)
frmPic001.oCP03 = SystemNumber(lbl1(3), 3)
frmPic001.oCP04 = SystemNumber(lbl1(3), 4)
frmPic001.StrMenu
frmPic001.SetSeekCmdok 'Add by Amy 2018/07/19
frmPic001.Show vbModal
'add by nickc 2007/08/03 檢查有無代表圖
'Modify by Amy 2018/07/19  改寫至function
'strSql = "SELECT ibf01,ibf02 FROM imgbytefile WHERE ibf01='" & SystemNumber(lbl1(3), 1) & "' and ibf02='" & SystemNumber(lbl1(3), 2) & "' and ibf03='" & SystemNumber(lbl1(3), 3) & "' and ibf04='" & SystemNumber(lbl1(3), 4) & "' and ibf05='1' "
'CheckOC2
'adoRecordset1.CursorLocation = adUseClient
'adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
'If adoRecordset1.RecordCount <> 0 And adoRecordset1.RecordCount > 0 Then
If ChkImgByteFile(SystemNumber(lbl1(3), 1), SystemNumber(lbl1(3), 2), SystemNumber(lbl1(3), 3), SystemNumber(lbl1(3), 4)) = True Then
    cmdPic.Caption = "已設定代表圖(&I)"
    cmdPic.BackColor = &HC0FFC0
Else
    cmdPic.Caption = "未設定代表圖(&I)"
    cmdPic.BackColor = &HC0C0FF
End If
'CheckOC2
'end 2018/07/19
End Sub
'Add by Morgan 2009/10/5
'台灣專利申請案上會稿完成日時,若該案同時有辦大陸案,則"墨天"欄位註記為"+"
Private Function GetSign(stCP09 As String) As String
   Dim stSQL As String, adoRst As ADODB.Recordset, iR As Integer
   stSQL = "select 1 from engineerprogress,caseprogress c1 where ep02='" & stCP09 & "' and ep08>0" & _
      " and cp09(+)=ep02 and instr('" & CaseMapIn & "',cp10)>0" & _
      " and exists(select * from patent where pa01=cp01 and pa02=cp02 and pa03=cp03 and pa04=cp04 and pa09='000')" & _
      " and exists(select * from casemap,patent where cm10='0' and cm05=cp01 and cm06=cp02 and cm07=cp03 and cm08=cp04" & _
      " and pa01(+)=cm01 and pa02(+)=cm02 and pa03(+)=cm03 and pa04(+)=cm04 and pa09='020' and pa57 is null)"
   iR = 1
   Set adoRst = ClsLawReadRstMsg(iR, stSQL)
   If iR = 1 Then
      GetSign = "+"
   End If
End Function

'Add By Sindy 2013/6/7
Private Sub cmdQuery_Click()
   If QueryData(True) = False Then ShowNoData
End Sub

'Add By Sindy 2013/6/7
Public Function QueryData(bolFirst As Boolean) As Boolean
Dim rsTmp As New ADODB.Recordset
Dim strSql As String
Dim strQuyDate As String 'Add By Sindy 2013/9/17
Dim strVal As String 'Add By Sindy 2013/10/22
   
   m_blnColOrderAsc = True
   QueryData = True
   
   'Add By Sindy 2013/9/17
   If Combo5.ListIndex = 0 Then
      strQuyDate = CompWorkDay(3, strSrvDate(1), 1) '不含當天,3個工作天
   ElseIf Combo5.ListIndex = 1 Then
      strQuyDate = CompWorkDay(5, strSrvDate(1), 1) '不含當天,5個工作天
   ElseIf Combo5.ListIndex = 2 Then
      strQuyDate = CompWorkDay(7, strSrvDate(1), 1) '不含當天,7個工作天
   Else
      '全部
   End If
   '2013/9/17 END
   
   grd2.Clear
   SetGrd2
   
   Screen.MousePointer = vbHourglass
   
   'Modify By Sindy 2013/10/22
'   strVal = "(select * from EmpElectronProcess where eep01||eep02 in(select eep01||max(eep02) from EmpElectronProcess,caseprogress where eep01=cp09(+) and cp27 is null and cp57 is null and (EEP04 not in(" & EMP_待辦歷程查詢除外的狀態 & ") or (EEP04='" & EMP_聯絡 & "'" & IIf(strQuyDate <> "", " And EEP06>=" & strQuyDate, "") & ")) group by eep01) and EEP09 is null and EEP05='" & Trim(Left(Combo1.Text, 6)) & "') EmpElectronProcess,"
   'Modify By Sindy 2015/5/12 調整SQL增加查詢速度
'   strVal = "(select * from EmpElectronProcess where eep01||eep02 in(select eep01||max(eep02) from EmpElectronProcess,caseprogress where eep01=cp09(+) and cp27 is null and cp57 is null and EEP04 not in(" & EMP_待辦歷程查詢除外的狀態 & ") group by eep01) and EEP09 is null and EEP05='" & Trim(Left(Combo1.Text, 6)) & "'" & _
'            " union select e1.* from EmpElectronProcess e1,caseprogress where e1.eep01=cp09(+) and cp27 is null and cp57 is null and e1.EEP02 in (select max(eep02) from EmpElectronProcess where eep01=e1.eep01) and e1.EEP05='" & Trim(Left(Combo1.Text, 6)) & "' And e1.EEP04 in('" & EMP_聯絡 & "')" & IIf(strQuyDate <> "", " And EEP06>=" & strQuyDate, "") & _
'            ") EmpElectronProcess,"
   'Modify By Sindy 2016/3/9 增加EEP13='Y'
   strVal = "(select * from EmpElectronProcess where EEP13='Y' and EEP04 not in(" & EMP_待辦歷程查詢除外的狀態 & ") and EEP09 is null and EEP05='" & Trim(Left(Combo1.Text, 6)) & "'" & _
      " union select * from EmpElectronProcess where EEP13='Y' and EEP05='" & Trim(Left(Combo1.Text, 6)) & "' And EEP04='" & EMP_聯絡 & "'" & IIf(strQuyDate <> "", " And EEP06>=" & strQuyDate, "") & _
            ") EmpElectronProcess,"
'   strVal = "(select e1.* from EmpElectronProcess e1,caseprogress where e1.eep01=cp09(+) and cp27 is null and cp57 is null and e1.EEP02 =(select max(eep02) from EmpElectronProcess where eep01=e1.eep01 and EEP04 not in(" & EMP_待辦歷程查詢除外的狀態 & ")) and e1.EEP09 is null and e1.EEP05='" & Trim(Left(Combo1.Text, 6)) & "'" & _
'      " union select e1.* from EmpElectronProcess e1,caseprogress where e1.eep01=cp09(+) and cp27 is null and cp57 is null and e1.EEP02 =(select max(eep02) from EmpElectronProcess where eep01=e1.eep01) and e1.EEP05='" & Trim(Left(Combo1.Text, 6)) & "' And e1.EEP04 in('" & EMP_聯絡 & "')" & IIf(strQuyDate <> "", " And EEP06>=" & strQuyDate, "") & _
'            ") EmpElectronProcess,"
   '2015/5/12 END
   '2013/10/22 END
   'Modify By Sindy 2016/3/9 +不顯示
   'Modify By Sindy 2016/9/5 and cp57 is null and cp27 is null ==> and cp158=0 and cp159=0
   strSql = "Select ' ' as V,EP01 as 目次,SqlDateT(EEP06)||' '||sqltime(EEP07) as 流程日期,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as 本所案號,PA05||PA06||PA07 as 案件名稱," & _
            "NA03 as 國家,Decode(PA09,'000',PTM03,PTM04) as 種類,Decode(PA09,'000',CPM03,CPM04) as 案件性質," & _
            "SqlDateT (cp06) as 本所期限, s1.ST02 as 承辦人, SqlDateT(cp48) as 承辦期限, s2.ST02 as 智權人員, ac03 as 目前流程狀態," & _
            "EEP01 as 總收文號,EEP02 as 序號,' ' as 不顯示 From " & strVal & _
            "CaseProgress,EngineerProgress,Patent," & _
            "staff s1,staff s2,nation,CasePropertyMap,PatentTradeMarkMap,allcode" & _
            " Where EEP01=CP09(+)" & _
            " and cp158=0 and cp159=0" & _
            " And EEP01=EP02(+)" & _
            " And CP01=PA01(+) And CP02=PA02(+) And CP03=PA03(+) And CP04=PA04(+)" & _
            " And CP14=s1.ST01(+) And CP13=s2.ST01(+)" & _
            " And PA09=NA01(+)" & _
            " And CP01=CPM01(+) And CP10=CPM02(+)" & _
            " And '1'=PTM01(+) AND PA08=PTM02(+)" & _
            " And ac01='09' And EEP04=ac02(+)" & _
            " And EP13='" & Trim(Left(Combo1.Text, 6)) & "'"
   'Modify By Sindy 2013/11/21
   'strSql = strSql & " order by EP01 desc"
   strSql = strSql & " order by EEP06 desc,EEP07 desc"
   '2013/11/21 END
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      Set grd2.Recordset = rsTmp
   Else
      QueryData = False
      Screen.MousePointer = vbDefault
      rsTmp.Close
      Set rsTmp = Nothing
      Exit Function
   End If
   
   '若有資料時游標停在第一筆
   grd2.Visible = False
   grd2.col = 0
   grd2.row = 1
   If bolFirst = True Then
      If rsTmp.RecordCount > 0 Then
         dblPrevRow = grd2.row
         grd2.Text = "V"
         For i = 0 To grd2.Cols - 1
            grd2.col = i
            grd2.CellBackColor = &HFFC0C0
         Next i
      End If
   End If
   grd2.Visible = True
   
   rsTmp.Close
   Screen.MousePointer = vbDefault
   
EXITSUB:
   Set rsTmp = Nothing
End Function

'Add By Sindy 2013/6/7
Private Sub SetGrd2()
   Dim arrGridHeadText, arrGridHeadWidth
   Dim iRow As Integer
   
   'Modify By Sindy 2016/3/9 +不顯示
   arrGridHeadText = Array("V", "目次", "流程日期", "本所案號", "案件名稱", _
                           "國家", "種類", "案件性質", "本所期限", "承辦人", _
                           "承辦期限", "智權人員", "目前流程狀態", _
                           "總收文號", "序號", "不顯示")
   arrGridHeadWidth = Array(200, 0, 800, 1200, 1000, _
                            800, 450, 1000, 800, 600, _
                            800, 600, 800, _
                            0, 0, 600)
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

'Add By Sindy 2016/3/9 增加不顯示功能
Private Sub grd2_Click()
Dim intRow As Integer, intCol As Integer
   
   intRow = grd2.MouseRow
   intCol = grd2.MouseCol
   If intRow <> 0 Then
      If intCol = 15 Then '不顯示
         'Modify By Sindy 2020/9/10 送標號只有標號或繪圖判發時，才可以上不顯示，繪圖人員不可以手動上不顯示，也不會因為有其他歷程而消失。
         If grd2.TextMatrix(intRow, 13) <> "" And _
            grd2.TextMatrix(intRow, 12) <> "退回" And _
            grd2.TextMatrix(intRow, 12) <> "草修" And _
            grd2.TextMatrix(intRow, 12) <> "草核完" And _
            grd2.TextMatrix(intRow, 12) <> "送標號" Then
            grd2.TextMatrix(intRow, 15) = "V"
            If MsgBox("請再次確定不顯示 " & vbCrLf & grd2.TextMatrix(intRow, 3) & " " & grd2.TextMatrix(intRow, 12) & " 嗎？", vbExclamation + vbYesNo + vbDefaultButton2, "重要訊息！") = vbNo Then
               grd2.TextMatrix(intRow, 15) = ""
            Else
               strExc(0) = "update EmpElectronProcess set eep13=null" & _
                           " where eep01='" & grd2.TextMatrix(intRow, 13) & "'" & _
                             " and eep02=" & grd2.TextMatrix(intRow, 14)
               Pub_SeekTbLog strExc(0) 'Add By Sindy 2018/8/27
               cnnConnection.Execute strExc(0)
               grd2.RowHeight(intRow) = 0
            End If
         End If
      End If
   End If
End Sub

'Add By Sindy 2013/6/7
Private Sub grd2_DblClick()
Dim nFrm As Form
   
   'For i = 1 To grd2.Rows - 1
      If grd2.TextMatrix(dblPrevRow, 0) = "V" Then
         If lbl1(1) <> grd2.TextMatrix(dblPrevRow, 13) Then
            For ii = 1 To grd1.Rows - 1
               If grd1.TextMatrix(ii, 23) = grd2.TextMatrix(dblPrevRow, 13) Then
                  SWPRow = ii
                  MouseClick Val(SWPRow)
                  Exit For
               End If
            Next ii
         Else
            If Me.cmd(1).Enabled = True Then
               If SetColTag(False) = False Then
                  Call cmdOK_Click(2)
                  If m_chkcmdok1 = False Then Exit Sub
               End If
            End If
         End If
         If Me.cmd(1).Enabled = True Then
'            'Add By Sindy 2017/9/19
'            '檢查表單是否已開啟，若是，則關閉
'            For Each nFrm In Forms
'               If StrComp(nFrm.Name, "frm090202_2", vbTextCompare) = 0 Then
'                  Unload frm090202_2
'               End If
'            Next
'            '2017/9/19 END
            If PUB_ChkFormIsClose("frm090202_2") = False Then Exit Sub 'Add By Sindy 2020/1/17
            intBackTab = 2
            frm090202_2.Hide
            frm090202_2.m_EEP01 = grd2.TextMatrix(dblPrevRow, 13) '總收文號
            frm090202_2.m_FlowUserNum = Trim(Left("" & Combo1.Text, 6)) 'Add By Sindy 2013/9/12 案件流程所屬人員
            frm090202_2.intReceiveKind = 3 '3.繪圖人員工作進度
            frm090202_2.SetParent Me
            If frm090202_2.QueryData = True Then
               frm090202_2.Show
               Me.Hide
            End If
            'Exit For
         Else
            Me.SSTab1.Tab = 1
         End If
      End If
   'Next i
End Sub

'Add By Sindy 2013/6/7
Private Sub GRD2_SelChange()
grd2.Visible = False
'Add By Sindy 2016/3/9
If grd2.MouseRow = 0 Then
   '已選取的資料列清除反白
   For j = 1 To grd2.Rows - 1
      If grd2.TextMatrix(j, 0) = "V" Then
         grd2.col = 0
         grd2.row = j
         grd2.Text = ""
         For i = 0 To grd2.Cols - 1
            grd2.col = i
            grd2.CellBackColor = QBColor(15)
         Next i
         Exit For
      End If
   Next j
Else
'2016/3/9 END
   '上一筆資料列清除反白
   'Modify By Sindy 2016/5/9
   'If dblPrevRow > 0 Then
   If dblPrevRow > 0 And dblPrevRow <= (grd2.Rows - 1) Then
   '2016/5/9 END
      grd2.col = 0
      grd2.row = dblPrevRow
      grd2.Text = ""
      For i = 0 To grd2.Cols - 1
         grd2.col = i
         grd2.CellBackColor = QBColor(15)
      Next i
   End If
   '目前資料列反白
   grd2.col = 0
   grd2.row = grd2.MouseRow
   dblPrevRow = grd2.row
'   If GRD2.Text = "V" Then
'      GRD2.Text = ""
'      For i = 0 To GRD2.Cols - 1
'         GRD2.col = i
'         GRD2.CellBackColor = QBColor(15)
'      Next i
'   Else
      If grd2.TextMatrix(grd2.row, 1) <> "" Then
         grd2.Text = "V"
         For i = 0 To grd2.Cols - 1
            grd2.col = i
            grd2.CellBackColor = &HFFC0C0
         Next i
      End If
'   End If
End If
grd2.Visible = True
End Sub

'Add By Sindy 2013/6/7
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

'Add by Amy 2018/10/08 取得繪圖束度及月考核資料(參考frm090201_2.StrMenu1)
Private Sub GetMonthAssess()
    Dim RsQ As New ADODB.Recordset, RsQ2 As New ADODB.Recordset
    Dim strQ As String, strQ2 As String
    Dim intQ As Integer
    Dim theWorkDay As Integer '截到目前工作天
    
    theWorkDay = 0
    Me.lblCal1(0).Caption = "0.00 基數"
    Me.lblCal1(1).Caption = "0.00 %"
    Me.lblCal1(2).Caption = "0.00 基數"
    Me.lblCal1(3).Caption = "0.00 基數"
    
    Select Case ProState
        'Modify by Amy 2021/03/02 96021 查工作進度資料查詢會錯,因ProState=4,故加4
        Case "1", "4" '個人工作進度資料維護
            strQ = "Select Count(*) From WorkDay Where WD01>='" & Mid(strSrvDate(1), 1, 6) & "01' And WD01<='" & CompWorkDay(2, strSrvDate(1), 1) & "' Having Count(*)>0"
            strQ2 = strQ2 & " ma01='" & Trim(Left("" & Combo1.Text, 6)) & "' "
            strQ2 = strQ2 & " And ma02='" & Mid(strSrvDate(1), 1, 6) & "' And ma03 ='2' "
        'Modify by Amy 2021/03/02 避免其他隻ProState=3會錯,故加3
        Case "2", "3" '管理工作進度資料查詢
            strQ = "Select Count(*) From WorkDay Where WD01>='" & Trim((Val(frm090706.Txt1(3)) + 1911)) & Trim(Right(ChgNumByNick(frm090706.Txt1(4)), 2)) & "01' And WD01<='" & IIf(Trim((Val(frm090706.Txt1(3)) + 1911)) & Trim(Right(ChgNumByNick(frm090706.Txt1(4)), 2)) = Mid(strSrvDate(1), 1, 6), CompWorkDay(2, strSrvDate(1), 1), Trim((Val(frm090706.Txt1(3)) + 1911)) & Trim(Right(ChgNumByNick(frm090706.Txt1(4)), 2)) & "31") & "' Having Count(*)>0"
            strQ2 = strQ2 & " ma01='" & Trim(Left("" & Combo1.Text, 6)) & "'  "
            strQ2 = strQ2 & " and ma02='" & Trim((Val(frm090706.Txt1(3)) + 1911)) & Trim(Right(ChgNumByNick(frm090706.Txt1(4)), 2)) & "' and ma03='2' "
    End Select
    intQ = 1
    Set RsQ = ClsLawReadRstMsg(intQ, strQ)
    If intQ = 1 Then
        theWorkDay = Val(RsQ.Fields(0))
        strQ2 = "Select * From MonthAssess Where " & strQ2 & " Order by ma36 Desc"
        RsQ2.CursorLocation = adUseClient
        RsQ2.Open strQ2, cnnConnection, adOpenStatic, adLockReadOnly
        If RsQ2.RecordCount > 0 Then
            RsQ2.MoveFirst
            Do While RsQ2.EOF = False
                '墨圖
                If RsQ2.Fields("ma36") = "2" Then
                    lblCal1(0).Caption = RsQ2.Fields("ma33") & " 基數"
                '草圖
                Else
                    lblCal1(1).Caption = RsQ2.Fields("ma33") & " 基數" '累計草圖完成量
                    lblCal1(3).Caption = RsQ2.Fields("ma34") & " %" '草圖累計達成比例
                    If Val(RsQ2.Fields("ma04")) <> 0 Then
                        strExc(0) = Format(Val(RsQ2.Fields("ma33")) - Val(RsQ2.Fields("ma04")) / Val(RsQ2.Fields("ma05")) * theWorkDay, "0.00")
                        lblCal1(2).Caption = IIf(Val(strExc(0)) >= 0, "+", "") & strExc(0) & " 基數" '草圖目前進度
                    Else
                        lblCal1(2).Caption = "尚無目標"
                    End If
                End If
                RsQ2.MoveNext
            Loop
        End If
        RsQ2.Close 'Modify by Amy 2018/11/01
    End If
    RsQ.Close
End Sub
