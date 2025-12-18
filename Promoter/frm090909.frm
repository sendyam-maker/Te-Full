VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm090909 
   Caption         =   "外專工作進度資料維護"
   ClientHeight    =   6490
   ClientLeft      =   3400
   ClientTop       =   2950
   ClientWidth     =   10040
   ControlBox      =   0   'False
   FillColor       =   &H0080FFFF&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   6490
   ScaleWidth      =   10040
   Begin VB.CheckBox Check1 
      Caption         =   "發文後補分割建議"
      Height          =   250
      Left            =   5130
      TabIndex        =   108
      Top             =   6620
      Width           =   2200
   End
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   23
      Left            =   4710
      MaxLength       =   1
      TabIndex        =   80
      Top             =   330
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "結束(&X)"
      Height          =   400
      Index           =   2
      Left            =   8190
      TabIndex        =   11
      Top             =   50
      Width           =   1125
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "確定"
      Height          =   400
      Index           =   1
      Left            =   7530
      TabIndex        =   10
      Top             =   50
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   720
      TabIndex        =   13
      Text            =   "Text1"
      Top             =   120
      Visible         =   0   'False
      Width           =   915
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5960
      Left            =   30
      TabIndex        =   14
      Top             =   510
      Width           =   9990
      _ExtentX        =   17604
      _ExtentY        =   10513
      _Version        =   393216
      Style           =   1
      Tab             =   1
      TabHeight       =   520
      TabCaption(0)   =   "瀏覽資料"
      TabPicture(0)   =   "frm090909.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "cmdRead"
      Tab(0).Control(1)=   "cmdok2(0)"
      Tab(0).Control(2)=   "cmdok2(1)"
      Tab(0).Control(3)=   "Combo3"
      Tab(0).Control(4)=   "grd1"
      Tab(0).Control(5)=   "Combo1"
      Tab(0).Control(6)=   "Label3"
      Tab(0).Control(7)=   "Label1(0)"
      Tab(0).Control(8)=   "Label5"
      Tab(0).ControlCount=   9
      TabCaption(1)   =   "詳細資料"
      TabPicture(1)   =   "frm090909.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "lblCMboth"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "lblCM10"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label18"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label1(47)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "lblEApp"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Label1(41)"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Label1(35)"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "lblClose"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "Label1(12)"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "lbl1(21)"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "Label1(1)"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "lbl1(1)"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "lbl1(0)"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "Label1(8)"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "Label1(21)"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "Label1(20)"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "Label1(19)"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "Label1(18)"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).Control(18)=   "Label1(17)"
      Tab(1).Control(18).Enabled=   0   'False
      Tab(1).Control(19)=   "Label1(16)"
      Tab(1).Control(19).Enabled=   0   'False
      Tab(1).Control(20)=   "Label1(15)"
      Tab(1).Control(20).Enabled=   0   'False
      Tab(1).Control(21)=   "Label1(14)"
      Tab(1).Control(21).Enabled=   0   'False
      Tab(1).Control(22)=   "Label1(13)"
      Tab(1).Control(22).Enabled=   0   'False
      Tab(1).Control(23)=   "lbl1(29)"
      Tab(1).Control(23).Enabled=   0   'False
      Tab(1).Control(24)=   "lbl1(10)"
      Tab(1).Control(24).Enabled=   0   'False
      Tab(1).Control(25)=   "lbl1(16)"
      Tab(1).Control(25).Enabled=   0   'False
      Tab(1).Control(26)=   "lbl1(30)"
      Tab(1).Control(26).Enabled=   0   'False
      Tab(1).Control(27)=   "Label1(32)"
      Tab(1).Control(27).Enabled=   0   'False
      Tab(1).Control(28)=   "lbl1(3)"
      Tab(1).Control(28).Enabled=   0   'False
      Tab(1).Control(29)=   "lbl1(5)"
      Tab(1).Control(29).Enabled=   0   'False
      Tab(1).Control(30)=   "lbl1(7)"
      Tab(1).Control(30).Enabled=   0   'False
      Tab(1).Control(31)=   "lbl1(9)"
      Tab(1).Control(31).Enabled=   0   'False
      Tab(1).Control(32)=   "lbl1(11)"
      Tab(1).Control(32).Enabled=   0   'False
      Tab(1).Control(33)=   "lbl1(13)"
      Tab(1).Control(33).Enabled=   0   'False
      Tab(1).Control(34)=   "lbl1(15)"
      Tab(1).Control(34).Enabled=   0   'False
      Tab(1).Control(35)=   "lbl1(17)"
      Tab(1).Control(35).Enabled=   0   'False
      Tab(1).Control(36)=   "lbl1(19)"
      Tab(1).Control(36).Enabled=   0   'False
      Tab(1).Control(37)=   "lbl1(23)"
      Tab(1).Control(37).Enabled=   0   'False
      Tab(1).Control(38)=   "Label1(26)"
      Tab(1).Control(38).Enabled=   0   'False
      Tab(1).Control(39)=   "Label1(2)"
      Tab(1).Control(39).Enabled=   0   'False
      Tab(1).Control(40)=   "Label1(5)"
      Tab(1).Control(40).Enabled=   0   'False
      Tab(1).Control(41)=   "Label1(25)"
      Tab(1).Control(41).Enabled=   0   'False
      Tab(1).Control(42)=   "Label2"
      Tab(1).Control(42).Enabled=   0   'False
      Tab(1).Control(43)=   "Label1(7)"
      Tab(1).Control(43).Enabled=   0   'False
      Tab(1).Control(44)=   "LblCP113"
      Tab(1).Control(44).Enabled=   0   'False
      Tab(1).Control(45)=   "Label1(4)"
      Tab(1).Control(45).Enabled=   0   'False
      Tab(1).Control(46)=   "Label1(9)"
      Tab(1).Control(46).Enabled=   0   'False
      Tab(1).Control(47)=   "lblEP39"
      Tab(1).Control(47).Enabled=   0   'False
      Tab(1).Control(48)=   "lblEP42"
      Tab(1).Control(48).Enabled=   0   'False
      Tab(1).Control(49)=   "Combo4"
      Tab(1).Control(49).Enabled=   0   'False
      Tab(1).Control(50)=   "Combo2"
      Tab(1).Control(50).Enabled=   0   'False
      Tab(1).Control(51)=   "Combo6"
      Tab(1).Control(51).Enabled=   0   'False
      Tab(1).Control(52)=   "Label1(23)"
      Tab(1).Control(52).Enabled=   0   'False
      Tab(1).Control(53)=   "lbl1(8)"
      Tab(1).Control(53).Enabled=   0   'False
      Tab(1).Control(54)=   "Label1(31)"
      Tab(1).Control(54).Enabled=   0   'False
      Tab(1).Control(55)=   "Label1(29)"
      Tab(1).Control(55).Enabled=   0   'False
      Tab(1).Control(56)=   "Label1(6)"
      Tab(1).Control(56).Enabled=   0   'False
      Tab(1).Control(57)=   "lbl1(28)"
      Tab(1).Control(57).Enabled=   0   'False
      Tab(1).Control(58)=   "Label1(11)"
      Tab(1).Control(58).Enabled=   0   'False
      Tab(1).Control(59)=   "Label1(3)"
      Tab(1).Control(59).Enabled=   0   'False
      Tab(1).Control(60)=   "Label1(22)"
      Tab(1).Control(60).Enabled=   0   'False
      Tab(1).Control(61)=   "Label1(46)"
      Tab(1).Control(61).Enabled=   0   'False
      Tab(1).Control(62)=   "txtCP64"
      Tab(1).Control(62).Enabled=   0   'False
      Tab(1).Control(63)=   "txtEP12"
      Tab(1).Control(63).Enabled=   0   'False
      Tab(1).Control(64)=   "LblCP114"
      Tab(1).Control(64).Enabled=   0   'False
      Tab(1).Control(65)=   "Lbl926"
      Tab(1).Control(65).Enabled=   0   'False
      Tab(1).Control(66)=   "cboCP14"
      Tab(1).Control(66).Enabled=   0   'False
      Tab(1).Control(67)=   "SSTab2"
      Tab(1).Control(67).Enabled=   0   'False
      Tab(1).Control(68)=   "cmd(1)"
      Tab(1).Control(68).Enabled=   0   'False
      Tab(1).Control(69)=   "txt1(19)"
      Tab(1).Control(69).Enabled=   0   'False
      Tab(1).Control(70)=   "txt1(13)"
      Tab(1).Control(70).Enabled=   0   'False
      Tab(1).Control(71)=   "txt1(12)"
      Tab(1).Control(71).Enabled=   0   'False
      Tab(1).Control(72)=   "cmd1"
      Tab(1).Control(72).Enabled=   0   'False
      Tab(1).Control(73)=   "txt1(3)"
      Tab(1).Control(73).Enabled=   0   'False
      Tab(1).Control(74)=   "txt1(5)"
      Tab(1).Control(74).Enabled=   0   'False
      Tab(1).Control(75)=   "txt1(6)"
      Tab(1).Control(75).Enabled=   0   'False
      Tab(1).Control(76)=   "txt1(7)"
      Tab(1).Control(76).Enabled=   0   'False
      Tab(1).Control(77)=   "txt1(2)"
      Tab(1).Control(77).Enabled=   0   'False
      Tab(1).Control(78)=   "txt1(0)"
      Tab(1).Control(78).Enabled=   0   'False
      Tab(1).Control(79)=   "txt1(8)"
      Tab(1).Control(79).Enabled=   0   'False
      Tab(1).Control(80)=   "txt1(9)"
      Tab(1).Control(80).Enabled=   0   'False
      Tab(1).Control(81)=   "txt1(1)"
      Tab(1).Control(81).Enabled=   0   'False
      Tab(1).Control(82)=   "cmdLetter"
      Tab(1).Control(82).Enabled=   0   'False
      Tab(1).ControlCount=   83
      TabCaption(2)   =   "待辦歷程"
      TabPicture(2)   =   "frm090909.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label17"
      Tab(2).Control(1)=   "Label1(48)"
      Tab(2).Control(2)=   "Label16"
      Tab(2).Control(3)=   "Label7"
      Tab(2).Control(4)=   "grd2"
      Tab(2).Control(5)=   "Combo5"
      Tab(2).Control(6)=   "cmdDetail"
      Tab(2).Control(7)=   "cmdQuery"
      Tab(2).ControlCount=   8
      Begin VB.CommandButton cmdLetter 
         Caption         =   "撰寫信函(&L)"
         Height          =   345
         Left            =   7464
         TabIndex        =   118
         Top             =   1608
         Width           =   1680
      End
      Begin VB.CommandButton cmdRead 
         Height          =   300
         Left            =   -71670
         Picture         =   "frm090909.frx":0054
         Style           =   1  '圖片外觀
         TabIndex        =   117
         Top             =   360
         Visible         =   0   'False
         Width           =   350
      End
      Begin VB.TextBox txt1 
         Height          =   300
         Index           =   1
         Left            =   8650
         MaxLength       =   6
         TabIndex        =   100
         Top             =   2940
         Width           =   530
      End
      Begin VB.TextBox txt1 
         Height          =   300
         Index           =   9
         Left            =   8040
         MaxLength       =   1
         TabIndex        =   88
         Top             =   4320
         Width           =   360
      End
      Begin VB.TextBox txt1 
         Enabled         =   0   'False
         Height          =   300
         Index           =   8
         Left            =   5730
         MaxLength       =   7
         TabIndex        =   87
         Top             =   4320
         Width           =   915
      End
      Begin VB.TextBox txt1 
         Height          =   300
         Index           =   0
         Left            =   8650
         MaxLength       =   6
         TabIndex        =   9
         Top             =   3240
         Width           =   530
      End
      Begin VB.TextBox txt1 
         Height          =   300
         Index           =   2
         Left            =   5010
         MaxLength       =   7
         TabIndex        =   1
         Top             =   720
         Width           =   915
      End
      Begin VB.TextBox txt1 
         Height          =   300
         Index           =   7
         Left            =   5010
         MaxLength       =   7
         TabIndex        =   3
         Top             =   1320
         Width           =   915
      End
      Begin VB.TextBox txt1 
         Height          =   300
         Index           =   6
         Left            =   5010
         MaxLength       =   6
         TabIndex        =   24
         Top             =   2250
         Visible         =   0   'False
         Width           =   900
      End
      Begin VB.TextBox txt1 
         Height          =   300
         Index           =   5
         Left            =   5010
         MaxLength       =   6
         TabIndex        =   4
         Top             =   1620
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.TextBox txt1 
         Height          =   300
         Index           =   3
         Left            =   5010
         MaxLength       =   7
         TabIndex        =   2
         Top             =   1020
         Width           =   915
      End
      Begin VB.CommandButton cmdok2 
         Caption         =   "當月資料"
         Default         =   -1  'True
         Height          =   400
         Index           =   0
         Left            =   -67656
         TabIndex        =   23
         Top             =   348
         Width           =   972
      End
      Begin VB.CommandButton cmdok2 
         Caption         =   "未發文"
         Height          =   400
         Index           =   1
         Left            =   -66648
         TabIndex        =   22
         Top             =   348
         Width           =   852
      End
      Begin VB.CommandButton cmd1 
         Caption         =   "專利相關案件"
         Height          =   345
         Left            =   7470
         TabIndex        =   21
         Top             =   2040
         Width           =   1680
      End
      Begin VB.TextBox txt1 
         Enabled         =   0   'False
         Height          =   300
         Index           =   12
         Left            =   5010
         MaxLength       =   7
         TabIndex        =   0
         Top             =   420
         Width           =   915
      End
      Begin VB.ComboBox Combo3 
         Height          =   260
         ItemData        =   "frm090909.frx":0156
         Left            =   -70260
         List            =   "frm090909.frx":0163
         TabIndex        =   19
         Top             =   390
         Width           =   2430
      End
      Begin VB.TextBox txt1 
         Height          =   300
         Index           =   13
         Left            =   7290
         MaxLength       =   7
         TabIndex        =   8
         Top             =   1020
         Width           =   915
      End
      Begin VB.TextBox txt1 
         Height          =   300
         Index           =   19
         Left            =   5010
         MaxLength       =   7
         TabIndex        =   7
         Top             =   2580
         Width           =   900
      End
      Begin VB.CommandButton cmd 
         Caption         =   "承辦歷程(&E)"
         Height          =   285
         Index           =   1
         Left            =   7470
         TabIndex        =   18
         Top             =   600
         Width           =   1320
      End
      Begin VB.CommandButton cmdQuery 
         Caption         =   "畫面更新(&Q)"
         Height          =   360
         Left            =   -66930
         TabIndex        =   17
         Top             =   300
         Width           =   1125
      End
      Begin VB.CommandButton cmdDetail 
         Caption         =   "明細資料(&D)"
         Height          =   360
         Left            =   -68100
         TabIndex        =   16
         Top             =   300
         Width           =   1125
      End
      Begin VB.ComboBox Combo5 
         Height          =   260
         ItemData        =   "frm090909.frx":01A3
         Left            =   -69120
         List            =   "frm090909.frx":01B3
         Style           =   2  '單純下拉式
         TabIndex        =   15
         Top             =   330
         Width           =   960
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grd1 
         Height          =   5100
         Left            =   -74940
         TabIndex        =   20
         Top             =   750
         Width           =   9825
         _ExtentX        =   17339
         _ExtentY        =   8996
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
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grd2 
         Height          =   4940
         Left            =   -74940
         TabIndex        =   25
         Top             =   870
         Width           =   9830
         _ExtentX        =   17339
         _ExtentY        =   8714
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
      Begin TabDlg.SSTab SSTab2 
         Height          =   2210
         Left            =   60
         TabIndex        =   99
         Top             =   3720
         Width           =   4520
         _ExtentX        =   7973
         _ExtentY        =   3898
         _Version        =   393216
         Style           =   1
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   176
         TabCaption(0)   =   "Tab 0"
         TabPicture(0)   =   "frm090909.frx":01D2
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "txtDST05"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Label1(24)"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "Label1(27)"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "txtPA162"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "FramePA162"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).ControlCount=   5
         TabCaption(1)   =   "Tab 1"
         TabPicture(1)   =   "frm090909.frx":01EE
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Label6"
         Tab(1).Control(1)=   "txtAMD05"
         Tab(1).ControlCount=   2
         Begin VB.Frame FramePA162 
            Height          =   970
            Left            =   0
            TabIndex        =   109
            Top             =   1170
            Width           =   4450
            Begin VB.CommandButton Command1 
               Caption         =   "複製前次內容"
               Height          =   285
               Left            =   2910
               TabIndex        =   110
               Top             =   0
               Width           =   1365
            End
            Begin MSForms.TextBox txtDST05Old 
               Height          =   710
               Left            =   0
               TabIndex        =   112
               Top             =   300
               Width           =   4490
               VariousPropertyBits=   -1466941409
               BackColor       =   -2147483638
               ScrollBars      =   2
               Size            =   "7920;1252"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin VB.Label Label1 
               Alignment       =   1  '靠右對齊
               AutoSize        =   -1  'True
               Caption         =   "前次內容:"
               Height          =   180
               Index           =   10
               Left            =   60
               TabIndex        =   111
               Top             =   90
               Width           =   770
            End
         End
         Begin VB.TextBox txtPA162 
            Height          =   270
            Left            =   1950
            MaxLength       =   1
            TabIndex        =   104
            TabStop         =   0   'False
            Top             =   90
            Width           =   255
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "是否提供核准分割建議:        ( Y: 是 N:否 )"
            Height          =   180
            Index           =   27
            Left            =   30
            TabIndex        =   107
            Top             =   130
            Width           =   3350
         End
         Begin VB.Label Label1 
            Alignment       =   1  '靠右對齊
            AutoSize        =   -1  'True
            Caption         =   "核准分割建議定稿文字:"
            Height          =   180
            Index           =   24
            Left            =   30
            TabIndex        =   106
            Top             =   370
            Width           =   1850
         End
         Begin MSForms.TextBox txtDST05 
            Height          =   590
            Left            =   0
            TabIndex        =   105
            Top             =   580
            Width           =   4490
            VariousPropertyBits=   -1466941413
            MaxLength       =   1000
            ScrollBars      =   2
            Size            =   "7920;1041"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin VB.Label Label6 
            Caption         =   "中說請款修正定稿文字"
            Height          =   200
            Left            =   -74940
            TabIndex        =   103
            Top             =   150
            Width           =   2030
         End
         Begin MSForms.TextBox txtAMD05 
            Height          =   1820
            Left            =   -75000
            TabIndex        =   102
            Top             =   360
            Width           =   4460
            VariousPropertyBits=   -1466941413
            MaxLength       =   2000
            ScrollBars      =   2
            Size            =   "7867;3210"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
      End
      Begin MSForms.ComboBox cboCP14 
         Height          =   290
         Left            =   2010
         TabIndex        =   116
         Top             =   330
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
      Begin VB.Label Label7 
         Caption         =   "註：在”不顯示”欄位上點一下(V)，可以取消顯示聯絡歷程。"
         ForeColor       =   &H000000C0&
         Height          =   200
         Left            =   -70110
         TabIndex        =   115
         Top             =   690
         Width           =   4970
      End
      Begin VB.Label Lbl926 
         Caption         =   "(一核 or 二核)"
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
         Left            =   5970
         TabIndex        =   114
         Top             =   480
         Visible         =   0   'False
         Width           =   1220
      End
      Begin VB.Label LblCP114 
         Alignment       =   1  '靠右對齊
         Caption         =   "核稿時數："
         ForeColor       =   &H00FF0000&
         Height          =   180
         Left            =   7670
         TabIndex        =   101
         Top             =   2970
         Width           =   960
      End
      Begin MSForms.TextBox txtEP12 
         Height          =   510
         Left            =   5730
         TabIndex        =   98
         Top             =   3810
         Width           =   4020
         VariousPropertyBits=   -1466941413
         MaxLength       =   2000
         ScrollBars      =   2
         Size            =   "7091;900"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCP64 
         Height          =   510
         Left            =   5730
         TabIndex        =   97
         Top             =   4950
         Width           =   4020
         VariousPropertyBits=   -1466941409
         MaxLength       =   2000
         ScrollBars      =   2
         Size            =   "7091;900"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         Caption         =   "進度備註："
         Height          =   180
         Index           =   46
         Left            =   4830
         TabIndex        =   96
         Top             =   4980
         Width           =   900
      End
      Begin VB.Label Label1 
         Caption         =   "(N:  不通知)"
         Height          =   210
         Index           =   22
         Left            =   8480
         TabIndex        =   95
         ToolTipText     =   "(N:  不通知, 自動內部收文)"
         Top             =   4350
         Width           =   1080
      End
      Begin VB.Label Label1 
         Alignment       =   1  '靠右對齊
         Caption         =   "是否通知客戶："
         Height          =   180
         Index           =   3
         Left            =   6680
         TabIndex        =   94
         Top             =   4380
         Width           =   1350
      End
      Begin VB.Label Label1 
         Caption         =   "點數："
         Height          =   260
         Index           =   11
         Left            =   2180
         TabIndex        =   93
         Top             =   3510
         Width           =   540
      End
      Begin MSForms.Label lbl1 
         Height          =   260
         Index           =   28
         Left            =   5730
         TabIndex        =   92
         Top             =   4680
         Width           =   1590
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "2805;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         Alignment       =   1  '靠右對齊
         Caption         =   "發文日："
         Height          =   180
         Index           =   6
         Left            =   4980
         TabIndex        =   91
         Top             =   4370
         Width           =   740
      End
      Begin VB.Label Label1 
         Alignment       =   1  '靠右對齊
         Caption         =   "取消收文日："
         Height          =   180
         Index           =   29
         Left            =   4340
         TabIndex        =   90
         Top             =   4680
         Width           =   1380
      End
      Begin VB.Label Label1 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         Caption         =   "作業備註："
         Height          =   180
         Index           =   31
         Left            =   4820
         TabIndex        =   89
         Top             =   3840
         Width           =   900
      End
      Begin MSForms.Label lbl1 
         Height          =   240
         Index           =   8
         Left            =   5010
         TabIndex        =   86
         Top             =   790
         Visible         =   0   'False
         Width           =   860
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "1508;423"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         Alignment       =   1  '靠右對齊
         Caption         =   "齊備日："
         Height          =   180
         Index           =   23
         Left            =   4260
         TabIndex        =   85
         Top             =   790
         Width           =   740
      End
      Begin MSForms.ComboBox Combo6 
         Height          =   320
         Left            =   5010
         TabIndex        =   5
         Top             =   2910
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
      Begin MSForms.ComboBox Combo2 
         Height          =   320
         Left            =   5010
         TabIndex        =   84
         Top             =   1620
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
      Begin MSForms.ComboBox Combo4 
         Height          =   320
         Left            =   5010
         TabIndex        =   6
         Top             =   2250
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
      Begin MSForms.Label lblEP42 
         Height          =   260
         Left            =   4980
         TabIndex        =   83
         Top             =   3300
         Width           =   890
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "1570;459"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lblEP39 
         Height          =   260
         Left            =   5010
         TabIndex        =   82
         Top             =   2010
         Width           =   890
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "1570;459"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.ComboBox Combo1 
         Height          =   315
         Left            =   -74100
         TabIndex        =   26
         Top             =   360
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
      Begin VB.Label Label1 
         Alignment       =   1  '靠右對齊
         Caption         =   "判發完成日："
         Height          =   180
         Index           =   9
         Left            =   3810
         TabIndex        =   79
         Top             =   3300
         Width           =   1190
      End
      Begin VB.Label Label1 
         Alignment       =   1  '靠右對齊
         Caption         =   "核稿完成日："
         Height          =   180
         Index           =   4
         Left            =   3810
         TabIndex        =   78
         Top             =   2010
         Width           =   1190
      End
      Begin VB.Label LblCP113 
         Alignment       =   1  '靠右對齊
         Caption         =   "工作時數："
         ForeColor       =   &H00FF0000&
         Height          =   180
         Left            =   7670
         TabIndex        =   77
         Top             =   3270
         Width           =   960
      End
      Begin MSForms.Label Label3 
         Height          =   290
         Left            =   -74040
         TabIndex        =   76
         Top             =   390
         Width           =   2450
         VariousPropertyBits=   27
         Size            =   "4313;494"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         Alignment       =   1  '靠右對齊
         Caption         =   "承辦期限："
         Height          =   180
         Index           =   7
         Left            =   4040
         TabIndex        =   75
         Top             =   500
         Width           =   960
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
         TabIndex        =   74
         Top             =   30
         Width           =   3225
      End
      Begin VB.Label Label1 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         Caption         =   "外文核稿人："
         Height          =   180
         Index           =   25
         Left            =   3920
         TabIndex        =   73
         Top             =   2310
         Width           =   1080
      End
      Begin VB.Label Label1 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         Caption         =   "核稿人："
         Height          =   180
         Index           =   5
         Left            =   3810
         TabIndex        =   72
         Top             =   1680
         Width           =   1190
      End
      Begin VB.Label Label1 
         Alignment       =   1  '靠右對齊
         Caption         =   "核稿期限："
         Height          =   180
         Index           =   2
         Left            =   3900
         TabIndex        =   71
         Top             =   1380
         Width           =   1110
      End
      Begin VB.Label Label1 
         Alignment       =   1  '靠右對齊
         Caption         =   "完稿日："
         Height          =   180
         Index           =   26
         Left            =   4260
         TabIndex        =   70
         Top             =   1080
         Width           =   740
      End
      Begin MSForms.Label lbl1 
         Height          =   260
         Index           =   23
         Left            =   2820
         TabIndex        =   69
         Top             =   3510
         Width           =   770
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "1358;459"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   260
         Index           =   19
         Left            =   990
         TabIndex        =   68
         Top             =   3240
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
         Height          =   260
         Index           =   17
         Left            =   990
         TabIndex        =   67
         Top             =   2900
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
         Height          =   260
         Index           =   15
         Left            =   990
         TabIndex        =   66
         Top             =   2580
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
         Height          =   260
         Index           =   13
         Left            =   1440
         TabIndex        =   65
         Top             =   2280
         Width           =   1410
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "2487;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   260
         Index           =   11
         Left            =   1470
         TabIndex        =   64
         Top             =   1950
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
         Height          =   260
         Index           =   9
         Left            =   1020
         TabIndex        =   63
         Top             =   1640
         Width           =   2780
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "4904;459"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   260
         Index           =   7
         Left            =   1020
         TabIndex        =   62
         Top             =   1330
         Width           =   1740
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "3069;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   260
         Index           =   5
         Left            =   1020
         TabIndex        =   61
         Top             =   1010
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
         Height          =   260
         Index           =   3
         Left            =   1020
         TabIndex        =   60
         Top             =   700
         Width           =   1170
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "2064;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         Caption         =   "(N：不算)"
         Height          =   180
         Index           =   32
         Left            =   2200
         TabIndex        =   59
         Top             =   1950
         Width           =   1070
      End
      Begin MSForms.Label lbl1 
         Height          =   495
         Index           =   30
         Left            =   6270
         TabIndex        =   58
         Top             =   7437
         Visible         =   0   'False
         Width           =   1600
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "2822;873"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   225
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   240
         Index           =   16
         Left            =   5940
         TabIndex        =   57
         Top             =   2370
         Width           =   1050
         VariousPropertyBits=   27
         Caption         =   "EP03"
         Size            =   "1852;423"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   180
         Index           =   10
         Left            =   5010
         TabIndex        =   56
         Top             =   1080
         Visible         =   0   'False
         Width           =   1010
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "1773;317"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   495
         Index           =   29
         Left            =   6615
         TabIndex        =   55
         Top             =   7587
         Visible         =   0   'False
         Width           =   1590
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "2408;556"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   225
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "法定期限："
         Height          =   260
         Index           =   13
         Left            =   80
         TabIndex        =   54
         Top             =   3210
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "本所期限："
         Height          =   260
         Index           =   14
         Left            =   80
         TabIndex        =   53
         Top             =   2890
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "案件性質："
         Height          =   260
         Index           =   15
         Left            =   80
         TabIndex        =   52
         Top             =   2580
         Width           =   900
      End
      Begin VB.Label Label1 
         Caption         =   "專利/商標種類："
         Height          =   260
         Index           =   16
         Left            =   80
         TabIndex        =   51
         Top             =   2270
         Width           =   1370
      End
      Begin VB.Label Label1 
         Caption         =   "是否算案件數："
         Height          =   260
         Index           =   17
         Left            =   80
         TabIndex        =   50
         Top             =   1950
         Width           =   1320
      End
      Begin VB.Label Label1 
         Caption         =   "案件名稱："
         Height          =   260
         Index           =   18
         Left            =   80
         TabIndex        =   49
         Top             =   1640
         Width           =   960
      End
      Begin VB.Label Label1 
         Caption         =   "本所案號："
         Height          =   260
         Index           =   19
         Left            =   80
         TabIndex        =   48
         Top             =   1330
         Width           =   900
      End
      Begin VB.Label Label1 
         Caption         =   "收文日："
         Height          =   260
         Index           =   20
         Left            =   80
         TabIndex        =   47
         Top             =   1010
         Width           =   740
      End
      Begin VB.Label Label1 
         Caption         =   "總收文號："
         Height          =   260
         Index           =   21
         Left            =   80
         TabIndex        =   46
         Top             =   700
         Width           =   930
      End
      Begin VB.Label Label1 
         Alignment       =   1  '靠右對齊
         Caption         =   "承辦人："
         Height          =   260
         Index           =   8
         Left            =   1280
         TabIndex        =   45
         Top             =   390
         Width           =   740
      End
      Begin MSForms.Label lbl1 
         Height          =   260
         Index           =   0
         Left            =   650
         TabIndex        =   44
         Top             =   390
         Width           =   630
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "1111;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   260
         Index           =   1
         Left            =   2010
         TabIndex        =   43
         Top             =   390
         Width           =   810
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "1429;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         Caption         =   "目次："
         Height          =   260
         Index           =   1
         Left            =   80
         TabIndex        =   42
         Top             =   390
         Width           =   540
      End
      Begin MSForms.Label lbl1 
         Height          =   260
         Index           =   21
         Left            =   990
         TabIndex        =   41
         Top             =   3510
         Width           =   800
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "1402;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "智權人員："
         Height          =   260
         Index           =   12
         Left            =   80
         TabIndex        =   40
         Top             =   3520
         Width           =   900
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
         Left            =   2970
         TabIndex        =   39
         Top             =   1350
         Width           =   950
      End
      Begin VB.Label Label1 
         Caption         =   "承辦人： "
         Height          =   180
         Index           =   0
         Left            =   -74904
         TabIndex        =   38
         Top             =   432
         Width           =   792
      End
      Begin VB.Label Label5 
         Caption         =   "顏色說明："
         Height          =   225
         Left            =   -71160
         TabIndex        =   37
         Top             =   432
         Width           =   915
      End
      Begin VB.Label Label1 
         Alignment       =   1  '靠右對齊
         Caption         =   "Claims完稿日："
         Height          =   180
         Index           =   35
         Left            =   5970
         TabIndex        =   36
         Top             =   1050
         Width           =   1310
      End
      Begin VB.Label Label1 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         Caption         =   "外文核完日："
         Height          =   180
         Index           =   41
         Left            =   3920
         TabIndex        =   35
         Top             =   2640
         Width           =   1080
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
         Left            =   2970
         TabIndex        =   34
         Top             =   860
         Visible         =   0   'False
         Width           =   890
      End
      Begin VB.Label Label16 
         Caption         =   "註：雙擊選取時，開啟承辦歷程"
         ForeColor       =   &H000000C0&
         Height          =   230
         Left            =   -74880
         TabIndex        =   33
         Top             =   420
         Width           =   2900
      End
      Begin VB.Label Label1 
         Alignment       =   1  '靠右對齊
         Caption         =   "判發人："
         Height          =   180
         Index           =   47
         Left            =   4200
         TabIndex        =   32
         Top             =   2990
         Width           =   740
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "最近聯絡："
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   48
         Left            =   -70020
         TabIndex        =   31
         Top             =   390
         Width           =   900
      End
      Begin VB.Label Label17 
         Caption         =   "已確認過會完日，在會完流程狀態前加註Y／N。"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   8.5
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   225
         Left            =   -74910
         TabIndex        =   30
         Top             =   665
         Width           =   3915
      End
      Begin VB.Label Label18 
         Caption         =   "可不跑承辦歷程"
         ForeColor       =   &H00C00000&
         Height          =   225
         Left            =   7860
         TabIndex        =   29
         Top             =   360
         Width           =   1335
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
         Left            =   2970
         TabIndex        =   28
         Top             =   1110
         Visible         =   0   'False
         Width           =   830
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
         Left            =   2970
         TabIndex        =   27
         Top             =   630
         Width           =   950
      End
   End
   Begin VB.Label LblCnt 
      AutoSize        =   -1  'True
      Caption         =   "LblCnt"
      Height          =   180
      Left            =   60
      TabIndex        =   113
      Top             =   60
      Width           =   470
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "核稿語文：   (1.英2.日)"
      Height          =   180
      Index           =   50
      Left            =   3990
      TabIndex        =   81
      Top             =   390
      Visible         =   0   'False
      Width           =   1790
   End
   Begin VB.Label Label4 
      Caption         =   "申請國家："
      Enabled         =   0   'False
      Height          =   180
      Left            =   2715
      TabIndex        =   12
      Top             =   495
      Width           =   900
   End
End
Attribute VB_Name = "frm090909"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/12/28 改成Form2.0 ; grd1改字型=新細明體-ExtB、grd2改字型=新細明體-ExtB、Combo1、Combo4、Combo6、lbl1(index)、txt1(10)改為txtEP12、txtCP64
'Memo By Morgan 2012/12/10 智權人員欄已修改
'Memo by Morgan2010/12/28 申請案號欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan 2010/8/16 日期欄已修改
Option Explicit

Public ManaGrp As String
Public TextOk As Boolean
Public Combo1_String As String
Public Combo1_Name As String 'Add By Sindy 2024/3/5
Dim strSql As String, i As Integer, s As Integer, k As Integer
Dim SWPRow As String, strTemp(0 To 27) As String
Dim Tmp001 As String, Tmp002 As String, Tmp003 As String, Tmp004 As String
Dim SeekTmpBk As String
Dim ChkNoData As Boolean
Dim StrGrp090201 As String, ChkData As Boolean
Dim m_SqlGrpStr1 As String, m_SqlGrpStr5 As String
Dim m_strCP09 As String '總收文號
Dim m_blnColOrderAsc As Boolean '欄位資料由小到大排序
Dim m_CP01 As String, m_CP02 As String, m_CP03 As String, m_CP04 As String '本所案號
Dim m_CP14 As String
Dim m_CP10 As String
Dim StrSQL6 As String
Dim StrSQL61 As String
Dim StrSQL62 As String
Dim StrSQL63 As String
Dim StrSQL64 As String
Dim StrSPa As String
Dim StrSSP As String
'add by nick 2005/01/27
Dim m_Country As String
' 宣告欄位內容結構
Private Type FIELDITEM
   fiName As String
   fiOldData As String
   fiNewData As String
   fiType As Integer
End Type
'edit by nick 2006/02/27 已不用常數
Dim m_FieldList() As FIELDITEM
Dim skMail() As SeekMails
Dim bolInsert As Boolean, bolUpdate As Boolean, bolDelete As Boolean, bolSelect As Boolean, bolPrint As Boolean
Dim m_CP43 As String 'Add by Morgan 2009/12/3
'Add By Sindy 2013/6/7
Public m_chkcmdok1 As Boolean '記錄確定鍵是否存檔成功
Dim dblPrevRow As Double
Public intBackTab As Integer
'2013/6/7 End
Dim m_CPM28 As String 'Add By Sindy 2013/9/18
Dim m_CPM29 As String 'Add By Sindy 2013/9/30
Dim m_CP27 As String '發文日
Dim m_PP04 As String 'Add By Sindy 2013/10/14 預設核稿人
Dim m_PP05 As String 'Add By Sindy 2013/10/14 預設判發人
Public m_Flow As String 'Add By Sindy 2013/10/14 欲新增的下一流程
Dim lngFormWidth As Long, lngFormHeight As Long 'Added by Morgan 2016/2/18
Dim m_intRow As Integer, m_intCol As Integer 'Add By Sindy 2016/3/7
Dim m_ProState As String 'Add By Sindy 2017/8/10 記錄目前權限
Dim m_EP41 As String 'Add By Sindy 2015/3/13 核稿語文
Dim m_EMPST16 As String '承辦人是屬那一組人員
Dim strR110035 As String
Dim m_strRefVal As String 'Add By Sindy 2024/5/2


'Add By Sindy 2024/3/6
Private Sub CboCP14_GotFocus()
   cboCP14.SelStart = 0
   cboCP14.SelLength = Len(cboCP14.Text)
End Sub
Private Sub CboCP14_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub CboCP14_LostFocus()
   If Trim(cboCP14.Text) <> "" Then
      cmd(1).Enabled = True '顯示承辦歷程按鍵
      '有變更承辦人時,才檢查
      If m_CP14 <> Trim(Left(cboCP14.Text, 5)) Then
         cmd(1).Enabled = False '不可操作承辦歷程
      End If
   End If
End Sub
Private Sub CboCP14_Validate(Cancel As Boolean)
Dim m_Team As String
Dim strText As String
Dim bolRunOK As Boolean
   
   If Trim(cboCP14.Text) <> "" Then
      '有變更承辦人時,才檢查
      If m_CP14 <> Trim(Left(cboCP14.Text, 5)) Then
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
      If cboCP14.Visible = True Then cboCP14.SetFocus
      Call CboCP14_GotFocus
   End If
   
   Cancel = False
End Sub
'2024/3/6 END

Private Sub cmd_Click(Index As Integer)
On Error GoTo ErrHnd
   Select Case Index
   'Modify By Sindy 2013/4/16
   Case 1 '承辦歷程
      'Add By Sindy 2024/3/6
      If cboCP14.Visible = True And m_CP14 <> Trim(Left(cboCP14.Text, 5)) Then
         MsgBox "承辦人有異動，不可在此時同時操做歷程！" & vbCrLf & "請更新後再操作", vbExclamation
         Exit Sub
      End If
      'Modify By Sindy 2025/3/18 +外專繪圖同新案翻譯歷程操作方式
      If Combo2.Visible = True And Combo2.Enabled = True And (m_CP10 = "201" Or m_CP10 = "931") And _
         Trim(Left(Combo2.Tag, 5)) <> Trim(Left(Combo2.Text, 5)) Then
         MsgBox "核稿工程師有異動，不可在此時同時操做歷程！" & vbCrLf & "請更新後再操作", vbExclamation
         Exit Sub
      End If
      '2024/3/6 END
      
      'Add By Sindy 2013/9/16
      If ProState = "2" Then
         If frm090614.txt1(8) = "N" Then MsgBox "不可從（不區分個人）的資料查詢中來執行承辦歷程作業！": Exit Sub
      End If
      '2013/9/16 END
      
      'Add By Sindy 2017/8/3 個人案件不可用主管權限操作
      If ProState = "2" And m_CP14 = strUserNum Then '2.主管
         MsgBox "個人案件不可用主管權限操作！", vbExclamation
         Exit Sub
      End If
      '2017/8/3 END
      
      'Add By Sindy 2015/12/3
      '重新檢查欄位有效性
      If TxtValidate = True Then
      '2015/12/3 END
         'Add By Sindy 2013/6/10
         If SetColTag(False) = False Then
            cmdok(1).Enabled = False 'Add By Sindy 2017/9/21
            cmd(1).Tag = "Y" 'Add By Sindy 2024/12/11 代表有按 承辦歷程(&E)
            Call cmdok_Click(1)
            cmd(1).Tag = "" 'Add By Sindy 2024/12/11 取消Tag註記
            cmdok(1).Enabled = True 'Add By Sindy 2017/9/21
            If m_chkcmdok1 = False Then Exit Sub
         Else
            Call Process(lbl1(3)) '要重新查詢資料 Add By Sindy 2018/10/4
         End If
         
         If PUB_ChkFormIsClose("frm090202_2") = False Then Exit Sub 'Add By Sindy 2020/1/17
         intBackTab = 1
         '2013/6/10 End
         frm090202_2.Hide
         frm090202_2.m_EEP01 = lbl1(3) '總收文號
         frm090202_2.m_FlowUserNum = Trim(Left("" & Combo1.Text, 6)) 'Add By Sindy 2013/9/12 案件流程所屬人員
         frm090202_2.intReceiveKind = 0 '0.承辦人工作進度
         frm090202_2.SetParent Me
         If frm090202_2.QueryData = True Then
            frm090202_2.Show
            Me.Hide
         End If
      End If
   Case Else
   End Select

   Exit Sub

ErrHnd:
   If Err.Number <> 0 Then
      MsgBox Err.Description, vbCritical
   End If
End Sub

'strType=1.核稿人
'        2.判發人
Private Sub SetComboData(mCombo As Object, strType As String)
Dim blnMatch As Boolean
Dim ii As Integer
   
   If strType <> "" Then
      If strType = "1" Then
         strExc(0) = "SELECT st01||' ==> '||st02 FROM STAFF" & _
                     " WHERE ST93='" & PUB_GetST93(Trim(Left(Combo1.Text, 6))) & "'" & _
                     " AND ST04='1' AND substr(ST01,1,1)<>'F' AND substr(ST01,4,1)<>'9'" & _
                     " ORDER BY ST01"
      Else
         'ST52~ST55
         strExc(0) = "SELECT s2.st01||' ==> '||s2.st02, 1 as sort FROM STAFF s1,STAFF s2" & _
                     " WHERE s1.ST01='" & Trim(Left(Combo1.Text, 6)) & "'" & _
                     " AND s1.ST52=s2.st01(+) AND s1.ST52 is not null"
         strExc(0) = strExc(0) & " union" & _
                     " SELECT s2.st01||' ==> '||s2.st02, 2 as sort FROM STAFF s1,STAFF s2" & _
                     " WHERE s1.ST01='" & Trim(Left(Combo1.Text, 6)) & "'" & _
                     " AND s1.ST53=s2.st01(+) AND s1.ST53 is not null"
         strExc(0) = strExc(0) & " union" & _
                     " SELECT s2.st01||' ==> '||s2.st02, 3 as sort FROM STAFF s1,STAFF s2" & _
                     " WHERE s1.ST01='" & Trim(Left(Combo1.Text, 6)) & "'" & _
                     " AND s1.ST54=s2.st01(+) AND s1.ST54 is not null"
         If m_EMPST16 = "3" Then '日文組
            strExc(0) = strExc(0) & " union" & _
                        " SELECT s2.st01||' ==> '||s2.st02, 4 as sort FROM STAFF s1,STAFF s2" & _
                        " WHERE s1.ST01='" & Trim(Left(Combo1.Text, 6)) & "'" & _
                        " AND s1.ST55=s2.st01(+) AND s1.ST55 is not null"
         End If
         strExc(0) = strExc(0) & " order by sort asc"
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

Private Sub GetPP04PP05(strEP04 As String)
Dim strEmpUser As String
   
   m_PP04 = "" '核判表設定的核稿人
   m_PP05 = "" '核判表設定的判發人
   'Modify By Sindy 2025/3/18 +外專繪圖同新案翻譯歷程操作方式
   If (m_CP10 = "201" Or m_CP10 = "931") Then
      Label1(5).Caption = "核稿工程師："
      strEmpUser = strEP04
   Else
      Label1(5).Caption = "核稿人："
      strEmpUser = m_CP14
   End If
   'Modify By Sindy 2024/6/26 +m_Country
   Call PUB_ChkIsSetPromoterReader(strEmpUser, m_CP01, m_CP10, m_PP04, m_PP05, m_strCP09, m_Country)
   If m_PP04 = strEmpUser Then m_PP04 = "" '為自行核稿,不需再將自己ID放入核稿人欄位
   If m_PP05 = strEmpUser Then m_PP05 = "" '為自行判發,不需再將自己ID放入判發人欄位
   'Add By Sindy 2024/1/10 926=核對已准專利 一核 時,預帶核稿人
   If m_PP04 = "" And (m_CP10 = "926" And txt1(12) <> "" And lbl1(17) = "") Then
      m_PP04 = m_PP05 '核稿人=判發人
   End If
   '2024/1/10 END
End Sub

Public Sub Process(strText As String)
Dim arrCaseNo '本所案號
Dim stVTB As String
Dim oLbl As Object
Dim oTxt1 As Object
Dim strPA158 As String
Dim strEP40 As String
Dim ii As Integer
Dim blnMatch As Boolean
Dim rsTmp As New ADODB.Recordset
   
   stVTB = " SELECT CP01,CP02,CP03,CP04,CP05,CP06,CP07,CP09,CP10,CP13,CP14,CP15,CP18,CP26,CP27,CP31,cp44,CP48,CP57,cp64,CP97,CP98,CP99,CP106,cp107,cp111,cp112,cp113" & _
            ",NVL(PA05,NVL(PA06,PA07)) C10,DECODE(PA09,'020',PTM04,PTM03) C14,'' C26,'' C28,PA57,'*' C33,pa09 as m_country,pa26 as cuno,CP43,CP147,PA158,CP118,PA08,cp144,pa27 as cuno2,pa28 as cuno3,pa29 as cuno4,pa30 as cuno5,cp114,pa150" & _
            " FROM CASEPROGRESS,PATENT,PATENTTRADEMARKMAP WHERE CP09='" & strText & "' and cp01 in (" & m_SqlGrpStr1 & ") AND PA01(+)=CP01 AND PA02(+)=CP02 AND PA03(+)=CP03 AND PA04(+)=CP04 AND PTM01(+)='1' AND PTM02(+)=PA08"
   stVTB = stVTB + " UNION all SELECT CP01,CP02,CP03,CP04,CP05,CP06,CP07,CP09,CP10,CP13,CP14,CP15,CP18,CP26,CP27,CP31,cp44,CP48,CP57,cp64,CP97,CP98,CP99,CP106,cp107,cp111,cp112,cp113" & _
            ",NVL(SP05,NVL(SP06,SP07)) C10,'' C14,'' C26,'' C28,SP15,'*' C33,sp09 as m_country,sp08 as cuno,CP43,CP147,'' pa158,CP118,'' PA08,cp144,sp58 as cuno2,sp59 as cuno3,sp65 as cuno4,sp66 as cuno5,cp114,sp79 pa150" & _
            " FROM CASEPROGRESS,SERVICEPRACTICE WHERE CP09='" & strText & "' and cp01 in (" & m_SqlGrpStr5 & ") AND SP01(+)=CP01 AND SP02(+)=CP02 AND SP03(+)=CP03 AND SP04(+)=CP04"
      
   strSql = "SELECT EP01,S1.ST02 C2,sqldateT(CP48) C3,CP09,EP13,sqldateT(cp05) C6,EP34,CP01||'-'||CP02||'-'||CP03||'-'||CP04 C8" & _
      ",EP06,C10,EP09,CP26,EP07,C14,EP04,decode(na01,'000',cpm03,cpm04) C16,EP03,sqldateT(CP06) C18,EP08,sqldateT(CP07) C20,CP27" & _
      ",S5.ST02 C22,EP11,CP18,EP12,C26,Nvl(EP35,0) C27,C28,sqldateT(CP57) C29,CP10,CP15,PA57,C33,EP27,EP31,cp13,ep05,m_country,cp31,S5.st06 as Area,cp107,cp97,nvl(cp98,0) as cp98,cp99" & _
      ",cp106,cuno,cp111,cp112,ep28,ep32,ep33,na03,cp64,cpm05,cp44,ibf01,S3.ST02 EP04N,pp04,s6.st02 pp04N,s2.st02 EP13N,s4.st02 EP03N" & _
      ",NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)) CuName,CP43,CP147,pa158,CP118,EP38,EP39,pp05,EP40,cpm28,cpm29,PA08,cpm23,EP41,cp144,cp113,EP42,cp114,pa150" & _
      " from (" & stVTB & ") X,ENGINEERPROGRESS,CASEPROPERTYMAP,nation" & _
      ",STAFF S1,STAFF S2,STAFF S3,STAFF S4,STAFF S5,customer,imgbytefile,promoterproofreader,staff S6" & _
      " where EP02(+)=CP09 and cpm01(+)=CP01 and cpm02(+)=CP10 AND na01(+)=m_country" & _
      " AND S1.ST01(+)=EP05 AND S2.ST01(+)=EP13 AND S3.ST01(+)=EP04 AND S4.ST01(+)=EP03 AND S5.ST01(+)=CP13" & _
      " and cu01(+)=substr(cuno,1,8) and cu02(+)=substr(cuno,9) and pp01(+)=cp01 and pp02(+)=cp14 and pp03(+)=cp10 and s6.st01(+)=pp04" & _
      " and ibf01(+)=cp01 and ibf02(+)=cp02 and ibf03(+)=cp03 and ibf04(+)=cp04 and ibf05(+)='1'"
      
   CheckOC
   With adoRecordset
   .CursorLocation = adUseClient
   .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   
'   intI = 1
'   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
'   If intI = 1 Then
'   End If

   '***** 清除欄位值 *****
   For Each oLbl In lbl1
      oLbl.Caption = ""
   Next
   Me.lblClose.Caption = ""
   For Each oTxt1 In txt1
      oTxt1.Text = ""
      oTxt1.Enabled = True
   Next
   txtEP12.Enabled = True
   Me.Combo2.Enabled = True '核稿人
   Me.Combo4.Enabled = True '外文核稿人
   Me.Combo6.Enabled = True '判發人
   txtCP64.Text = ""
   Combo2.Clear: Combo2.Tag = ""
   Combo6.Clear: Combo6.Tag = ""
   m_EMPST16 = ""
   m_CPM28 = ""
   m_CPM29 = ""
   lblEP39 = ""
   lblEP42 = ""
   cboCP14.Text = "": cboCP14.Visible = False 'Add By Sindy 2024/3/6 預設值
   cmdLetter.Visible = False 'Added by Morgan 2024/3/29
   
   '***** END *****
   If .RecordCount <> 0 And .RecordCount > 0 Then
      .MoveFirst
      'Modify By Sindy 2024/3/27
      '承辦人是屬那一組人員
      'Modify By Sindy 2025/3/18 +外專繪圖同新案翻譯歷程操作方式
      If ("" & .Fields("cp10").Value = 翻譯 Or "" & .Fields("cp10").Value = "931") Then
         Call GetPrjSalesNM("" & .Fields("EP04"), , , "st16", m_EMPST16) '核稿工程師
      Else
         Call GetPrjSalesNM(Trim(Left(Combo1.Text, 6)), , , "st16", m_EMPST16)
      End If
      '2024/3/27 END
      m_CP43 = "" & .Fields("CP43")
      m_CP01 = SystemNumber(Trim(.Fields("C8")), 1)
      m_CP02 = SystemNumber(Trim(.Fields("C8")), 2)
      m_CP03 = SystemNumber(Trim(.Fields("C8")), 3)
      m_CP04 = SystemNumber(Trim(.Fields("C8")), 4)
      strPA158 = "" & .Fields("PA158")
      lblEP39 = ChangeWStringToTDateString("" & .Fields("EP39")) '核稿完成日
      lblEP42 = ChangeWStringToTDateString("" & .Fields("EP42")) '判發完成日
      m_CPM28 = "" & .Fields("CPM28")
      
      '電子送件
      If Not IsNull(.Fields("CP118")) Then
         lblEApp.Visible = True
      Else
         lblEApp.Visible = False
      End If
      
      For i = 0 To 29
         '外文核稿人(EP03)
         If i = 16 Then
            txt1(6).Text = CheckStr(.Fields(i))
         '核稿期限(EP08)
         ElseIf i = 18 Then
            txt1(7).Text = ChangeWStringToTString(CheckStr(.Fields(i)))
         '發文日(CP27)
         ElseIf i = 20 Then
            txt1(8).Text = ChangeWStringToTString(CheckStr(.Fields(i)))
         '是否通知客戶(EP11)
         ElseIf i = 22 Then
            txt1(9).Text = CheckStr(.Fields(i))
         '承辦備註(EP12)
         ElseIf i = 24 Then
            txtEP12.Text = CheckStr(.Fields(i))
         '承辦期限(CP48)
         ElseIf i = 2 Then
            txt1(12).Text = ChangeTDateStringToTString(CheckStr(.Fields(i)))
         ElseIf i = 15 Then 'decode(na01,'000',cpm03,cpm04) C16
            If Not IsNull(.Fields("CP43")) Then
               lbl1(i) = CheckStr(.Fields(i)) & PUB_GetRelateCasePropertyName(strText, "1")
            Else
               lbl1(i) = CheckStr(.Fields(i))
            End If
         '核稿人(EP04)
         ElseIf i = 14 Then
            txt1(5).Text = CheckStr(.Fields(i))
         Else
            'i=12=EP07
            'i=26=C26
            If i <> 25 And i <> 26 And i <> 27 And i <> 12 And i <> 4 And i <> 6 Then
               lbl1(i) = CheckStr(.Fields(i))
            End If
         End If
      Next i
      '齊備日
      txt1(2).Text = ChangeWStringToTString(lbl1(8).Caption)
      '完稿日
      txt1(3).Text = ChangeWStringToTString(lbl1(10).Caption)
      'Claims完稿日
      txt1(13) = ChangeWStringToTString(CheckStr(.Fields("EP31")))
      '外文核完日
      txt1(19) = ChangeWStringToTString(CheckStr(.Fields("EP33")))
      m_CP14 = "" & .Fields("ep05").Value
      m_CP10 = "" & .Fields("cp10").Value
      txtCP64 = CheckStr(.Fields("cp64"))
      txt1(0) = CheckStr(.Fields("cp113")) '工作時數
      txt1(1) = CheckStr(.Fields("cp114")) '核稿時數
      m_Country = "" & .Fields("m_country").Value
      m_CPM29 = "" & .Fields("CPM29") '是否不電子簽核
      'Add By Sindy 2024/3/6 承辦人下拉選單
      '承辦人工作進度管理，開放改承辦人功能，僅限ST03='F21'的主管及99097=李柏翰才能改
      'Modify By Sindy 2025/3/18 +外專繪圖同新案翻譯歷程操作方式
      If m_CP10 <> "201" And m_CP10 <> "931" And ProState = "2" And Val(txt1(8).Text) = 0 And _
         (Pub_StrUserSt03 = "F21" Or strUserNum = "99097") Then '99097=李柏翰
         cboCP14.Visible = True
         Call Frm060101_1_SetCboCP14(m_CP10, "" & .Fields("pa150"), cboCP14)
         If "" & .Fields("ep05") <> "" Then
            cboCP14.Text = "" & .Fields("ep05")
            Call CboCP14_Validate(False)
         Else
            cboCP14.Text = ""
         End If
      End If
      '2024/3/6 END
      
      'Added by Morgan 2024/3/29
      '二核定搞
      If .Fields("pa150") <> "3" And .Fields("cp10") = "926" Then
         cmdLetter.Visible = True
      End If
      'end 2024/3/29
      
      '一案兩請
      strSql = "select * from casemap where cm01='" & m_CP01 & "' and cm02='" & m_CP02 & "' and cm03='" & m_CP03 & "' and cm04='" & m_CP04 & "' and cm10='3'" & _
               " Union select * from casemap where cm05='" & m_CP01 & "' and cm06='" & m_CP02 & "' and cm07='" & m_CP03 & "' and cm08='" & m_CP04 & "' and cm10='3'"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strSql)
      If intI = 1 Then
         lblCM10.Visible = True
      Else
         lblCM10.Visible = False
      End If
      
      '台灣大陸案件提示
      lblCMboth.Caption = ""
      If (strExc(1) = "P" Or strExc(1) = "FCP") And "" & .Fields("m_country") = 台灣國家代號 Then
         If PUB_GetRefCaseChk(strExc(1), strExc(2), strExc(3), strExc(4), "CASEMAP", "0", "A", 大陸國家代號) Then
            lblCMboth.Caption = "有大陸案"
         End If
      ElseIf strExc(1) = "P" And "" & .Fields("m_country") = 大陸國家代號 Then
         If PUB_GetRefCaseChk(strExc(1), strExc(2), strExc(3), strExc(4), "CASEMAP", "0", "A", 台灣國家代號) Then
            lblCMboth.Caption = "有台灣案"
         End If
      End If
      
      'FMP的CPM29.是否不電子簽核
      If m_CP01 = "P" Or m_CP01 = "PS" Or m_CP01 = "CFP" Or m_CP01 = "CPS" Then
         strExc(0) = "select * from CasePropertyMap where CPM01='" & m_CP01 & "' and CPM02='" & m_CP10 & "'"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            m_CPM29 = "" & RsTemp.Fields("CPM36")
         End If
      End If
      
      '記錄收文號
      m_strCP09 = Me.lbl1(3).Caption
      If Len(Trim(CheckStr(.Fields(20)))) <> 0 Then
         m_CP27 = .Fields(20) '發文日
      Else
         m_CP27 = "" '發文日
      End If
      
      If IsNull(.Fields(31).Value) <> 0 Then
          Me.lblClose.Caption = ""
      Else
          Me.lblClose.Caption = "已閉卷"
      End If
      
      'Add By Sindy 2025/9/18 預設核稿語文
'      'Add By Sindy 2024/1/8 顯示核稿語文
'      m_EP41 = "" & .Fields("EP41")
'      If m_EP41 = "" Or PUB_ChkEmpFlowExists(lbl1(3), EMP_送英核) = False Then '尚未送英核時才預設核稿語文
'         txt1(23) = "1" '預設英文
'         If m_EMPST16 = "3" Then
'            txt1(23) = "2" '日文
'         End If
'      Else
'         txt1(23) = m_EP41
'      End If
      Call SetEP41("" & .Fields("EP41"), "" & .Fields("EP03"))
      '2025/9/18 END
      
      SetEngChecker '設定外文核稿人選單
      
      Call GetPP04PP05("" & .Fields("EP04")) 'Modify By Sindy 2024/3/27
      '核稿人:
      Combo2.AddItem "", 0
      'Modify By Sindy 2025/3/18 +外專繪圖同新案翻譯歷程操作方式
      If m_CP10 = "201" Or m_CP10 = "931" Then
         Combo2.Tag = "" & .Fields("EP04")
      Else
         '不用完稿日為預設核稿人的基準點,改用檢查有無送核歷程
         '已有核稿人不會再重新預設(同內商內專)
         If PUB_ChkEmpFlowExists(lbl1(3), EMP_送核) = False And _
            Len("" & .Fields("EP04")) = 0 Then
            Combo2.Tag = m_PP04
         Else
            If m_CP27 = "" And "" & .Fields("EP04") = "" And m_PP04 <> "" Then
               Combo2.Tag = m_PP04
            Else
               Combo2.Tag = "" & .Fields("EP04")
            End If
         End If
      End If
      If Combo2.Tag <> "" Then
         Combo2.AddItem Combo2.Tag & " ==> " & GetPrjSalesNM(Combo2.Tag), 1
      End If
      'Modify By Sindy 2025/3/18 +外專繪圖同新案翻譯歷程操作方式
      If (m_CP10 = "201" Or m_CP10 = "931") Then
         'Modify By Sindy 2024/3/11
         'Call SetComboData(Combo2, "")
         Call Frm060101_1_SetCboCP14("", "" & .Fields("pa150"), Combo2, "==>")
         If Combo2.Tag <> "" Then
            For ii = 0 To Me.Combo2.ListCount - 1
                blnMatch = False
                If Trim(Left(Me.Combo2.List(ii), 6)) = Trim(Left(Combo2.Tag, 6)) Then
                    Me.Combo2.ListIndex = ii
                    blnMatch = True
                    Exit For
                End If
            Next ii
            If blnMatch = False Then
               strExc(10) = Combo2.Tag & " ==> " & GetPrjSalesNM(Combo2.Tag)
               Combo2.AddItem strExc(10), 1
               Me.Combo2.Text = strExc(10)
            End If
         End If
         '2024/3/11 END
      Else
         Call SetComboData(Combo2, "1")
      End If
      '判發人:
      Combo6.AddItem "", 0
      '不用完稿日為預設判發人的基準點,改用檢查有無送判或判發歷程
      '已有判發人不會再重新預設(同內商內專)
      If PUB_ChkEmpFlowExists(lbl1(3), EMP_送判) = False And _
         PUB_ChkEmpFlowExists(lbl1(3), EMP_判發) = False And _
         Len("" & .Fields("EP40")) = 0 Then
         Combo6.Tag = m_PP05
      Else
         Combo6.Tag = "" & .Fields("EP40")
      End If
      If Combo6.Tag <> "" Then
         Combo6.AddItem Combo6.Tag & " ==> " & GetPrjSalesNM(Combo6.Tag), 1
      End If
      Call SetComboData(Combo6, "2")
   End If
   End With
   
'*************************************
'從 frm090901_1 Move過來
'*************************************
   'Added by Morgan 2012/11/30
   txtPA162.Enabled = False
   txtPA162 = "": txtPA162.Tag = "" 'Added by Morgan 2022/8/1
   txtDST05.Locked = True
   Command1.Enabled = False
   'Added by Lydia 2015/04/24 +中說請款修正定稿文字
   'Modified by Lydia 2015/08/27 為了能拉動卷軸,改成locked
   'txtAMD05.Enabled = False
   txtAMD05.Locked = True
   SSTab2.Visible = False 'lydia 無初審的高度
   txtDST05 = "": txtDST05.Tag = "": Check1.Tag = "": Check1.Visible = False
   If (m_CP10 = "204" Or m_CP10 = "205" Or m_CP10 = "203" Or m_CP10 = "107") _
      And (m_CP01 = "FCP" Or m_CP01 = "FG") _
      And Val(txt1(8).Text) = 0 And Val(lbl1(28).Caption) = 0 Then
'      '最後一道107,203,204,205
'      strExc(0) = "select cp09 from caseprogress" & _
'         " where cp01='" & m_CP01 & "' and cp02='" & m_CP02 & "' and cp03='" & m_CP03 & "' and cp04='" & m_CP04 & "'" & _
'         " and cp10 in ('107','203','204','205') and cp27>" & DBDATE(txt1(8).Text)
'      intI = 1
'      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'      If intI = 0 Then
         txtDST05.Tag = "Y"
'      End If
   End If
   'Modified by Morgan 2019/12/30
   If txtDST05.Tag = "Y" _
      Or (m_CP10 = "1001" And Val(txt1(8).Text) = 0 And Val(lbl1(28).Caption) = 0 And m_CP01 = "FCP") Then
      
      Me.SSTab2.TabVisible(1) = False: Me.SSTab2.TabCaption(0) = "" 'Added by Lydia 2015/04/24
      Me.SSTab2.TabVisible(0) = True: Me.SSTab2.Tab = 0 'Add By Sindy 2023/11/20
      strExc(0) = "select pa162,DST05,DST09 from caseprogress a,patent,divsugtext" & _
         " where cp09='" & m_strCP09 & "'" & _
         " and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04" & _
         " and pa08 in ('1','2')" & _
         " and dst01(+)=pa01 and dst02(+)=pa02 and dst03(+)=pa03 and dst04(+)=pa04"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         txtPA162 = "" & RsTemp(0)
         txtPA162.Tag = txtPA162.Text 'Added by Morgan 2022/8/1
         If strUserNum = m_CP14 Then
            If "" & RsTemp("DST09") = m_strCP09 And txtPA162 = "Y" Then txtDST05 = "" & RsTemp(1) 'Add By Sindy 2023/11/24
            'Memo by Lydia 承辦人可輸入建議定稿文字,欄位清空,下方保留上次記錄文字
            txtDST05Old = "" & RsTemp(1)
            txtPA162.Enabled = True
            txtDST05.Locked = False
            Command1.Enabled = True
            'Modified by Lydia 2015/04/24
            SSTab2.Visible = True
'            If m_CP10 = "1001" Then
'               txtEP09.Enabled = False
'            End If
            txtDST05.Height = 590: FramePA162.Visible = True 'Add By Sindy 2023/11/23
         Else
            'Memo by Lydia 非承辦人只show建議定稿文字
            txtDST05 = "" & RsTemp(1)
            'Modified by Lydia 2015/04/24
            SSTab2.Visible = True
            'SSTab2.Height = 1160
            txtDST05.Height = 590 * 2: FramePA162.Visible = False
         End If
      End If
   'Added by Lydia 2015/04/24 +中說請款修正定稿文字
   'Added by Lydia 2015/06/25 +主動修正203
   ElseIf InStr("201,209,210,235,203", m_CP10) > 0 And m_CP01 = "FCP" And Val(lbl1(28).Caption) = 0 Then
      Check1.Tag = "Y"
      'Added by Lydia 2015/06/25 判斷未經過"補輸中說"的主動修正,是否符合
      If m_CP10 = "203" Then
         strExc(0) = "select cp09 from caseprogress where cp01='" & m_CP01 & "' and cp02='" & m_CP02 & "' and cp03='" & m_CP03 & "' and cp04='" & m_CP04 & "' " & _
               " and cp57 is null and cp10 in ('201','209','210','235') "
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 0 Then
            Check1.Tag = ""
         End If
      End If
      'end 2015/06/25
      If Check1.Tag = "Y" Then
         If txtDST05.Tag = "Y" Then Check1.Visible = True
         Me.SSTab2.TabVisible(0) = False: Me.SSTab2.TabCaption(1) = ""
         Me.SSTab2.TabVisible(1) = True: Me.SSTab2.Tab = 1 'Add By Sindy 2023/11/20
         'Modified by Lydia 2015/11/26 AMD05長度已達2000字,Text高度拉長
         SSTab2.Visible = True
         'Modified by Lydia 2015/06/04 +CP27
         strExc(0) = "select AMD05,nvl(CP27,0) CP27 from caseprogress a,patent,Amendedtext" & _
                     " where cp09='" & m_strCP09 & "' and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04 " & _
                     " and amd01(+)=pa01 and amd01(+)=pa01 and amd02(+)=pa02 and amd03(+)=pa03 and amd04(+)=pa04"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            txtAMD05.Tag = "" & RsTemp(0)
            txtAMD05.Text = txtAMD05.Tag
         End If
         'Modified by Lydia 2015/06/12 除了承辦人外,核稿人也可輸入中說請款備註
         'If strUserNum = lblCP14 Or Pub_StrUserSt03 = "M51" Then
         'MODIFY BY SONIA 2015/6/17 再加下班翻譯的所內編號也可以輸 FCP-051573
         'If strUserNum = lblCP14 Or strUserNum = lblEP04 Or Pub_StrUserSt03 = "M51" Then
         If strUserNum = m_CP14 Or strUserNum = txt1(5) Or PUB_GetMapID(strUserNum, 0) = m_CP14 Or Pub_StrUserSt03 = "M51" Then
            'Modified by Lydia 2015/08/27 為了能拉動卷軸,改成locked
            'txtAMD05.Enabled = True
            txtAMD05.Locked = False
         End If
      End If
   'end 2015/04/24
   End If
   'end 2012/11/30
   '926.核對已准專利銷承辦期限
   Lbl926.Visible = False
   If m_CP10 = "926" Then
      If Val(txt1(8).Text) = 0 And Val(lbl1(28).Caption) = 0 And Val(txt1(12)) > 0 _
         And m_CP14 <> strUserNum Then
         txt1(12).Enabled = True
      Else
         txt1(12).Enabled = False
      End If
      'Modify By Sindy 2024/1/5
      If txt1(12) <> "" And lbl1(17) <> "" Then '判斷"有"承辦期限"有"本所期限
         Lbl926.Caption = "(二核)"
         Lbl926.Visible = True
      ElseIf txt1(12) <> "" And lbl1(17) = "" Then
         Lbl926.Caption = "(一核)"
         Lbl926.Visible = True
      End If
      '2024/1/5 END
   End If
'************************************* END
   
   CheckOC
   InitialField
   Set rsTmp = New ADODB.Recordset
   strSql = "SELECT * FROM CASEPROGRESS " & _
            "WHERE CP09 = '" & lbl1(3).Caption & "' "
   If rsTmp.State = 1 Then rsTmp.Close
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
       UpdateFieldOldData rsTmp
   End If
   If rsTmp.State = 1 Then rsTmp.Close
   
   Dim tmpInti As Integer
   For tmpInti = 0 To Combo4.ListCount - 1
       If Trim(txt1(6).Text) = Trim(Mid(Combo4.List(tmpInti), 1, InStr(1, Combo4.List(tmpInti), "=") - IIf(InStr(1, Combo4.List(tmpInti), "=") = 0, 0, 1))) Then
           Combo4.Text = Combo4.List(tmpInti)
       End If
   Next tmpInti
   'Add By Sindy 2025/9/18
   If Trim(txt1(6).Text) <> "" Then
      If Trim(txt1(6).Text) = Trim(Combo4.Text) Or Trim(Combo4.Text) = "" Then
         Combo4.Text = Trim(txt1(6).Text) & " ==> " & GetPrjSalesNM(Trim(txt1(6).Text))
      End If
   End If
   '2025/9/18 END
'*************************************
   If m_CPM29 = "N" Then
      Label18.Visible = False 'Modify By Sindy 2014/8/28 先不顯示
   Else
      Label18.Visible = False
   End If
   
   '若為個人工作管理及承辦人下拉選單為操作者
   If ProState = "1" Or Trim(Left("" & Combo1.Text, 6)) = strUserNum Then
      '未發文
      If Val(txt1(8).Text) = 0 Then
         If m_CPM29 = "" Then '要電子簽核的案件性質
            For Each oTxt1 In Me.txt1
               oTxt1.Enabled = False
            Next
         End If
         txt1(2).Enabled = True '齊備日
         'Add By Sindy 2024/3/8 核稿人
         If ProState = "1" Then
         '2024/3/8 END
            'Modify By Sindy 2025/3/18 +外專繪圖同新案翻譯歷程操作方式
            If (m_CP10 = "201" Or m_CP10 = "931") Then
               Me.Combo2.Enabled = False
            Else
               Me.Combo2.Enabled = True
            End If
         End If
         '2024/3/8 END
         Me.Combo4.Enabled = True '外文核稿人
         Me.Combo6.Enabled = True '判發人
         '日期有輸入過，就鎖住
         If Trim(txt1(2).Text) <> "" Then '完稿日
            txt1(2).Enabled = False
         End If
         If Trim(txt1(3).Text) <> "" Then '完稿日
            txt1(3).Enabled = False
         End If
         If Trim(txt1(7).Text) <> "" Then '核稿期限
            txt1(7).Enabled = False
         End If
         If Trim(txt1(19).Text) <> "" Then '外文核完日
            txt1(19).Enabled = False
         End If
         If Trim(txt1(13).Text) <> "" Then 'Claims完稿日
            txt1(13).Enabled = False
         End If
         
         'Add By Sindy 2024/1/8 927=其他翻譯
         If m_CP10 = "201" Or m_CP10 = "927" Then
            txt1(0).Enabled = False '工作時數
            txt1(1).Enabled = True '核稿時數
         Else
            txt1(0).Enabled = True
            txt1(1).Enabled = False
         End If
         '2024/1/8 END
      End If
      
   ElseIf ProState = "2" Then '主管權限
      frm090614.TextOk = True
      '第二級期限管制人有修改權限
      'Modify By Sindy 2024/3/5 + 協助機械組內專主管
      If Me.txt1(0).Visible And Me.txt1(0).Enabled = True Then Me.txt1(0).SetFocus
      If (Pub_StrUserSt03 = "M51" Or _
         PUB_GetST52(Trim(Left("" & Combo1.Text, 6)), strUserNum) = True Or _
         InStr(Pub_GetSpecMan("承辦人工作管理可修改資料人員"), strUserNum) > 0 Or _
         InStr(Pub_GetSpecMan("協助機械組內專主管"), strUserNum) > 0) _
         And Trim(Left("" & Combo1.Text, 6)) <> strUserNum Then
         
         '有修改權限
         'Add By Sindy 2024/1/8 927=其他翻譯
         If m_CP10 = "201" Or m_CP10 = "927" Then
            txt1(0).Enabled = False '工作時數
            txt1(1).Enabled = True '核稿時數
            'Add By Sindy 2024/3/27 無完稿日時,核稿工程師不可修改
            If m_CP10 = "201" And Trim(txt1(3).Text) = "" Then
               Me.Combo2.Enabled = False
            Else
               Me.Combo2.Enabled = True
            End If
            '2024/3/27 END
         Else
            txt1(0).Enabled = True
            txt1(1).Enabled = False
         End If
         '2024/1/8 END
      Else
         '無權限
         For Each oTxt1 In Me.txt1
            oTxt1.Enabled = False
         Next
         txtEP12.Enabled = False
         Me.Combo2.Enabled = False '核稿人
         Me.Combo4.Enabled = False '外文核稿人
         Me.Combo6.Enabled = False '判發人
      End If
   End If
   '已發文
   If Val(txt1(8).Text) > 0 Then
      txtEP12.Enabled = False
      For Each oTxt1 In Me.txt1
         oTxt1.Enabled = False
      Next
      Me.Combo2.Enabled = False '核稿人
      Me.Combo4.Enabled = False '外文核稿人
      Me.Combo6.Enabled = False '判發人
   End If
   
   'Add By Sindy 2024/8/30 已有送英核歷程,外文核稿人鎖住
   If PUB_ChkEmpFlowExists(lbl1(3), EMP_送英核) = True Then
      Combo4.Enabled = False
   End If
   '2024/8/30 END
   
   Call SetColTag(True)
End Sub

'Add By Sindy 2025/9/18 預設核稿語文
Private Sub SetEP41(strEP41 As String, strEP03 As String)
   'Add By Sindy 2024/1/8 顯示核稿語文
   m_EP41 = strEP41
   If m_EP41 = "" Or PUB_ChkEmpFlowExists(lbl1(3), EMP_送英核) = False Then '尚未送英核時才預設核稿語文
      txt1(23) = "1" '預設英文
      If m_EMPST16 = "3" Then
         txt1(23) = "2" '日文
         'Add By Sindy 2025/9/18
         If PUB_GetST93(strEP03) = "F41" Then
            txt1(23) = "1" '英文
         End If
         '2025/9/18 END
      End If
   Else
      txt1(23) = m_EP41
   End If
   'Add By Sindy 2025/9/18 不同要存檔,設m_EP41 = ""
   If m_EP41 <> txt1(23) Then
      m_EP41 = ""
   End If
End Sub

'Add By Sindy 2013/6/10
'bolSetTag=true : 將輸入欄位值記錄至.tag裡面
'bolSetTag=false : 比較輸入欄位值.Tag與畫面上資料是否一致
Private Function SetColTag(bolSetTag As Boolean) As Boolean
Dim oTxt1 As Object
   
   If bolSetTag = True Then
      For Each oTxt1 In txt1
         oTxt1.Tag = oTxt1.Text
      Next
      txtEP12.Tag = txtEP12
      txtCP64.Tag = txtCP64
      Combo2.Tag = Combo2.Text '核稿人
      Combo6.Tag = Combo6.Text '判發人
      Combo4.Tag = Combo4.Text
      txtPA162.Tag = txtPA162.Text
      txtDST05.Tag = txtDST05.Text
      txtAMD05.Tag = txtAMD05.Text
      'Add By Sindy 2024/1/9
      If m_EP41 = "" Then
         txt1(23).Tag = ""
      End If
      '2024/1/9 END
   Else
      SetColTag = True
      For Each oTxt1 In txt1
         If oTxt1.Tag <> oTxt1.Text Then SetColTag = False: Exit Function
      Next
      If txtEP12.Tag <> txtEP12 Then SetColTag = False: Exit Function
      If txtCP64.Tag <> txtCP64 Then SetColTag = False: Exit Function
      If Left(Combo2.Tag, 5) <> Left(Combo2.Text, 5) Then SetColTag = False: Exit Function '核稿人
      If Left(Combo6.Tag, 5) <> Left(Combo6.Text, 5) Then SetColTag = False: Exit Function '判發人
      If Left(Combo4.Tag, 5) <> Left(Combo4.Text, 5) Then SetColTag = False: Exit Function
      If txtPA162.Tag <> txtPA162 Then SetColTag = False: Exit Function
      If txtDST05.Tag <> txtDST05 Then SetColTag = False: Exit Function
      If txtAMD05.Tag <> txtAMD05 Then SetColTag = False: Exit Function
   End If
End Function

'91.08.14  nick  加畫面顯示其他國外案
Private Sub Cmd1_Click()
Dim iMouse As Integer
iMouse = Screen.MousePointer

Me.Hide
Screen.MousePointer = vbHourglass
frm090201_2_1.SetParent Me 'Add By Sindy 2014/1/14
frm090201_2_1.Show
frm090201_2_1.StrMenu (lbl1(7).Caption)
'Modify by Morgan 2009/11/12
'Screen.MousePointer = vbDefault
Screen.MousePointer = iMouse
End Sub

'Add By Sindy 2013/8/19
Private Sub cmdDetail_Click()
   Call grd2_DblClick
End Sub
'Added by Morgan 2024/3/29
'撰寫信函
Private Sub cmdLetter_Click()
   Me.Enabled = False
   Screen.MousePointer = vbHourglass
   Call Forms(0).SetTmpfrm090401
   With Tmpfrm090401
   .OutCallCP09 = m_strCP09
   .Text1 = m_CP01
   .Text2 = m_CP02
   .Text3 = m_CP03
   .Text4 = m_CP04
   .Option1(1).Value = True '點選英文
   .Option2.Value = True '點選FC代理人
   .Command1.Value = True
   .Command2.Value = True
   End With
   Set Tmpfrm090401 = Nothing
   Screen.MousePointer = vbDefault
   Me.Enabled = True
End Sub

'Modify By Sindy 2013/5/21
'Private Sub cmdOK_Click(index As Integer)
Public Sub cmdok_Click(Index As Integer)
'2013/5/21 End
   '***2008/11/21 加註BY SONIA 按確定後很快按結束會因為DoEvents造成錯誤,因使用者未反應故暫不取消DoEvents
   Dim iMouse As Integer
'   Dim bolUpdDate As Boolean 'Add By Sindy 2013/10/4
   
   iMouse = Screen.MousePointer
   
   Select Case Index
   
   Case 1 '確定
'         'Add By Sindy 2017/8/3 個人案件不可用主管權限操作
'         If ProState = "2" And m_CP14 = strUserNum Then '2.主管
'            MsgBox "個人案件不可用主管權限操作！", vbExclamation
'            Exit Sub
'         End If
'         '2017/8/3 END
         
         Select Case ProState
         Case "1", "2"
            m_chkcmdok1 = False 'Add By Sindy 2013/6/7 進入承辦歷程時會先執行一次確定鍵,因有可能已在此畫面先修改資料,且有些日期檢查條件須先執行
            
            'add by nickc 2007/12/28 加入修正
            If SSTab1.Tab = 0 Then Exit Sub '*****
            
            Screen.MousePointer = vbHourglass
            'Modify By Sindy 2017/9/15
            'If SSTab1.Tab = 1 Then
            'If SSTab1.Tab = 1 Or (SSTab1.Tab = 2 And Me.m_Flow <> "") Then
            If SSTab1.Tab = 1 Or Me.m_Flow <> "" Then 'Modify By Sindy 2017/9/20
            '2017/9/15 END
               If ChkNoData = False Then
                  '重新檢查欄位有效性
                  If TxtValidate = True Then
                     'DoEvents 'Modify By Sindy 2024/3/13 mark
                     'Me.Enabled = False
                     If FormSave = True Then
                        '集中發信
                        If m_Flow = "" Then
                           BatctMail
'                        'Add By Sindy 2024/1/25
'                        Else
'                           Call Process(lbl1(3)) '要重新查詢資料
'                        '2024/1/25 END
                        End If
                        'Add By Sindy 2024/3/6
                        'Modify By Sindy 2025/3/18 +外專繪圖同新案翻譯歷程操作方式
                        If (cboCP14.Visible = True And m_CP14 <> Trim(Left(cboCP14.Text, 5))) Or _
                           (Combo2.Visible = True And Combo2.Enabled = True And (m_CP10 = "201" Or m_CP10 = "931") And Trim(Left(Combo2.Tag, 5)) <> Trim(Left(Combo2.Text, 5))) Then
                           StrMenu1
                           StrMenu
                        Else
                        '2024/3/6 END
                           '更新mdb暫存資料及第一畫面的Grid內容
                           UpdEngMdb
                           TextOk = False
                           'add by nickc 2005/07/11 讓存完檔的變色正確
                           'ChgGrdColor 'Remove by Morgan 2009/11/12 UpdEngMdb內做就好
                           Call SetColTag(True) 'Add By Sindy 2013/6/10
                           m_chkcmdok1 = True 'Add By Sindy 2013/6/7
                        End If
'                     Else
'                        MsgBox "存檔失敗!" & Err.Number & Err.Description
                     End If
                     'Me.Enabled = True
                     'Add By Sindy 2017/9/21
                     If cmdok(1).Enabled = True Then
                     '2017/9/21 END
                        SSTab1.Tab = 0
                     End If
'                     'Modify By Sindy 2013/6/7
'                     If intBackTab = 2 Then
'                        Call QueryData(True)
'                     End If
'                     SSTab1.Tab = intBackTab
'                     intBackTab = 0
'                     '2013/6/7 End
                  End If
               End If
            Else
               SSTab1.Tab = 1
            End If
            'Modify by Morgan 2009/11/12
            'Screen.MousePointer = vbDefault
            Me.m_Flow = "" 'Add By Sindy 2017/8/28
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
        Case Else
        End Select
   Case Else
   End Select
End Sub

Private Sub cmdok2_Click(Index As Integer)
Dim iMouse As Integer
Dim bolReadAllEmp As String 'Add By Sindy 2024/6/3

iMouse = Screen.MousePointer
Screen.MousePointer = vbHourglass
grd1.Visible = False

'Add By Sindy 2024/6/3
bolReadAllEmp = False
If ProState = "2" Then
   If frm090614.txt1(8) = "N" Then
      bolReadAllEmp = True
   End If
End If
'2024/6/3 END

Select Case Index
Case 0 '當月資料
      'Modify By Sindy 2023/12/26 +,R110033 取消=,format(R110026 * r110027 * iif(r110030='N',1,iif(isnull(r110028) or r110028=0,1,r110028)),'0.00')
      'Modify By Sindy 2024/2/23 +,R110035
      'Modify By Sindy 2025/3/18 +外專繪圖同新案翻譯歷程操作方式
      If bolReadAllEmp = True Then
         strSql = "SELECT R110002,R110003,R110004,R110005,R110006,R110023,R110008,R110009,R110007,R110011,R110033,R110010,R110012,R110013,R110014,R110031,R110015,R110016,R110017,R110018,R110019,R110020,R110021,R110022, R110024,r110029,r110025,r110030,r110032,R110035,R110005" & _
                  " FROM R090614" & _
                  " WHERE ID='" & strUserNum & "'" & _
                  " AND (R110001 IN (" & Combo1_String & ") or (instr('" & Combo1_Name & "',R110016)>0 and R110034 in('201','931')))" & _
                  " ORDER BY R110002 desc,R110003,R110004 "
      Else
         strSql = "SELECT R110002,R110003,R110004,R110005,R110006,R110023,R110008,R110009,R110007,R110011,R110033,R110010,R110012,R110013,R110014,R110031,R110015,R110016,R110017,R110018,R110019,R110020,R110021,R110022, R110024,r110029,r110025,r110030,r110032,R110035,R110005" & _
                  " FROM R090614" & _
                  " WHERE ID='" & strUserNum & "'" & _
                  " AND (R110001='" & Trim(Left("" & Combo1.Text, 6)) & "' or (instr('" & Trim("" & Combo1.Text) & "',R110016)>0 and R110034 in('201','931')))" & _
                  " ORDER BY R110002 desc,R110003,R110004 "
      End If
      CheckOC
      With adoRecordset
          .CursorLocation = adUseClient
          .Open strSql, adoEng, adOpenStatic, adLockReadOnly
          If .RecordCount <> 0 And .RecordCount > 0 Then
               If InStr(UCase(Pub_GetModuleFileName), "VB6.EXE") <> 0 Then LblCnt.Caption = LblCnt.Caption & "/" & .RecordCount 'Add By Sindy 2023/12/7
               Set grd1.Recordset = adoRecordset
               ChgGrdColor
          Else
               grd1.Clear
               grd1.Rows = 2
          End If
      End With
      CheckOC
      SWPRow = 1
Case 1 '未發文
      'Modify By Sindy 2023/12/26 +,R110033 取消=,format(R110026 * r110027 * iif(r110030='N',1,iif(isnull(r110028) or r110028=0,1,r110028)),'0.00')
      'Modify By Sindy 2024/2/23 +,R110035
      'Modify By Sindy 2025/3/18 +外專繪圖同新案翻譯歷程操作方式
      If bolReadAllEmp = True Then
         strSql = "SELECT R110002,R110003,R110004,R110005,R110006,R110023,R110008,R110009,R110007,R110011,R110033,R110010,R110012,R110013,R110014,R110031,R110015,R110016,R110017,R110018,R110019,R110020,R110021,R110022, R110024,r110029,r110025,r110030,r110032,R110035,R110005" & _
                  " FROM R090614" & _
                  " WHERE ID='" & strUserNum & "'" & _
                  " AND (R110001 IN (" & Combo1_String & ") or (instr('" & Combo1_Name & "',R110016)>0 and R110034 in('201','931')))" & _
                  " AND (R110018='' or R110018='0') and (R110024='' or R110024='0' or R110034='908')" & _
                  " order by R110002 desc "
      Else
         strSql = "SELECT R110002,R110003,R110004,R110005,R110006,R110023,R110008,R110009,R110007,R110011,R110033,R110010,R110012,R110013,R110014,R110031,R110015,R110016,R110017,R110018,R110019,R110020,R110021,R110022, R110024,r110029,r110025,r110030,r110032,R110035,R110005" & _
                  " FROM R090614" & _
                  " WHERE ID='" & strUserNum & "'" & _
                  " AND (R110001='" & Trim(Left("" & Combo1.Text, 6)) & "' or (instr('" & Trim("" & Combo1.Text) & "',R110016)>0 and R110034 in('201','931')))" & _
                  " AND (R110018='' or R110018='0') and (R110024='' or R110024='0' or R110034='908')" & _
                  " order by R110002 desc "
      End If
      CheckOC
      With adoRecordset
          .CursorLocation = adUseClient
          .Open strSql, adoEng, adOpenStatic, adLockReadOnly
          If .RecordCount <> 0 And .RecordCount > 0 Then
               If InStr(UCase(Pub_GetModuleFileName), "VB6.EXE") <> 0 Then LblCnt.Caption = LblCnt.Caption & "/" & .RecordCount  'Add By Sindy 2023/12/7
               Set grd1.Recordset = adoRecordset
               ChgGrdColor
          Else
               grd1.Clear
               grd1.Rows = 2
               'SetGrd1
          End If
      End With
      CheckOC
      SWPRow = 1
Case Else
End Select
SetGrd1
Call MouseClick(1)
grd1.Visible = True
Screen.MousePointer = iMouse
End Sub

'Add By Sindy 2024/3/13
Private Sub CmdRead_Click()
   cmdRead.Enabled = False
   
   '查詢再會當掉的話,就用此按鍵
   
   cmdRead.Enabled = True
End Sub

'Add By Sindy 2024/3/13
Private Sub Combo1_Click()
   'cmdRead_Click
End Sub
Private Sub Combo1_GotFocus()
   Combo1.SelStart = 0
   Combo1.SelLength = Len(Combo1.Text)
End Sub

'Modify By Sindy 2014/1/16
'Private Sub Combo1_Click()
'Modified by Lydia 2021/12/28 Form2.0點選同一人不會觸發Click事件，改用DropButtonClick事件但要控制第2次才執行
'Public Sub Combo1_Click()
''2014/1/16 END
Public Sub Combo1_DropButtonClick()
   Static bClick As Boolean
   If bClick = False Then
      bClick = True
      Exit Sub
   End If
   bClick = False
'end 2021/12/28
   
   Call QueryCombo1Data 'Modify By Sindy 2025/4/10 改為共用函數
'   Me.Enabled = False 'Add By Sindy 2024/3/13
'
'   Dim iMouse As Integer
'   iMouse = Screen.MousePointer
'
'   Me.GRD1.Visible = False
'   Screen.MousePointer = vbHourglass
'   Me.MousePointer = vbHourglass
'   GRD1.MousePointer = flexArrowHourGlass
'   Me.Enabled = False
'   Combo1.Enabled = False
'
'   Label3.Caption = GetPrjSalesNM(Trim(Left("" & Combo1.Text, 6)))
'   'Add By Sindy 2013/9/16 在切換承辦人時,會出現”陣列索引超出範圍”
''   If Combo1.Tag <> Combo1.Text Then
'      SWPRow = 0
'      dblPrevRow = 0
'      Combo1.Tag = Combo1.Text
'      StrMenu1
'      StrMenu
''   End If
'   '2013/9/16
''    StrMenu1
''    StrMenu
'
''   If ChkNoData = True Then
''      'Modified by Lydia 2021/12/28 10=>9
''      For s = 0 To 9
''         If s <> 0 And s <> 1 And s <> 4 Then 'Add By Sindy 2023/10/2 + if
''            Txt1(s).Enabled = False
''         End If
''      Next s
''   Else
''      'Modified by Lydia 2021/12/28 10=>9
''      For s = 0 To 9
''         If s <> 0 And s <> 1 And s <> 4 Then 'Add By Sindy 2023/10/2 + if
''            Txt1(s).Enabled = True
''         End If
''      Next s
''   End If
''   SetGrd1
'
'   'DoEvents 'Modify By Sindy 2024/3/13 mark
'   'cmdok2(0).SetFocus
'   'If cmdok2(0).Visible = True And cmdok2(0).Enabled = True Then cmdok2(0).SetFocus
'
'   Combo1.Enabled = True
'   Me.Enabled = True
'   GRD1.MousePointer = flexDefault
'   Me.MousePointer = vbDefault
'   'Modify by Morgan 2009/11/12
'   'Screen.MousePointer = vbDefault
'   Screen.MousePointer = iMouse
'   Me.GRD1.Visible = True
'
'   If Me.GRD1.Visible = True And Me.GRD1.Enabled = True Then Me.GRD1.SetFocus 'Add By Sindy 2024/3/13
'
'   Me.Enabled = True 'Add By Sindy 2024/3/13
End Sub

'Modify By Sindy 2025/4/10
Public Sub QueryCombo1Data()
   Me.Enabled = False 'Add By Sindy 2024/3/13
   
   Dim iMouse As Integer
   iMouse = Screen.MousePointer
   
   Me.grd1.Visible = False
   Screen.MousePointer = vbHourglass
   Me.MousePointer = vbHourglass
   grd1.MousePointer = flexArrowHourGlass
   Me.Enabled = False
   Combo1.Enabled = False
   
   Label3.Caption = GetPrjSalesNM(Trim(Left("" & Combo1.Text, 6)))
   'Add By Sindy 2013/9/16 在切換承辦人時,會出現”陣列索引超出範圍”
'   If Combo1.Tag <> Combo1.Text Then
      SWPRow = 0
      dblPrevRow = 0
      Combo1.Tag = Combo1.Text
      StrMenu1
      StrMenu
'   End If
   '2013/9/16
'    StrMenu1
'    StrMenu

'   If ChkNoData = True Then
'      'Modified by Lydia 2021/12/28 10=>9
'      For s = 0 To 9
'         If s <> 0 And s <> 1 And s <> 4 Then 'Add By Sindy 2023/10/2 + if
'            Txt1(s).Enabled = False
'         End If
'      Next s
'   Else
'      'Modified by Lydia 2021/12/28 10=>9
'      For s = 0 To 9
'         If s <> 0 And s <> 1 And s <> 4 Then 'Add By Sindy 2023/10/2 + if
'            Txt1(s).Enabled = True
'         End If
'      Next s
'   End If
'   SetGrd1
   
   'DoEvents 'Modify By Sindy 2024/3/13 mark
   'cmdok2(0).SetFocus
   'If cmdok2(0).Visible = True And cmdok2(0).Enabled = True Then cmdok2(0).SetFocus
   
   Combo1.Enabled = True
   Me.Enabled = True
   grd1.MousePointer = flexDefault
   Me.MousePointer = vbDefault
   'Modify by Morgan 2009/11/12
   'Screen.MousePointer = vbDefault
   Screen.MousePointer = iMouse
   Me.grd1.Visible = True
   
   If Me.grd1.Visible = True And Me.grd1.Enabled = True Then Me.grd1.SetFocus 'Add By Sindy 2024/3/13
   
   Me.Enabled = True 'Add By Sindy 2024/3/13
End Sub
'Sindy 2025/4/10 END

'核稿人
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

Private Sub Combo4_Click()
   txt1(6).Text = Trim(Left(Me.Combo4.Text, 6))
End Sub

Private Sub Combo4_LostFocus()
   Dim ii As Integer
   Dim blnMatch As Boolean
   
   For ii = 0 To Me.Combo4.ListCount - 1
       blnMatch = False
       If Trim(Left(Me.Combo4.List(ii), 6)) = Trim(Left(Me.Combo4.Text, 6)) Then
           Me.Combo4.ListIndex = ii
           Call SetEP41(txt1(23), txt1(6)) 'Add By Sindy 2025/9/18 預設核稿語文
           blnMatch = True
           Exit For
       End If
   Next ii
   'Modify By Sindy 2025/9/18
   'Me.Combo4.ListIndex = 0
   If blnMatch = False Then
      strExc(10) = GetPrjSalesNM(Left(Combo4.Text, 5))
      If strExc(10) = "" Then
         Me.Combo4.ListIndex = 0
      Else
         Combo4.Text = Left(Combo4.Text, 5) & " ==> " & strExc(10)
         txt1(6) = Left(Combo4.Text, 5)
         Call SetEP41(txt1(23), txt1(6)) '預設核稿語文
      End If
   End If
   '2025/9/18 END
End Sub

Private Sub Combo5_Click()
   If Me.Visible = True Then
      If QueryData(True) = False Then ShowNoData 'Add By Sindy 2023/4/12
   End If
End Sub

'Add By Sindy 2015/5/21 判發人
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
         'Modify By Sindy 2024/1/3
         If Combo6.ListCount > 0 Then
         '2024/1/3 END
            Me.Combo6.ListIndex = 0
         End If
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

Private Sub Form_Activate()
'Dim nFrm As Form
   ProState = m_ProState 'Add By Sindy 2017/8/10 重新設定權限
   If PUB_ChkFormIsClose("frm090202_2", "承辦") = False Then Exit Sub 'Add By Sindy 2020/1/21
'   'Add By Sindy 2017/8/30
'   '檢查表單是否已開啟，若是，則關閉
'   If Me.Visible = True Then
'      For Each nFrm In Forms
'         If StrComp(nFrm.Name, "frm090202_2", vbTextCompare) = 0 Then
'            'If frm090202_2.intReceiveKind = 0 Then '0.承辦人工作進度
'            'If frm090202_2.lblCP09.Caption = "" Then Unload frm090202_2: Exit For
'            If UCase(frm090202_2.m_PrevForm.Name) <> UCase(Me.Name) Then Exit For
'            If Not (frm090202_2.cmdAdd.Visible = False And frm090202_2.cmdSend.Enabled = False) Then
'               Unload frm090202_2
'            End If
'         End If
'      Next
'   End If
'   '2017/8/30 END
End Sub

Private Sub Form_Load()
Dim iMouse As Integer
Dim nFrm As Form 'Add By Sindy 2018/1/24
   
   iMouse = Screen.MousePointer
   
'   'Add By Sindy 2018/1/24
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
'   '2018/1/24 END
   If PUB_ChkFormIsClose("frm090202_2") = False Then Exit Sub 'Add By Sindy 2020/1/21
   
   m_ProState = ProState 'Add By Sindy 2017/8/10 記錄目前權限
   
   'add by nick 2006/02/27 重新定義
   ReDim m_FieldList(TF_CP)
   InitialField
   
   'Added by Morgan 2016/2/18
   '配合主畫面,調整表單起始大小( 預設大小 >= 起始大小 <=1024 * Screen.TwipsPerPixelX )
   lngFormWidth = 9435
   lngFormHeight = 6950 'Modified by Lydia 2021/12/28 Height 6120=> 6950
   If Forms(0).Width >= 1024 * Screen.TwipsPerPixelX Then
      lngFormWidth = 1024 * Screen.TwipsPerPixelX - 200
   ElseIf Forms(0).Width >= Me.Width Then
      lngFormWidth = Forms(0).Width - 200
   End If
   Me.Width = lngFormWidth
   Me.Height = lngFormHeight
   'end 2016/2/18
   
   MoveFormToCenter Me
   ReDim skMail(0) As SeekMails
   
   '讀取各基本檔可用系統別
   m_SqlGrpStr1 = SQLGrpStr("", 1)
   m_SqlGrpStr5 = SQLGrpStr("", 5)
   
   Combo5.Text = Combo5.List(3) 'Add By Sindy 2013/9/17
   
   Select Case ProState
   Case "1" '個人
      'add by nickc 2007/12/14
      '讀取使用權限
      Me.Caption = "外專工作進度資料維護 (個人)" 'Add By Sindy 2024/2/23
      bolInsert = IsUserHasRightOfFunction("frm090201_4", strAdd, False)
      bolUpdate = IsUserHasRightOfFunction("frm090201_4", strEdit, False)
      bolDelete = IsUserHasRightOfFunction("frm090201_4", strDel, False)
      bolSelect = IsUserHasRightOfFunction("frm090201_4", strFind, False)
      bolPrint = IsUserHasRightOfFunction("frm090201_4", strPrint, False)
      
      TextOk = True
      '統計年月(個人抓系統日的年月)
      Text1.Text = Mid(strSrvDate(1), 1, 6)
      
   Case "2" '主管 承辦人管理工作進度資料查詢
      'add by nickc 2007/12/14
      Me.Caption = "外專工作進度資料維護 (主管)" 'Add By Sindy 2024/2/23
      bolInsert = IsUserHasRightOfFunction("frm090614", strAdd, False)
      bolUpdate = IsUserHasRightOfFunction("frm090614", strEdit, False)
      bolDelete = IsUserHasRightOfFunction("frm090614", strDel, False)
      bolSelect = IsUserHasRightOfFunction("frm090614", strFind, False)
      bolPrint = IsUserHasRightOfFunction("frm090614", strPrint, False)
      
      frm090614.TextOk = True
      cmdok(2).Caption = "回前畫面"
      '統計年月(管理抓查詢畫面輸入的年月)
      Text1.Text = Trim((Val(frm090614.txt1(3)) + 1911)) & Trim(Right(ChgNumByNick(frm090614.txt1(4)), 2))
      
   Case "3" '分所
      'add by nickc 2007/12/14
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
         'Add By Sindy 2013/9/17
         Combo1.AddItem strUserNum & " " & "(" & strUserName & ")", 0
         Combo1.Text = Combo1.List(0)
         '2013/9/17 END
         'StrMenu1 'Modify By Sindy 2016/9/6 因前句Combo1就會run 到 StrMenu1
         StrMenu1  'Added by Lydia 2021/12/28 因為Combo1改成Form 2.0不使用Combo1_Click，所以預設先執行承辦人選單
         
         SetEngineer '設定承辦人選單
         'Add By Sindy 2013/9/16 檢查當時是否需要為他人職代
         Call Pub_SetForOthersEmpCombo(strUserNum, Combo1, False)
         '2013/9/16 END
   Case "2" '承辦人管理工作進度資料查詢
         frm090614.Process4
         If Combo1_String <> "" And frm090614.txt1(8) = "" Then Combo1.ListIndex = 0 'Add By Sindy 2024/3/5
         StrMenu1
   Case Else
   End Select
   Label3.Caption = GetPrjSalesNM(Trim(Left("" & Combo1.Text, 6)))
   
   SetEngChecker '設定外文核稿人選單

   StrMenu
   
   Select Case ProState
   Case "1"
      If TextOk = False Then Screen.MousePointer = iMouse: GoTo EXITSUB
      'Add By Sindy 2013/9/16
      'Combo1.Enabled = False
      Combo1.Enabled = True
      '2013/9/16 END
   Case "2"
      If frm090614.TextOk = False Then Screen.MousePointer = iMouse: TextOk = True: GoTo EXITSUB
      Combo1.Enabled = True
   Case Else
   End Select
   
   'SetGrd1
   Call MouseClick(1)
   Screen.MousePointer = iMouse
   Me.SSTab1.Tab = 0
   Me.Combo3.ListIndex = 0
   'Add By Sindy 2013/5/16
'   If Pub_StrUserSt03 = "M51" Or InStr(UCase(Pub_GetModuleFileName), "VB6.EXE") <> 0 Then
      Me.cmd(1).Visible = True '承辦歷程
      'Add By Sindy 2013/6/7
'      If ProState = "1" Then '個人
         Me.SSTab1.TabVisible(2) = True '待辦歷程
         SSTab1.Tab = 2
         If QueryData(True) = False Then
            SSTab1.Tab = 0
         End If
'      Else
'         Me.SSTab1.TabVisible(2) = False
'      End If
      '2013/6/7 End
'   Else
'      Me.cmd(1).Visible = False
'      Me.SSTab1.TabVisible(2) = False
'   End If
   '2013/5/16 End
   If bolUpdate = False Then
      cmdok(1).Visible = False
   End If
   
   'Add By Sindy 2023/12/29
   LblCnt.Visible = False
   If Pub_StrUserSt03 = "M51" Then
      LblCnt.Visible = True
   End If
   
   Exit Sub

EXITSUB:
   Me.Hide
   Select Case ProState
   Case "1"
        Me.Hide
   Case "2"
        frm090614.Show
        Me.Hide
   Case Else
   End Select
End Sub

'Add By Sindy 2013/6/7
Private Sub cmdQuery_Click()
   If QueryData(True) = False Then ShowNoData
End Sub

'Add By Sindy 2013/6/7
Public Function QueryData(bolFirst As Boolean) As Boolean
Dim rsTmp As New ADODB.Recordset
Dim strSql As String
Dim i As Integer
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
''   strVal = "(select * from EmpElectronProcess where eep01||eep02 in(select eep01||max(eep02) from EmpElectronProcess,caseprogress where eep01=cp09(+) and cp27 is null and cp57 is null and (EEP04 not in(" & EMP_待辦歷程查詢除外的狀態 & ") or (EEP04='" & EMP_聯絡 & "'" & IIf(strQuyDate <> "", " And EEP06>=" & strQuyDate, "") & ")) group by eep01) and EEP09 is null and EEP05='" & Trim(Left("" & Combo1.Text, 6)) & "'" & _
''            " union select EmpElectronProcess.* from EmpElectronProcess,caseprogress where eep01=cp09(+) and cp27 is null and cp57 is null and eep04='" & EMP_退件 & "' and eep09='Y' and EEP05='" & Trim(Left("" & Combo1.Text, 6)) & "'" & _
''            ") EmpElectronProcess,"
'   strVal = "(select * from EmpElectronProcess where eep01||eep02 in(select eep01||max(eep02) from EmpElectronProcess,caseprogress where eep01=cp09(+) and cp27 is null and cp57 is null and EEP04 not in(" & EMP_待辦歷程查詢除外的狀態 & ") group by eep01) and EEP09 is null and EEP05='" & Trim(Left("" & Combo1.Text, 6)) & "'" & _
'            " union select e1.* from EmpElectronProcess e1,caseprogress where e1.eep01=cp09(+) and cp27 is null and cp57 is null and e1.EEP02 in (select max(eep02) from EmpElectronProcess where eep01=e1.eep01) and e1.EEP05='" & Trim(Left(Combo1.text, 6)) & "' And e1.EEP04 in('" & EMP_聯絡 & "')" & IIf(strQuyDate <> "", " And EEP06>=" & strQuyDate, "") & _
'            " union select EmpElectronProcess.* from EmpElectronProcess,caseprogress where eep01=cp09(+) and cp27 is null and cp57 is null and eep04='" & EMP_退件 & "' and eep09='Y' and EEP05='" & Trim(Left("" & Combo1.Text, 6)) & "'" & _
'            ") EmpElectronProcess,"
   '2013/10/22 END
   'Modify By Sindy 2016/3/3 取消此句,因退件不會上待回覆Y " union select EmpElectronProcess.* from EmpElectronProcess,caseprogress where eep01=cp09(+) and cp27 is null and cp57 is null and eep04='" & EMP_退件 & "' and eep09='Y' and EEP05='" & Trim(Left("" & Combo1.Text, 6)) & "'"
   '                         增加EEP13='Y'
   strVal = "select eep01,eep02 from EmpElectronProcess where EEP13='Y' and EEP09 is null and EEP05='" & Trim(Left("" & Combo1.Text, 6)) & "' and EEP04 not in(" & EMP_待辦歷程查詢除外的狀態 & ")" & _
            " union select eep01,eep02 from EmpElectronProcess where EEP13='Y' and EEP05='" & Trim(Left("" & Combo1.Text, 6)) & "' and EEP04='" & EMP_聯絡 & "'" & IIf(strQuyDate <> "", " And EEP06>=" & strQuyDate, "")
   '2016/3/3 END
   'Modify By Sindy 2015/9/30 +IIf(Pub_StrUserSt15 = "P12", " And EEP04 not in('" & EMP_判發 & "','" & EMP_退件重送 & "')", "")
   'Modify By Sindy 2016/3/3 +不顯示
   'Modify By Sindy 2016/9/2 And cp27 is null And cp57 is null -> and cp158=0 and cp159=0
   strSql = "Select ' ' as V,EP01 as 目次,SqlDateT(EEP06)||' '||sqltime(EEP07) as 流程日期,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as 本所案號,PA05||PA06||PA07 as 案件名稱," & _
            "NA03 as 國家,Decode(PA09,'000',PTM03,PTM04) as 種類,Decode(PA09,'000',CPM03,CPM04) as 案件性質," & _
            "SqlDateT (cp06) as 本所期限, s1.ST02 as 承辦人, SqlDateT(cp48) as 承辦期限, s2.ST02 as 智權人員, ac03 as 目前流程狀態," & _
            "EEP01 as 總收文號,EEP02 as 序號,EP08,EP38,' ' as 不顯示,EEP06 a,EEP07 b" & _
            " From EmpElectronProcess,CaseProgress,EngineerProgress,Patent," & _
            "staff s1,staff s2,nation,CasePropertyMap,PatentTradeMarkMap,allcode" & _
            " Where (EEP01,EEP02) in(" & strVal & ")" & _
            " and EEP01=CP09(+)" & _
            " and cp158=0 and cp159=0" & _
            " And EEP01=EP02(+)" & _
            " And CP01=PA01 And CP02=PA02 And CP03=PA03 And CP04=PA04" & _
            " And CP14=s1.ST01(+) And CP13=s2.ST01(+)" & _
            " And PA09=NA01(+)" & _
            " And CP01=CPM01(+) And CP10=CPM02(+)" & _
            " And '1'=PTM01(+) AND PA08=PTM02(+)" & _
            " And ac01='09' And EEP04=ac02(+)"
   'If ProState = "1" Then '個人
      'Modify By Sindy 2024/4/29 + and ep09 is not null
      'Modify By Sindy 2025/3/18 +外專繪圖同新案翻譯歷程操作方式
      strSql = strSql & " And ((CP14='" & Trim(Left("" & Combo1.Text, 6)) & "' and cp10<>'201' and cp10<>'931') or (EP04='" & Trim(Left("" & Combo1.Text, 6)) & "' and cp10 in('201','931') and ep09 is not null))"
   'End If
   'Add By Sindy 2018/4/17
   strSql = strSql & " union " & _
            "Select ' ' as V,EP01 as 目次,SqlDateT(EEP06)||' '||sqltime(EEP07) as 流程日期,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as 本所案號,SP05||SP06||SP07 as 案件名稱," & _
            "NA03 as 國家,'' as 種類,Decode(SP09,'000',CPM03,CPM04) as 案件性質," & _
            "SqlDateT (cp06) as 本所期限, s1.ST02 as 承辦人, SqlDateT(cp48) as 承辦期限, s2.ST02 as 智權人員, ac03 as 目前流程狀態," & _
            "EEP01 as 總收文號,EEP02 as 序號,EP08,EP38,' ' as 不顯示,EEP06 a,EEP07 b" & _
            " From EmpElectronProcess,CaseProgress,EngineerProgress,servicepractice," & _
            "staff s1,staff s2,nation,CasePropertyMap,allcode" & _
            " Where (EEP01,EEP02) in(" & strVal & ")" & _
            " and EEP01=CP09(+)" & _
            " and cp158=0 and cp159=0" & _
            " And EEP01=EP02(+)" & _
            " And CP01=SP01 And CP02=SP02 And CP03=SP03 And CP04=SP04" & _
            " And CP14=s1.ST01(+) And CP13=s2.ST01(+)" & _
            " And SP09=NA01(+)" & _
            " And CP01=CPM01(+) And CP10=CPM02(+)" & _
            " And ac01='09' And EEP04=ac02(+)"
   'If ProState = "1" Then '個人
      'Modify By Sindy 2024/4/29 + and ep09 is not null
      'Modify By Sindy 2025/3/18 +外專繪圖同新案翻譯歷程操作方式
      strSql = strSql & " And ((CP14='" & Trim(Left("" & Combo1.Text, 6)) & "' and cp10<>'201' and cp10<>'931') or (EP04='" & Trim(Left("" & Combo1.Text, 6)) & "' and cp10 in('201','931') and ep09 is not null))"
   'End If
   'Modify By Sindy 2013/11/21
   'strSql = strSql & " order by EP01 desc"
   '2018/4/17 END
   strSql = strSql & " order by a desc,b desc"
   '2013/11/21 END
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      Set grd2.Recordset = rsTmp
      
      'Add By Sindy 2013/10/18
      For i = 1 To grd2.Rows - 1
         Call SetColColor(i)
      Next i
      '2013/10/18 END
      
   Else
      QueryData = False
      Screen.MousePointer = vbDefault
      rsTmp.Close
      Set rsTmp = Nothing
      Exit Function
   End If
      
ExitQuery:
'   '若有資料時游標停在第一筆
'   If bolFirst = True Then
'      grd2.Visible = False
'      grd2.col = 0
'      grd2.row = 1
'      If rsTmp.RecordCount > 0 Then
'         dblPrevRow = grd2.row
'         grd2.Text = "V"
'         m_intRow = 1: m_intCol = 0 'Add By Sindy 2016/3/10
'         For i = 0 To grd2.Cols - 1
'            grd2.col = i
'            'Modify By Sindy 2013/10/29
'            If grd2.CellBackColor <> &H8080FF Then
'               grd2.CellBackColor = &HFFC0C0
'            End If
'         Next i
'      End If
'      grd2.Visible = True
'   End If
   
   rsTmp.Close
   Screen.MousePointer = vbDefault
   'Me.SSTab1.Tab = 2 'Modify By Sindy 2017/8/21 Mark
   
EXITSUB:
   Set rsTmp = Nothing
End Function

'Add By Sindy 2013/10/18
Private Sub SetColColor(intRow As Integer)
Dim i As Integer
   
   grd2.row = intRow
   'Add By Sindy 2016/3/7 柏翰:繪圖判發跟退件要淺紅色表示
   If grd2.TextMatrix(intRow, 12) = "繪圖判發" Or _
          grd2.TextMatrix(intRow, 12) = "退件" Then
      grd2.col = 12
      grd2.CellBackColor = &HC0C0FF
   End If
End Sub

'Add By Sindy 2013/6/7
Private Sub SetGrd2()
   Dim arrGridHeadText, arrGridHeadWidth
   Dim iRow As Integer
   
   'Modify By Sindy 2016/3/3 +不顯示
   arrGridHeadText = Array("V", "目次", "流程日期", "本所案號", "案件名稱", _
                           "國家", "種類", "案件性質", "本所期限", "承辦人", _
                           "承辦期限", "智權人員", "目前流程狀態", _
                           "總收文號", "序號", "EP08", "EP38", "不顯示", "EEP06 a", "EEP07 b")
   arrGridHeadWidth = Array(200, 400, 800, 1400, 1000, _
                            700, 450, 900, 800, 600, _
                            800, 600, 600, _
                            0, 0, 0, 0, 600, 0, 0)
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

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'   'Add By Sindy 2020/3/19
'   If MsgBox("未完成稿件是否已上傳暫存區？", vbExclamation + vbYesNo + vbDefaultButton2, Me.Caption & " 重要訊息！") = vbNo Then
'      Cancel = True
'      Exit Sub
'   End If
'   '2020/3/19 END
End Sub

'Added by Morgan 2016/2/18
Private Sub Form_Resize()
   'Modified by Morgan 2016/3/11 不可最大化
   If Me.WindowState = 2 Then
      Me.WindowState = 0
   'end 2016/3/11
   ElseIf Me.WindowState = 0 Then
      If Me.Width < lngFormWidth Then Me.Width = lngFormWidth
      Me.Height = lngFormHeight '高度固定
      Me.SSTab1.Width = Me.Width - 200
      Me.grd1.Width = Me.SSTab1.Width - 200
      Me.grd2.Width = Me.SSTab1.Width - 200
   End If
End Sub

'Add By Sindy 2024/2/23
Private Sub GRD1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Static iRow As Integer, iCol As Integer
   
   If grd1.MouseRow <> 0 And (grd1.MouseCol = 14 Or grd1.MouseCol = 29) Then
      If iRow <> grd1.MouseRow Or iCol <> grd1.MouseCol Then
         strExc(10) = ""
         If grd1.MouseCol = 14 Then
            If grd1.TextMatrix(grd1.MouseRow, 29) <> "" Then
               strExc(10) = "最近歷程:" & grd1.TextMatrix(grd1.MouseRow, 29)
            End If
         ElseIf grd1.MouseCol = 29 Then
            strExc(10) = "本所案號:" & grd1.TextMatrix(grd1.MouseRow, 3) & "(" & grd1.TextMatrix(grd1.MouseRow, 7) & ")"
         End If
         CreateToolTip GetHWndForToolTip(grd1), strExc(10)
         iRow = grd1.MouseRow
         iCol = grd1.MouseCol
      End If
   End If
End Sub

'Add By Sindy 2016/3/3 增加不顯示功能
Private Sub grd2_Click()
   m_intRow = grd2.MouseRow
   m_intCol = grd2.MouseCol
   If m_intRow <> 0 Then
      If m_intCol = 17 Then '不顯示
         'Modify By Sindy 2024/1/4 + 轉檔完成
         '926=核對已准專利
         If grd2.TextMatrix(m_intRow, 13) <> "" And _
            grd2.TextMatrix(m_intRow, 12) <> "核修" And _
            (grd2.TextMatrix(m_intRow, 12) <> "核完" Or (m_CP10 = "926" And grd2.TextMatrix(m_intRow, 10) = "" And grd2.TextMatrix(m_intRow, 12) = "核完")) And _
            grd2.TextMatrix(m_intRow, 12) <> "會修" And _
            InStr(grd2.TextMatrix(m_intRow, 12), "會完") = 0 And _
            grd2.TextMatrix(m_intRow, 12) <> "繪圖判發" And _
            (grd2.TextMatrix(m_intRow, 12) <> "判發" Or (m_CP10 = "926" And grd2.TextMatrix(m_intRow, 10) = "" And grd2.TextMatrix(m_intRow, 12) = "判發")) And _
            grd2.TextMatrix(m_intRow, 12) <> "退回" And _
            grd2.TextMatrix(m_intRow, 12) <> "退件" And _
            grd2.TextMatrix(m_intRow, 12) <> "圖修" And _
            InStr(grd2.TextMatrix(m_intRow, 12), "圖完") = 0 And _
            grd2.TextMatrix(m_intRow, 12) <> "轉檔完成" Then
            
            grd2.TextMatrix(m_intRow, 17) = "V"
            If MsgBox("請再次確定不顯示 " & vbCrLf & grd2.TextMatrix(m_intRow, 3) & " " & grd2.TextMatrix(m_intRow, 12) & " 嗎？", vbExclamation + vbYesNo + vbDefaultButton2, "重要訊息！") = vbNo Then
               grd2.TextMatrix(m_intRow, 17) = ""
            Else
               strExc(0) = "update EmpElectronProcess set eep13=null" & _
                           " where eep01='" & grd2.TextMatrix(m_intRow, 13) & "'" & _
                             " and eep02=" & grd2.TextMatrix(m_intRow, 14)
               Pub_SeekTbLog strExc(0) 'Add By Sindy 2018/8/27
               cnnConnection.Execute strExc(0)
               grd2.RowHeight(m_intRow) = 0
            End If
         End If
      End If
   End If
End Sub

'Add By Sindy 2013/6/7
Private Sub grd2_DblClick()
Dim nFrm As Form
   
   'Modify By Sindy 2016/3/3
   If m_intRow <> 0 Then
      If m_intCol <> 17 Then
   '2016/3/3 END
         'Add By Sindy 2024/3/6
         If cboCP14.Visible = True And m_CP14 <> Trim(Left(cboCP14.Text, 5)) Then
            MsgBox "承辦人有異動，不可在此時同時操做歷程！" & vbCrLf & "請更新後再操作", vbExclamation
            Exit Sub
         End If
         'Modify By Sindy 2025/3/18 +外專繪圖同新案翻譯歷程操作方式
         If Combo2.Visible = True And Combo2.Enabled = True And (m_CP10 = "201" Or m_CP10 = "931") And Trim(Left(Combo2.Tag, 5)) <> Trim(Left(Combo2.Text, 5)) Then
            MsgBox "核稿工程師有異動，不可在此時同時操做歷程！" & vbCrLf & "請更新後再操作", vbExclamation
            Exit Sub
         End If
         '2024/3/6 END
         
         'Add By Sindy 2013/9/16
         If ProState = "2" Then
            If frm090614.txt1(8) = "N" Then MsgBox "不可從（不區分個人）的資料查詢中來執行承辦歷程作業！": Exit Sub
         End If
         '2013/9/16 END
         
         'Add By Sindy 2018/1/2 個人案件不可用主管權限操作
         If ProState = "2" And m_CP14 = strUserNum Then '2.主管
            MsgBox "個人案件不可用主管權限操作！", vbExclamation
            Exit Sub
         End If
         '2018/1/2 END
         
         'For i = 1 To grd2.Rows - 1
            'Add By Sindy 2017/11/9
            If dblPrevRow = 0 Then
               MsgBox "請點選一筆資料列!", vbExclamation
               Exit Sub
            End If
            '2017/11/9 END
            If grd2.TextMatrix(dblPrevRow, 0) = "V" Then
      '         If lbl1(3) <> grd2.TextMatrix(dblPrevRow, 13) Then
                  Call Process(grd2.TextMatrix(dblPrevRow, 13)) 'Modify By Sindy 2013/10/28 要重新查詢資料,因核稿人及判發人有預設問題 ex.P106408品薇在新增下一流程會變自行判發
      '         Else
'Modify By Sindy 2017/9/15 Mark
'                  If Me.cmd(1).Enabled = True Then
'                     If SetColTag(False) = False Then
'                        Call cmdok_Click(1)
'                        If m_chkcmdok1 = False Then Exit Sub
'                     End If
'                  End If
      '         End If
               If Me.cmd(1).Enabled = True Then
                  'Add By Sindy 2015/12/3
                  '重新檢查欄位有效性
                  If TxtValidate = True Then
                  '2015/12/3 END
                     
'                     'Add By Sindy 2017/9/19
'                     '檢查表單是否已開啟，若是，則關閉
'                     For Each nFrm In Forms
'                        If StrComp(nFrm.Name, "frm090202_2", vbTextCompare) = 0 Then
'                           Unload frm090202_2
'                           Exit For
'                        End If
'                     Next
'                     '2017/9/19 END
                     If PUB_ChkFormIsClose("frm090202_2") = False Then Exit Sub 'Add By Sindy 2020/1/17
                     intBackTab = 2
                     frm090202_2.Hide
                     frm090202_2.m_EEP01 = grd2.TextMatrix(dblPrevRow, 13) '總收文號
                     frm090202_2.m_FlowUserNum = Trim(Left("" & Combo1.Text, 6)) 'Add By Sindy 2013/9/12 案件流程所屬人員
                     frm090202_2.intReceiveKind = 0 '0.承辦人工作進度
                     frm090202_2.SetParent Me
                     If frm090202_2.QueryData = True Then
                        frm090202_2.Show
                        Me.Hide
                     End If
                     'Exit For
                  End If
               Else
                  Me.SSTab1.Tab = 1
               End If
            End If
         'Next i
      End If
   End If
End Sub

'Add By Sindy 2013/6/7
Private Sub GRD2_SelChange()
Dim j As Integer 'Add By Sindy 2016/3/4

grd2.Visible = False
'Add By Sindy 2016/3/4
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
         Call SetColColor(j)
         Exit For
      End If
   Next j
Else
'2016/3/4 END
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
         'Modify By Sindy 2013/10/29
         If grd2.CellBackColor <> &H8080FF Then
         '2013/10/29 END
            grd2.CellBackColor = QBColor(15)
         End If
      Next i
      Call SetColColor(CStr(dblPrevRow))
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
            'Modify By Sindy 2013/10/29
            If grd2.CellBackColor <> &H8080FF Then
            '2013/10/29 END
               grd2.CellBackColor = &HFFC0C0
            End If
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

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next

'add by nickc 2006/02/27
ClearFieldList

Set frm090909 = Nothing
End Sub

Sub StrMenu1()
Me.Enabled = False
'DoEvents 'Modify By Sindy 2024/3/13 mark
On Error GoTo ErrHnd 'Add By Sindy 2024/3/14
adoEng.Execute "drop table R090614 "
'Modify By Sindy 2023/12/26 +,R110033 text:指定送件日
'Modify By Sindy 2024/2/20 +,R110034 text:CP10
'Modify By Sindy 2024/2/23 +,R110035 text:GetEEPCurState(cp09)最近歷程
RunCreateTable: 'Add By Sindy 2024/3/14
adoEng.Execute "create table R090614 (R110001 text,R110002 double,R110003 text,R110004 text,R110005 text" & _
               ",R110006 text,R110007 text,R110008 text,R110009 text,R110010 text" & _
               ",R110011 text,R110012 text,R110013 text,R110014 text,R110015 text" & _
               ",R110016 text,R110017 text,R110018 text,R110019 double,R110020 memo" & _
               ",R110021 text,R110022 text,ID text,R110023 text, R110024 text,R110025 text" & _
               ",R110026 double,R110027 double,R110028 double,R110029 text,R110030 text" & _
               ",R110031 text,R110032 double,R110033 text,R110034 text,R110035 text)"
On Error GoTo 0 'Add by Sindy 2024/3/14 還原錯誤控制

Select Case ProState
Case "1" '承辦人個人工作進度資料維護
      StrGrp090201 = ""
      StrSQL6 = ""
      StrSQL61 = ""
      StrSQL62 = ""
      StrSQL63 = ""
      StrSQL64 = ""
      StrSPa = ""
      StrSSP = ""
       '含齊備日為當月, 發文日為19221111的資料
       'Modify By Sindy 2024/4/29 + and ep09 is not null
       'Modify By Sindy 2025/3/18 +外專繪圖同新案翻譯歷程操作方式
       StrSQL6 = StrSQL6 & " and (CP14='" & Trim(Left("" & Combo1.Text, 6)) & "' or (cp10 in('201','931') and EP04='" & Trim(Left("" & Combo1.Text, 6)) & "' and ep09 is not null)) and cp05>=19980101 "
       StrSQL61 = StrSQL61 & " and cp158=0 and cp159=0 "
       StrSQL62 = StrSQL62 & " and CP27>=" & Mid(strSrvDate(1), 1, 6) & "01 AND CP27<=" & Mid(strSrvDate(1), 1, 6) & "31 "
       StrSQL63 = StrSQL63 & " and CP57>=" & Mid(strSrvDate(1), 1, 6) & "01 AND CP57<=" & Mid(strSrvDate(1), 1, 6) & "31 and cp27 is null "
       StrSQL64 = StrSQL64 & " and CP05>=" & Mid(strSrvDate(1), 1, 6) & "01 AND CP05<=" & Mid(strSrvDate(1), 1, 6) & "31 and cp57 is null And CP05>CP27 "
       'C類未發文的不管閉不閉卷都要能在系統中作業
       'Modify By Sindy 2024/2/20 908=代辦退費未發文的不管閉不閉卷都要能在系統中作業
       StrSPa = StrSPa & " and ((pa58>=" & Mid(strSrvDate(1), 1, 6) & "01 AND pa58<=" & Mid(strSrvDate(1), 1, 6) & "31) or pa58 is null or cp27>0 or (cp01 in('P','FCP') and cp09>'C') or (cp01 in('P','FCP') and cp10 in('908'))) "
       StrSSP = StrSSP & " and ((sp16>=" & Mid(strSrvDate(1), 1, 6) & "01 AND sp16<=" & Mid(strSrvDate(1), 1, 6) & "31) or sp16 is null or cp27>0 or (cp01='FG' and cp10 in('908'))) "

Case "2" '承辦人管理工作進度資料查詢
   'Modify By Sindy 2024/3/5
   If Trim(Left(Combo1.Text, 6)) <> "" Then
      strExc(9) = Trim(Left(Combo1.Text, 6))
   ElseIf Combo1_String <> "" Then
      strExc(10) = Replace(Combo1_String, "','", ",")
      strExc(10) = Replace(Combo1_String, "'", "")
      tmpBol = Split(strExc(10), ",")
      strExc(9) = tmpBol(0)
   End If
   'If ManaGrp = "" Then
      ManaGrp = "P,PS,CFP,CPS,"
      CheckOC
      With adoRecordset
         .CursorLocation = adUseClient
         .Open "select distinct sg02 from staff_group ,staff where st01='" & strExc(9) & "' and st11=sg01(+) ", cnnConnection, adOpenStatic, adLockReadOnly
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
      If ManaGrp <> "" And Right(ManaGrp, 1) = "," Then
         ManaGrp = Left(ManaGrp, Len(ManaGrp) - 1)
      End If
      frm090614.ManaGrp = ManaGrp
   'End If
   '2024/3/5 END
   StrGrp090201 = frm090614.ManaGrp
      
      '改成收文日要小於等於發文年月當月的最後一天
      StrSQL6 = " and cp05<=" & Trim((Val(frm090614.txt1(3)) + 1911)) & Trim(Right(ChgNumByNick(frm090614.txt1(4)), 2)) & "31 "
'      If Len(Trim(frm090614.txt1(6))) <> 0 Then
'         StrSQL6 = StrSQL6 & " AND S1.ST03>='" & frm090614.txt1(6) & "' "
'      End If
'      If Len(Trim(frm090614.txt1(7))) <> 0 Then
'         StrSQL6 = StrSQL6 & " AND S1.ST03<='" & frm090614.txt1(7) & "' "
'      End If
      StrSQL61 = ""
      StrSQL62 = ""
      StrSQL63 = ""
      StrSQL64 = ""
      'C類未發文的不管閉不閉卷都要能在系統中作業
      'Modify By Sindy 2024/2/20 908=代辦退費未發文的不管閉不閉卷都要能在系統中作業
      StrSPa = " and ((pa58>=" & Trim((Val(frm090614.txt1(3)) + 1911)) & Trim(Right(ChgNumByNick(frm090614.txt1(4)), 2)) & "01 and pa58<=" & Trim((Val(frm090614.txt1(3)) + 1911)) & Trim(Right(ChgNumByNick(frm090614.txt1(4)), 2)) & "31) or pa58 is null or cp27>0 or (cp01 in('P','FCP') and cp09>'C') or (cp01 in('P','FCP') and cp10 in('908'))) "
      StrSSP = " and ((sp16>=" & Trim((Val(frm090614.txt1(3)) + 1911)) & Trim(Right(ChgNumByNick(frm090614.txt1(4)), 2)) & "01 and sp16<=" & Trim((Val(frm090614.txt1(3)) + 1911)) & Trim(Right(ChgNumByNick(frm090614.txt1(4)), 2)) & "31) or sp16 is null or cp27>0 or (cp01='FG' and cp10 in('908'))) "
      If frm090614.txt1(8) = "N" Then
         '不限制發文日止日及取消收文日止日
         'Modify By Sindy 2024/4/29 + and ep09 is not null
         'Modify By Sindy 2025/3/18 +外專繪圖同新案翻譯歷程操作方式
         StrSQL6 = StrSQL6 & " and (CP14 IN (" & Combo1_String & ") or (cp10 in('201','931') and EP04 IN (" & Combo1_String & ") and ep09 is not null)) and cp05>=19980101 "
         StrSQL61 = StrSQL61 & " and cp158=0 and cp159=0 "
         StrSQL62 = StrSQL62 & " and CP27>=" & Trim((Val(frm090614.txt1(3)) + 1911)) & Trim(Right(ChgNumByNick(frm090614.txt1(4)), 2)) & "01 "
         StrSQL63 = StrSQL63 & " and CP57>=" & Trim((Val(frm090614.txt1(3)) + 1911)) & Trim(Right(ChgNumByNick(frm090614.txt1(4)), 2)) & "01 and cp27 is null "
         StrSQL64 = StrSQL64 & " and CP05>=" & Trim((Val(frm090614.txt1(3)) + 1911)) & Trim(Right(ChgNumByNick(frm090614.txt1(4)), 2)) & "01 and CP05<=" & Trim((Val(frm090614.txt1(3)) + 1911)) & Trim(Right(ChgNumByNick(frm090614.txt1(4)), 2)) & "31 and cp57 is null And CP05>CP27 "
      Else
         '不限制發文日止日及取消收文日止日
         'Modify By Sindy 2024/4/29 + and ep09 is not null
         'Modify By Sindy 2025/3/18 +外專繪圖同新案翻譯歷程操作方式
         StrSQL6 = StrSQL6 & " and (CP14='" & Trim(Left("" & Combo1.Text, 6)) & "' or (cp10 in('201','931') and EP04='" & Trim(Left("" & Combo1.Text, 6)) & "' and ep09 is not null)) and cp05>=19980101 "
         StrSQL61 = StrSQL61 & " and cp158=0 and cp159=0 "
         StrSQL62 = StrSQL62 & " and CP27>=" & Trim((Val(frm090614.txt1(3)) + 1911)) & Trim(Right(ChgNumByNick(frm090614.txt1(4)), 2)) & "01 "
         StrSQL63 = StrSQL63 & " and CP57>=" & Trim((Val(frm090614.txt1(3)) + 1911)) & Trim(Right(ChgNumByNick(frm090614.txt1(4)), 2)) & "01 and cp27 is null "
         StrSQL64 = StrSQL64 & " and CP05>=" & Trim((Val(frm090614.txt1(3)) + 1911)) & Trim(Right(ChgNumByNick(frm090614.txt1(4)), 2)) & "01 and CP05<=" & Trim((Val(frm090614.txt1(3)) + 1911)) & Trim(Right(ChgNumByNick(frm090614.txt1(4)), 2)) & "31 and cp57 is null And CP05>CP27 "
      End If
Case Else
End Select

CheckOC

'StrSPa= and ((pa58>=20231001 AND pa58<=20231031) or pa58 is null or cp27>0) :PA58=閉卷日期
'StrSSP= and ((sp16>=20231001 AND sp16<=20231031) or sp16 is null or cp27>0) :SP16=閉卷日期
'StrSQL6= and (CP14='B2027' or (cp10='201' and EP04='B2027')) and cp05>=19980101
'StrSQL61= and cp158=0 and cp159=0
'StrSQL62= and CP27>=20231001 AND CP27<=20231031 :這個月發文
'StrSQL63= and CP57>=20231001 AND CP57<=20231031 and cp27 is null :這個月取消收文
'StrSQL64= and CP05>=20231001 AND CP05<=20231031 and cp57 is null And CP05>CP27 :這個月收文,但收文日大於發文日未取消收文
'Modify By Sindy 2023/12/26 +,SQLDateT2(CP142)||decode(CP164,'1','當天','2','之前','3','之後',cp164) CP142 =>R110033
'Modify By Sindy 2024/1/17 EP12=>CP64
'Modify By Sindy 2024/2/23 +,GetEEPCurState(cp09) as EEP04New =>R110035
'Modify By Sindy 2024/3/18 +排除D類收文號
'第一次 PGMID.CheckCuInAD(PA01,PA09,PA26,PA27,PA28,PA29,PA30)||
strSql = "SELECT CP14,EP01,SUBSTR(CP09,1,1),SQLDateT2(CP05),CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(PA57,'Y','＊',''),NVL(PA05,NVL(PA06,PA07)),CP26,DECODE(PA09,'020',PTM04,PTM03),nvl(DECODE(PA09,'000',CPM03,CPM04),cp10),SQLDateT2(CP48),SQLDateT2(CP06),SQLDateT2(CP07),SQLDateT2(EP06),SQLDateT2( EP09 ),SQLDateT2(EP07),nvl(S3.ST02,ep04),SQLDateT2( EP08 ),SQLDateT2(CP27),Nvl(EP35,0),CP64,nvl(S2.ST02,cp13),CP09,S1.ST05,CP01,PA09,CP10,NVL(NA03,NA04),SQLDateT2(NVL(CP57,PA58)), S1.ST02,cp97,cp98,cp111,ep34,cp112,SQLDateT2(EP28) ep28,SQLDateT2(CP142)||decode(CP164,'1','當天','2','之前','3','之後',cp164) CP142,GetEEPCurState(cp09) as EEP04New,pa150" & _
         " FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,PATENT,CASEPROPERTYMAP,PATENTTRADEMARKMAP,NATION " & _
         " WHERE CP09=EP02(+) AND CP01=pa01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND EP04=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND '1'=PTM01(+) AND PA08=PTM02(+) AND PA09=NA01(+) " & StrSQL6 & StrSQL61 & StrSPa & _
         " AND CP01 IN (" & SQLGrpStr(StrGrp090201, 1) & ") AND substr(CP09,1,1)<>'D'"
strSql = strSql + " UNION SELECT CP14,EP01,SUBSTR(CP09,1,1),SQLDateT2(CP05),CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(SP15,'Y','＊',''),NVL(SP05,NVL(SP06,SP07)),CP26,'',nvl(DECODE(SP09,'000',CPM03,CPM04),cp10),SQLDateT2(CP48),SQLDateT2(CP06),SQLDateT2(CP07),SQLDateT2(EP06),SQLDateT2( EP09 ),SQLDateT2(EP07),nvl(S3.ST02,ep04),SQLDateT2( EP08 ),SQLDateT2(CP27),Nvl(EP35,0),CP64,nvl(S2.ST02,cp13),CP09,S1.ST05,CP01,SP09,CP10,NVL(NA03,NA04),SQLDateT2(NVL(CP57,sp16)), S1.ST02,cp97,cp98,cp111,ep34,cp112,SQLDateT2(EP28) ep28,SQLDateT2(CP142)||decode(CP164,'1','當天','2','之前','3','之後',cp164) CP142,GetEEPCurState(cp09) as EEP04New,sp79 pa150" & _
         " FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,SERVICEPRACTICE,CASEPROPERTYMAP,NATION " & _
         " WHERE CP09=EP02(+) AND cp01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND EP04=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND SP09=NA01(+) " & StrSQL6 & StrSQL61 & StrSSP & _
         " AND CP01 IN (" & SQLGrpStr(StrGrp090201, 5) & ") AND substr(CP09,1,1)<>'D'"
'第二次
strSql = strSql + " UNION SELECT CP14,EP01,SUBSTR(CP09,1,1),SQLDateT2(CP05),CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(PA57,'Y','＊',''),NVL(PA05,NVL(PA06,PA07)),CP26,DECODE(PA09,'020',PTM04,PTM03),nvl(DECODE(PA09,'000',CPM03,CPM04),cp10),SQLDateT2(CP48),SQLDateT2(CP06),SQLDateT2(CP07),SQLDateT2(EP06),SQLDateT2( EP09 ),SQLDateT2(EP07),nvl(S3.ST02,ep04),SQLDateT2( EP08 ),SQLDateT2(CP27),Nvl(EP35,0),CP64,nvl(S2.ST02,cp13),CP09,S1.ST05,CP01,PA09,CP10,NVL(NA03,NA04),SQLDateT2(NVL(CP57,PA58)), S1.ST02,cp97,cp98,cp111,ep34,cp112,SQLDateT2(EP28) ep28,SQLDateT2(CP142)||decode(CP164,'1','當天','2','之前','3','之後',cp164) CP142,GetEEPCurState(cp09) as EEP04New,pa150" & _
         " FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,PATENT,CASEPROPERTYMAP,PATENTTRADEMARKMAP,NATION " & _
         " WHERE CP09=EP02(+) AND CP01=pa01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND EP04=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND '1'=PTM01(+) AND PA08=PTM02(+) AND PA09=NA01(+) " & Replace(UCase(StrSQL6), "CP05", "CP05+0") & StrSQL62 & StrSPa & _
         " AND CP01 IN (" & SQLGrpStr(StrGrp090201, 1) & ") AND substr(CP09,1,1)<>'D'"
strSql = strSql + " UNION SELECT CP14,EP01,SUBSTR(CP09,1,1),SQLDateT2(CP05),CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(SP15,'Y','＊',''),NVL(SP05,NVL(SP06,SP07)),CP26,'',nvl(DECODE(SP09,'000',CPM03,CPM04),cp10),SQLDateT2(CP48),SQLDateT2(CP06),SQLDateT2(CP07),SQLDateT2(EP06),SQLDateT2( EP09 ),SQLDateT2(EP07),nvl(S3.ST02,ep04),SQLDateT2( EP08 ),SQLDateT2(CP27),Nvl(EP35,0),CP64,nvl(S2.ST02,cp13),CP09,S1.ST05,CP01,SP09,CP10,NVL(NA03,NA04),SQLDateT2(NVL(CP57,sp16)), S1.ST02,cp97,cp98,cp111,ep34,cp112,SQLDateT2(EP28) ep28,SQLDateT2(CP142)||decode(CP164,'1','當天','2','之前','3','之後',cp164) CP142,GetEEPCurState(cp09) as EEP04New,sp79 pa150" & _
         " FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,SERVICEPRACTICE,CASEPROPERTYMAP,NATION " & _
         " WHERE CP09=EP02(+) AND cp01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND EP04=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND SP09=NA01(+) " & Replace(UCase(StrSQL6), "CP05", "CP05+0") & StrSQL62 & StrSSP & _
         " AND CP01 IN (" & SQLGrpStr(StrGrp090201, 5) & ") AND substr(CP09,1,1)<>'D'"
'第三次
strSql = strSql + " UNION SELECT CP14,EP01,SUBSTR(CP09,1,1),SQLDateT2(CP05),CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(PA57,'Y','＊',''),NVL(PA05,NVL(PA06,PA07)),CP26,DECODE(PA09,'020',PTM04,PTM03),nvl(DECODE(PA09,'000',CPM03,CPM04),cp10),SQLDateT2(CP48),SQLDateT2(CP06),SQLDateT2(CP07),SQLDateT2(EP06),SQLDateT2( EP09 ),SQLDateT2(EP07),nvl(S3.ST02,ep04),SQLDateT2( EP08 ),SQLDateT2(CP27),Nvl(EP35,0),CP64,nvl(S2.ST02,cp13),CP09,S1.ST05,CP01,PA09,CP10,NVL(NA03,NA04),SQLDateT2(NVL(CP57,PA58)), S1.ST02,cp97,cp98,cp111,ep34,cp112,SQLDateT2(EP28) ep28,SQLDateT2(CP142)||decode(CP164,'1','當天','2','之前','3','之後',cp164) CP142,GetEEPCurState(cp09) as EEP04New,pa150" & _
         " FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,PATENT,CASEPROPERTYMAP,PATENTTRADEMARKMAP,NATION " & _
         " WHERE CP09=EP02(+) AND CP01=pa01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND EP04=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND '1'=PTM01(+) AND PA08=PTM02(+) AND PA09=NA01(+) " & Replace(UCase(StrSQL6), "CP05", "CP05+0") & StrSQL63 & StrSPa & _
         " AND CP01 IN (" & SQLGrpStr(StrGrp090201, 1) & ") AND substr(CP09,1,1)<>'D'"
strSql = strSql + " UNION SELECT CP14,EP01,SUBSTR(CP09,1,1),SQLDateT2(CP05),CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(SP15,'Y','＊',''),NVL(SP05,NVL(SP06,SP07)),CP26,'',nvl(DECODE(SP09,'000',CPM03,CPM04),cp10),SQLDateT2(CP48),SQLDateT2(CP06),SQLDateT2(CP07),SQLDateT2(EP06),SQLDateT2( EP09 ),SQLDateT2(EP07),nvl(S3.ST02,ep04),SQLDateT2( EP08 ),SQLDateT2(CP27),Nvl(EP35,0),CP64,nvl(S2.ST02,cp13),CP09,S1.ST05,CP01,SP09,CP10,NVL(NA03,NA04),SQLDateT2(NVL(CP57,sp16)), S1.ST02,cp97,cp98,cp111,ep34,cp112,SQLDateT2(EP28) ep28,SQLDateT2(CP142)||decode(CP164,'1','當天','2','之前','3','之後',cp164) CP142,GetEEPCurState(cp09) as EEP04New,sp79 pa150" & _
         " FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,SERVICEPRACTICE,CASEPROPERTYMAP,NATION " & _
         " WHERE CP09=EP02(+) AND cp01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND EP04=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND SP09=NA01(+) " & Replace(UCase(StrSQL6), "CP05", "CP05+0") & StrSQL63 & StrSSP & _
         " AND CP01 IN (" & SQLGrpStr(StrGrp090201, 5) & ") AND substr(CP09,1,1)<>'D'"
'第四次
strSql = strSql + " UNION SELECT CP14,EP01,SUBSTR(CP09,1,1),SQLDateT2(CP05),CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(PA57,'Y','＊',''),NVL(PA05,NVL(PA06,PA07)),CP26,DECODE(PA09,'020',PTM04,PTM03),nvl(DECODE(PA09,'000',CPM03,CPM04),cp10),SQLDateT2(CP48),SQLDateT2(CP06),SQLDateT2(CP07),SQLDateT2(EP06),SQLDateT2( EP09 ),SQLDateT2(EP07),nvl(S3.ST02,ep04),SQLDateT2( EP08 ),SQLDateT2(CP27),Nvl(EP35,0),CP64,nvl(S2.ST02,cp13),CP09,S1.ST05,CP01,PA09,CP10,NVL(NA03,NA04),SQLDateT2(NVL(CP57,PA58)), S1.ST02,cp97,cp98,cp111,ep34,cp112,SQLDateT2(EP28) ep28,SQLDateT2(CP142)||decode(CP164,'1','當天','2','之前','3','之後',cp164) CP142,GetEEPCurState(cp09) as EEP04New,pa150" & _
         " FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,PATENT,CASEPROPERTYMAP,PATENTTRADEMARKMAP,NATION " & _
         " WHERE CP09=EP02(+) AND CP01=pa01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND EP04=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND '1'=PTM01(+) AND PA08=PTM02(+) AND PA09=NA01(+) " & StrSQL6 & StrSQL64 & StrSPa & _
         " AND CP01 IN (" & SQLGrpStr(StrGrp090201, 1) & ") AND substr(CP09,1,1)<>'D'"
strSql = strSql + " UNION SELECT CP14,EP01,SUBSTR(CP09,1,1),SQLDateT2(CP05),CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(SP15,'Y','＊',''),NVL(SP05,NVL(SP06,SP07)),CP26,'',nvl(DECODE(SP09,'000',CPM03,CPM04),cp10),SQLDateT2(CP48),SQLDateT2(CP06),SQLDateT2(CP07),SQLDateT2(EP06),SQLDateT2( EP09 ),SQLDateT2(EP07),nvl(S3.ST02,ep04),SQLDateT2( EP08 ),SQLDateT2(CP27),Nvl(EP35,0),CP64,nvl(S2.ST02,cp13),CP09,S1.ST05,CP01,SP09,CP10,NVL(NA03,NA04),SQLDateT2(NVL(CP57,sp16)), S1.ST02,cp97,cp98,cp111,ep34,cp112,SQLDateT2(EP28) ep28,SQLDateT2(CP142)||decode(CP164,'1','當天','2','之前','3','之後',cp164) CP142,GetEEPCurState(cp09) as EEP04New,sp79 pa150" & _
         " FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,SERVICEPRACTICE,CASEPROPERTYMAP,NATION " & _
         " WHERE CP09=EP02(+) AND cp01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND EP04=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND SP09=NA01(+) " & StrSQL6 & StrSQL64 & StrSSP & _
         " AND CP01 IN (" & SQLGrpStr(StrGrp090201, 5) & ") AND substr(CP09,1,1)<>'D'"
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
        'DoEvents
        strCP09s = "''"
        Do While .EOF = False
'         If .Fields("cp09") = "AB2044260" Then
'            intI = intI
'         End If
            strCP09s = strCP09s & ",'" & .Fields("cp09") & "'" 'Add by Morgan 2011/8/3
'            If .Fields("cp09") = "CB3000020" Then
'               MsgBox .Fields("cp09")
'            End If
            For i = 0 To 27 '26  'Modify By Sindy 2023/12/26 26 => 27 edit by nickc 2007/11/27 21
                strTemp(i) = CheckStr(.Fields(i))
                'Modify by Morgan 2011/1/3 修正日期欄位排序問題(小於100年的前面補空白)
                If Len(strTemp(i)) = 8 Then
                  If Mid(strTemp(i), 3, 1) = "/" And Mid(strTemp(i), 6, 1) = "/" Then
                     strTemp(i) = " " & strTemp(i)
                  End If
                End If
            Next i
            
            'Add By Sindy 2024/3/6 承辦人非機械組,但案件是機械組
            strExc(10) = ""
            'Modify By Sindy 2025/3/18 +外專繪圖同新案翻譯歷程操作方式
            If (.Fields("cp10") = "201" Or .Fields("cp10") = "931") Then
               If Mid(Trim(Left("" & Combo1.Text, 6)), 4, 1) <> "9" Then
                  strExc(9) = GetPrjSalesNM_2(strTemp(15))
                  If strExc(9) <> "" Then
                     Call GetPrjSalesNM(strExc(9), , , "st16", strExc(10))
                  End If
               End If
            Else
               Call GetPrjSalesNM(strTemp(0), , , "st16", strExc(10))
            End If
            If strExc(10) <> "" And strExc(10) <> "4" And "" & .Fields("pa150") = "4" Then
               strTemp(4) = "▲" & strTemp(4)
            End If
            '2024/3/6 END
            
            'edit by nickc 2007/11/27
            'strSQL = "INSERT INTO R090614 VALUES ('" & strTemp(0) & "'," & Val(strTemp(1)) & ",'" & strTemp(2) & "','" & strTemp(3) & "','" & strTemp(4) & "','" & ChgSQL(strTemp(5)) & "','" & strTemp(6) & "','" & strTemp(7) & "','" & strTemp(8) & "','" & strTemp(9) & "','" & strTemp(10) & "','" & strTemp(11) & "','" & strTemp(12) & "','" & strTemp(13) & "','" & strTemp(14) & "','" & strTemp(15) & "','" & strTemp(16) & "','" & strTemp(17) & "'," & Val(strTemp(18)) & ",'" & strTemp(19) & "','" & strTemp(20) & "','" & strTemp(21) & "','" & strUserNum & "','" & "" & .Fields(26).Value & "','" & .Fields(27).Value & "','" & .Fields(28).Value & "' ) "
            'Modify by Morgan 2009/7/14 +EP28
            'Modify by Sindy 2023/12/26 +CP142
            'Modify By Sindy 2024/2/20 +,R110034(CP10)
            'Modify By Sindy 2024/2/23 +,R110035(EEP04New)
            strSql = "INSERT INTO R090614 VALUES (" & _
               "'" & strTemp(0) & "'," & Val(strTemp(1)) & ",'" & strTemp(2) & "','" & strTemp(3) & "','" & strTemp(4) & "'," & _
               "'" & ChgSQL(strTemp(5)) & "','" & strTemp(6) & "','" & strTemp(7) & "','" & strTemp(8) & "','" & strTemp(9) & "'," & _
               "'" & strTemp(10) & "','" & strTemp(11) & "','" & strTemp(12) & "','" & strTemp(13) & "','" & strTemp(14) & "'," & _
               "'" & strTemp(15) & "','" & strTemp(16) & "','" & strTemp(17) & "'," & Val(strTemp(18)) & ",'" & ChgSQL(strTemp(19)) & "'," & _
               "'" & strTemp(20) & "','" & strTemp(21) & "','" & strUserNum & "','" & "" & .Fields(26).Value & "','" & .Fields(27).Value & "'," & _
               "'" & .Fields(28).Value & "'," & Val("" & .Fields("cp97")) & "," & Val("" & .Fields("cp98")) & "," & Val("" & .Fields("cp111")) & ",'" & "" & .Fields("ep34").Value & "'," & _
               "'" & "" & .Fields("cp112").Value & "','" & .Fields("ep28").Value & "',0,'" & .Fields("CP142").Value & "','" & .Fields("CP10").Value & "','" & .Fields("EEP04New").Value & "')"
            adoEng.Execute strSql
            .MoveNext
            'DoEvents
        Loop
'        'Add by Morgan 2011/8/3
'        '更新支援+修改+衍生基數
'        'Modified by Morgan 2014/3/20 2014/4/1 起支援改每小時折計0.2基數
'        'strExc(0) = "select SH12,sum(pp) pp from (" & _
'         " select SH12,Round(Decode(SH06, 'CFP', Nvl(SH05,0)/3, Nvl(SH05,0)/4) ,2) pp from supporthour where SH12 in (" & strCP09s & ") And SH11='V' and SH05>0" & _
'         " Union All Select MH12,Round(Nvl(MH05,0)*0.2 ,2) pp From ModifyHour Where MH12 in (" & strCP09s & ") And MH11='V' and MH05>0" & _
'         " Union All Select EH12,Round(Nvl(EH05,0)*0.2 ,2) pp From ExtendHour Where EH12 in (" & strCP09s & ") And EH11='V'and EH05>0) X group by SH12"
'        'Modified by Morgan 2019/4/9 108考核支援時數轉換要除組別參數
'        strExc(0) = "select SH12,sum(pp) pp from (" & _
'         " select SH12,Round(" & Sh2EPtCode & " / GetDivNum(st70,sh01) ,2) pp from supporthour,staff where st01(+)=sh02 and SH12 in (" & strCP09s & ") And SH11='V' and SH05>0" & _
'         " Union All Select MH12,Round(Nvl(MH05,0)*0.2 ,2) pp From ModifyHour Where MH12 in (" & strCP09s & ") And MH11='V' and MH05>0" & _
'         " Union All Select EH12,Round(Nvl(EH05,0)*0.2 ,2) pp From ExtendHour Where EH12 in (" & strCP09s & ") And EH11='V'and EH05>0) X group by SH12"
'         'end 2014/3/20
'        intI = 1
'        Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'        If intI = 1 Then
'           With RsTemp
'           Do While Not .EOF
'               strSql = "update R090614 set R110032=" & .Fields("pp") & " where R110022='" & .Fields("SH12") & "'"
'               adoEng.Execute strSql, intI
'               .MoveNext
'           Loop
'           End With
'        End If
'        'end 2011/8/3
    End If
End With
CheckOC
End Sub

Sub ChgGrdColor(Optional iRow As Integer = -1)
Dim ColorFlag As String
Dim iStart As Integer, iEnd As Integer
Dim i As Integer 'Add By Sindy 2024/3/13

With grd1
   If iRow >= 0 Then
      iStart = iRow
      iEnd = iRow
   Else
      iStart = 1
      iEnd = .Rows - 1
   End If
   
   For i = iStart To iEnd
      'DoEvents 'Modify By Sindy 2024/3/13 mark
      'Add By Sindy 2024/3/13
      If i > grd1.Rows - 1 Or i < 0 Then
         MsgBox "不正確的資料列值。( .Row=" & i & " )"
         Exit Sub
      ElseIf i = 1 And iEnd = 1 Then
         .row = i
         If .Text = "" Then
            Exit Sub
         End If
      End If
      '2024/3/13 END
      
      .row = i 'i值若有誤, 會出現的訊息為 "不正確的資料列值"
      
      'Add By Sindy 2024/3/6
      .col = 30 '本所案號為sort使用
      .Text = Replace(Replace(.Text, "▲", ""), "＊", "")
      '2024/3/6 END
      
      .col = 24 '取消收文日
      Tmp003 = Trim(.Text)
      '若有取消收文日期
      If Tmp003 <> "" Then
         '灰色
         .col = 3 '本所案號
         .CellBackColor = QBColor(8)
         .col = 10 '指定送件日
         .CellBackColor = QBColor(8)
         .col = 11 '承辦期限
         .CellBackColor = QBColor(8)
         .col = 13 '齊備日
         .CellBackColor = QBColor(8)
      Else
         .col = 19 '發文日
         '若無發文日
         If .Text = "" Then
            .col = 9 '本所期限
            '若系統日大於等於本所期限且本所期限有值(逾本所期限未發文)
            If Val(ChangeTStringToWString(ChangeTDateStringToTString(Trim(.Text)))) <= Val(strSrvDate(1)) And Trim(.Text) <> "" Then
               '淺紅色
               .col = 3 '本所案號
               .CellBackColor = &HC0C0FF '在淺一點 &H8080FF
               .col = 10 '指定送件日
               .CellBackColor = &HC0C0FF
               .col = 11 '承辦期限
               .CellBackColor = &HC0C0FF
               .col = 13 '齊備日
               .CellBackColor = &HC0C0FF
            End If
         End If
      End If
   Next i
'   '預設目前在第一筆的位置
'   With Me.grd1
'      .row = 1
'      .col = 0
'      .CellBackColor = &HFFC0C0
'      .col = 12 '法定期限
'      .CellBackColor = &HFFC0C0
'      SWPRow = .row
'   End With
   'SetGrd1
End With
End Sub

Sub StrMenu()
Dim iMouse As Integer

iMouse = Screen.MousePointer
'Modify by Morgan 2011/8/3 +R110032(支援+修改+衍生基數)
Select Case ProState
Case "1"
      'Modify by Morgan 2009/7/14 加欄位 R110031
      'Modify by Morgan 2010/11/4 +r110025,r110030
      'Modify By Sindy 2023/12/26 +,R110033 取消=,format(R110026 * r110027 * iif(r110030='N',1,iif(isnull(r110028)  or r110028=0,1,r110028))+R110032,'0.00')
      'Modify By Sindy 2024/2/23 +,R110035
      'Modify By Sindy 2025/3/18 +外專繪圖同新案翻譯歷程操作方式
      strSql = "SELECT R110002,R110003,R110004,R110005,R110006,R110023,R110008,R110009,R110007,R110011,R110033,R110010,R110012,R110013,R110014,R110031,R110015,R110016,R110017,R110018,R110019,R110020,R110021,R110022, R110024,r110029,r110025,r110030,R110032,R110035,R110005" & _
               " FROM R090614" & _
               " WHERE ID='" & strUserNum & "'" & _
               " AND (R110001='" & Trim(Left("" & Combo1.Text, 6)) & "' or (instr('" & Trim("" & Combo1.Text) & "',R110016)>0 and R110034 in('201','931')))"
'      'Modify By Sindy 2013/6/7 依會稿完成日大至小+目次大至小排序
'      If Pub_StrUserSt03 = "M51" Or InStr(UCase(Pub_GetModuleFileName), "VB6.EXE") <> 0 Then
'         strSql = strSql & " ORDER BY R110017 desc,R110002 desc,R110003,R110004"
'      Else
         strSql = strSql & " ORDER BY R110002 desc,R110003,R110004"
'      End If
'      '2013/6/7 End
Case "2"
      If frm090614.txt1(8) = "N" Then
         'Modify by Morgan 2009/7/14 加欄位 R110031
         'Modify by Morgan 2010/11/4 +r110025,r110030
         'Modify By Sindy 2023/12/26 +,R110033 取消=,format(R110026 * r110027 * iif(r110030='N',1,iif(isnull(r110028)  or r110028=0,1,r110028))+R110032,'0.00')
         'Modify By Sindy 2024/2/23 +,R110035
         'Modify By Sindy 2025/3/18 +外專繪圖同新案翻譯歷程操作方式
         strSql = "SELECT R110002,R110003,R110004,R110005,R110006,R110023,R110008,R110009,R110007,R110011,R110033,R110010,R110012,R110013,R110014,R110031,R110015,R110016,R110017,R110018,R110019,R110020,R110021,R110022, R110024,r110029,r110025,r110030,R110032,R110035,R110005" & _
                  " FROM R090614" & _
                  " WHERE ID='" & strUserNum & "'" & _
                  " AND (R110001 IN (" & Combo1_String & ") or (instr('" & Combo1_Name & "',R110016)>0 and R110034 in('201','931')))" & _
                  " ORDER BY R110005,R110002 desc,  R110004 "
      Else
         'Modify by Morgan 2009/7/14 加欄位 R110031
         'Modify by Morgan 2010/11/4 +r110025,r110030
         'Modify By Sindy 2023/12/26 +,R110033 取消=,format(R110026 * r110027 * iif(r110030='N',1,iif(isnull(r110028) or r110028=0,1,r110028))+R110032,'0.00')
         'Modify By Sindy 2024/2/23 +,R110035
         'Modify By Sindy 2025/3/18 +外專繪圖同新案翻譯歷程操作方式
         strSql = "SELECT R110002,R110003,R110004,R110005,R110006,R110023,R110008,R110009,R110007,R110011,R110033,R110010,R110012,R110013,R110014,R110031,R110015,R110016,R110017,R110018,R110019,R110020,R110021,R110022, R110024,r110029,r110025,r110030,R110032,R110035,R110005" & _
                  " FROM R090614" & _
                  " WHERE ID='" & strUserNum & "'" & _
                  " AND (R110001='" & Trim(Left("" & Combo1.Text, 6)) & "' or (instr('" & Trim("" & Combo1.Text) & "',R110016)>0 and R110034 in('201','931')))" & _
                  " ORDER BY R110002 desc, R110003, R110004 "
      End If
Case Else
End Select
CheckOC
LblCnt.Caption = "" 'Add By Sindy 2023/12/7
With adoRecordset
    .CursorLocation = adUseClient
    .Open strSql, adoEng, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        If InStr(UCase(Pub_GetModuleFileName), "VB6.EXE") <> 0 Then LblCnt.Caption = .RecordCount  'Add By Sindy 2023/12/7
        If ProState = "2" Then
            InsertQueryLog (.RecordCount) 'Add By Sindy 2010/12/17
        End If
        Set grd1.Recordset = adoRecordset
        ChkNoData = False
    Else
        If ProState = "2" Then
            InsertQueryLog (0) 'Add By Sindy 2010/12/17
        End If
        ChkNoData = True
        grd1.Clear
        grd1.Rows = 2
        'Modify by Morgan 2009/11/12
        'Screen.MousePointer = vbDefault
        Screen.MousePointer = iMouse
        'Exit Sub
    End If
End With
CheckOC
ChgGrdColor
SWPRow = "1"
grd1.row = Val(SWPRow)
If ChkNoData = False Then grd1.col = 1
Call cmdok2_Click(1) '未發文 Add By Sindy 2023/11/21
End Sub

Private Sub SetGrd1()
With grd1
    .Visible = False
    'Modify by Morgan 2009/7/13 加欄位:預會日 15
    '.Cols = 27 'edit by nickc 2007/11/27 加欄位 25
    .Cols = 31 '30
    
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
    .ColWidth(5) = 0 '450
    .CellAlignment = flexAlignCenterCenter
    .col = 6:   .Text = "種類"
    .ColWidth(6) = 450
    .CellAlignment = flexAlignCenterCenter
    .col = 7:   .Text = "案件性質"
    .ColWidth(7) = 795
    .CellAlignment = flexAlignCenterCenter
    'Add By Cheng 2002/04/16
    .col = 8:   .Text = "Y/N"
    .ColWidth(8) = 0
    .CellAlignment = flexAlignCenterCenter
    .col = 9:   .Text = "本所期限"
    .ColWidth(9) = 795
    .ColAlignment(9) = flexAlignRightCenter
    .CellAlignment = flexAlignCenterCenter
    .col = 10:  .Text = "指定送件日"
    .ColWidth(10) = 1000
    .ColAlignment(10) = flexAlignRightCenter
    .CellAlignment = flexAlignCenterCenter
    .col = 11:   .Text = "承辦期限"
    .ColWidth(11) = 795
    .ColAlignment(11) = flexAlignRightCenter
    .CellAlignment = flexAlignCenterCenter
'edit by nickc 2007/11/27 加欄位
    '.col = 12:   .Text = "確認"
    '.ColWidth(12) = 300
    '.CellAlignment = flexAlignCenterCenter
    'edit by nickc 2007/11/27 加欄位  以下都往後退
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
    'Modify by Morgan 2009/7/13 加欄位:預會日 15
    .col = 15:  .Text = "預會日"
    .ColWidth(15) = 0
    .ColAlignment(15) = flexAlignRightCenter
    .CellAlignment = flexAlignCenterCenter
    .col = 16:  .Text = "會稿日"
    .ColWidth(16) = 0
    .ColAlignment(16) = flexAlignRightCenter
    .CellAlignment = flexAlignCenterCenter
    .col = 17:  .Text = "核稿人"
    .ColWidth(17) = 700
    .CellAlignment = flexAlignCenterCenter
    .col = 18:  .Text = "核稿期限" '會稿完成日
    .ColWidth(18) = 795
    .ColAlignment(18) = flexAlignRightCenter
    .CellAlignment = flexAlignCenterCenter
    .col = 19:  .Text = "發文日"
    .ColWidth(19) = 795
    .ColAlignment(19) = flexAlignRightCenter
    .CellAlignment = flexAlignCenterCenter
    .col = 20:  .Text = "承辦天數"
    .ColWidth(20) = 0 '800
    .CellAlignment = flexAlignCenterCenter
    .col = 21:  .Text = "進度備註" '"承辦備註" Modify By Sindy 2024/1/17
    .ColWidth(21) = 2000
    .CellAlignment = flexAlignCenterCenter
    .col = 22:  .Text = "智權人員" 'R110021
    .ColWidth(22) = 800
    .CellAlignment = flexAlignCenterCenter
    .col = 23:  .Text = "" '總收文號=R110022
    .ColWidth(23) = 0
    .CellAlignment = flexAlignCenterCenter
    .col = 24:  .Text = "" '取消收文日 or 閉卷日期=R110024
    .ColWidth(24) = 0
    .CellAlignment = flexAlignCenterCenter
    .col = 25:  .Text = "" 'EP34=R110029
    .ColWidth(25) = 0
    .CellAlignment = flexAlignCenterCenter
    .col = 26:  .Text = "" 'CP14=R110025
    .ColWidth(26) = 0
    .CellAlignment = flexAlignCenterCenter
    .col = 27:  .Text = "" 'CP112=R110030
    .ColWidth(27) = 0
    .CellAlignment = flexAlignCenterCenter
    'Add By Sindy 2024/2/23
    .col = 28:  .Text = "" '0=R110032
    .ColWidth(28) = 0
    .CellAlignment = flexAlignCenterCenter
    .col = 29:  .Text = "最近歷程" 'R110035
    .ColWidth(29) = 1000
    .CellAlignment = flexAlignCenterCenter
    '2024/2/23 END
    
    For intI = 30 To .Cols - 1
      .ColWidth(intI) = 0
    Next
    
    .Visible = True
    
End With
'   'Add By Cheng 2002/10/23
'   '預設目前在第一筆的位置
'   With Me.grd1
'      .row = 1
'      .col = 0
'      .CellBackColor = &HFFC0C0
'      .col = 12
'      .CellBackColor = &HFFC0C0
'      SWPRow = .row
'   End With
End Sub

Private Sub GRD1_DblClick()
    'Modify By Cheng 2004/03/08
    If Me.grd1.MouseRow > 0 Then
        'Add By Cheng 2003/04/28
        '若有資料
        If Me.grd1.Rows > 1 Then
            SWPRow = str(grd1.MouseRow)
            'Modify By Cheng 2003/05/05
            '若點選的那筆無資料, 則退出函式
            If Me.grd1.TextMatrix(SWPRow, 1) = "" Then Exit Sub
    '        MouseClick Val(SWPRow)
            SSTab1.Tab = 1
        End If
    End If
End Sub

Private Sub grd1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'Add By Cheng 2004/03/08
If Me.grd1.MouseRow <= 0 Then Exit Sub
'End
If Button = 1 Then
   Call MouseClick(grd1.MouseRow, False)
End If
End Sub

Sub ClearMouseClick()
Dim ii As Integer
   
   With grd1
      If SWPRow <> "" Then
         .row = SWPRow
         For ii = 0 To grd1.Cols - 1
            If ii <> 3 And Not (ii >= 10 And ii <= 13) Then
               .col = ii
               .CellBackColor = QBColor(15)
            End If
         Next ii
         SWPRow = ""
      End If
   End With
End Sub

Sub MouseClick(Strindex As Integer, Optional ByVal bolQuery As Boolean = True)
    Dim iMouse As Integer
    Dim ii As Integer
    
    iMouse = Screen.MousePointer
    Screen.MousePointer = vbHourglass
    
    With grd1
        'DoEvents 'Modify By Sindy 2024/3/13 mark
        .Visible = False
        Call ClearMouseClick
'        If SWPRow <> "" Then
'           .row = SWPRow
'           .col = 0
'           .CellBackColor = QBColor(15)
'           .col = 12
'           .CellBackColor = QBColor(15)
'        End If
        .col = 0
        If Strindex <> 0 Then
            .row = Strindex
        Else
            .row = .MouseRow
        End If
        If .row = 0 Then
            .row = 1
        End If
        
        If bolQuery = True Then
            'Modify by Morgan 2009/7/13 加欄位:預會日 15
            '.col = 22 'edit by nickc 2007/11/27 加欄位  以下都往後退21
            .col = 23 '總收文號 =>R110022
            Call Process(.Text)
        End If
        
'        .col = 0
'        .CellBackColor = &HFFC0C0
'        .col = 12
'        .CellBackColor = &HFFC0C0
        For ii = 0 To grd1.Cols - 1
         If ii <> 3 And Not (ii >= 10 And ii <= 13) Then
            .col = ii
            .CellBackColor = &HFFC0C0
         End If
        Next ii
        SWPRow = .row
        .Visible = True
    End With
    'Modify by Morgan 2009/11/12
    'Screen.MousePointer = vbDefault
    Screen.MousePointer = iMouse
End Sub

Private Sub grd1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim nCol As Long, nRow As Long
   getGrdColRow grd1, x, y, nCol, nRow
   If nCol < 0 Or nRow < 0 Then Exit Sub
   'grd1.col = nCol
   grd1.row = nRow
   If Me.grd1.row < 1 Then
      Call ClearMouseClick 'Add By Sindy 2024/2/26
      If nCol = 3 Then
         nCol = 30
         grd1.col = nCol
      End If
      Select Case Me.grd1.MouseCol
         Case 0
            If m_blnColOrderAsc = True Then
                Me.grd1.Sort = 3 '昇冪
                m_blnColOrderAsc = False
            Else
                Me.grd1.Sort = 4 '降冪
                m_blnColOrderAsc = True
            End If
         Case Else
            If m_blnColOrderAsc = True Then
                Me.grd1.Sort = 5 '昇冪
                m_blnColOrderAsc = False
            Else
                Me.grd1.Sort = 6 '降冪
                m_blnColOrderAsc = True
            End If
      End Select
   End If
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
   Dim ii As Integer
   'Add By Sindy 2013/6/7
   'Modify By Sindy 2017/8/21
'   If SSTab1.Tab = 2 And Me.cmd(1).Visible = True Then
'      Call QueryData(False)
'   End If
   If SSTab1.Tab = 2 Then
      Call QueryData(True)
   Else
      Call QueryData(False)
   End If
   '2017/8/21 END
   
   If PreviousTab = 2 Then
      '若有資料
      If (Me.grd2.Rows - 1) < dblPrevRow Then dblPrevRow = 0 'Add By Sindy 2024/7/10
      If Me.grd2.Rows > 1 And dblPrevRow > 0 Then
         If Me.grd2.TextMatrix(dblPrevRow, 1) <> "" Then
            For i = 1 To Me.grd1.Rows - 1
               If Me.grd2.TextMatrix(dblPrevRow, 1) = Me.grd1.TextMatrix(i, 0) Then
                  'SWPRow = i
                  MouseClick i
                  Exit For
               End If
            Next i
            'MouseClick Val(SWPRow)
         End If
      End If
   End If
   '2013/6/7 End
   'Add By Cheng 2003/04/28
   If PreviousTab = 0 Or PreviousTab = 1 Then
      '若有資料
      If Me.grd1.Rows > 1 Then
         'Modify By Cheng 2003/05/05
         '若點選的那筆無資料, 則退出函式
         If Me.grd1.TextMatrix(Val("0" & SWPRow), 1) = "" Then SSTab1.Tab = 0: Exit Sub
         'Add by Sindy 2013/6/7
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
                  'Modify By Sindy 2013/10/29
                  If grd2.CellBackColor <> &H8080FF Then
                  '2013/10/29 END
                     grd2.CellBackColor = QBColor(15)
                  End If
               Next ii
               dblPrevRow = 0
               Call SetColColor(CStr(dblPrevRow))
            End If
            For i = 1 To Me.grd2.Rows - 1
               If Me.grd2.TextMatrix(i, 1) = Me.grd1.TextMatrix(Val("0" & SWPRow), 0) Then
                  '目前資料列反白
                  dblPrevRow = i
                  grd2.col = 0
                  grd2.row = dblPrevRow
                  If grd2.TextMatrix(grd2.row, 1) <> "" Then
                     grd2.Text = "V"
                     For ii = 0 To grd2.Cols - 1
                        grd2.col = ii
                        'Modify By Sindy 2013/10/29
                        If grd2.CellBackColor <> &H8080FF Then
                           grd2.CellBackColor = &HFFC0C0
                        End If
                     Next ii
                  End If
                  Exit For
               End If
            Next i
         End If
         '2013/6/7 End
         If PreviousTab = 0 And Me.SSTab1.Tab = 1 Then
            MouseClick Val(SWPRow)
         End If
      End If
   End If
End Sub

Private Sub txt1_Change(Index As Integer)
'   Select Case Index
'      'Add By Sindy 2015/3/13
'      Case 23 '核稿語文
'         If txt1(Index) = "2" Then
'            Label1(25).Caption = "日文核稿人："
'            Label1(41).Caption = "日文核完日："
'         Else
'            Label1(25).Caption = "英文核稿人："
'            Label1(41).Caption = "英文核完日："
'         End If
'   End Select
End Sub

Private Sub txt1_GotFocus(Index As Integer)
txt1(Index).SelStart = 0
txt1(Index).SelLength = Len(txt1(Index))
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = UpperCase(KeyAscii)
'add by nickc 2007/11/28 協理說，有輸入完稿日才可以輸入不會稿
Select Case Index
   Case 0, 1, 12
      KeyAscii = Pub_NumAscii(KeyAscii, True)
   Case 9
      If KeyAscii <> 8 And KeyAscii <> Asc("N") Then
         KeyAscii = 0
         Beep
      End If
   Case 23
      If KeyAscii <> Asc("1") And KeyAscii <> Asc("2") Then
         KeyAscii = 0
         Beep
      End If
End Select
End Sub

Private Sub txt1_LostFocus(Index As Integer)
   'Add by Morgan 2010/9/6 若回第一頁籤時不檢查,否則若有錯誤時會無窮回圈
   If Me.SSTab1.Tab = 0 Then Exit Sub
   
Dim tmpInti As Integer

Select Case Index
Case 12 '承辦期限
      If Len(txt1(Index)) <> 0 Then
         If txt1(Index) <> txt1(Index).Tag And m_CP10 = "926" Then '926=核對已准專利
            MsgBox "承辦期限只能取消不可修改！"
            txt1(Index) = txt1(Index).Tag
            txt1(Index).SetFocus
            txt1(Index).SelLength = Len(txt1(Index))
            Exit Sub
         End If
      End If
      
Case 2 '齊備日
      If Len(txt1(Index)) <> 0 Then
         If Not ChkWorkDay(ChangeTStringToWString(txt1(Index))) Then
            ShowDateErr
            txt1(Index).SetFocus
            txt1(Index).SelLength = Len(txt1(Index))
            Exit Sub
         End If
      End If
     
Case 3 '完稿日
      If Len(txt1(Index)) <> 0 Then
         If Not ChkWorkDay(ChangeTStringToWString(txt1(Index))) Then
            ShowDateErr
            txt1(Index).SetFocus
            txt1(Index).SelLength = Len(txt1(Index))
            Exit Sub
         End If
      End If
         
Case 7 '核稿期限
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
         'Add By Cheng 2003/01/27
         '若發文日為111111則不檢查是否為工作日
         If Me.txt1(Index).Text <> "111111" Then
             If Not ChkWorkDay(ChangeTStringToWString(txt1(Index))) Then
                ShowDateErr
                txt1(Index).SetFocus
                txt1(Index).SelLength = Len(txt1(Index))
                Exit Sub
             End If
         End If
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
       
Case Else
End Select
End Sub

'Sub ChkTxt()
Sub ChkTxt(Strindex As String)
    ChkData = False
    
    '完稿日
    If Strindex = "3" Or Strindex = "" Then
        If Len(txt1(3)) = 0 Then
            If Len(txt1(7)) <> 0 Or Len(txt1(8)) <> 0 Then
                ShowDateRanErr
                txt1(3).SetFocus
                Exit Sub
            End If
        End If
        If Not ChkDateRanPro(txt1(2), txt1(3), 2) And Len(txt1(3)) <> 0 Then
            txt1(3).SetFocus
            txt1_GotFocus (3)
            Exit Sub
        End If
    End If
    
    '核稿期限
    If Strindex = "7" Or Strindex = "" Then
        If Len(txt1(7)) = 0 Then
            If Len(txt1(8)) <> 0 Then
                ShowDateRanErr
                txt1(7).SetFocus
                Exit Sub
            End If
        End If
    End If
    
    'Modify By Cheng 2003/02/05
    '若發文日為111111則不檢查
    If Strindex = "8" Or Strindex = "" Then
        If Me.txt1(8).Text <> "111111" Then
            If Not ChkDateRanPro(txt1(7), txt1(8), 5) And Len(txt1(8)) <> 0 Then
                'Modify By Cheng 2003/01/03
                If Me.txt1(8).Enabled And Me.txt1(8).Visible Then
                    '游標設在發文日
                    txt1(8).SetFocus
                    txt1_GotFocus (8)
                    Exit Sub
                Else
                    '游標設在完稿日
                    txt1(3).SetFocus
                    txt1_GotFocus (3)
                    Exit Sub
                End If
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
    
    ChkData = True
'End If
End Sub

Public Sub txt1_Validate(Index As Integer, Cancel As Boolean)
'add by nickc 2006/03/07
Dim P_DateLine As String
Dim CFP_DateLine As String
'Add By Cheng 2003/05/08
'If Index <> 2 And Index <> 3 And Index <> 4 And Index <> 7 And Index <> 8 And Index <> 9 And Index <> 10 Then Exit Sub
Select Case Index
Case 0 '工作時數
   'Modify By Sindy 2024/3/1 + And txt1(Index) <> "0" And Not (m_Flow = EMP_附加流程)
   'Modify By Sindy 2024/5/2 剔除 (txt1(Index).Tag = txt1(Index).Text And m_strRefVal = "Y")
   If txt1(Index) <> "" And txt1(Index) <> "0" And Not (m_Flow = EMP_附加流程) And _
      Not (txt1(Index).Tag = txt1(Index).Text And m_strRefVal = "Y") Then
      
      Cancel = Not PUB_CheckCP113(txt1(Index), m_CP01, m_CP10, m_CP14, , , m_strRefVal)
      If Cancel = True Then
         If txt1(Index).Enabled = True And txt1(Index).Visible = True Then
            txt1(Index).SetFocus
            txt1(Index).SelLength = Len(txt1(Index))
         End If
         Exit Sub
      End If
   End If
   
'Modified by Lydia 2021/12/28 去掉txt1(10)
Case 2, 3, 7, 8, 9, 12
    'Add By Sindy 2022/6/23 王錦寬.副總:凡C類來函，承辦人為工程師者，該道程序之齊備日僅能更改，請禁止取消。
    If Index = 2 Then '齊備日
      If Val(txt1(Index).Tag) > 0 And Left(m_strCP09, 1) = "C" And (PUB_GetST03(m_CP14) = "P10" Or PUB_GetST03(m_CP14) = "P11") Then
         If Val(txt1(Index).Text) = 0 Then
            MsgBox "C類來函齊備日僅能更改，禁止取消！", , "錯誤！"
            txt1(Index).Text = txt1(Index).Tag
            Cancel = True
            Exit Sub
         End If
      End If
    End If
    '2022/6/23 END
    
    'Add By Cheng 20036/04/28
    '若欄位無資料則不檢查
    If Me.txt1(Index).Text = "" Then Exit Sub
    'Modify By Cheng 2003/05/06
    'ChkTxt
    'Modify By Cheng 2004/02/03
'    ChkTxt "" & Index: DoEvents
    ChkTxt "" & Index
    'End
    If ChkData = False Then
        Cancel = True
        Exit Sub
    End If
    
'add by nickc 2007/08/16
Case 19   '英文核完日
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
Dim Cancel As Boolean
Dim strEmp As String 'Add By Sindy 2024/3/27

TxtValidate = False

'Added by Sindy 2024/04/11 分案和工作進度維護點選不可查閱工程師需要彈訊息
If cboCP14.Visible = True And m_CP14 <> Trim(Left(cboCP14.Text, 5)) Then
   If Left(Trim(cboCP14), 1) > "6" And Left(Trim(cboCP14), 1) < "F" Then
      If PUB_ChkCufaByCaseNo(Trim(Left(cboCP14, 6)), Me.Name, m_CP01 & m_CP02 & m_CP03 & m_CP04, "2") = False Then
         SSTab1.Tab = 1
         cboCP14.SetFocus
         Exit Function
      End If
   End If
End If
'2024/04/11 END

'Added by Morgan 2012/12/3
If ProState = "1" Then '個人權限
   If txtPA162.Enabled = True Then
      If txtPA162 = "" Then
         MsgBox "請設定是否要加註核准分割建議！", vbExclamation
         SSTab1.Tab = 1 'Add By Sindy 2024/3/15
         txtPA162.SetFocus
         Exit Function
      ElseIf txtPA162 = "Y" And Trim(txtDST05) = "" Then
         MsgBox "請輸入建議定稿文字！", vbExclamation
         SSTab1.Tab = 1 'Add By Sindy 2024/3/15
         txtDST05.SetFocus
         Exit Function
      ElseIf txtPA162 = "N" And Trim(txtDST05) <> "" Then
         MsgBox "當設定為""不要""加註核准分割建議時，不可輸入建議定稿文字！", vbExclamation
         SSTab1.Tab = 1 'Add By Sindy 2024/3/15
         txtDST05.SetFocus
         Exit Function
      End If
      
      'Added by Morgan 2012/12/26
      If txtDST05 <> "" Then
         strExc(0) = PUB_StringFilter(txtDST05)
         If strExc(0) <> txtDST05 Then
            If MsgBox("建議定稿文字內發現有跳行符號，存檔時將自動清除。是否要繼續??", vbYesNo + vbDefaultButton2) = vbYes Then
               txtDST05 = strExc(0)
            Else
               SSTab1.Tab = 1 'Add By Sindy 2024/3/15
               txtDST05.SetFocus
               Exit Function
            End If
         End If
      End If
      'end 2012/12/26
   End If
End If

'核稿人
'Modify By Sindy 2024/1/4 排除新案翻譯
'Modify By Sindy 2024/3/26 排除修改承辦人時,不需檢查核稿人
If Combo2.Enabled = True And Val(m_CP27) = 0 _
   And m_CP10 <> 翻譯 _
   And Not (cboCP14.Visible = True And m_CP14 <> Trim(Left(cboCP14.Text, 5))) Then
   
   '若核判表有設定核稿人時只可以修改但不可以空白
   If Trim(m_PP04) <> "" And Trim(Left(Combo2.Text, 6)) = "" Then
      MsgBox "核稿人不可空白!!!", vbExclamation + vbOKOnly
      SSTab1.Tab = 1 'Add By Sindy 2024/3/15
      Combo2.SetFocus
      Exit Function
   End If
   If Trim(Left(Combo2.Text, 6)) <> "" Then
      '增加檢查核稿人是否離職
      If PUB_ChkEmpFlowExists(lbl1(3), EMP_送核) = True And lblEP39 = "" Then
         If ChkStaffST04(Trim(Left(Combo2.Text, 6))) = True Then
            SSTab1.Tab = 1
            Combo2.SetFocus
            Exit Function
         End If
      End If
      '核稿人不可與承辦人相同
      'Modify By Sindy 2024/2/26 +And lblEP39.Caption = "":若已有核稿完成日就不控管; 因主管休假職代操作時,核稿人會改為操作人員
      If UCase(Trim(Left("" & Combo1.Text, 6))) = UCase(Trim(Left(Combo2.Text, 6))) And lblEP39.Caption = "" Then
         MsgBox "核稿人不可與承辦人相同!!!", vbExclamation + vbOKOnly
         SSTab1.Tab = 1 'Add By Sindy 2024/3/15
         Combo2.SetFocus
         Exit Function
      End If
      '檢查同組
      If PUB_GetST03(Trim(Left(Combo1.Text, 6))) = "F21" Then '外專工程師
         Call GetPrjSalesNM(Trim(Left(Combo2.Text, 6)), , , "st16", strExc(10))
         'Modify By Sindy 2024/3/5 排除員編第4號是9的人員(支援人員)
         'Modify By Sindy 2024/3/22 排除機械組
         'Modify By Sindy 2024/4/17 改共用函數
         'If m_EMPST16 <> strExc(10) And Mid(Trim(Left(Combo1.Text, 6)), 4, 1) <> "9" And m_EMPST16 <> "4" Then
         'Modify By Sindy 2025/1/21 開放法律所律師
         If PUB_GetST03(Trim(Left(Combo2.Text, 6))) <> "L01" Then
         '2025/1/21 END
            If m_EMPST16 <> strExc(10) And PUB_NeedChkFCPST16(IIf(m_CP10 = 翻譯, Trim(Left(Combo2.Text, 6)), Trim(Left(Combo1.Text, 6)))) = True Then
               MsgBox "核稿人與承辦人不同組別!!!", vbExclamation + vbOKOnly
               SSTab1.Tab = 1 'Add By Sindy 2024/3/15
               Combo2.SetFocus
               Exit Function
            End If
         End If
      End If
      'Modify By Sindy 2024/1/3 簡經理說核稿人不鎖必須是主任以上的人員,判發人員鎖住即可
'      '檢查52=代主任以上
'      If Trim(m_PP04) <> Trim(Left(Combo2.Text, 6)) Then
'         strExc(0) = "SELECT * FROM STAFF" & _
'                     " WHERE ST01='" & Trim(Left(Combo2.Text, 6)) & "'" & _
'                     " AND ST04='1' AND ST20<=52 AND ST20 is not null"
'         intI = 1
'         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'         If intI = 0 Then
'            MsgBox "核稿人必須是主任以上的人員!!!", vbExclamation + vbOKOnly
'            Combo2.SetFocus
'            Exit Function
'         End If
'      End If
      '2024/1/3 END
   End If
End If

'判發人
'Modify By Sindy 2024/3/26 排除修改承辦人時,不需檢查判發人
'Modify By Sindy 2024/3/26 排除修改核稿工程師時,不需檢查判發人
'Modify By Sindy 2025/3/18 +外專繪圖同新案翻譯歷程操作方式
If Combo6.Enabled = True And Val(m_CP27) = 0 _
   And Not (cboCP14.Visible = True And m_CP14 <> Trim(Left(cboCP14.Text, 5))) _
   And Not ((m_CP10 = 翻譯 Or m_CP10 = "931") And Combo2.Enabled = True And Trim(Left(Combo2.Tag, 5)) <> Trim(Left(Combo2.Text, 5))) Then
   
   'Add By Sindy 2024/3/27
   '承辦人
   'Modify By Sindy 2025/3/18 +外專繪圖同新案翻譯歷程操作方式
   If (m_CP10 = 翻譯 Or m_CP10 = "931") Then
      strEmp = Trim(Combo2.Text) '核稿工程師
      'Call GetPrjSalesNM(Trim(Left(strEmp, 6)), , , "st16", m_EMPST16)
      'Call GetPP04PP05(Trim(Left(strEmp, 6)))
   Else
      strEmp = Trim(Combo1.Text)
   End If
   '2024/3/27 END
   
   '若核判表有設定判發人時只可以修改但不可以空白
   If Trim(m_PP05) <> "" And Trim(Left(Combo6.Text, 6)) = "" Then
      MsgBox "判發人不可空白!!!", vbExclamation + vbOKOnly
      SSTab1.Tab = 1 'Add By Sindy 2024/3/15
      Combo6.SetFocus
      Exit Function
   End If
   If Trim(Left(Combo6.Text, 6)) <> "" Then
      '增加檢查判發人是否離職
      If PUB_ChkEmpFlowExists(lbl1(3), EMP_送判) = True And _
         PUB_ChkEmpFlowExists(lbl1(3), EMP_判發) = False Then
         If ChkStaffST04(Trim(Left(Combo6.Text, 6))) = True Then
            SSTab1.Tab = 1
            Combo6.SetFocus
            Exit Function
         End If
      End If
      '判發人不可與承辦人相同
      'Modify By Sindy 2024/2/26 +And LblEP42.Caption = "":若已有判發完成日就不控管; 因主管休假職代操作時,判發人會改為操作人員
      If UCase(Trim(Left(strEmp, 6))) = UCase(Trim(Left(Combo6.Text, 6))) And lblEP42.Caption = "" Then
         MsgBox "判發人不可與承辦人相同!!!", vbExclamation + vbOKOnly
         SSTab1.Tab = 1 'Add By Sindy 2024/3/15
         Combo6.SetFocus
         Exit Function
      End If
      '檢查同組
      If PUB_GetST03(Trim(Left(strEmp, 6))) = "F21" Then '外專工程師
         Call GetPrjSalesNM(Trim(Left(Combo6.Text, 6)), , , "st16", strExc(10))
         'Modify By Sindy 2024/3/5 排除員編第4號是9的人員(支援人員)
         'Modify By Sindy 2024/3/22 排除機械組
         'Modify By Sindy 2024/4/17 改共用函數
         'If m_EMPST16 <> strExc(10) And Mid(Trim(Left(strEmp, 6)), 4, 1) <> "9" And m_EMPST16 <> "4" Then
         'Modify By Sindy 2025/1/21 開放法律所律師
         If PUB_GetST03(Trim(Left(Combo6.Text, 6))) <> "L01" Then
         '2025/1/21 END
            If m_EMPST16 <> strExc(10) And PUB_NeedChkFCPST16(IIf(m_CP10 = 翻譯, Trim(Left(Combo2.Text, 6)), Trim(Left(Combo1.Text, 6)))) = True Then
               MsgBox "判發人與承辦人不同組別!!!", vbExclamation + vbOKOnly
               SSTab1.Tab = 1 'Add By Sindy 2024/3/15
               Combo6.SetFocus
               Exit Function
            End If
         End If
      End If
      '檢查52=代主任以上
      'Modify By Sindy 2024/3/22 排除機械組
      If Trim(m_PP05) <> Trim(Left(Combo6.Text, 6)) And m_EMPST16 <> "4" Then
         strExc(0) = "SELECT * FROM STAFF" & _
                     " WHERE ST01='" & Trim(Left(Combo6.Text, 6)) & "'" & _
                     " AND ST04='1' AND ST20<=52 AND ST20 is not null"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 0 Then
            MsgBox "判發人必須是主任以上的人員!!!", vbExclamation + vbOKOnly
            SSTab1.Tab = 1 'Add By Sindy 2024/3/15
            Combo6.SetFocus
            Exit Function
         End If
      End If
   End If
End If

Cancel = False
Call txt1_Validate(0, Cancel)
If Cancel = True Then
   'txt1(0).SetFocus
   SSTab1.Tab = 1 'Add By Sindy 2024/3/15
   Exit Function
End If

'檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
If txtEP12 <> "" Then
    If PUB_ChkUniText(Me, , True, "TextBox") = False Then
          Exit Function
    End If
End If

TxtValidate = True
End Function

'Add By Cheng 2003/04/28
Sub StrMenuOneRec(Optional ByVal Strindex As Integer = 1)
Dim ii As Integer
    For ii = 1 To Me.grd1.Rows - 1
        '若目次相同, 收文號也相同
        If Me.grd1.TextMatrix(ii, 0) = Me.lbl1(0).Caption And Me.grd1.TextMatrix(ii, 23) = m_strCP09 Then
            '承辦期限
            'Me.GRD1.TextMatrix(ii, 10) = ChangeTStringToTDateString(Me.txt1(12).Text)
            Me.grd1.TextMatrix(ii, 11) = ChangeTStringToTDateString(Me.txt1(12).Text)
            '齊備日
            Me.grd1.TextMatrix(ii, 13) = ChangeTStringToTDateString(Me.txt1(2).Text)
            '完稿日
            Me.grd1.TextMatrix(ii, 14) = ChangeTStringToTDateString(Me.txt1(3).Text)
            '預會日
            'Me.grd1.TextMatrix(ii, 15) = ChangeTStringToTDateString(Me.txt1(18).Text)
            '會稿日
            'Me.grd1.TextMatrix(ii, 16) = ChangeTStringToTDateString(Me.txt1(4).Text)
            '核稿人
            Me.grd1.TextMatrix(ii, 17) = IIf(Combo2.Text = "", "", IIf(InStr(Combo2, "==>") > 0, Trim(Mid(Combo2, 10)), Trim(Mid(Combo2, 6))))
            '核稿期限
            Me.grd1.TextMatrix(ii, 18) = ChangeTStringToTDateString(Me.txt1(7).Text)
            '發文日
            Me.grd1.TextMatrix(ii, 19) = ChangeTStringToTDateString(Me.txt1(8).Text)
            'Modify By Sindy 2024/1/17 改顯示進度備註
'            '承辦備註
'            Me.GRD1.TextMatrix(ii, 21) = Me.txtEP12.Text
            'Add By Sindy 2024/2/23
            '最近歷程
            Me.grd1.TextMatrix(ii, 29) = strR110035
            '2024/2/23 END
            
            '修正日期欄位排序問題(小於100年的前面補空白)
            'For intI = 10 To 21
            For intI = 11 To 20 '21
               If Len(grd1.TextMatrix(ii, intI)) = 8 Then
                 If Mid(grd1.TextMatrix(ii, intI), 3, 1) = "/" And Mid(grd1.TextMatrix(ii, intI), 6, 1) = "/" Then
                    grd1.TextMatrix(ii, intI) = " " & grd1.TextMatrix(ii, intI)
                 End If
               End If
            Next
            
            ChgGrdColor ii
            Exit For
        End If
    Next ii
    
    SWPRow = Strindex
    grd1.row = Val(SWPRow)
    grd1.col = 1
End Sub

'add by nickc 2006/02/27 控制只跟 DB 溝通一次
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
            'add by nickc 2006/03/14
            m_FieldList(nIndex).fiNewData = Empty
         End If
      End If
   Next nIndex
EXITSUB:
End Sub

'Add by Morgan 2008/12/3 自cmdok_Click抽出
Private Function FormSave() As Boolean
'Dim strToM As String, strSub As String, strTo As String, strContent As String
'Dim strCP01 As String, strCP02 As String, strCP03 As String, strCP04 As String
Dim strTo As String, strCC As String 'Add By Sindy 2024/3/11
Dim strMC07 As String, strMC08 As String

On Error GoTo ErrHnd
   
   'strToM = PUB_GetFCPEngSup(strUserNum) 'Added by Lydia 2020/08/24 外專工程師主管
   
   'Modify By Sindy 2018/6/20 + If m_Flow = "" Then
   If m_Flow = "" Then cnnConnection.BeginTrans
   
   cnnConnection.Execute "begin user_data.user_formname:='" & Me.Name & "';end;" 'Add by Morgan 2010/10/29
   
   '目次
   SeekTmpBk = Trim(lbl1(0).Caption)
   strSql = "Update EngineerProgress Set EP06=" & IIf(ChangeTStringToWString(txt1(2)) = "", "NULL", ChangeTStringToWString(txt1(2))) & _
            ",EP40='" & Trim(Left("" & Combo6.Text, 6)) & "'" & _
            ",EP04='" & Trim(Left("" & Combo2.Text, 6)) & "'" & _
            ",EP03='" & txt1(6) & "'" & _
            ",EP31=" & IIf(ChangeTStringToWString(txt1(13)) = "", "NULL", ChangeTStringToWString(txt1(13))) & _
            ",EP34='N'" & _
            ",EP41=" & CNULL(txt1(23)) & _
            ",EP11='" & txt1(9) & "'" & _
            ",EP12='" & txtEP12 & "'" & _
            ",EP08=" & IIf(ChangeTStringToWString(txt1(7)) = "", "NULL", ChangeTStringToWString(txt1(7)))
   'Modify By Sindy 2013/12/18 防止簽核流程已存入日期,但此處又更新到日期,如英文核完日 ex.CFP-023734
   '完稿日
   If Val(txt1(3).Tag) <> Val(txt1(3)) Then
      strSql = strSql & ",EP09=" & IIf(ChangeTStringToWString(txt1(3)) = "", "NULL", ChangeTStringToWString(txt1(3)))
   End If
   '外文核完日
   If Val(txt1(19).Tag) <> Val(ChangeTStringToWString(txt1(19))) Then
      strSql = strSql & ",ep33=" & IIf(ChangeTStringToWString(txt1(19)) = "", "NULL", ChangeTStringToWString(txt1(19)))
   End If
   '2013/12/18 END
   strSql = strSql & " Where EP02='" & lbl1(3).Caption & "' "
   cnnConnection.Execute strSql
   
   '加入 齊備日異動時，紀錄
   If DBDATE(Trim(lbl1(8))) <> DBDATE(Trim(txt1(2))) Then
       Pub_SaveLog strUserNum, "齊備日異動：" & DBDATE(lbl1(8)) & "==>" & DBDATE(txt1(2)) & " ", m_CP01, m_CP02, m_CP03, m_CP04, lbl1(3).Caption
   End If
   '加入 核稿人異動時，紀錄
   If Combo2.Tag <> Combo2.Text Then
       Pub_SaveLog strUserNum, "核稿人異動：" & Combo2.Tag & "==>" & Combo2.Text & " ", m_CP01, m_CP02, m_CP03, m_CP04, lbl1(3).Caption
   End If
   '判發人異動時，紀錄
   If Combo6.Tag <> Combo6.Text Then
       Pub_SaveLog strUserNum, "判發人異動：" & Combo6.Tag & "==>" & Combo6.Text & " ", m_CP01, m_CP02, m_CP03, m_CP04, lbl1(3).Caption
   End If
   '外文核稿人異動時，紀錄
   If Combo4.Tag <> Combo4.Text Then
      If txt1(23) = "2" Then
         Pub_SaveLog strUserNum, "日文核稿人異動：" & Combo4.Tag & "==>" & Combo4.Text & " ", m_CP01, m_CP02, m_CP03, m_CP04, lbl1(3).Caption
      Else
         Pub_SaveLog strUserNum, "英文核稿人異動：" & Combo4.Tag & "==>" & Combo4.Text & " ", m_CP01, m_CP02, m_CP03, m_CP04, lbl1(3).Caption
      End If
   End If
   
   '更新案件進度檔
   strSql = ""
   If txt1(0).Tag <> txt1(0).Text Then '工作時數
      strSql = IIf(strSql <> "", strSql & ",", "") & "cp113=" & CNULL(txt1(0).Text, True)
   End If
   If txt1(1).Tag <> txt1(1).Text Then '核稿時數
      strSql = IIf(strSql <> "", strSql & ",", "") & "cp114=" & CNULL(txt1(1).Text, True)
   End If
   'Add By Sindy 2024/1/2
   If txt1(12).Tag <> txt1(12).Text Then '承辦期限
      strSql = IIf(strSql <> "", strSql & ",", "") & "cp48=" & CNULL(txt1(12).Text, True)
   End If
   '2024/1/2 END
   'Add By Sindy 2024/3/6
   strTo = ""
   If cboCP14.Visible = True And m_CP14 <> Trim(Left(cboCP14.Text, 5)) Then
      strSql = IIf(strSql <> "", strSql & ",", "") & "cp14=" & CNULL(Trim(Left(cboCP14.Text, 5)))
      strTo = Trim(Left(cboCP14.Text, 5))
      strCC = IIf(m_CP14 <> "", m_CP14, "")
   'Modify By Sindy 2025/3/18 +外專繪圖同新案翻譯歷程操作方式
   ElseIf Combo2.Visible = True And Combo2.Enabled = True And (m_CP10 = "201" Or m_CP10 = "931") And _
      Trim(Left(Combo2.Tag, 5)) <> Trim(Left(Combo2.Text, 5)) Then
      strTo = Trim(Left(Combo2.Text, 5))
      strCC = Trim(Left(Combo2.Tag, 5))
   End If
   If strTo <> "" Then
      strExc(0) = "select * from caseprogress where cp09='" & lbl1(3) & "'"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
'         strCP01 = RsTemp.Fields("CP01")
'         strCP02 = RsTemp.Fields("CP02")
'         strCP03 = RsTemp.Fields("CP03")
'         strCP04 = RsTemp.Fields("CP04")
         If Trim("" & RsTemp.Fields("CP60")) <> "" Then
            '若已開請款單則換承辦人或核稿人時發Mail通知相關人員
            If Trim("" & RsTemp.Fields("CP60")) > "X" Then
               'Modify By Sindy 2025/3/18 +外專繪圖同新案翻譯歷程操作方式
               If (m_CP10 = "201" Or m_CP10 = "931") Then
                  Call PUB_PointReAssignInform(m_CP01 & "-" & m_CP02 & IIf(m_CP03 & m_CP04 = "000", "", "-" & m_CP03 & "-" & m_CP04), "" & RsTemp.Fields("CP60"), , , Trim(Left(Combo2.Tag, 5)), Trim(Left(Combo2.Text, 5)))
               Else
                  Call PUB_PointReAssignInform(m_CP01 & "-" & m_CP02 & IIf(m_CP03 & m_CP04 = "000", "", "-" & m_CP03 & "-" & m_CP04), "" & RsTemp.Fields("CP60"), m_CP14, Trim(Left(cboCP14.Text, 5)))
               End If
            End If
         End If
      End If
      'FCP-XXX變更承辦人，發給原承辦人及新承辦人，不必限制承辦人為虛帳號人員時才加發EMAIL。
      'Modify By Sindy 2024/3/14 FCP變更承辦人發EMAIL請再加控制 (因為此文要給外翻, 淑華要加收其他翻譯)
      '  若為C類來函，新的承辦人的屬於「協助機械組內專工程師」人員且同時屬於「協助機械組內專主管」人員時，
      '  副本加發系統特殊設定人員「M」(目前為淑華)。
      If Left(lbl1(3), 1) = "C" Then
         If InStr(Pub_GetSpecMan("協助機械組工程師"), Trim(Left(cboCP14.Text, 5))) > 0 _
            And InStr(Pub_GetSpecMan("協助機械組內專主管"), Trim(Left(cboCP14.Text, 5))) > 0 Then
            If strCC <> "" Then strCC = strCC & ";"
            strCC = strCC & Pub_GetSpecMan("M")
         End If
      End If
      '2024/3/14 END
      'Add By Sindy 2024/5/23 新增副本人員: 國外部對接主管
      If InStr(Pub_GetSpecMan("協助機械組工程師"), strTo) > 0 Then
         Call GetPrjSalesNM(strTo, , , "st52", strExc(10))
         If strExc(10) <> "" Then
            If strCC <> "" Then strCC = strCC & ";"
            strCC = strCC & strExc(10)
         End If
      End If
      '2024/5/23 END
      strExc(9) = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08,mc09,mc13)" & _
                  " VALUES ( '" & strUserNum & "','" & strTo & "',to_char(sysdate,'yyyymmdd')" & _
                  ",to_char(sysdate,'hh24miss'),'" & m_CP01 & "-" & m_CP02 & IIf(m_CP03 & m_CP04 <> "000", "-" & m_CP03 & "-" & m_CP04, "") & "變更承辦人'," & _
                  CNULL(IIf(m_CP14 <> "", "原承辦人：" & GetStaffName(m_CP14, True) & vbCrLf & "新承辦人：" & GetStaffName(Trim(Left(cboCP14.Text, 6)), True) & vbCrLf, "")) & _
                  ",'" & strCC & "','" & lbl1(3) & "')"
      cnnConnection.Execute strExc(9), intI
   End If
   '2024/3/6 END
   If strSql <> "" Then
      strSql = "UPDATE CASEPROGRESS SET " & strSql & _
               " WHERE CP09 = '" & Me.lbl1(3).Caption & "' "
      Pub_SeekTbLog strSql
      cnnConnection.Execute strSql, intI
      'Added by Lydia 2024/03/27 一併修改相關收文號之承辦人 ---Wilison
      If cboCP14.Visible = True And m_CP14 <> Trim(Left(cboCP14.Text, 5)) Then
         Call PUB_SaveFCPcp14Ex(m_CP01, m_CP02, m_CP03, m_CP04, Me.lbl1(3).Caption, m_CP10, Trim(Left(cboCP14.Text, 5)))
      End If
      'end 2024/03/27
   End If
  
'*************************************
'從 frm090901_1 Move過來
'*************************************
'   '完稿日
'   If Val(txt1(3).Tag) <> Val(txt1(3)) Then
'      'Added by Morgan 2012/5/17
'      '電話聯絡單完搞日輸入自動發Mail給主管
'      If txt1(3).Tag = "" Then
'         If m_CP10 = "945" Then
'             'Modified by Lydia 2020/08/24 改用模組取得
'             'strSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08)" & _
'               " select '" & strUserNum & "' mc01,oMan mc02,to_char(sysdate,'yyyymmdd') mc03,to_char(sysdate,'hh24miss') mc04" & _
'               ",cp01||'-'||cp02||decode(cp03||cp04,'000','','-'||cp03||'-'||cp04)||'('||cp09||')電話聯絡單已完稿請發文!(完搞日：" & txtEP09 & ") ' mc07" & _
'               ",'如旨' mc08 from caseprogress,staff,SetSpecMan" & _
'               " where cp09='" & lblCP09 & "' and cp27 is null and st01(+)=cp14 and OCODE=decode(st16,'1','T','2','R','3','S','4','T1')"
'             strSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08)" & _
'               " select '" & strUserNum & "' mc01,'" & strToM & "' mc02,to_char(sysdate,'yyyymmdd') mc03,to_char(sysdate,'hh24miss') mc04" & _
'               ",cp01||'-'||cp02||decode(cp03||cp04,'000','','-'||cp03||'-'||cp04)||'('||cp09||')電話聯絡單已完稿請發文!(完搞日：" & txt1(3) & ") ' mc07" & _
'               ",'如旨' mc08 from caseprogress " & _
'               " where cp09='" & m_strCP09 & "' and cp27 is null "
'            cnnConnection.Execute strSql, intI
'         End If
'      End If
'   End If
   If Me.SSTab2.TabVisible(0) = True Then
      'Added by Morgan 2012/12/3
      If txtPA162.Enabled = True Then
         strSql = "update patent set pa162='" & Me.txtPA162.Text & "' where pa01='" & m_CP01 & "' and pa02='" & m_CP02 & "' and pa03='" & m_CP03 & "' and pa04='" & m_CP04 & "'"
         Pub_SeekTbLog strSql
         cnnConnection.Execute strSql, intI
         
         strSql = "delete divsugtext where dst01='" & m_CP01 & "' and dst02='" & m_CP02 & "' and dst03='" & m_CP03 & "' and dst04='" & m_CP04 & "'"
         cnnConnection.Execute strSql, intI
         
         strSql = "insert into divsugtext(dst01,dst02,dst03,dst04,dst05,dst06,dst07,dst08,dst09) values " & _
            "('" & m_CP01 & "','" & m_CP02 & "','" & m_CP03 & "','" & m_CP04 & "'" & _
            ",'" & ChgSQL(txtDST05) & "','" & strUserNum & "',to_char(sysdate,'yyyymmdd'),to_char(sysdate,'hh24miss'),'" & m_strCP09 & "')"
            
         Pub_SeekTbLog strSql
         cnnConnection.Execute strSql, intI
         
         'Add By Sindy 2024/12/11 當工程師判斷，不需加註分割建議，
         '   並至工作進度資料維護將Y改為N以及空白改成N ->點選確定並存檔後 ->設定自動發email通知各區程序上核准發文。
         If cmd(1).Tag = "" Then '代表直接按"確定"鍵做存檔
            If txtPA162.Tag <> txtPA162.Text And txtPA162.Text = "N" Then
               'Add By Sindy 2024/12/19 增加判斷,有通知告准未發文者才發mail通知 ex:FCP-072781
               strExc(0) = "select cp09 from caseprogress" & _
                           " where cp01='" & m_CP01 & "' and cp02='" & m_CP02 & "' and cp03='" & m_CP03 & "' and cp04='" & m_CP04 & "'" & _
                           " and cp10='1917' and cp158=0 and cp159=0"
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
               If intI = 1 Then
               '2024/12/19 END
                  strTo = PUB_GetFCPHandler(m_CP01, m_CP02, m_CP03, m_CP04) '外專程序管制人
                  strCC = Pub_GetSpecMan("外專告准程序") & ";" & strUserNum & ";" & PUB_GetFCPEngSup(strUserNum) & ";backup"
                  strMC07 = "【工程師不做分割加註】請進行告准 Our Ref: " & m_CP01 & "-" & m_CP02 & IIf(m_CP03 & m_CP04 <> "000", "-" & m_CP03 & "-" & m_CP04, "")
                  strMC08 = "本案工程師已設定不做分割加註，並已回存基本檔，請進行下列事宜：" & vbCrLf & vbCrLf & _
                            "1. 請程序上核准發文日" & vbCrLf & _
                            "2. 請上發文日後通知" & GetPrjSalesNM(Pub_GetSpecMan("外專告准程序")) & "進行告准作業。"
                  strExc(9) = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08,mc09)" & _
                              " VALUES ( '" & strUserNum & "','" & strTo & "',to_char(sysdate,'yyyymmdd')" & _
                              ",to_char(sysdate,'hh24miss'),'" & strMC07 & "'," & CNULL(strMC08) & _
                              ",'" & strCC & "')"
                  cnnConnection.Execute strExc(9), intI
               End If
            End If
         End If
         '2024/12/11 END
         
'         If txtDST05 <> "" Then
'            strExc(1) = "'本所案號：'||pa01||'-'||pa02||decode(pa03||pa04,'000','','-'||pa03||'-'||pa04)||chr(13)||chr(10)" & _
'                        "||'案件名稱：'||pa05||chr(13)||chr(10)" & _
'                        "||'申請人：'||cu04||chr(13)||chr(10)" & _
'                        "||'承辦期限：'||sqldatet(cp48)||chr(13)||chr(10)" & _
'                        "||'核准分割建議定稿文字：'||dst05||chr(13)||chr(10)"
'
'            If m_CP10 = "1001" Then
'               strExc(2) = "cp01||'-'||cp02||decode(cp03||cp04,'000','','-'||cp03||'-'||cp04)||'('||cp09||')已輸入核准分割建議定稿文字，請審核後至系統上完稿日，再將卷宗交各區程序上發文日!'"
'            Else
'               strExc(2) = "cp01||'-'||cp02||decode(cp03||cp04,'000','','-'||cp03||'-'||cp04)||'('||cp09||')已輸入核准分割建議定稿文字，請審核!'"
'            End If
'
'            'Modified by Lydia 2020/08/24 改用模組取得
'            'strSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08)" & _
'               " select '" & strUserNum & "' mc01,decode(oMan,st01,B0102,oMan) mc02,to_char(sysdate,'yyyymmdd') mc03,to_char(sysdate,'hh24miss') mc04" & _
'               "," & strExc(2) & " mc07," & strExc(1) & " mc08" & _
'               " from caseprogress,patent,customer,divsugtext,staff,SetSpecMan,ABS001" & _
'               " where cp09='" & lblCP09 & "' and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04 and cu01(+)=substr(pa26,1,8) and cu02(+)=substr(pa26,9)" & _
'               " and dst01(+)=cp01 and dst02(+)=cp02 and dst03(+)=cp03 and dst04(+)=cp04" & _
'               " and st01(+)=cp14 and OCODE=decode(st16,'1','T','2','R','3','S','4','T1') and B0101(+)=st01"
'            strSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08)" & _
'               " select '" & strUserNum & "' mc01,'" & strToM & "' mc02,to_char(sysdate,'yyyymmdd') mc03,to_char(sysdate,'hh24miss') mc04" & _
'               "," & strExc(2) & " mc07," & strExc(1) & " mc08" & _
'               " from caseprogress,patent,customer,divsugtext " & _
'               " where cp09='" & m_strCP09 & "' and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04 and cu01(+)=substr(pa26,1,8) and cu02(+)=substr(pa26,9)" & _
'               " and dst01(+)=cp01 and dst02(+)=cp02 and dst03(+)=cp03 and dst04(+)=cp04"
'            cnnConnection.Execute strSql, intI
'         End If
      End If
      'end 2012/12/3
      
'      'Added by Morgan 2022/8/1--Winfrey
'      If m_CP10 = "1001" Then
'         strSub = ""
'         '1.不需加註分割建議，email通知各區程序上核准發文。
'         '2.分割建議主管上完稿日，email通知各區程序上核准發文。
'         If (txtPA162.Enabled = True And txtPA162 = "N") Then
'            strSub = "【工程師已確認不須分割加註】請進行告准 Our Ref: "
'         ElseIf (txtPA162 = "Y" And Val(txt1(3).Tag) <> Val(txt1(3)) And txt1(3) <> "") Then
'            'Memo by Morgan 2022/10/11 因為日文定稿還是要給工程師核稿,主旨保留以作識別
'            If PUB_GetLanguage(m_CP01, m_CP02, m_CP03, m_CP04) = "3" Then
'               strSub = "【工程師已完成分割加註(日文定稿)】請進行告准 Our Ref: "
'            Else
'               strSub = "【工程師已完成分割加註】請進行告准 Our Ref: "
'            End If
'         End If
'         If strSub <> "" Then
'            strTo = PUB_GetFCPHandler(m_CP01, m_CP02, m_CP03, m_CP04)
'            strContent = "1.請程序上發文日。" & vbCrLf & "2.請告准人員進行後續告准，感謝您。"
'            '1917=通知告准
'            strSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08,mc09)" & _
'               " select '" & strUserNum & "' mc01,'" & strTo & "' mc02,to_char(sysdate,'yyyymmdd') mc03,to_char(sysdate,'hh24miss') mc04" & _
'               ",'" & strSub & "'||cp01||'-'||cp02||decode(cp03||cp04,'000','','-'||cp03||'-'||cp04) mc07,'" & strContent & "' mc08,cp14 mc09" & _
'               " from caseprogress  where cp43='" & m_strCP09 & "' and cp10='1917' and cp27 is null"
'            cnnConnection.Execute strSql, intI
'         End If
'      End If
'      'end 2022/8/1
   End If
   'Added by Lydia 2015/04/24 +中說請款修正定稿文字
   'Modified by Lydia 2015/08/27 為了能拉動卷軸,改成locked
   'If txtAMD05.Enabled = True And txtAMD05.Text <> txtAMD05.Tag Then
   If txtAMD05.Locked = False And txtAMD05.Text <> txtAMD05.Tag And Me.SSTab2.TabVisible(1) = True Then
      strSql = "delete AmendedText where AMD01='" & m_CP01 & "' and AMD02='" & m_CP02 & "' and AMD03='" & m_CP03 & "' and AMD04='" & m_CP04 & "'"
      cnnConnection.Execute strSql, intI
      
      strSql = "insert into AmendedText(AMD01,AMD02,AMD03,AMD04,AMD05,AMD06,AMD07,AMD08,AMD09) values " & _
         "('" & m_CP01 & "','" & m_CP02 & "','" & m_CP03 & "','" & m_CP04 & "'" & _
         ",'" & ChgSQL(txtAMD05) & "','" & strUserNum & "',to_char(sysdate,'yyyymmdd'),to_char(sysdate,'hh24miss'),'" & m_strCP09 & "')"
         
      Pub_SeekTbLog strSql
      cnnConnection.Execute strSql, intI
   End If
'************************************* END

   cnnConnection.Execute "begin user_data.user_formname:=Null;end;" 'Add by Morgan 2010/10/29
   
   'Modify By Sindy 2018/6/20 + If m_Flow = "" Then
   If m_Flow = "" Then cnnConnection.CommitTrans
   
   FormSave = True
   Exit Function
   
ErrHnd:
   cnnConnection.Execute "begin user_data.user_formname:=Null;end;" 'Add by Morgan 2010/10/29
   'Modify By Sindy 2018/6/20 + If m_Flow = "" Then
   If m_Flow = "" Then cnnConnection.RollbackTrans
   MsgBox Err.Description, vbCritical
         
End Function

'批次發Mail
Public Sub BatctMail()
   Dim i As Integer
   For i = 1 To UBound(skMail)
        PUB_SendMail skMail(i).fiSender, skMail(i).fiReceiver, skMail(i).fiRecriverNo, skMail(i).fiSubject, skMail(i).fiContent
   Next i
   ReDim skMail(0) As SeekMails
   'Trigger 也會產生待發郵件
   'Modify By Sindy 2024/6/6
   'PUB_SendMailCache
   Call PUB_SendMailCache(, , , True)
   '2024/6/6 END
End Sub

'更新mdb暫存資料及第一畫面的Grid內容
Public Sub UpdEngMdb()

On Error GoTo ErrHnd
   
   'R110013.齊備日
   'R110014.完稿日
   'R110015.會稿日
   'R110017.核稿期限
   'R110016.核稿人
   'R110010.承辦期限
   'R110018.發文日
   'R110020.承辦備註
   'Add By Sindy 2024/2/23
   'R110035.最近歷程
   strExc(0) = "select GetEEPCurState('" & lbl1(3).Caption & "') from dual"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   strR110035 = ""
   If intI = 1 Then
      strR110035 = "" & RsTemp.Fields(0)
   End If
   '2024/2/23 END
   
   strSql = "UPDATE R090614 SET " & _
      "R110013='" & IIf(txt1(2) = "", "", Right(" " & ChangeTStringToTDateString(txt1(2)), 9)) & "'," & _
      "R110014='" & IIf(txt1(3) = "", "", Right(" " & ChangeTStringToTDateString(txt1(3)), 9)) & "'," & _
      "R110017='" & IIf(txt1(7) = "", "", Right(" " & ChangeTStringToTDateString(txt1(7)), 9)) & "'," & _
      "R110016='" & IIf(Combo2.Text = "", "", IIf(InStr(Combo2, "==>") > 0, Trim(Mid(Combo2, 10)), Trim(Mid(Combo2, 6)))) & "'," & _
      "R110010='" & IIf(txt1(12) = "", "", Right(" " & ChangeTStringToTDateString(txt1(12)), 9)) & "'," & _
      "R110018='" & IIf(txt1(8) = "", "", Right(" " & ChangeTStringToTDateString(txt1(8)), 9)) & "'," & _
      "R110020='" & txtEP12 & "',R110035='" & strR110035 & "'" & _
       " WHERE ID='" & strUserNum & "' AND R110022='" & lbl1(3).Caption & "' "
   adoEng.Execute strSql, intI
   
   For i = 1 To grd1.Rows - 1
      grd1.row = i
      grd1.col = 0
      '若目次相同, 收文號也相同
      If grd1.Text = SeekTmpBk And Me.grd1.TextMatrix(i, 23) = m_strCP09 Then
         Call MouseClick(i, False)
         StrMenuOneRec SWPRow
         Exit For
      End If
   Next i
   
ErrHnd:
   If Err.Number <> 0 Then
      MsgBox Err.Description
   End If
End Sub

'Add by Morgan 2008/12/8 從Form_Load抽出
'設定承辦人選單
Private Sub SetEngineer()
   strSql = "SELECT Distinct (R110001&' '&'(' & R110025&')') FROM R090614 WHERE ID='" & strUserNum & "'" & _
            " AND R110001='" & strUserNum & "'" & _
            " ORDER BY (R110001&' '&'(' & R110025&')') "
   CheckOC
   i = 0
   Combo1.Clear
   Combo1_String = "" '92.6.26 ADD BY SONIA
   With adoRecordset
      .CursorLocation = adUseClient
      .Open strSql, adoEng, adOpenStatic, adLockReadOnly
      If .RecordCount <> 0 And .RecordCount > 0 Then
         Do While .EOF = False
            Combo1.AddItem "" & .Fields(0), i
            i = i + 1
            '92.6.26 ADD BY SONIA
            If Combo1_String = "" Then
               Combo1_String = "'" & Trim(Left("" & .Fields(0), 6)) & "'"
            Else
               Combo1_String = Combo1_String + ",'" & Trim(Left("" & .Fields(0), 6)) & "'"
            End If
            '92.6.26 END
            .MoveNext
         Loop
      'Modify By Sindy 2024/4/10
         'Combo1.Text = Combo1.List(0)
      Else
         Combo1.AddItem strUserNum & " " & "(" & strUserName & ")", i
         Combo1_String = "'" & strUserNum & "'"
      End If
      Combo1.Text = Combo1.List(0)
      '2024/4/10 END
   End With
End Sub

'設定外文核稿人選單
Private Sub SetEngChecker()
   Combo4.Clear
   Combo4.AddItem "", 0
   If m_EMPST16 = "3" Then '日文
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

'承辦備註
Private Sub txtEP12_GotFocus()
    TextInverse txtEP12
End Sub

Private Sub txtCP64_GotFocus()
    TextInverse txtCP64
End Sub

'Add By Sindy 2023/10/16
Private Sub Command1_Click()
   txtDST05.Text = txtDST05Old.Text
End Sub
Private Sub txtPA162_GotFocus()
   CloseIme
   TextInverse txtPA162
End Sub
Private Sub txtPA162_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 8 And KeyAscii <> Asc("Y") And KeyAscii <> Asc("N") Then
      KeyAscii = 0
      Beep
   End If
End Sub
'Added by Lydia 2021/09/23 Form 2.0的TextBox增加右鍵選單功能
Private Sub txtDST05_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then Forms(0).PopupMenu2 txtDST05
End Sub
Private Sub txtAMD05_GotFocus()
   TextInverse txtAMD05
End Sub
'Added by Lydia 2021/09/23 Form 2.0的TextBox增加右鍵選單功能
Private Sub txtAMD05_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then Forms(0).PopupMenu2 txtAMD05
End Sub
'Private Sub Check1_Click()
'   If Check1.Value = 1 Then '發文後補分割建議
'      Me.SSTab2.TabVisible(1) = False: Me.SSTab2.TabCaption(0) = ""
'      Me.SSTab2.TabVisible(0) = True
'      Me.SSTab2.Tab = 0 'Add By Sindy 2023/11/21
'   Else
'      Me.SSTab2.TabVisible(0) = False: Me.SSTab2.TabCaption(1) = ""
'      Me.SSTab2.TabVisible(1) = True
'      Me.SSTab2.Tab = 1 'Add By Sindy 2023/11/21
'   End If
'End Sub
'2023/10/16 END
