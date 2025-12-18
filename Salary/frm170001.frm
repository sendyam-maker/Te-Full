VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm170001 
   BorderStyle     =   1  '單線固定
   Caption         =   "薪資基本資料"
   ClientHeight    =   6420
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   9168
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6420
   ScaleWidth      =   9168
   Begin VB.TextBox txtSD 
      Alignment       =   2  '置中對齊
      Height          =   270
      Index           =   48
      Left            =   7812
      MaxLength       =   1
      TabIndex        =   10
      Top             =   1872
      Width           =   240
   End
   Begin VB.TextBox txtSD 
      Alignment       =   2  '置中對齊
      Height          =   270
      Index           =   49
      Left            =   2610
      MaxLength       =   1
      TabIndex        =   8
      Text            =   "N"
      Top             =   1890
      Width           =   240
   End
   Begin VB.TextBox txtSD 
      Height          =   270
      Index           =   46
      Left            =   8170
      MaxLength       =   5
      TabIndex        =   5
      Text            =   "99999"
      Top             =   1332
      Width           =   700
   End
   Begin VB.TextBox txtSD 
      Alignment       =   2  '置中對齊
      Height          =   270
      Index           =   11
      Left            =   5265
      MaxLength       =   1
      TabIndex        =   11
      Text            =   "Y"
      Top             =   1350
      Width           =   240
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4275
      Left            =   0
      TabIndex        =   92
      Top             =   2160
      Width           =   9150
      _ExtentX        =   16150
      _ExtentY        =   7535
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "使用中資料"
      TabPicture(0)   =   "frm170001.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1(10)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblDsp(8)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblDsp(7)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label1(36)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label1(27)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label1(25)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label1(24)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label1(23)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label1(22)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Line3"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "lblSD(45)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label1(41)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label1(40)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label1(37)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Label3"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Label2"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Label1(35)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Label1(34)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Label1(33)"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Label1(32)"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Label1(31)"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "Label1(30)"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "Label1(29)"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "Label1(28)"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "Label1(1)"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "Line2"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "Line1"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "Label1(21)"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "Label1(20)"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "Label1(19)"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "Label1(18)"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "Label1(17)"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "Label1(16)"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "Label1(14)"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "Label1(13)"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "Label1(11)"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "Label1(43)"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "Label1(45)"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "Label1(46)"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).Control(39)=   "Label1(48)"
      Tab(0).Control(39).Enabled=   0   'False
      Tab(0).Control(40)=   "lblSD(47)"
      Tab(0).Control(40).Enabled=   0   'False
      Tab(0).Control(41)=   "Label1(38)"
      Tab(0).Control(41).Enabled=   0   'False
      Tab(0).Control(42)=   "txtSD(2)"
      Tab(0).Control(42).Enabled=   0   'False
      Tab(0).Control(43)=   "txtSD(13)"
      Tab(0).Control(43).Enabled=   0   'False
      Tab(0).Control(44)=   "txtSD(12)"
      Tab(0).Control(44).Enabled=   0   'False
      Tab(0).Control(45)=   "txtSD(15)"
      Tab(0).Control(45).Enabled=   0   'False
      Tab(0).Control(46)=   "txtSD(14)"
      Tab(0).Control(46).Enabled=   0   'False
      Tab(0).Control(47)=   "txtSD(44)"
      Tab(0).Control(47).Enabled=   0   'False
      Tab(0).Control(48)=   "txtSD(43)"
      Tab(0).Control(48).Enabled=   0   'False
      Tab(0).Control(49)=   "txtSD(35)"
      Tab(0).Control(49).Enabled=   0   'False
      Tab(0).Control(50)=   "txtSD(33)"
      Tab(0).Control(50).Enabled=   0   'False
      Tab(0).Control(51)=   "txtSD(31)"
      Tab(0).Control(51).Enabled=   0   'False
      Tab(0).Control(52)=   "txtSD(30)"
      Tab(0).Control(52).Enabled=   0   'False
      Tab(0).Control(53)=   "txtSD(29)"
      Tab(0).Control(53).Enabled=   0   'False
      Tab(0).Control(54)=   "txtSD(28)"
      Tab(0).Control(54).Enabled=   0   'False
      Tab(0).Control(55)=   "txtSD(26)"
      Tab(0).Control(55).Enabled=   0   'False
      Tab(0).Control(56)=   "txtSD(24)"
      Tab(0).Control(56).Enabled=   0   'False
      Tab(0).Control(57)=   "txtSD(22)"
      Tab(0).Control(57).Enabled=   0   'False
      Tab(0).Control(58)=   "txtSD(21)"
      Tab(0).Control(58).Enabled=   0   'False
      Tab(0).Control(59)=   "txtSD(20)"
      Tab(0).Control(59).Enabled=   0   'False
      Tab(0).Control(60)=   "txtSD(19)"
      Tab(0).Control(60).Enabled=   0   'False
      Tab(0).Control(61)=   "txtSD(27)"
      Tab(0).Control(61).Enabled=   0   'False
      Tab(0).Control(62)=   "txtSD(36)"
      Tab(0).Control(62).Enabled=   0   'False
      Tab(0).Control(63)=   "txtSD(25)"
      Tab(0).Control(63).Enabled=   0   'False
      Tab(0).Control(64)=   "txtSD(23)"
      Tab(0).Control(64).Enabled=   0   'False
      Tab(0).Control(65)=   "txtSD(34)"
      Tab(0).Control(65).Enabled=   0   'False
      Tab(0).Control(66)=   "txtSD(32)"
      Tab(0).Control(66).Enabled=   0   'False
      Tab(0).Control(67)=   "txtSD(10)"
      Tab(0).Control(67).Enabled=   0   'False
      Tab(0).Control(68)=   "txtSD(9)"
      Tab(0).Control(68).Enabled=   0   'False
      Tab(0).Control(69)=   "txtSD(45)"
      Tab(0).Control(69).Enabled=   0   'False
      Tab(0).Control(70)=   "txtSD(8)"
      Tab(0).Control(70).Enabled=   0   'False
      Tab(0).Control(71)=   "txtSD(47)"
      Tab(0).Control(71).Enabled=   0   'False
      Tab(0).Control(72)=   "txtSD(52)"
      Tab(0).Control(72).Enabled=   0   'False
      Tab(0).ControlCount=   73
      TabCaption(1)   =   "待更新資料"
      TabPicture(1)   =   "frm170001.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lblSL(33)"
      Tab(1).Control(1)=   "Label1(47)"
      Tab(1).Control(2)=   "lblSL(11)"
      Tab(1).Control(3)=   "lblSL(12)"
      Tab(1).Control(4)=   "lblSL(13)"
      Tab(1).Control(5)=   "lblSL(14)"
      Tab(1).Control(6)=   "lblSL(15)"
      Tab(1).Control(7)=   "lblSL(16)"
      Tab(1).Control(8)=   "lblSL(17)"
      Tab(1).Control(9)=   "Line4"
      Tab(1).Control(10)=   "Line5"
      Tab(1).Control(11)=   "lblSL(34)"
      Tab(1).Control(12)=   "Label1(56)"
      Tab(1).Control(13)=   "lblSL(19)"
      Tab(1).Control(14)=   "lblSL(20)"
      Tab(1).Control(15)=   "lblSL(21)"
      Tab(1).Control(16)=   "lblSL(22)"
      Tab(1).Control(17)=   "lblSL(23)"
      Tab(1).Control(18)=   "lblSL(24)"
      Tab(1).Control(19)=   "lblSL(25)"
      Tab(1).Control(20)=   "Label4"
      Tab(1).Control(21)=   "Label5"
      Tab(1).Control(22)=   "lblSL(37)"
      Tab(1).Control(23)=   "lblSL(38)"
      Tab(1).Control(24)=   "Line6"
      Tab(1).Control(25)=   "lblSL(9)"
      Tab(1).Control(26)=   "lblSL(10)"
      Tab(1).Control(27)=   "lblSL(5)"
      Tab(1).Control(28)=   "lblSL(4)"
      Tab(1).Control(29)=   "Label1(73)"
      Tab(1).Control(30)=   "lblSL(6)"
      Tab(1).Control(31)=   "lblDsp(9)"
      Tab(1).Control(32)=   "lblDsp(10)"
      Tab(1).Control(33)=   "lblSL(3)"
      Tab(1).Control(34)=   "lblSL(18)"
      Tab(1).Control(35)=   "lblSL(26)"
      Tab(1).Control(36)=   "lblSL(7)"
      Tab(1).Control(37)=   "lblSL(8)"
      Tab(1).Control(38)=   "lblSL(39)"
      Tab(1).Control(39)=   "lblSL(36)"
      Tab(1).Control(40)=   "txtSL(33)"
      Tab(1).Control(41)=   "txtSL(17)"
      Tab(1).Control(42)=   "txtSL(34)"
      Tab(1).Control(43)=   "txtSL(25)"
      Tab(1).Control(44)=   "txtSL(37)"
      Tab(1).Control(45)=   "txtSL(10)"
      Tab(1).Control(46)=   "txtSL(8)"
      Tab(1).Control(47)=   "txtSL(3)"
      Tab(1).Control(48)=   "txtSL(18)"
      Tab(1).Control(49)=   "txtSL(26)"
      Tab(1).Control(50)=   "txtSL(13)"
      Tab(1).Control(51)=   "txtSL(11)"
      Tab(1).Control(52)=   "txtSL(15)"
      Tab(1).Control(53)=   "txtSL(12)"
      Tab(1).Control(54)=   "txtSL(16)"
      Tab(1).Control(55)=   "txtSL(14)"
      Tab(1).Control(56)=   "txtSL(24)"
      Tab(1).Control(57)=   "txtSL(22)"
      Tab(1).Control(58)=   "txtSL(23)"
      Tab(1).Control(59)=   "txtSL(20)"
      Tab(1).Control(60)=   "txtSL(21)"
      Tab(1).Control(61)=   "txtSL(19)"
      Tab(1).Control(62)=   "txtSL(7)"
      Tab(1).Control(63)=   "txtSL(9)"
      Tab(1).Control(64)=   "txtSL(6)"
      Tab(1).Control(65)=   "txtSL(5)"
      Tab(1).Control(66)=   "txtSL(38)"
      Tab(1).Control(67)=   "txtSL(4)"
      Tab(1).Control(68)=   "txtSL(39)"
      Tab(1).Control(69)=   "txtSL(36)"
      Tab(1).ControlCount=   70
      TabCaption(2)   =   "勞健保補助及健保眷屬異動"
      TabPicture(2)   =   "frm170001.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label20"
      Tab(2).Control(1)=   "Label1(50)"
      Tab(2).Control(2)=   "Label18"
      Tab(2).Control(3)=   "Label1(44)"
      Tab(2).Control(4)=   "cboST56"
      Tab(2).Control(5)=   "cboST50"
      Tab(2).Control(6)=   "txtSD(16)"
      Tab(2).Control(7)=   "txtSD(17)"
      Tab(2).Control(8)=   "Frame1"
      Tab(2).ControlCount=   9
      Begin VB.TextBox txtSL 
         Alignment       =   1  '靠右對齊
         Height          =   285
         Index           =   36
         Left            =   -69870
         MaxLength       =   7
         TabIndex        =   52
         Text            =   "9999999"
         Top             =   1860
         Width           =   735
      End
      Begin VB.TextBox txtSD 
         Alignment       =   1  '靠右對齊
         Height          =   285
         Index           =   52
         Left            =   7380
         MaxLength       =   7
         TabIndex        =   21
         Text            =   "9999999"
         Top             =   1590
         Width           =   735
      End
      Begin VB.TextBox txtSL 
         Alignment       =   1  '靠右對齊
         Height          =   285
         Index           =   39
         Left            =   -67560
         MaxLength       =   7
         TabIndex        =   51
         Text            =   "9999999"
         Top             =   1590
         Width           =   735
      End
      Begin VB.TextBox txtSD 
         Alignment       =   1  '靠右對齊
         Height          =   285
         Index           =   47
         Left            =   5130
         Locked          =   -1  'True
         MaxLength       =   7
         TabIndex        =   186
         Text            =   "9999999"
         Top             =   3930
         Width           =   735
      End
      Begin VB.Frame Frame1 
         Caption         =   "眷屬資料"
         ForeColor       =   &H8000000D&
         Height          =   2865
         Left            =   -74865
         TabIndex        =   175
         Top             =   1140
         Width           =   8880
         Begin VB.ComboBox cboHL05 
            Enabled         =   0   'False
            Height          =   300
            ItemData        =   "frm170001.frx":0054
            Left            =   2655
            List            =   "frm170001.frx":0056
            Style           =   2  '單純下拉式
            TabIndex        =   181
            Top             =   585
            Width           =   4005
         End
         Begin VB.ComboBox textSR03 
            Enabled         =   0   'False
            Height          =   300
            ItemData        =   "frm170001.frx":0058
            Left            =   870
            List            =   "frm170001.frx":006B
            TabIndex        =   180
            Top             =   210
            Width           =   1785
         End
         Begin VB.CheckBox chkSR08 
            Caption         =   "健保眷屬"
            Enabled         =   0   'False
            Height          =   285
            Left            =   180
            TabIndex        =   178
            Top             =   600
            Width           =   1035
         End
         Begin VB.CommandButton cmdLog 
            Caption         =   "健保異動資料"
            Height          =   315
            Left            =   7110
            TabIndex        =   176
            Top             =   570
            Width           =   1635
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid grd1 
            Height          =   1875
            Left            =   90
            TabIndex        =   177
            Top             =   930
            Width           =   8715
            _ExtentX        =   15367
            _ExtentY        =   3302
            _Version        =   393216
            BackColor       =   16777215
            Cols            =   1
            FixedCols       =   0
            ForeColorSel    =   16777215
            ScrollTrack     =   -1  'True
            HighLight       =   0
            SelectionMode   =   1
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
         Begin MSForms.TextBox textSR04 
            Height          =   300
            Left            =   3570
            TabIndex        =   179
            Top             =   225
            Width           =   1725
            VariousPropertyBits=   671105049
            MaxLength       =   12
            Size            =   "7223;529"
            FontName        =   "新細明體-ExtB"
            FontEffects     =   1073750016
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "姓名："
            Height          =   180
            Left            =   2835
            TabIndex        =   184
            Top             =   270
            Width           =   540
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "健保補助類別："
            Height          =   180
            Left            =   1350
            TabIndex        =   183
            Top             =   645
            Width           =   1260
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "稱謂："
            Height          =   180
            Left            =   210
            TabIndex        =   182
            Top             =   270
            Width           =   540
         End
      End
      Begin VB.TextBox txtSD 
         Alignment       =   1  '靠右對齊
         Height          =   270
         Index           =   17
         Left            =   -73560
         MaxLength       =   2
         TabIndex        =   174
         Text            =   "99.99"
         Top             =   780
         Width           =   600
      End
      Begin VB.TextBox txtSD 
         Height          =   270
         Index           =   16
         Left            =   -73560
         MaxLength       =   1
         TabIndex        =   173
         Text            =   "Y"
         Top             =   510
         Width           =   285
      End
      Begin VB.ComboBox cboST50 
         Height          =   276
         ItemData        =   "frm170001.frx":009C
         Left            =   -70140
         List            =   "frm170001.frx":009E
         Style           =   2  '單純下拉式
         TabIndex        =   168
         Top             =   795
         Width           =   4005
      End
      Begin VB.ComboBox cboST56 
         Height          =   276
         ItemData        =   "frm170001.frx":00A0
         Left            =   -70140
         List            =   "frm170001.frx":00A2
         Style           =   2  '單純下拉式
         TabIndex        =   167
         Top             =   495
         Width           =   4005
      End
      Begin VB.TextBox txtSL 
         Alignment       =   1  '靠右對齊
         Height          =   285
         Index           =   4
         Left            =   -67590
         MaxLength       =   2
         TabIndex        =   71
         Text            =   "99.99"
         Top             =   3930
         Width           =   735
      End
      Begin VB.TextBox txtSD 
         Alignment       =   1  '靠右對齊
         Height          =   285
         Index           =   8
         Left            =   7380
         MaxLength       =   5
         TabIndex        =   41
         Text            =   "99.99"
         Top             =   3930
         Width           =   735
      End
      Begin VB.TextBox txtSL 
         Alignment       =   1  '靠右對齊
         Height          =   285
         Index           =   38
         Left            =   -69840
         Locked          =   -1  'True
         MaxLength       =   7
         TabIndex        =   70
         Text            =   "9999999"
         Top             =   3930
         Width           =   735
      End
      Begin VB.TextBox txtSD 
         Alignment       =   1  '靠右對齊
         Height          =   285
         Index           =   45
         Left            =   2880
         Locked          =   -1  'True
         MaxLength       =   7
         TabIndex        =   40
         Text            =   "9999999"
         Top             =   3930
         Width           =   735
      End
      Begin VB.TextBox txtSD 
         Alignment       =   1  '靠右對齊
         Height          =   285
         Index           =   9
         Left            =   7380
         MaxLength       =   7
         TabIndex        =   36
         Text            =   "9999999"
         Top             =   3390
         Width           =   735
      End
      Begin VB.TextBox txtSD 
         Alignment       =   1  '靠右對齊
         Height          =   285
         Index           =   10
         Left            =   7380
         MaxLength       =   7
         TabIndex        =   39
         Text            =   "9999999"
         Top             =   3660
         Width           =   735
      End
      Begin VB.TextBox txtSD 
         Alignment       =   1  '靠右對齊
         Height          =   285
         Index           =   32
         Left            =   7368
         MaxLength       =   7
         TabIndex        =   27
         Text            =   "9999999"
         Top             =   2460
         Width           =   735
      End
      Begin VB.TextBox txtSD 
         Alignment       =   1  '靠右對齊
         Height          =   285
         Index           =   34
         Left            =   7380
         MaxLength       =   7
         TabIndex        =   30
         Text            =   "9999999"
         Top             =   2730
         Width           =   735
      End
      Begin VB.TextBox txtSD 
         Alignment       =   1  '靠右對齊
         Height          =   285
         Index           =   23
         Left            =   7380
         MaxLength       =   7
         TabIndex        =   16
         Text            =   "9999999"
         Top             =   1050
         Width           =   735
      End
      Begin VB.TextBox txtSD 
         Alignment       =   1  '靠右對齊
         Height          =   285
         Index           =   25
         Left            =   7380
         MaxLength       =   7
         TabIndex        =   19
         Text            =   "9999999"
         Top             =   1320
         Width           =   735
      End
      Begin VB.TextBox txtSL 
         Alignment       =   1  '靠右對齊
         Height          =   285
         Index           =   5
         Left            =   -67590
         MaxLength       =   7
         TabIndex        =   66
         Text            =   "9999999"
         Top             =   3390
         Width           =   735
      End
      Begin VB.TextBox txtSL 
         Alignment       =   1  '靠右對齊
         Height          =   285
         Index           =   6
         Left            =   -67590
         MaxLength       =   7
         TabIndex        =   69
         Text            =   "9999999"
         Top             =   3660
         Width           =   735
      End
      Begin VB.TextBox txtSL 
         Alignment       =   1  '靠右對齊
         Height          =   285
         Index           =   9
         Left            =   -69840
         MaxLength       =   7
         TabIndex        =   65
         Text            =   "9999999"
         Top             =   3390
         Width           =   735
      End
      Begin VB.TextBox txtSL 
         Alignment       =   1  '靠右對齊
         Height          =   285
         Index           =   7
         Left            =   -72090
         MaxLength       =   7
         TabIndex        =   64
         Text            =   "9999999"
         Top             =   3390
         Width           =   735
      End
      Begin VB.TextBox txtSL 
         Alignment       =   1  '靠右對齊
         Height          =   285
         Index           =   19
         Left            =   -72495
         MaxLength       =   7
         TabIndex        =   55
         Text            =   "9999999"
         Top             =   2460
         Width           =   735
      End
      Begin VB.TextBox txtSL 
         Alignment       =   1  '靠右對齊
         Height          =   285
         Index           =   21
         Left            =   -72495
         MaxLength       =   7
         TabIndex        =   58
         Text            =   "9999999"
         Top             =   2730
         Width           =   735
      End
      Begin VB.TextBox txtSL 
         Alignment       =   1  '靠右對齊
         Height          =   285
         Index           =   20
         Left            =   -69840
         MaxLength       =   7
         TabIndex        =   56
         Text            =   "9999999"
         Top             =   2460
         Width           =   735
      End
      Begin VB.TextBox txtSL 
         Alignment       =   1  '靠右對齊
         Height          =   285
         Index           =   23
         Left            =   -69840
         MaxLength       =   7
         TabIndex        =   59
         Text            =   "9999999"
         Top             =   2730
         Width           =   735
      End
      Begin VB.TextBox txtSL 
         Alignment       =   1  '靠右對齊
         Height          =   285
         Index           =   22
         Left            =   -67590
         MaxLength       =   7
         TabIndex        =   57
         Text            =   "9999999"
         Top             =   2460
         Width           =   735
      End
      Begin VB.TextBox txtSL 
         Alignment       =   1  '靠右對齊
         Height          =   285
         Index           =   24
         Left            =   -67590
         MaxLength       =   7
         TabIndex        =   60
         Text            =   "9999999"
         Top             =   2730
         Width           =   735
      End
      Begin VB.TextBox txtSL 
         Alignment       =   1  '靠右對齊
         Height          =   285
         Index           =   14
         Left            =   -67590
         MaxLength       =   7
         TabIndex        =   46
         Text            =   "9999999"
         Top             =   1050
         Width           =   735
      End
      Begin VB.TextBox txtSL 
         Alignment       =   1  '靠右對齊
         Height          =   285
         Index           =   16
         Left            =   -67590
         MaxLength       =   7
         TabIndex        =   49
         Text            =   "9999999"
         Top             =   1320
         Width           =   735
      End
      Begin VB.TextBox txtSL 
         Alignment       =   1  '靠右對齊
         Height          =   285
         Index           =   12
         Left            =   -69870
         MaxLength       =   7
         TabIndex        =   45
         Text            =   "9999999"
         Top             =   1050
         Width           =   735
      End
      Begin VB.TextBox txtSL 
         Alignment       =   1  '靠右對齊
         Height          =   285
         Index           =   15
         Left            =   -69870
         MaxLength       =   7
         TabIndex        =   48
         Text            =   "9999999"
         Top             =   1320
         Width           =   735
      End
      Begin VB.TextBox txtSL 
         Alignment       =   1  '靠右對齊
         Height          =   285
         Index           =   11
         Left            =   -72495
         MaxLength       =   7
         TabIndex        =   44
         Text            =   "9999999"
         Top             =   1050
         Width           =   735
      End
      Begin VB.TextBox txtSL 
         Alignment       =   1  '靠右對齊
         Height          =   285
         Index           =   13
         Left            =   -72495
         MaxLength       =   7
         TabIndex        =   47
         Text            =   "9999999"
         Top             =   1320
         Width           =   735
      End
      Begin VB.TextBox txtSL 
         Alignment       =   1  '靠右對齊
         Height          =   285
         Index           =   26
         Left            =   -69840
         MaxLength       =   7
         TabIndex        =   63
         Text            =   "9999999"
         Top             =   3000
         Width           =   735
      End
      Begin VB.TextBox txtSL 
         Alignment       =   1  '靠右對齊
         Height          =   285
         Index           =   18
         Left            =   -69870
         MaxLength       =   7
         TabIndex        =   53
         Text            =   "9999999"
         Top             =   1590
         Width           =   735
      End
      Begin VB.TextBox txtSL 
         Height          =   285
         Index           =   3
         Left            =   -69645
         MaxLength       =   1
         TabIndex        =   42
         Text            =   "1"
         Top             =   420
         Width           =   285
      End
      Begin VB.TextBox txtSL 
         Alignment       =   1  '靠右對齊
         Height          =   285
         Index           =   8
         Left            =   -72090
         MaxLength       =   7
         TabIndex        =   67
         Text            =   "9999999"
         Top             =   3660
         Width           =   735
      End
      Begin VB.TextBox txtSL 
         Alignment       =   1  '靠右對齊
         Height          =   285
         Index           =   10
         Left            =   -69840
         MaxLength       =   7
         TabIndex        =   68
         Text            =   "9999999"
         Top             =   3660
         Width           =   735
      End
      Begin VB.TextBox txtSL 
         Alignment       =   1  '靠右對齊
         Height          =   285
         Index           =   37
         Left            =   -67590
         MaxLength       =   7
         TabIndex        =   62
         Text            =   "9999999"
         Top             =   3000
         Width           =   735
      End
      Begin VB.TextBox txtSL 
         Alignment       =   1  '靠右對齊
         Height          =   285
         Index           =   25
         Left            =   -72495
         MaxLength       =   7
         TabIndex        =   61
         Text            =   "9999999"
         Top             =   3000
         Width           =   735
      End
      Begin VB.TextBox txtSL 
         Height          =   285
         Index           =   34
         Left            =   -73440
         MaxLength       =   1
         TabIndex        =   54
         Text            =   "2"
         Top             =   2190
         Width           =   285
      End
      Begin VB.TextBox txtSL 
         Alignment       =   1  '靠右對齊
         Height          =   285
         Index           =   17
         Left            =   -72495
         MaxLength       =   7
         TabIndex        =   50
         Text            =   "9999999"
         Top             =   1590
         Width           =   735
      End
      Begin VB.TextBox txtSL 
         Height          =   285
         Index           =   33
         Left            =   -73515
         MaxLength       =   1
         TabIndex        =   43
         Text            =   "1"
         Top             =   780
         Width           =   285
      End
      Begin VB.TextBox txtSD 
         Alignment       =   1  '靠右對齊
         Height          =   285
         Index           =   36
         Left            =   5130
         MaxLength       =   7
         TabIndex        =   33
         Text            =   "9999999"
         Top             =   3000
         Width           =   735
      End
      Begin VB.TextBox txtSD 
         Alignment       =   1  '靠右對齊
         Height          =   285
         Index           =   27
         Left            =   5100
         MaxLength       =   7
         TabIndex        =   23
         Text            =   "9999999"
         Top             =   1590
         Width           =   735
      End
      Begin VB.TextBox txtSD 
         Height          =   285
         Index           =   19
         Left            =   1485
         MaxLength       =   1
         TabIndex        =   13
         Text            =   "1"
         Top             =   780
         Width           =   285
      End
      Begin VB.TextBox txtSD 
         Alignment       =   1  '靠右對齊
         Height          =   285
         Index           =   20
         Left            =   2475
         MaxLength       =   7
         TabIndex        =   14
         Text            =   "9999999"
         Top             =   1050
         Width           =   735
      End
      Begin VB.TextBox txtSD 
         Alignment       =   1  '靠右對齊
         Height          =   285
         Index           =   21
         Left            =   5100
         MaxLength       =   7
         TabIndex        =   15
         Text            =   "9999999"
         Top             =   1050
         Width           =   735
      End
      Begin VB.TextBox txtSD 
         Alignment       =   1  '靠右對齊
         Height          =   285
         Index           =   22
         Left            =   2475
         MaxLength       =   7
         TabIndex        =   17
         Text            =   "9999999"
         Top             =   1320
         Width           =   735
      End
      Begin VB.TextBox txtSD 
         Alignment       =   1  '靠右對齊
         Height          =   285
         Index           =   24
         Left            =   5100
         MaxLength       =   7
         TabIndex        =   18
         Text            =   "9999999"
         Top             =   1320
         Width           =   735
      End
      Begin VB.TextBox txtSD 
         Alignment       =   1  '靠右對齊
         Height          =   285
         Index           =   26
         Left            =   2475
         MaxLength       =   7
         TabIndex        =   20
         Text            =   "9999999"
         Top             =   1590
         Width           =   735
      End
      Begin VB.TextBox txtSD 
         Height          =   285
         Index           =   28
         Left            =   1530
         MaxLength       =   1
         TabIndex        =   24
         Text            =   "2"
         Top             =   2190
         Width           =   285
      End
      Begin VB.TextBox txtSD 
         Alignment       =   1  '靠右對齊
         Height          =   285
         Index           =   29
         Left            =   2475
         MaxLength       =   7
         TabIndex        =   25
         Text            =   "9999999"
         Top             =   2460
         Width           =   735
      End
      Begin VB.TextBox txtSD 
         Alignment       =   1  '靠右對齊
         Height          =   285
         Index           =   30
         Left            =   5130
         MaxLength       =   7
         TabIndex        =   26
         Text            =   "9999999"
         Top             =   2460
         Width           =   735
      End
      Begin VB.TextBox txtSD 
         Alignment       =   1  '靠右對齊
         Height          =   285
         Index           =   31
         Left            =   2475
         MaxLength       =   7
         TabIndex        =   28
         Text            =   "9999999"
         Top             =   2730
         Width           =   735
      End
      Begin VB.TextBox txtSD 
         Alignment       =   1  '靠右對齊
         Height          =   285
         Index           =   33
         Left            =   5130
         MaxLength       =   7
         TabIndex        =   29
         Text            =   "9999999"
         Top             =   2730
         Width           =   735
      End
      Begin VB.TextBox txtSD 
         Alignment       =   1  '靠右對齊
         Height          =   285
         Index           =   35
         Left            =   2475
         MaxLength       =   7
         TabIndex        =   31
         Text            =   "9999999"
         Top             =   3000
         Width           =   735
      End
      Begin VB.TextBox txtSD 
         Alignment       =   1  '靠右對齊
         Height          =   285
         Index           =   43
         Left            =   5100
         MaxLength       =   7
         TabIndex        =   22
         Text            =   "9999999"
         Top             =   1860
         Width           =   735
      End
      Begin VB.TextBox txtSD 
         Alignment       =   1  '靠右對齊
         Height          =   285
         Index           =   44
         Left            =   7380
         MaxLength       =   7
         TabIndex        =   32
         Text            =   "9999999"
         Top             =   3000
         Width           =   735
      End
      Begin VB.TextBox txtSD 
         Alignment       =   1  '靠右對齊
         Height          =   285
         Index           =   14
         Left            =   5130
         MaxLength       =   7
         TabIndex        =   35
         Text            =   "9999999"
         Top             =   3390
         Width           =   735
      End
      Begin VB.TextBox txtSD 
         Alignment       =   1  '靠右對齊
         Height          =   285
         Index           =   15
         Left            =   5130
         MaxLength       =   7
         TabIndex        =   38
         Text            =   "9999999"
         Top             =   3660
         Width           =   735
      End
      Begin VB.TextBox txtSD 
         Alignment       =   1  '靠右對齊
         Height          =   285
         Index           =   12
         Left            =   2880
         MaxLength       =   7
         TabIndex        =   34
         Text            =   "9999999"
         Top             =   3390
         Width           =   735
      End
      Begin VB.TextBox txtSD 
         Alignment       =   1  '靠右對齊
         Height          =   285
         Index           =   13
         Left            =   2880
         MaxLength       =   7
         TabIndex        =   37
         Text            =   "9999999"
         Top             =   3660
         Width           =   735
      End
      Begin VB.TextBox txtSD 
         Height          =   285
         Index           =   2
         Left            =   5355
         MaxLength       =   1
         TabIndex        =   12
         Text            =   "1"
         Top             =   420
         Width           =   285
      End
      Begin VB.Label lblSL 
         AutoSize        =   -1  'True
         Caption         =   "退休金投保薪資："
         Height          =   195
         Index           =   36
         Left            =   -71310
         TabIndex        =   148
         Top             =   1920
         Width           =   1455
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "證照津貼："
         Height          =   180
         Index           =   38
         Left            =   6495
         TabIndex        =   191
         Top             =   1620
         Width           =   900
      End
      Begin VB.Label lblSL 
         AutoSize        =   -1  'True
         Caption         =   "證照津貼："
         Height          =   180
         Index           =   39
         Left            =   -68445
         TabIndex        =   190
         Top             =   1635
         Width           =   900
      End
      Begin VB.Label lblSD 
         AutoSize        =   -1  'True
         Caption         =   "健保投保金額："
         Height          =   180
         Index           =   47
         Left            =   3885
         TabIndex        =   187
         Top             =   3960
         Width           =   1260
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "勞退自提費率：               %"
         Height          =   180
         Index           =   44
         Left            =   -74820
         TabIndex        =   172
         Top             =   840
         Width           =   2070
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "員工健保補助類別："
         Height          =   180
         Left            =   -71775
         TabIndex        =   171
         Top             =   855
         Width           =   1620
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "適用勞退新制：        ( Y:適用 )"
         Height          =   180
         Index           =   50
         Left            =   -74820
         TabIndex        =   170
         Top             =   540
         Width           =   2355
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "員工勞保補助類別："
         Height          =   180
         Left            =   -71775
         TabIndex        =   169
         Top             =   555
         Width           =   1620
      End
      Begin VB.Label lblSL 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         Caption         =   "特殊健保投保薪資："
         Height          =   195
         Index           =   8
         Left            =   -73710
         TabIndex        =   165
         Top             =   3690
         Width           =   1635
      End
      Begin VB.Label lblSL 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         Caption         =   "特殊勞保投保薪資："
         Height          =   195
         Index           =   7
         Left            =   -73710
         TabIndex        =   164
         Top             =   3420
         Width           =   1635
      End
      Begin VB.Label Label1 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         Caption         =   "特殊健保投保薪資："
         Height          =   195
         Index           =   48
         Left            =   1260
         TabIndex        =   163
         Top             =   3690
         Width           =   1635
      End
      Begin VB.Label Label1 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         Caption         =   "特殊勞保投保薪資："
         Height          =   195
         Index           =   46
         Left            =   1260
         TabIndex        =   162
         Top             =   3420
         Width           =   1635
      End
      Begin VB.Label lblSL 
         AutoSize        =   -1  'True
         Caption         =   "特殊退休金投保薪資："
         Height          =   180
         Index           =   26
         Left            =   -71640
         TabIndex        =   161
         Top             =   3060
         Width           =   1800
      End
      Begin VB.Label lblSL 
         AutoSize        =   -1  'True
         Caption         =   "特殊退休金投保薪資："
         Height          =   180
         Index           =   18
         Left            =   -71655
         TabIndex        =   160
         Top             =   1650
         Width           =   1800
      End
      Begin VB.Label lblSL 
         AutoSize        =   -1  'True
         Caption         =   "編　　制：        (R:正  T:試  S:留  P:內兼  F:外兼)"
         Height          =   195
         Index           =   3
         Left            =   -70545
         TabIndex        =   159
         Top             =   480
         Width           =   3750
      End
      Begin VB.Label lblDsp 
         Caption         =   "台一國際專利法律事務所"
         Height          =   195
         Index           =   10
         Left            =   -73125
         TabIndex        =   158
         Top             =   2235
         Width           =   4065
      End
      Begin VB.Label lblDsp 
         Caption         =   "台一國際專利商標事務所"
         Height          =   195
         Index           =   9
         Left            =   -73155
         TabIndex        =   157
         Top             =   825
         Width           =   4065
      End
      Begin VB.Label lblSL 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         Caption         =   "喪事互助："
         Height          =   195
         Index           =   6
         Left            =   -68505
         TabIndex        =   156
         Top             =   3690
         Width           =   915
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "扣繳項目"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   11.4
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   73
         Left            =   -74790
         TabIndex        =   155
         Top             =   3390
         Width           =   960
      End
      Begin VB.Label lblSL 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         Caption         =   "所得稅率："
         Height          =   195
         Index           =   4
         Left            =   -68505
         TabIndex        =   154
         Top             =   3990
         Width           =   915
      End
      Begin VB.Label lblSL 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         Caption         =   "婚事互助："
         Height          =   195
         Index           =   5
         Left            =   -68505
         TabIndex        =   153
         Top             =   3450
         Width           =   915
      End
      Begin VB.Label lblSL 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         Caption         =   "健  保  費："
         Height          =   195
         Index           =   10
         Left            =   -70740
         TabIndex        =   152
         Top             =   3720
         Width           =   915
      End
      Begin VB.Label lblSL 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         Caption         =   "勞  保  費："
         Height          =   195
         Index           =   9
         Left            =   -70740
         TabIndex        =   151
         Top             =   3450
         Width           =   915
      End
      Begin VB.Line Line6 
         BorderWidth     =   2
         X1              =   -74910
         X2              =   -66810
         Y1              =   3330
         Y2              =   3330
      End
      Begin VB.Label lblSL 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         Caption         =   "健保投保金額："
         Height          =   180
         Index           =   38
         Left            =   -71085
         TabIndex        =   150
         Top             =   3960
         Width           =   1260
      End
      Begin VB.Label lblSL 
         AutoSize        =   -1  'True
         Caption         =   "退休金投保薪資："
         Height          =   195
         Index           =   37
         Left            =   -69030
         TabIndex        =   149
         Top             =   3060
         Width           =   1455
      End
      Begin VB.Label Label5 
         Alignment       =   2  '置中對齊
         Caption         =   "第二家"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   15.6
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   1095
         Left            =   -74865
         TabIndex        =   147
         Top             =   2220
         Width           =   375
      End
      Begin VB.Label Label4 
         Alignment       =   2  '置中對齊
         Caption         =   "第一家"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   15.6
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   1125
         Left            =   -74865
         TabIndex        =   146
         Top             =   780
         Width           =   375
      End
      Begin VB.Label lblSL 
         AutoSize        =   -1  'True
         Caption         =   "特  支  費："
         Height          =   195
         Index           =   25
         Left            =   -73395
         TabIndex        =   145
         Top             =   3060
         Width           =   915
      End
      Begin VB.Label lblSL 
         AutoSize        =   -1  'True
         Caption         =   "房租津貼："
         Height          =   195
         Index           =   24
         Left            =   -68490
         TabIndex        =   144
         Top             =   2790
         Width           =   915
      End
      Begin VB.Label lblSL 
         AutoSize        =   -1  'True
         Caption         =   "差旅津貼："
         Height          =   195
         Index           =   23
         Left            =   -70740
         TabIndex        =   143
         Top             =   2790
         Width           =   915
      End
      Begin VB.Label lblSL 
         AutoSize        =   -1  'True
         Caption         =   "午餐津貼："
         Height          =   195
         Index           =   22
         Left            =   -68490
         TabIndex        =   142
         Top             =   2520
         Width           =   915
      End
      Begin VB.Label lblSL 
         AutoSize        =   -1  'True
         Caption         =   "技術津貼："
         Height          =   180
         Index           =   21
         Left            =   -73380
         TabIndex        =   141
         Top             =   2790
         Width           =   900
      End
      Begin VB.Label lblSL 
         AutoSize        =   -1  'True
         Caption         =   "職務津貼："
         Height          =   195
         Index           =   20
         Left            =   -70740
         TabIndex        =   140
         Top             =   2520
         Width           =   915
      End
      Begin VB.Label lblSL 
         AutoSize        =   -1  'True
         Caption         =   "基本薪資："
         Height          =   195
         Index           =   19
         Left            =   -73395
         TabIndex        =   139
         Top             =   2520
         Width           =   915
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "所得項目"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   11.4
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   56
         Left            =   -74430
         TabIndex        =   138
         Top             =   2460
         Width           =   960
      End
      Begin VB.Label lblSL 
         AutoSize        =   -1  'True
         Caption         =   "公  司  別："
         Height          =   195
         Index           =   34
         Left            =   -74430
         TabIndex        =   137
         Top             =   2235
         Width           =   915
      End
      Begin VB.Line Line5 
         BorderWidth     =   2
         X1              =   -74910
         X2              =   -66780
         Y1              =   2160
         Y2              =   2160
      End
      Begin VB.Line Line4 
         BorderWidth     =   2
         X1              =   -74910
         X2              =   -66810
         Y1              =   750
         Y2              =   750
      End
      Begin VB.Label lblSL 
         AutoSize        =   -1  'True
         Caption         =   "特  支  費："
         Height          =   195
         Index           =   17
         Left            =   -73395
         TabIndex        =   136
         Top             =   1650
         Width           =   915
      End
      Begin VB.Label lblSL 
         AutoSize        =   -1  'True
         Caption         =   "房租津貼："
         Height          =   195
         Index           =   16
         Left            =   -68490
         TabIndex        =   135
         Top             =   1380
         Width           =   915
      End
      Begin VB.Label lblSL 
         AutoSize        =   -1  'True
         Caption         =   "差旅津貼："
         Height          =   195
         Index           =   15
         Left            =   -70770
         TabIndex        =   134
         Top             =   1380
         Width           =   915
      End
      Begin VB.Label lblSL 
         AutoSize        =   -1  'True
         Caption         =   "午餐津貼："
         Height          =   195
         Index           =   14
         Left            =   -68490
         TabIndex        =   133
         Top             =   1080
         Width           =   915
      End
      Begin VB.Label lblSL 
         AutoSize        =   -1  'True
         Caption         =   "技術津貼："
         Height          =   180
         Index           =   13
         Left            =   -73380
         TabIndex        =   132
         Top             =   1380
         Width           =   900
      End
      Begin VB.Label lblSL 
         AutoSize        =   -1  'True
         Caption         =   "職務津貼："
         Height          =   195
         Index           =   12
         Left            =   -70770
         TabIndex        =   131
         Top             =   1080
         Width           =   915
      End
      Begin VB.Label lblSL 
         AutoSize        =   -1  'True
         Caption         =   "基本薪資："
         Height          =   195
         Index           =   11
         Left            =   -73395
         TabIndex        =   130
         Top             =   1080
         Width           =   915
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "所得項目"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   11.4
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   47
         Left            =   -74460
         TabIndex        =   129
         Top             =   1080
         Width           =   960
      End
      Begin VB.Label lblSL 
         AutoSize        =   -1  'True
         Caption         =   "公  司  別："
         Height          =   195
         Index           =   33
         Left            =   -74460
         TabIndex        =   128
         Top             =   825
         Width           =   915
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "特殊退休金投保薪資："
         Height          =   180
         Index           =   45
         Left            =   3330
         TabIndex        =   127
         Top             =   3052
         Width           =   1800
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "特殊退休金投保薪資："
         Height          =   180
         Index           =   43
         Left            =   3300
         TabIndex        =   126
         Top             =   1650
         Width           =   1800
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "公  司  別："
         Height          =   195
         Index           =   11
         Left            =   540
         TabIndex        =   125
         Top             =   825
         Width           =   915
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "所得項目"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   11.4
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   13
         Left            =   540
         TabIndex        =   124
         Top             =   1080
         Width           =   960
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "基本薪資："
         Height          =   195
         Index           =   14
         Left            =   1575
         TabIndex        =   123
         Top             =   1080
         Width           =   915
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "職務津貼："
         Height          =   195
         Index           =   16
         Left            =   4200
         TabIndex        =   122
         Top             =   1080
         Width           =   915
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "技術津貼："
         Height          =   180
         Index           =   17
         Left            =   1590
         TabIndex        =   121
         Top             =   1380
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "午餐津貼："
         Height          =   195
         Index           =   18
         Left            =   6480
         TabIndex        =   120
         Top             =   1080
         Width           =   915
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "差旅津貼："
         Height          =   195
         Index           =   19
         Left            =   4200
         TabIndex        =   119
         Top             =   1380
         Width           =   915
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "房租津貼："
         Height          =   195
         Index           =   20
         Left            =   6480
         TabIndex        =   118
         Top             =   1380
         Width           =   915
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "特  支  費："
         Height          =   195
         Index           =   21
         Left            =   1575
         TabIndex        =   117
         Top             =   1650
         Width           =   915
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         X1              =   90
         X2              =   8190
         Y1              =   750
         Y2              =   750
      End
      Begin VB.Line Line2 
         BorderWidth     =   2
         X1              =   90
         X2              =   8220
         Y1              =   2160
         Y2              =   2160
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "公  司  別："
         Height          =   195
         Index           =   1
         Left            =   540
         TabIndex        =   116
         Top             =   2235
         Width           =   915
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "所得項目"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   11.4
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   28
         Left            =   540
         TabIndex        =   115
         Top             =   2460
         Width           =   960
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "基本薪資："
         Height          =   195
         Index           =   29
         Left            =   1575
         TabIndex        =   114
         Top             =   2505
         Width           =   915
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "職務津貼："
         Height          =   195
         Index           =   30
         Left            =   4230
         TabIndex        =   113
         Top             =   2505
         Width           =   915
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "技術津貼："
         Height          =   180
         Index           =   31
         Left            =   1590
         TabIndex        =   112
         Top             =   2782
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "午餐津貼："
         Height          =   195
         Index           =   32
         Left            =   6480
         TabIndex        =   111
         Top             =   2505
         Width           =   915
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "差旅津貼："
         Height          =   195
         Index           =   33
         Left            =   4230
         TabIndex        =   110
         Top             =   2775
         Width           =   915
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "房租津貼："
         Height          =   195
         Index           =   34
         Left            =   6480
         TabIndex        =   109
         Top             =   2775
         Width           =   915
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "特  支  費："
         Height          =   195
         Index           =   35
         Left            =   1575
         TabIndex        =   108
         Top             =   3045
         Width           =   915
      End
      Begin VB.Label Label2 
         Alignment       =   2  '置中對齊
         Caption         =   "第一家"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   15.6
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   1125
         Left            =   135
         TabIndex        =   107
         Top             =   780
         Width           =   375
      End
      Begin VB.Label Label3 
         Alignment       =   2  '置中對齊
         Caption         =   "第二家"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   15.6
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   1095
         Left            =   135
         TabIndex        =   106
         Top             =   2220
         Width           =   375
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "兼職人員資料以時薪輸入"
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   37
         Left            =   5970
         TabIndex        =   105
         Top             =   810
         Width           =   1995
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "退休金投保薪資："
         Height          =   195
         Index           =   40
         Left            =   3660
         TabIndex        =   104
         Top             =   1920
         Width           =   1455
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "退休金投保薪資："
         Height          =   195
         Index           =   41
         Left            =   5940
         TabIndex        =   103
         Top             =   3045
         Width           =   1455
      End
      Begin VB.Label lblSD 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         Caption         =   "勞健保投保薪資："
         Height          =   180
         Index           =   45
         Left            =   1455
         TabIndex        =   102
         Top             =   3960
         Width           =   1440
      End
      Begin VB.Line Line3 
         BorderWidth     =   2
         X1              =   90
         X2              =   8190
         Y1              =   3330
         Y2              =   3330
      End
      Begin VB.Label Label1 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         Caption         =   "勞  保  費："
         Height          =   195
         Index           =   22
         Left            =   4230
         TabIndex        =   101
         Top             =   3450
         Width           =   915
      End
      Begin VB.Label Label1 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         Caption         =   "健  保  費："
         Height          =   195
         Index           =   23
         Left            =   4230
         TabIndex        =   100
         Top             =   3720
         Width           =   915
      End
      Begin VB.Label Label1 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         Caption         =   "婚事互助："
         Height          =   195
         Index           =   24
         Left            =   6465
         TabIndex        =   99
         Top             =   3450
         Width           =   915
      End
      Begin VB.Label Label1 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         Caption         =   "所得稅率："
         Height          =   195
         Index           =   25
         Left            =   6465
         TabIndex        =   98
         Top             =   3990
         Width           =   915
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "扣繳項目"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   11.4
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   27
         Left            =   180
         TabIndex        =   97
         Top             =   3390
         Width           =   960
      End
      Begin VB.Label Label1 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         Caption         =   "喪事互助："
         Height          =   195
         Index           =   36
         Left            =   6465
         TabIndex        =   96
         Top             =   3690
         Width           =   915
      End
      Begin VB.Label lblDsp 
         Caption         =   "台一國際專利商標事務所"
         Height          =   195
         Index           =   7
         Left            =   1845
         TabIndex        =   95
         Top             =   825
         Width           =   4065
      End
      Begin VB.Label lblDsp 
         Caption         =   "台一國際專利法律事務所"
         Height          =   195
         Index           =   8
         Left            =   1845
         TabIndex        =   94
         Top             =   2235
         Width           =   4065
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "編　　制：        (R:正  T:試  S:留  P:內兼  F:外兼)"
         Height          =   195
         Index           =   10
         Left            =   4455
         TabIndex        =   93
         Top             =   480
         Width           =   3750
      End
   End
   Begin VB.TextBox txtSD 
      Alignment       =   2  '置中對齊
      Height          =   270
      Index           =   3
      Left            =   3915
      MaxLength       =   1
      TabIndex        =   2
      Text            =   "Y"
      Top             =   1110
      Width           =   285
   End
   Begin VB.TextBox txtSD 
      Alignment       =   2  '置中對齊
      BackColor       =   &H00FFFFFF&
      Height          =   270
      Index           =   4
      Left            =   6975
      MaxLength       =   1
      TabIndex        =   3
      Text            =   "Y"
      Top             =   1110
      Width           =   285
   End
   Begin VB.TextBox txtSD 
      Height          =   270
      Index           =   7
      Left            =   1020
      MaxLength       =   1
      TabIndex        =   4
      Text            =   "9"
      Top             =   1350
      Width           =   285
   End
   Begin VB.TextBox txtSD 
      Height          =   270
      Index           =   6
      Left            =   6060
      MaxLength       =   30
      TabIndex        =   7
      Text            =   "123456789012345678901234567890"
      Top             =   1596
      Width           =   2835
   End
   Begin VB.TextBox txtSD 
      Height          =   270
      Index           =   5
      Left            =   1020
      MaxLength       =   1
      TabIndex        =   6
      Text            =   "1"
      Top             =   1620
      Width           =   285
   End
   Begin VB.TextBox txtSD 
      Height          =   270
      Index           =   1
      Left            =   1020
      MaxLength       =   6
      TabIndex        =   1
      Text            =   "123456"
      Top             =   630
      Width           =   735
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7605
      Top             =   -30
      _ExtentX        =   995
      _ExtentY        =   995
      BackColor       =   -2147483643
      ImageWidth      =   20
      ImageHeight     =   20
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm170001.frx":00A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm170001.frx":03C0
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm170001.frx":06DC
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm170001.frx":08B8
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm170001.frx":0BD4
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm170001.frx":0EF0
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm170001.frx":120C
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm170001.frx":1528
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm170001.frx":1844
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm170001.frx":1B60
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm170001.frx":1E7C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar TBar1 
      Align           =   1  '對齊表單上方
      Height          =   576
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9168
      _ExtentX        =   16171
      _ExtentY        =   1016
      ButtonWidth     =   1101
      ButtonHeight    =   974
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      DisabledImageList=   "ImageList1"
      HotImageList    =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   14
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "新增"
            Key             =   "keyInsert"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "修改"
            Key             =   "keyUpdate"
            ImageIndex      =   2
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   1
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "刪除"
            Key             =   "keyDelete"
            ImageIndex      =   3
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   1
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "查詢"
            Key             =   "keyQuery"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   1
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "第一筆"
            Key             =   "keyFirst"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "前一筆"
            Key             =   "keyPrevious"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "後一筆"
            Key             =   "keyNext"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "最後筆"
            Key             =   "keyLast"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   1
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "確定"
            Key             =   "keyOk"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "取消"
            Key             =   "keyCancel"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   1
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "結束"
            Key             =   "keyExit"
            ImageIndex      =   11
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.TextBox txtSD 
      Alignment       =   2  '置中對齊
      Height          =   270
      Index           =   50
      Left            =   5196
      MaxLength       =   1
      TabIndex        =   9
      Top             =   1875
      Width           =   240
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "勞保是否無就保：     ( Y:無就保 )"
      Height          =   180
      Left            =   6408
      TabIndex        =   192
      Top             =   1944
      Width           =   2604
   End
   Begin MSForms.TextBox textCUID 
      Height          =   300
      Left            =   2700
      TabIndex        =   91
      TabStop         =   0   'False
      Top             =   600
      Width           =   5700
      VariousPropertyBits=   671105055
      Size            =   "7223;529"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "是否列印薪資單：     ( Y:要印 )"
      Height          =   180
      Left            =   3792
      TabIndex        =   189
      Top             =   1920
      Width           =   2400
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "兼職薪資所得是否扣補充保費：     ( N:不扣 )"
      Height          =   180
      Left            =   120
      TabIndex        =   188
      Top             =   1920
      Width           =   3480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "帳號通知銀行薪資年月："
      Height          =   180
      Index           =   39
      Left            =   6195
      TabIndex        =   185
      Top             =   1395
      Width           =   1980
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "勞健保是否以合夥人身分投保：     ( Y:是)"
      Height          =   180
      Index           =   26
      Left            =   2775
      TabIndex        =   166
      Top             =   1395
      Width           =   3255
   End
   Begin VB.Label lblDsp 
      AutoSize        =   -1  'True
      Caption         =   "4"
      Height          =   180
      Index           =   6
      Left            =   2280
      TabIndex        =   90
      Top             =   1395
      Width           =   90
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "( Y 應稅 )"
      Height          =   180
      Left            =   7335
      TabIndex        =   89
      Top             =   1155
      Width           =   735
   End
   Begin VB.Label lblDsp 
      AutoSize        =   -1  'True
      Caption         =   "1 北 "
      Height          =   180
      Index           =   5
      Left            =   1440
      TabIndex        =   88
      Top             =   1155
      Width           =   360
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "( Y已婚 )"
      Height          =   180
      Left            =   4275
      TabIndex        =   87
      Top             =   1155
      Width           =   690
   End
   Begin VB.Label lblDsp 
      AutoSize        =   -1  'True
      Caption         =   "L  本國"
      Height          =   180
      Index           =   4
      Left            =   6990
      TabIndex        =   86
      Top             =   915
      Width           =   555
   End
   Begin VB.Label lblDsp 
      AutoSize        =   -1  'True
      Caption         =   "B1 台一"
      Height          =   180
      Index           =   3
      Left            =   3960
      TabIndex        =   85
      Top             =   915
      Width           =   615
   End
   Begin VB.Label lblDsp 
      AutoSize        =   -1  'True
      Caption         =   "A1 台一部"
      Height          =   180
      Index           =   2
      Left            =   1050
      TabIndex        =   84
      Top             =   915
      Width           =   795
   End
   Begin MSForms.Label lblName 
      Height          =   285
      Left            =   1830
      TabIndex        =   83
      Top             =   660
      Width           =   780
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "1376;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "帳　　號："
      Height          =   180
      Index           =   9
      Left            =   5160
      TabIndex        =   82
      Top             =   1665
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "入帳類別：        (1:現金  2:北  3:匯款  4:中  5:南  6:高  7:其他)"
      Height          =   180
      Index           =   8
      Left            =   120
      TabIndex        =   81
      Top             =   1665
      Width           =   4665
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "眷口數："
      Height          =   180
      Index           =   7
      Left            =   1485
      TabIndex        =   80
      Top             =   1395
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "扶養人數："
      Height          =   180
      Index           =   6
      Left            =   120
      TabIndex        =   79
      Top             =   1395
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "應　　稅："
      Height          =   180
      Index           =   4
      Left            =   6030
      TabIndex        =   78
      Top             =   1155
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "已　　婚："
      Height          =   180
      Index           =   3
      Left            =   2955
      TabIndex        =   77
      Top             =   1155
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "國　　籍："
      Height          =   180
      Index           =   15
      Left            =   6030
      TabIndex        =   76
      Top             =   915
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "職　　稱："
      Height          =   180
      Index           =   12
      Left            =   2970
      TabIndex        =   75
      Top             =   915
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "部　　門："
      Height          =   180
      Index           =   2
      Left            =   120
      TabIndex        =   74
      Top             =   915
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "員工所屬所別："
      Height          =   180
      Index           =   5
      Left            =   120
      TabIndex        =   73
      Top             =   1155
      Width           =   1260
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "員工代號："
      Height          =   180
      Index           =   0
      Left            =   120
      TabIndex        =   72
      Top             =   675
      Width           =   900
   End
End
Attribute VB_Name = "frm170001"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/12/20 Form2.0已修改
'Memo By Sonia 2012/12/6 智權人員欄已修改
'Memo by Morgan 2010/12/2 員工編號欄已修改
'Memo by Morgan 2010/7/27 日期欄已修改
'Memo by Morgan 2024/1/31 新部門已修改
'Create by Morgan 2008/12/17
Option Explicit

Dim m_EditMode As Integer '1:新增 2:修改 3:刪除 4:查詢

Dim m_bInsert As Boolean
Dim m_bUpdate As Boolean
Dim m_bDelete As Boolean
Dim m_bQuery As Boolean

Dim m_FieldList() As FIELDITEM
Dim TF_SD As Integer '欄位數
Dim oText As Object, oLabel As Object
Dim idx As Integer
Dim iR(1 To 15) As String '勞保勞退健保費率資料
Dim m_bConfirmCheck As Boolean
Dim m_ST13 '到職日


Private Sub cmdLog_Click()
   Dim i As Integer, strSR08 As String, strHL05 As String
   For i = 1 To GRD1.Rows - 1
      GRD1.row = i
      GRD1.col = 0
      If GRD1.CellBackColor = &HFFC0C0 Then
         With frm160001_1
            .strHL01 = txtSD(1)
            .strHL02 = GRD1.TextMatrix(i, 12)
            .InitForm
            .Show vbModal
            GetNewHiData txtSD(1), GRD1.TextMatrix(i, 12), strSR08, strHL05
            GRD1.TextMatrix(i, 6) = strSR08
            If strSR08 = "Y" Then
               chkSR08.Value = 1
            Else
               chkSR08.Value = 0
            End If
            GRD1.TextMatrix(i, 13) = strHL05
            SelCombo cboHL05, strHL05
         End With
         Exit For
      End If
   Next
   If i = GRD1.Rows Then
      MsgBox "請先點選眷屬資料！"
   End If
End Sub

Private Sub Form_Load()
   '取得使用者執行各項功能的權限
   m_bInsert = IsUserHasRightOfFunction(Me.Name, strAdd, False)
   m_bUpdate = IsUserHasRightOfFunction(Me.Name, strEdit, False)
   m_bDelete = IsUserHasRightOfFunction(Me.Name, strDel, False)
   m_bQuery = IsUserHasRightOfFunction(Me.Name, strFind, False)

   MoveFormToCenter Me
   
   textCUID.BackColor = &H8000000F
   
   InitialField
   If ShowRecord(-2) = True Then
      m_EditMode = 0
   Else
      Form_KeyDown vbKeyF2, 0
   End If
   SetInputEntry
   UpdateToolbarState
   SSTab1.Tab = 0
   
   'Add by Morgan 2011/2/23 避免新舊制混淆，改不顯示
   Label1(40).Visible = False
   txtSD(43).Visible = False
   Label1(41).Visible = False
   txtSD(44).Visible = False
   lblSL(36).Visible = False
   txtSL(36).Visible = False
   lblSL(37).Visible = False
   txtSL(37).Visible = False

'Removed by Morgan 2013/1/21 sd11 改為 勞健保是否以合夥人身分投保
'   'Add by Morgan 2012/2/24 適用一般勞健保費率已無作用，改不顯示
'   Label1(26).Visible = False
'   txtSD(11).Visible = False
'end 2013/1/21

   'Add by Morgan 2012/2/24 勞健保投保薪資僅系統計算用，改不顯示
   If Pub_StrUserSt03 <> "M51" Then
      lblSD(45).Visible = False
      txtSD(45).Visible = False
      'Added by Morgan 2013/1/21
      lblSL(38).Visible = False
      txtSL(38).Visible = False
      lblSD(47).Visible = False
      txtSD(47).Visible = False
   End If
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm170001 = Nothing
End Sub

Private Sub SetIR()
   strExc(0) = "select * from InsuranceRate"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      For intI = 1 To 15
         iR(intI) = "" & RsTemp.Fields("IR" & Format(intI, "00"))
      Next
   End If
End Sub

' 初始化欄位陣列
Private Sub InitialField()
   strExc(0) = "select * from SalaryData where rownum<1"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI <> 2 Then
      With RsTemp
      TF_SD = .Fields.Count
      ReDim m_FieldList(TF_SD) As FIELDITEM
      For Each oText In txtSD
         idx = oText.Index
         m_FieldList(idx).fiName = "SD" & Format(idx, "00")
         'Modified by Lydia 2017/06/29 O12和O8的Type不同,統一做文字處理
         'If .Fields(m_FieldList(idx).fiName).Type = 200 Then
            m_FieldList(idx).fiType = 0
         'Else
         '   m_FieldList(idx).fiType = 1
         'End If
         'end 2017/06/29
      Next
      End With
   End If
   SetIR
   
   'Add by Morgan 2009/6/26
   cboHL05.Clear
   cboHL05.AddItem "無"
   cboST50.Clear
   cboST50.AddItem "無"
   cboST56.Clear
   cboST56.AddItem "無"
   strSql = "select HR01||' '||HR04 from HiReduce order by 1"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      With RsTemp
      While Not .EOF
         cboHL05.AddItem "" & .Fields(0).Value
         cboST50.AddItem "" & .Fields(0).Value
         .MoveNext
      Wend
      End With
   End If
   
   strSql = "select LR01||' '||LR04 from LiReduce order by 1"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      With RsTemp
      While Not .EOF
         cboST56.AddItem "" & .Fields(0).Value
         .MoveNext
      Wend
      End With
   End If
   'end 2009/6/26
   
End Sub
' 顯示資料
'p_iWay:0=尋找,-2=首筆,-1=前筆,+1=後筆,2=末筆
Private Function ShowRecord(Optional ByVal p_iWay As Integer = 0) As Boolean
   
   Dim stKey01 As String
   Dim adoRst As New ADODB.Recordset
   
   stKey01 = txtSD(1)
   
   Select Case p_iWay
      Case 0
         strExc(0) = "SELECT * FROM SalaryData" & _
            " WHERE SD01 = '" & stKey01 & "'"
      Case -2
         strExc(0) = "SELECT * FROM SalaryData order by 1 ASC"
      Case -1
         strExc(0) = "SELECT * FROM SalaryData" & _
            " WHERE SD01 <'" & stKey01 & "' order by 1 DESC"
      Case 1
         strExc(0) = "SELECT * FROM SalaryData" & _
            " WHERE SD01 >'" & stKey01 & "' order by 1 ASC"
      Case 2
         strExc(0) = "SELECT * FROM SalaryData order by 1 DESC"
   End Select
   intI = 1
   adoRst.MaxRecords = 1
   Set adoRst = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      ClearField
      UpdateCtrlData adoRst
      ShowRecord = True
   Else
      If p_iWay = -1 Then
         MsgBox "已經是第一筆！", vbInformation
      ElseIf p_iWay = 1 Then
         MsgBox "已經是最後筆！", vbInformation
      Else
         MsgBox "查無資料！", vbInformation
         ClearField
      End If
   End If
   
   If m_EditMode = 0 Then
      SetCtrlReadOnly True
   End If
   Set adoRst = Nothing
   If Me.Visible = True Then
      txtSD(1).SetFocus
      txtSD_GotFocus 1
   End If
End Function

Private Sub txtSD_Change(Index As Integer)
   If Index = 1 Then
      If txtSD(Index) = "" Then
         lblName = "" 'Modify By Sindy 2021/12/20
      End If
   End If
End Sub

Private Sub txtSD_GotFocus(Index As Integer)
   TextInverse txtSD(Index)
   CloseIme
End Sub

Private Sub ClearField()
   lblName.Caption = Empty 'Add By Sindy 2021/12/20
   For Each oText In txtSD
      oText.Text = Empty
   Next
   For Each oLabel In lblDsp
      oLabel.Caption = Empty
   Next
   For intI = 1 To TF_SD
      m_FieldList(intI).fiOldData = Empty
      m_FieldList(intI).fiNewData = Empty
   Next
   textCUID = ""
   m_bConfirmCheck = False
   If m_EditMode = 1 Then
      txtSD(4) = "Y"
      txtSD(16) = "Y"
   End If
   ClearSR
End Sub

' 將資料庫中的資料更新到所有欄位中
Private Sub UpdateCtrlData(ByRef p_Rst As ADODB.Recordset)
   Dim CUID(1 To 6) As String
   With p_Rst
   If .RecordCount > 0 Then
      For Each oText In txtSD
         idx = oText.Index
         '2010/11/11 add by sonia
         If idx = 46 Then
            If "" & .Fields(m_FieldList(idx).fiName) <> "" Then
               m_FieldList(idx).fiOldData = "" & Val(.Fields(m_FieldList(idx).fiName)) - 191100
            Else
               m_FieldList(idx).fiOldData = "" & .Fields(m_FieldList(idx).fiName)
            End If
         Else
         '2010/11/11 end
            m_FieldList(idx).fiOldData = "" & .Fields(m_FieldList(idx).fiName)
         End If
         m_FieldList(idx).fiNewData = m_FieldList(idx).fiOldData
         oText.Text = m_FieldList(idx).fiOldData
      Next
      
      CUID(1) = "" & .Fields("sd37")
      CUID(2) = "" & .Fields("sd38")
      CUID(3) = "" & .Fields("sd39")
      CUID(4) = "" & .Fields("sd40")
      CUID(5) = "" & .Fields("sd41")
      CUID(6) = "" & .Fields("sd42")
      SetRefData txtSD(1)
      If txtSD(19) <> "" Then lblDsp(7) = CompNameQuery(txtSD(19))
      If txtSD(28) <> "" Then lblDsp(8) = CompNameQuery(txtSD(28))
      txtSD(11).Tag = txtSD(11)
      txtSD(12).Tag = txtSD(12)
      txtSD(13).Tag = txtSD(13)
      txtSD(20).Tag = txtSD(20)
      txtSD(21).Tag = txtSD(21)
      txtSD(23).Tag = txtSD(23)
   End If
   End With
   UpdateCUID CUID, textCUID
   txtSD(1).Tag = txtSD(1)
   
   'Add by Morgan 2009/6/17 待更新資料
   For Each oText In txtSL
      oText.Locked = True
      oText.BackColor = &HE0E0E0
      oText.Visible = False
      lblSL(oText.Index).Visible = False
   Next
   '薪資
   strSql = "select * from salaryupdate,salarylog where su01='" & txtSD(1) & "'" & _
      " and su05 is null and su03='1' and sl01(+)=su01 and sl02(+)=su02"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      With RsTemp
         SSTab1.TabVisible(1) = True
         For Each oText In txtSL
            Select Case oText.Index
               'Modified by Morgan 2012/6/20 SL05,SL06 婚喪扣款除外
               'Modify By Sindy 2020/6/22 + 39
               Case 3, 4, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 33, 34, 39
                  lblSL(oText.Index).Visible = True
                  oText.Visible = True
                  oText.Text = "" & RsTemp.Fields("SL" & Format(oText.Index, "00"))
            End Select
         Next
      End With
   Else
      If SSTab1.Tab = 1 Then SSTab1.Tab = 0
      SSTab1.TabVisible(1) = False
   End If
   
   '勞退
   'Add By Sindy 2020/6/22 新制才需要顯示
   If txtSD(16) = "Y" Then
   '2020/6/22 END
      strSql = "select * from salaryupdate,salarylog where su01='" & txtSD(1) & "'" & _
         " and su05 is null and su03='2' and sl01(+)=su01 and sl02(+)=su02"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strSql)
      If intI = 1 Then
      With RsTemp
         SSTab1.TabVisible(1) = True
         For Each oText In txtSL
            Select Case oText.Index
               Case 18, 26, 33, 34
                  lblSL(oText.Index).Visible = True
                  oText.Visible = True
                  oText.Text = "" & RsTemp.Fields("SL" & Format(oText.Index, "00"))
            End Select
         Next
         '勞退投保薪資
         'Added by Morgan 2020/9/25
         Label1(40).Visible = True
         txtSD(43).Visible = True
         Label1(41).Visible = True
         txtSD(44).Visible = True
         'end 2020/9/25
         lblSL(36).Visible = True
         txtSL(36).Visible = True
         'Modified by Morgan 2020/9/25 +SL39
         'txtSL(36).Text = Val("" & RsTemp.Fields("SL11")) + Val("" & RsTemp.Fields("SL12")) + Val("" & RsTemp.Fields("SL14"))
         txtSL(36).Text = Val("" & RsTemp.Fields("SL11")) + Val("" & RsTemp.Fields("SL12")) + Val("" & RsTemp.Fields("SL14")) + Val("" & RsTemp.Fields("SL39"))
         'end 2020/9/25
         lblSL(37).Visible = True
         txtSL(37).Visible = True
         txtSL(37).Text = Val("" & RsTemp.Fields("SL19")) + Val("" & RsTemp.Fields("SL20")) + Val("" & RsTemp.Fields("SL22"))
      End With
      End If
   'Added by Morgan 2020/9/25
   Else
      Label1(40).Visible = False
      txtSD(43).Visible = False
      Label1(41).Visible = False
      txtSD(44).Visible = False
      lblSL(36).Visible = False
      txtSL(36).Visible = False
      lblSL(37).Visible = False
      txtSL(37).Visible = False
   'end 2020/9/25
   End If
   
   '勞健保
   strSql = "select * from salaryupdate,salarylog where su01='" & txtSD(1) & "'" & _
      " and su05 is null and su03='3' and sl01(+)=su01 and sl02(+)=su02"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
   With RsTemp
      SSTab1.TabVisible(1) = True
      For Each oText In txtSL
         Select Case oText.Index
            Case 7, 8, 9, 10, 33, 34
               lblSL(oText.Index).Visible = True
               oText.Visible = True
               oText.Text = "" & RsTemp.Fields("SL" & Format(oText.Index, "00"))
         End Select
      Next
      '勞健保投保薪資
      'Modified by Morgan 2020/9/25 +SL39
      'txtSL(38).Text = Val("" & RsTemp.Fields("SL11")) + Val("" & RsTemp.Fields("SL12")) + Val("" & RsTemp.Fields("SL14"))
      txtSL(38).Text = Val("" & RsTemp.Fields("SL11")) + Val("" & RsTemp.Fields("SL12")) + Val("" & RsTemp.Fields("SL14")) + Val("" & RsTemp.Fields("SL39"))
      'end 2020/9/25
   End With
   End If
   
   If txtSL(33).Text <> "" Then
      lblDsp(9) = CompNameQuery(txtSL(33))
   End If
   If txtSL(34).Text <> "" Then
      lblDsp(10) = CompNameQuery(txtSL(34))
   End If
   'END 2009/6/17
   
   'Add by Morgan 2009/6/26
   '讀取勞健保補助類別
   strSql = "select st50,st56 from staff where st01='" & txtSD(1) & "'"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      With RsTemp
      cboST50.Tag = "" & .Fields("st50")
      SelCombo cboST50, "" & .Fields("st50")
      cboST56.Tag = "" & .Fields("st56")
      SelCombo cboST56, "" & .Fields("st56")
      End With
   End If
        
   '眷屬
   strSql = "SELECT sr03||' '||decode(sr03,'1','父親','2','母親','3','配偶','4','子女','其他')" & _
      ",Sr04,DECODE(SR12,NULL,NULL,'刪') STATUS,sr05||' '||decode(sr05,'M','男','F','女','不詳')" & _
      ",sqldatet(sr06),sr07,sr08,sr13,sr09,sr10,sr11,sqldatet(sr12),sr02,hl05" & _
      " FROM staff_relation,(select hl02,hl05 from HIrelationlog a" & _
      " where hl01='" & txtSD(1) & "'" & _
      " and (hl02,hl03) in (select b.hl02,max(b.hl03) from hirelationlog b where b.hl01=a.hl01" & _
      " and b.hl03<=to_char(sysdate,'yyyymmdd') group by hl02)" & _
      ") X WHERE SR01 = '" & txtSD(1) & "' and hl02(+)=SR02 order by sr02 "
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
   Set GRD1.Recordset = RsTemp
   SetGrd
   'end 2009/6/26

End Sub

Private Sub SetCtrlReadOnly(ByVal bLocked As Boolean)
   For Each oText In txtSD
      Select Case oText.Index
      'Modified by Morgan 2012/6/20 +婚喪扣款sd09,sd10可修改
      'Modified by Morgan 2013/1/21 +sd11
      'Modified by Morgan 2013/1/24 +sd49
      'Modified by Morgan 2015/12/10 +sd50
      'Modified by Morgan 2023/6/29 +sd48
      'Modified by Morgan 2025/7/29 -sd09,sd10 114/7/28起廢止婚喪互助辦法
      Case 1, 3, 4, 5, 6, 7, 11, 46, 49, 50, 48
         oText.Locked = bLocked
      Case Else
         oText.Locked = True
         oText.BackColor = &HE0E0E0
      End Select
   Next
   
   cboST50.Enabled = Not bLocked
   cboST56.Enabled = Not bLocked
End Sub

' 更新 Create 及 Update 的人
Private Sub UpdateCUID(ByRef p_CUID() As String, ByRef oText As Object)
   Dim strTemp As String
   Dim strCName As String
   Dim strCDate As String
   Dim strCTime As String
   Dim strUName As String
   Dim strUDate As String
   Dim strUTime As String
   
   If p_CUID(1) <> "" Then
      strCName = GetStaffName(p_CUID(1), True)
   End If
   If p_CUID(2) <> "" Then
      strCDate = ChangeWStringToTDateString(p_CUID(2))
   End If
   
   If p_CUID(3) <> "" Then
      strCTime = Format(p_CUID(3), "##:##")
   End If
   
   If p_CUID(4) <> "" Then
      strUName = GetStaffName(p_CUID(4), True)
   End If
   If p_CUID(5) <> "" Then
      strUDate = ChangeWStringToTDateString(p_CUID(5))
   End If
   
   If p_CUID(6) <> "" Then
      strUTime = Format(p_CUID(6), "##:##")
   End If
      
   ' 設定CUID中的文字
   oText = "CREATE : " & strCName & " " & _
              " " & strCDate & " " & _
              " " & strCTime & String(10, " ") & _
              "UPDATE : " & strUName & " " & _
              " " & strUDate & " " & _
              " " & strUTime
              
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      ' 新增
      Case vbKeyF2:
         If m_bInsert Then
            If m_EditMode = 0 Then
               OnAction KeyCode
               KeyCode = 0
            End If
         End If
      ' 修改
      Case vbKeyF3:
         If m_bUpdate Then
            If m_EditMode = 0 Then
               OnAction KeyCode
               KeyCode = 0
            End If
         End If
      ' 查詢
      Case vbKeyF4:
         If m_bQuery Then
            If m_EditMode = 0 Then
               OnAction KeyCode
               KeyCode = 0
            End If
         End If
      ' 刪除
      Case vbKeyF5:
         If m_bDelete Then
            If m_EditMode = 0 Then
               OnAction KeyCode
               KeyCode = 0
            End If
         End If
      ' 第一筆, 上一筆, 下一筆, 最後一筆
      Case vbKeyHome, vbKeyPageUp, vbKeyPageDown, vbKeyEnd:
         If m_bQuery Then
            If m_EditMode = 0 Then
               OnAction KeyCode
               KeyCode = 0
            End If
         End If
      Case vbKeyF9, vbKeyF10:
         If m_EditMode <> 0 Then
            OnAction KeyCode
            KeyCode = 0
         End If
         
      Case vbKeyEscape:
         If TypeName(Me.ActiveControl) <> "ComboBox" Then
            If m_EditMode <> 0 Then
               OnAction vbKeyF10
            Else
               OnAction KeyCode
            End If
         End If
         
      Case vbKeyReturn
         '做完取消，不然 form 內其他物件有寫 keycode 或是 keyascii 事件的話，也會做到
         KeyCode = 0
         If m_EditMode <> 0 Then
            OnAction vbKeyF9
         End If

   End Select
End Sub

' 執行指令
Public Sub OnAction(ByVal KeyCode As Integer)
   Dim bCancel As Boolean
   
   Select Case KeyCode
      Case vbKeyF2 ' 新增
         m_EditMode = 1
         ClearField
         UpdateToolbarState
         SetInputEntry
         
      Case vbKeyF3 ' 修改
         m_EditMode = 2
         UpdateToolbarState
         SetInputEntry

      Case vbKeyF5 ' 刪除
         If MsgBox("是否要刪除此筆資料?", vbYesNo + vbCritical + vbDefaultButton2, "詢問") = vbYes Then
            m_EditMode = 3
            If OnWork = False Then
                Exit Sub
            End If
            UpdateToolbarState
         End If
         
      Case vbKeyF4 ' 查詢
         m_EditMode = 4
         ClearField
         SetInputEntry
         UpdateToolbarState
         
      Case vbKeyHome ' 第一筆
         ShowRecord -2
      Case vbKeyPageUp ' 前一筆
         ShowRecord -1
      Case vbKeyPageDown ' 後一筆
         ShowRecord 1
      Case vbKeyEnd ' 最後一筆
         ShowRecord 2
      Case vbKeyF9 ' 確定
         If OnWork = False Then
            Exit Sub
         End If
         SetInputEntry
         UpdateToolbarState
         
      Case vbKeyF10 ' 取消
         bCancel = False
         Select Case m_EditMode
            Case 1, 2:
               If MsgBox("你並未存檔, 確定離開嗎?", vbYesNo + vbQuestion + vbDefaultButton2, "詢問") = vbYes Then
                  bCancel = True
               End If
            Case Else
               bCancel = True
         End Select
         If bCancel = True Then
            txtSD(1) = txtSD(1).Tag
            m_EditMode = 0
            SetInputEntry
            ShowRecord
            UpdateToolbarState
         End If
         
      Case vbKeyEscape ' 離開
         Unload Me
   End Select
End Sub

'依照權限設定其工具列的按紐狀態
Private Sub UpdateToolbarState()
   Select Case m_EditMode
      Case 0 ' 無任何動作
         If m_bInsert Then
            TBar1.Buttons(1).Enabled = True
         Else
            TBar1.Buttons(1).Enabled = False
         End If
         If m_bUpdate And txtSD(1) <> "" Then
            TBar1.Buttons(2).Enabled = True
         Else
            TBar1.Buttons(2).Enabled = False
         End If
         If m_bDelete And txtSD(1) <> "" Then
            TBar1.Buttons(3).Enabled = True
         Else
            TBar1.Buttons(3).Enabled = False
         End If
         If m_bQuery Then
            TBar1.Buttons(4).Enabled = True
         Else
            TBar1.Buttons(4).Enabled = False
         End If
         If m_bQuery And txtSD(1) <> "" Then
            TBar1.Buttons(6).Enabled = True
            TBar1.Buttons(7).Enabled = True
            TBar1.Buttons(8).Enabled = True
            TBar1.Buttons(9).Enabled = True
         Else
            TBar1.Buttons(6).Enabled = False
            TBar1.Buttons(7).Enabled = False
            TBar1.Buttons(8).Enabled = False
            TBar1.Buttons(9).Enabled = False
         End If
         TBar1.Buttons(11).Enabled = False
         TBar1.Buttons(12).Enabled = False
         TBar1.Buttons(14).Enabled = True
      
      Case 1, 2, 3, 4 '維護
         TBar1.Buttons(1).Enabled = False
         TBar1.Buttons(2).Enabled = False
         TBar1.Buttons(3).Enabled = False
         TBar1.Buttons(4).Enabled = False
         TBar1.Buttons(6).Enabled = False
         TBar1.Buttons(7).Enabled = False
         TBar1.Buttons(8).Enabled = False
         TBar1.Buttons(9).Enabled = False
         TBar1.Buttons(11).Enabled = True
         TBar1.Buttons(12).Enabled = True
         TBar1.Buttons(14).Enabled = False
   End Select
End Sub

' 開始輸入資料
Private Sub SetInputEntry()
   Select Case m_EditMode
      Case 1
         SetCtrlReadOnly False
         If Me.Visible = True Then
            txtSD(1).SetFocus
         End If
      Case 2
         SetCtrlReadOnly False
         txtSD(1).Locked = True
         If Me.Visible = True Then
            txtSD(3).SetFocus
         End If
         'Added by Morgan 2013/1/24
         If Left(txtSD(1), 1) <> "F" Then
            txtSD(49).Locked = True
         End If
         'end 2013/1/24
      Case 4
         SetCtrlReadOnly True
         txtSD(1).Locked = False
         If Me.Visible = True Then
            txtSD(1).SetFocus
         End If
      Case Else
         SetCtrlReadOnly True
         If Me.Visible = True Then
            txtSD(1).SetFocus
         End If
   End Select
   PUB_ChangeCaption Me, m_EditMode
End Sub

Private Function OnWork() As Boolean
   Select Case m_EditMode
      Case 1: '新增
         '重新檢查欄位有效性
         If TxtValidate() = True Then
            UpdateFieldNewData
            If AddRecord = True Then
               OnWork = True
               m_EditMode = 0
               ShowRecord
            End If
         End If
         
      Case 2: '修改
         '重新檢查欄位有效性
         If TxtValidate() = True Then
            UpdateFieldNewData
            If ModRecord = True Then
               OnWork = True
               m_EditMode = 0
               ShowRecord
            End If
         End If
         
      Case 3: '刪除
         If DelRecord = True Then
            OnWork = True
            m_EditMode = 0
            ShowRecord 2
         End If
      
      Case 4: '查詢
         If TxtValidate() = True Then
            If ShowRecord = True Then
               OnWork = True
               m_EditMode = 0
            Else
               txtSD(1).SetFocus
               txtSD_GotFocus 1
            End If
         End If
         
   End Select
End Function


Private Function TxtValidate() As Boolean
   
   Dim bCancel As Boolean
   
   m_bConfirmCheck = True
   
   If txtSD(1) = "" Then
      ShowMsg "請輸入員工代號 !"
      txtSD(1).SetFocus
      txtSD_GotFocus 1
      GoTo EscPoint
   End If
      
   '維護
   If m_EditMode = 1 Or m_EditMode = 2 Then
      'Modified by Morgan 2012/9/12 F編號也要設定否則翻譯費會沒有公司別 Ex.F5644,F5645
      If txtSD(19) = "" Then
         MsgBox "第一家的公司別不可空白!"
         txtSD(19).SetFocus
         GoTo EscPoint
      End If
      
      If Left(txtSD(1), 1) <> "F" Then   '2009/12/10 ADD BY SONIA F編號不檢查
         If txtSD(20) = "" Then
            MsgBox "第一家的基本薪資不可空白!"
            txtSD(20).SetFocus
            GoTo EscPoint
         End If
         If txtSD(28) = "" And txtSD(29) & txtSD(30) & txtSD(31) & txtSD(32) & txtSD(33) & txtSD(34) & txtSD(35) & txtSD(36) <> "" Then
            MsgBox "有設定第二家的所得資料，公司別不可空白!"
            txtSD(28).SetFocus
            GoTo EscPoint
         End If
         
         If Trim(txtSD(16)) = "" And Val(txtSD(17)) > 0 Then
            MsgBox "適用勞退新制才能輸入自提費率!"
            txtSD(17).SetFocus
            GoTo EscPoint
         End If
         
         If Trim(txtSD(16)) = "" And (Trim(txtSD(27)) <> "" Or Trim(txtSD(36)) <> "") Then
            MsgBox "適用勞退新制才能輸入退休金投保薪資!"
            If Trim(txtSD(27)) <> "" Then
               txtSD(27).SetFocus
            Else
               txtSD(36).SetFocus
            End If
            GoTo EscPoint
         End If
         
      'Added by Morgan 2013/1/24
      '外翻要檢查兼職薪資所得是否扣補充保費的設定
      Else
         If m_FieldList(49).fiOldData <> txtSD(49) Then
            If PUB_ChkNotPaidNhiFee(txtSD(1), "50", strExc(1)) = True Then
               '改要扣(要看金額)或不扣但有補充保費(一定是錯的)
               If txtSD(49) = "" Or (txtSD(49) = "N" And Val(strExc(1)) > 0) Then
                  If MsgBox("此外翻員人尚有未申報之兼職薪資所得補充保費，變更設定將可能會造成原補充保費不正確，是否要確定要繼續？", vbYesNo + vbDefaultButton2 + vbQuestion) = vbNo Then
                     GoTo EscPoint
                  End If
               End If
            End If
         End If
      'end 2013/1/24
      
      End If    '2009/12/10 END
   End If
   
   For Each oText In txtSD
      If oText.Locked = False And oText.Visible = True And oText.Enabled = True Then
         idx = oText.Index
         bCancel = False
         txtSD_Validate idx, bCancel
         If bCancel = True Then
            txtSD(idx).SetFocus
            txtSD_GotFocus idx
            GoTo EscPoint
         End If
      End If
   Next
   
   TxtValidate = True
   
EscPoint:
   m_bConfirmCheck = False
    
End Function

Private Function AddRecord() As Boolean
   Dim stCols As String, stValues As String, stSQL As String
On Error GoTo ErrHand
   cnnConnection.BeginTrans
   '畫面有的欄位才更新
   stCols = "": stValues = ""
   For Each oText In txtSD
      idx = oText.Index
      If m_FieldList(idx).fiNewData <> "" Then
         stCols = stCols & "," & m_FieldList(idx).fiName
         '文字
         If m_FieldList(idx).fiType = 0 Then
            stValues = stValues & "," & CNULL(ChgSQL(m_FieldList(idx).fiNewData))
         '數字
         Else
            stValues = stValues & "," & CNULL(m_FieldList(idx).fiNewData, True)
         End If
      End If
   Next
   stCols = Mid(stCols, 2)
   stValues = Mid(stValues, 2)
   stSQL = "INSERT INTO SalaryData (" & stCols & ") Values (" & stValues & ")"
   
   Pub_SeekTbLog stSQL
   cnnConnection.Execute stSQL, intI
   
   
   '非翻譯人員要新增薪資異動資料
   If Left(txtSD(1), 1) <> "F" Then
      'Modify By Sindy 2020/6/22 + SL39
      stSQL = "insert into SALARYLOG(SL01,SL02,SL03,SL04" & _
         ",SL05,SL06,SL07,SL08,SL09,SL10" & _
         ",SL11,SL12,SL13,SL14,SL15,SL16,SL17,SL18" & _
         ",SL19,SL20,SL21,SL22,SL23,SL24,SL25,SL26,SL39)" & _
         " SELECT SD01," & m_ST13 & ",SD02,SD08" & _
         ",SD09,SD10,SD12,SD13,SD14,SD15" & _
         ",SD20,SD21,SD22,SD23,SD24,SD25,SD26,SD27" & _
         ",SD29,SD30,SD31,SD32,SD33,SD34,SD35,SD36,SD52" & _
         " FROM SalaryData WHERE SD01='" & txtSD(1) & "'"
         
      Pub_SeekTbLog stSQL
      cnnConnection.Execute stSQL, intI
   End If
   
   cnnConnection.CommitTrans
   
   AddRecord = True
   Exit Function
   
ErrHand:
    cnnConnection.RollbackTrans
    MsgBox Err.Description, vbCritical
    
End Function

Private Function ModRecord() As Boolean
   Dim stSQL As String, stSet As String, stCols As String, stValues As String
   Dim bDifference As Boolean, bAddNew As Boolean
   Dim stST50 As String, stST56 As String
   
On Error GoTo ErrHand
   cnnConnection.BeginTrans
   
   stSQL = "begin user_data.user_enabled:=1; UPDATE SalaryData SET "
   stSet = ""
   For Each oText In txtSD
      idx = oText.Index
      If m_FieldList(idx).fiNewData <> m_FieldList(idx).fiOldData Then
         bDifference = True
         '文字
         If m_FieldList(idx).fiType = 0 Then
            stSet = stSet & "," & m_FieldList(idx).fiName & "=" & CNULL(ChgSQL(m_FieldList(idx).fiNewData))
         '數字
         Else
            stSet = stSet & "," & m_FieldList(idx).fiName & "=" & CNULL(m_FieldList(idx).fiNewData, True)
         End If
      End If
   Next
   
   If bDifference = True Then
      stSet = Mid(stSet, 2)
      stSQL = stSQL & stSet & " where sd01='" & txtSD(1) & "'; end; "
      Pub_SeekTbLog stSQL
      
      cnnConnection.Execute stSQL, intI
   End If
   
   'Add by Morgan 2011/5/23
   If cboST50.ListIndex > 0 Then
      stST50 = Left(cboST50, 2)
   Else
      stST50 = ""
   End If
   If cboST56.ListIndex > 0 Then
      stST56 = Left(cboST56, 2)
   Else
      stST56 = ""
   End If
   
   If stST50 <> cboST50.Tag Or stST56 <> cboST56.Tag Then
      stSQL = "Update Staff set ST50='" & stST50 & "',ST56='" & stST56 & "' where st01='" & txtSD(1) & "'"
      Pub_SeekTbLog stSQL
      cnnConnection.Execute stSQL, intI
   End If
   'end 2011/5/23
   
   cnnConnection.CommitTrans
   
   ModRecord = True
   Exit Function
   
ErrHand:
    cnnConnection.RollbackTrans
    MsgBox Err.Description, vbCritical

End Function

Private Sub Tbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
   Select Case Button.Index
      ' 新增
      Case 1: OnAction vbKeyF2
      ' 修改
      Case 2: OnAction vbKeyF3
      ' 刪除
      Case 3: OnAction vbKeyF5
      ' 查詢
      Case 4: OnAction vbKeyF4
      ' 第一筆
      Case 6: OnAction vbKeyHome
      ' 前一筆
      Case 7: OnAction vbKeyPageUp
      ' 後一筆
      Case 8: OnAction vbKeyPageDown
      ' 最後一筆
      Case 9: OnAction vbKeyEnd
      ' 確定
      Case 11: OnAction vbKeyF9
      ' 取消
      Case 12: OnAction vbKeyF10
      ' 離開
      Case 14: OnAction vbKeyEscape
   End Select
End Sub

Private Sub UpdateFieldNewData()
   For Each oText In txtSD
      idx = oText.Index
      '2010/11/11 add by sonia
      If idx = 46 Then
         m_FieldList(idx).fiNewData = Val(oText.Text) + 191100
      Else
      '2010/11/11 end
         m_FieldList(idx).fiNewData = oText.Text
      End If
   Next
   
End Sub

Private Sub txtSD_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   Select Case Index
      Case 1, 19, 28
         '不控制
      Case 2
         If KeyAscii <> 8 And KeyAscii <> Asc("R") And KeyAscii <> Asc("T") And KeyAscii <> Asc("S") And KeyAscii <> Asc("P") And KeyAscii <> Asc("F") Then
            KeyAscii = 0
            Beep
         End If
      
      'Modified by Morgan 2023/6/29 +48
      Case 3, 4, 16, 50, 48
         If KeyAscii <> 8 And KeyAscii <> Asc("Y") Then
            KeyAscii = 0
            Beep
         End If
         
      Case 5
         If KeyAscii <> 8 And Not (KeyAscii >= Asc("1") And KeyAscii <= Asc("7")) Then
            KeyAscii = 0
            Beep
         End If
         
      Case 11
         'Modified by Morgan 2013/1/21 改為勞健保是否以合夥人身分投保,放Y
         'If KeyAscii <> 8 And KeyAscii <> Asc("N") Then
         If KeyAscii <> 8 And KeyAscii <> Asc("Y") Then
            KeyAscii = 0
            Beep
         End If
      
      'Added by Morgan 2013/1/24
      Case 49
         If KeyAscii <> 8 And KeyAscii <> Asc("N") Then
            KeyAscii = 0
            Beep
         End If
      
      Case Else
         If KeyAscii <> 8 And Not IsNumeric(Chr(KeyAscii)) Then
            KeyAscii = 0
            Beep
         End If
   End Select
End Sub

Private Sub txtSD_Validate(Index As Integer, Cancel As Boolean)
Dim stDate As String  '2010/11/11 add by sonia
   
   If m_EditMode = 1 Or m_EditMode = 2 Then
   
      Select Case Index
         Case 1
            If txtSD(Index) <> "" Then
               If ChkStaffID(txtSD(Index)) = True Then
                  Cancel = True
               End If
               If SetRefData(txtSD(Index)) = False Then
                  MsgBox "員工代碼輸入錯誤！"
                  Cancel = True
               End If
               If m_EditMode = 1 Then
                  If CheckExists(txtSD(Index)) Then
                     MsgBox "薪資資料已存在！"
                     Cancel = True
                  End If
                  '非翻譯人員不可無到職日
                  If m_ST13 = "" And Left(txtSD(Index), 1) <> "F" Then
                     MsgBox "該員工人事資料無【到職日】，不可新增！"
                     Cancel = True
                  End If
               End If
            End If
            
         Case 19
            If txtSD(Index) <> "" Then
               lblDsp(7) = CompNameQuery(txtSD(Index))
               If lblDsp(7) = "" Then
                  ShowMsg "公司別錯誤 !"
                  Cancel = True
               End If
            End If
            
         Case 28
            If txtSD(Index) <> "" Then
               lblDsp(8) = CompNameQuery(txtSD(Index))
               If lblDsp(8) = "" Then
                  ShowMsg "公司別錯誤 !"
                  Cancel = True
               End If
            End If
         
         Case 8
            If Val(txtSD(Index)) > 99 Then
               Cancel = True
               MsgBox "所得稅率不可大於 99！"
            End If
            
         Case 17
            If Val(txtSD(Index)) > 6 Then
               Cancel = True
               MsgBox "勞退自提費率不可大於 6！"
            End If
         
         '2010/11/11 add by sonia 加帳號通知銀行薪資年月
         Case 46
            stDate = CompDate("1", 1, strSrvDate(1))
            If txtSD(5) = "2" And txtSD(6) <> "" Then
               If txtSD(Index) = "" Then
                  Cancel = True
                  MsgBox "帳號通知銀行薪資年月不可空白！"
               ElseIf Val(Right(txtSD(Index), 2)) > 12 Then
                  Cancel = True
                  MsgBox "通知銀行薪資年月之月份錯誤！"
               ElseIf Val(txtSD(Index)) > Val(Left(stDate, 6)) - 191100 Then
                  Cancel = True
                  MsgBox "帳號通知銀行薪資年月錯誤！"
               End If
            End If
         '2010/11/11 end
         
      End Select
      
      If Cancel = True Then TextInverse txtSD(Index)
      
      '若是案確定的檢查時略過
      If Cancel = False And m_bConfirmCheck = False Then
         Select Case Index
            Case 1
               '新增時預設婚喪戶助金額
               If m_EditMode = 1 Then
                  If Left(txtSD(Index), 1) = "F" Then
                     'txtSD(11) = "N" 'Removed by Morgan 2013/1/21
                     txtSD(16) = ""
                  Else
                     'Removed by Morgan 2025/7/29 114/7/28起廢止婚喪互助辦法
                     'If PUB_GetHelpFee(txtSD(1), strExc(1)) = True Then
                     '   txtSD(9) = strExc(1) '婚事互助
                     '   txtSD(10) = txtSD(9) '喪事互助
                     'End If
                     'end 2025/7/29
                     
                  End If
               End If
            
            'Added by Morgan 2013/1/21
            Case 11 '勞健保是否以合夥人身分投保
               If txtSD(Index) <> txtSD(Index).Tag Then
                  SetInsureFee
               End If
               txtSD(Index).Tag = txtSD(Index)
            
            'Modified by Morgan 2023/6/29 +sd48 勞保是否無就保
            Case 12, 48 '勞保投保薪資,
               If txtSD(Index) <> txtSD(Index).Tag Then
                  SetInsureFee 1
               End If
               txtSD(Index).Tag = txtSD(Index)
               
            Case 13 '健保投保薪資
               If txtSD(Index) <> txtSD(Index).Tag Then
                  SetInsureFee 2
               End If
               txtSD(Index).Tag = txtSD(Index)
               
            Case 16 '適用勞退新制
               If txtSD(Index) = "" Then
                  txtSD(17) = ""
                  txtSD(27) = ""
                  txtSD(36) = ""
               End If
               
            Case 20, 21, 23
               'Modified by Morgan 2013/1/21 sd11 改為 勞健保是否以合夥人身分投保
               'If txtSD(11) = "" And (txtSD(12) = "" Or txtSD(13) = "") Then
               If Left(txtSD(1), 1) <> "F" And (txtSD(12) = "" Or txtSD(13) = "") Then
               'end 2013/1/21
                  strExc(1) = Val(txtSD(20)) + Val(txtSD(21)) + Val(txtSD(23))
                  strExc(2) = Val(txtSD(20).Tag) + Val(txtSD(21).Tag) + Val(txtSD(23).Tag)
                  If strExc(1) <> strExc(2) Then
                     If txtSD(12) = "" Then
                        SetInsureFee 1
                     End If
                     If txtSD(13) = "" Then
                        SetInsureFee 2
                     End If
                  End If
               End If
               txtSD(Index).Tag = txtSD(Index)
               
         End Select
      End If
   End If
End Sub
'設定勞健保費
'iOption:0=全部,1=勞保費,2=健保費
Private Sub SetInsureFee(Optional iOption As Integer)
   Dim lngInsureSalary As Long '投保薪資
   Dim lngInsureBase As Long '投保等級
   Dim dblInsureRate As Double '投保費率
   Dim dblFreeRate As Double '補助比率
   Dim dblInsureRate2 As Double '就業保險費率
   Dim intShareRate As Integer '負擔比例
   
   'Modified by Morgan 2013/1/21 sd11 改為 勞健保是否以合夥人身分投保
   'If txtSD(11) = "" Then '適用一般勞健保費率
      
      If iOption = 0 Or iOption = 1 Then
         '勞保投保薪資
         'Modified by Morgan 2016/3/31 特殊勞保投保薪資會輸0(63001,已退休)
         'lngInsureSalary = Val(txtSD(12))
         'If lngInsureSalary = 0 Then
         If txtSD(12) <> "" Then
            lngInsureSalary = Val(txtSD(12))
         Else
         'end 2016/3/31
            lngInsureSalary = Val(txtSD(20)) + Val(txtSD(21)) + Val(txtSD(23))
         End If
         '勞保等級
         lngInsureBase = PUB_GetInsureBase(lngInsureSalary, "L")
         'Modify by Morgan 2009/6/29
         '98/5/1 起外國人也有失業給付,費率改與本國人同,只有雇主(所長)沒有(未來加65歲以上)
         'If Left(lblDsp(4), 1) = "F" Then
         'Modify by Morgan 2010/10/26 勞保費率及就業保險費率需個別計算(四捨五入)
         'Modified by Morgan 2013/1/21 改判斷 sd11
         'If CheckExceptLiRate = True Then
         'Modified by Morgan 2023/6/29 +判斷 sd48勞保是否無就保
         If txtSD(11) = "Y" Or txtSD(48) = "Y" Then
         'end 2013/1/21
            dblInsureRate = Val(iR(1))
            dblInsureRate2 = 0 'Add
         
         'Added by Morgan 2015/1/28
         '超過65歲也沒有就保
         ElseIf PUB_ChkOver65(txtSD(1)) Then
            dblInsureRate = Val(iR(1))
            dblInsureRate2 = 0
         'end 2015/1/28
         
         Else
            dblInsureRate = Val(iR(1))
            dblInsureRate2 = Val(iR(2)) 'Add
         End If
         '勞保費=勞保等級*勞保費率*勞保個人負擔比例
         'txtSD(14) = Round(lngInsureBase * dblInsureRate / 100 * Val(iR(3)) / 100, 0)
         txtSD(14) = Round(lngInsureBase * dblInsureRate / 100 * Val(iR(3)) / 100, 0) + Round(lngInsureBase * dblInsureRate2 / 100 * Val(iR(3)) / 100, 0)
         'end 2010/10/27
      End If
      
      If iOption = 0 Or iOption = 2 Then
         '健保投保薪資
         'Modified by Morgan 2016/3/31 與勞保檢查一致
         'lngInsureSalary = Val(txtSD(13))
         'If lngInsureSalary = 0 Then
         If txtSD(13) <> "" Then
            lngInsureSalary = Val(txtSD(13))
         Else
         'end 2016/3/31
            lngInsureSalary = Val(txtSD(20)) + Val(txtSD(21)) + Val(txtSD(23))
         End If
         dblInsureRate = Val(iR(6))
         
         'Added by Morgan 2013/1/21
         '以合夥人身分投保 100% 個人負擔
         If txtSD(11) = "Y" Then
            intShareRate = 100
         Else
            intShareRate = Val(iR(7))
         End If
         'end 2013/1/21
         
         '健保費=健保等級*健保費率*健保個人負擔比例
         'Modify by Morgan 2010/4/15 健保費調整改用共用函數
         'lngInsureBase = GetInsureBase(lngInsureSalary, "H") '健保等級
         'txtSD(15) = Round(lngInsureBase * dblInsureRate / 100 * Val(IR(7)) / 100, 0)
         lngInsureBase = PUB_GetInsureBase(lngInsureSalary, "H", dblFreeRate) '健保等級
         'Modified by Morgan 2013/1/21
         'txtSD(15) = PUB_GetHIFee(lngInsureBase, dblInsureRate, Val(IR(7)), dblFreeRate)
         txtSD(15) = PUB_GetHIFee(lngInsureBase, dblInsureRate, intShareRate, dblFreeRate)
         txtSD(47) = lngInsureBase
      'End If
      
   End If
End Sub

Private Function SetRefData(stUserNo As String) As Boolean
   
   'Modified by Morgan 2024/1/31 新部門
   strExc(0) = "select * from staff,acc090,acc090new,allcode" & _
      ",(select SR01,count(*) FMC from Staff_Relation where sr01='" & stUserNo & "' and SR08 is null group by sr01) x" & _
      " where st01='" & stUserNo & "' and a0901(+)=st03 and a0921(+)=st93 and ac01(+)='01' and ac02(+)=st20 and sr01(+)=st01"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      With RsTemp
      '名稱
      lblName.Caption = "" & .Fields("st02") 'Modify By Sindy 2021/12/20
      '部門
      'Modified by Morgan 2024/1/31 優先顯示新部門
      If Not IsNull(.Fields("a0922")) Then
         lblDsp(2) = "" & .Fields("a0922")
      Else
         lblDsp(2) = "" & .Fields("a0902")
      End If
      '職稱
      lblDsp(3) = "" & .Fields("ac03")
      '國籍
      lblDsp(4) = "" & .Fields("st24")
      If lblDsp(4) = "L" Then
         lblDsp(4) = lblDsp(4) & " 本國"
      ElseIf lblDsp(4) = "F" Then
         lblDsp(4) = lblDsp(4) & " 外國"
      End If
      '所別
      Select Case "" & .Fields("st06")
         Case 1: lblDsp(5) = "北所"
         Case 2: lblDsp(5) = "中所"
         Case 3: lblDsp(5) = "南所"
         Case 4: lblDsp(5) = "高所"
         Case Else: lblDsp(5) = "其他"
      End Select
      '眷口數
      lblDsp(6) = Val("" & .Fields("FMC"))
      m_ST13 = "" & .Fields("st13")
      End With
      SetRefData = True
   End If
   
End Function

' 刪除記錄
Private Function DelRecord() As Boolean
   Dim stSQL As String
   
On Error GoTo ErrHand
   cnnConnection.BeginTrans
   
   '刪除
   stSQL = "delete from SALARYDATA where SD01='" & txtSD(1) & "'"
   Pub_SeekTbLog stSQL
   cnnConnection.Execute stSQL, intI
   
   cnnConnection.CommitTrans
   
   DelRecord = True
   txtSD(1).Tag = ""
   Exit Function
   
ErrHand:
    cnnConnection.RollbackTrans
    MsgBox Err.Description, vbCritical

End Function

'Remove by Morgan 2010/4/14 改抓公用函數
'Private Function GetInsureBase(pInsureSalary As Long, pKind As String) As Long
'   strExc(0) = "select si02 from SalaryInsurance" & _
'      " where si01='" & pKind & "' and si03<=" & pInsureSalary & " and si04>=" & pInsureSalary
'   intI = 1
'   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'   If intI = 1 Then
'      GetInsureBase = Val("" & RsTemp.Fields(0))
'   End If
'End Function

Private Function CheckExists(pstrUserNo As String) As Boolean
   CheckExists = True
   strExc(0) = "select 1 from salarydata where sd01='" & pstrUserNo & "'"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 0 Then
      CheckExists = False
   End If
End Function

Private Sub grd1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   Dim tmpMouseRow
   Dim i, j
   
   GRD1.Visible = False
   'tmpMouseRow = grd1.row
   tmpMouseRow = GRD1.MouseRow
   GRD1.Visible = True
   If tmpMouseRow <> 0 Then
       GRD1.row = tmpMouseRow
       GRD1.col = 0
       If GRD1.CellBackColor = QBColor(15) Then
            GRD1.Visible = False
            For j = 1 To GRD1.Rows - 1
                GRD1.row = j
                For i = 0 To GRD1.Cols - 1
                     GRD1.col = i
                     GRD1.CellBackColor = QBColor(15)
                Next i
            Next j
            GRD1.row = tmpMouseRow
            For i = 0 To GRD1.Cols - 1
                GRD1.col = i
                GRD1.CellBackColor = &HFFC0C0
            Next i
            textSR03.Text = GRD1.TextMatrix(tmpMouseRow, 0)
            textSR04.Text = GRD1.TextMatrix(tmpMouseRow, 1)
            chkSR08.Value = IIf(GRD1.TextMatrix(tmpMouseRow, 6) = "Y", vbChecked, vbUnchecked)
            SelCombo cboHL05, GRD1.TextMatrix(tmpMouseRow, 13)
            GRD1.Visible = True
       End If
   End If
End Sub
'讀取最新的健保資料
Private Sub GetNewHiData(ByVal stSR01 As String, ByVal stSR02 As String, ByRef stSR08 As String, ByRef stHL05 As String)
   Dim intR As Integer, stSQL As String
   stSQL = "select SR08,HL05 FROM staff_relation,(select hl02,hl05 from HIrelationlog a" & _
      " where hl01='" & stSR01 & "' and hl02=" & stSR02 & _
      " and hl03= (select max(b.hl03) from hirelationlog b where b.hl01=a.hl01 and b.hl02=a.hl02)" & _
      ") X WHERE SR01 = '" & stSR01 & "' and SR02=" & stSR02 & " and hl02(+)=SR02"
   intR = 1
   Set RsTemp = ClsLawReadRstMsg(intR, stSQL)
   If intR = 1 Then
      stSR08 = "" & RsTemp("SR08").Value
      stHL05 = "" & RsTemp("HL05").Value
   End If
End Sub

Private Sub SetGrd()
   Dim arrGridHeadText, arrGridHeadWidth
   Dim iRow As Integer
   arrGridHeadText = Array("稱謂", "姓名", "狀態", "性別", "出生日期", "身分證字號", "健保眷屬", "歿", "電話", "郵遞區號", "地址", "刪除日期", "序號", "健保補助類別")
   arrGridHeadWidth = Array(600, 850, 450, 500, 800, 1000, 800, 350, 1100, 820, 2000, 820, 0, 0)
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

'Add by Morgan 2009/6/24
'選取選單
Private Sub SelCombo(ByRef pCBO As ComboBox, ByVal pValue As String, Optional pLen As Integer = 2)
   Dim idx As Integer
   If pValue = "" Then
      pCBO.ListIndex = 0
   Else
      For idx = 1 To pCBO.ListCount - 1
         If Left(pCBO.List(idx), pLen) = pValue Then
            pCBO.ListIndex = idx
            Exit For
         End If
      Next
   End If
End Sub

Private Sub ClearSR()
   textSR03 = Empty
   textSR04 = Empty
   chkSR08.Value = vbUnchecked
   cboHL05.ListIndex = 0
End Sub

'Removed by Morgan 2013/1/21
''Add by Morgan 2009/6/23
''檢查勞保費是否特別,現在只有雇主(所長)
'Private Function CheckExceptLiRate() As Boolean
'   If txtSL(1) <> "" Then
'      strExc(0) = "select st02 from staff where st01='" & txtSD(1) & "' and st20='11'"
'      intI = 1
'      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'      If intI = 1 Then
'         CheckExceptLiRate = True
'      End If
'   End If
'End Function
