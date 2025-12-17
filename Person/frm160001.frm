VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm160001 
   BorderStyle     =   1  '³æ½u©T©w
   Caption         =   "­û¤u°ò¥»¸ê®Æ"
   ClientHeight    =   5880
   ClientLeft      =   50
   ClientTop       =   340
   ClientWidth     =   9170
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5880
   ScaleWidth      =   9170
   Begin VB.CommandButton cmdOK 
      Caption         =   "¬d¸ß¤U¤@µ§(&N)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   6
      Left            =   6990
      Style           =   1  '¹Ï¤ù¥~Æ[
      TabIndex        =   133
      Top             =   120
      Visible         =   0   'False
      Width           =   1320
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "µ²§ô(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   5
      Left            =   8355
      Style           =   1  '¹Ï¤ù¥~Æ[
      TabIndex        =   132
      Top             =   120
      Visible         =   0   'False
      Width           =   756
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4632
      Left            =   12
      TabIndex        =   71
      Top             =   936
      Width           =   9144
      _ExtentX        =   16122
      _ExtentY        =   8167
      _Version        =   393216
      Tabs            =   7
      Tab             =   2
      TabsPerRow      =   7
      TabHeight       =   520
      TabCaption(0)   =   "°ò¥»¸ê®Æ 1"
      TabPicture(0)   =   "frm160001.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "tmpPic"
      Tab(0).Control(1)=   "Command1"
      Tab(0).Control(2)=   "textST27"
      Tab(0).Control(3)=   "textST51"
      Tab(0).Control(4)=   "textST25"
      Tab(0).Control(5)=   "textST22"
      Tab(0).Control(6)=   "textST24"
      Tab(0).Control(7)=   "textST21"
      Tab(0).Control(8)=   "textST20"
      Tab(0).Control(9)=   "textST06"
      Tab(0).Control(10)=   "textST03"
      Tab(0).Control(11)=   "textST49"
      Tab(0).Control(12)=   "textST42"
      Tab(0).Control(13)=   "textST18"
      Tab(0).Control(14)=   "textST41"
      Tab(0).Control(15)=   "textST19"
      Tab(0).Control(16)=   "textST13"
      Tab(0).Control(17)=   "textST10"
      Tab(0).Control(18)=   "textST09"
      Tab(0).Control(19)=   "textST23"
      Tab(0).Control(20)=   "textST26"
      Tab(0).Control(21)=   "textST28"
      Tab(0).Control(22)=   "textST29"
      Tab(0).Control(23)=   "textST30"
      Tab(0).Control(24)=   "textST31"
      Tab(0).Control(25)=   "textST32"
      Tab(0).Control(26)=   "textST33"
      Tab(0).Control(27)=   "Label23"
      Tab(0).Control(28)=   "textST08"
      Tab(0).Control(29)=   "LabelST30"
      Tab(0).Control(30)=   "Label1(32)"
      Tab(0).Control(31)=   "Line1"
      Tab(0).Control(32)=   "Label1(29)"
      Tab(0).Control(33)=   "Label15"
      Tab(0).Control(34)=   "Label1(10)"
      Tab(0).Control(35)=   "Label1(4)"
      Tab(0).Control(36)=   "Label1(3)"
      Tab(0).Control(37)=   "Label22"
      Tab(0).Control(38)=   "Label13"
      Tab(0).Control(39)=   "Label4"
      Tab(0).Control(40)=   "Label1(9)"
      Tab(0).Control(41)=   "Label1(8)"
      Tab(0).Control(42)=   "Label1(7)"
      Tab(0).Control(43)=   "Label1(5)"
      Tab(0).Control(44)=   "Label1(2)"
      Tab(0).Control(45)=   "Label2"
      Tab(0).Control(46)=   "Label1(12)"
      Tab(0).Control(47)=   "Label6"
      Tab(0).Control(48)=   "Label1(13)"
      Tab(0).Control(49)=   "Label1(14)"
      Tab(0).Control(50)=   "Label1(15)"
      Tab(0).Control(51)=   "Label17"
      Tab(0).Control(52)=   "Label1(16)"
      Tab(0).Control(53)=   "Label1(17)"
      Tab(0).Control(54)=   "Label1(18)"
      Tab(0).Control(55)=   "Label1(19)"
      Tab(0).Control(56)=   "Label1(20)"
      Tab(0).Control(57)=   "Label1(21)"
      Tab(0).Control(58)=   "Label21"
      Tab(0).Control(59)=   "Label1(22)"
      Tab(0).Control(60)=   "Label1(23)"
      Tab(0).Control(61)=   "Label1(24)"
      Tab(0).ControlCount=   62
      TabCaption(1)   =   "°ò¥»¸ê®Æ 2"
      TabPicture(1)   =   "frm160001.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txtBackDate"
      Tab(1).Control(1)=   "textST37"
      Tab(1).Control(2)=   "textST40"
      Tab(1).Control(3)=   "textST39"
      Tab(1).Control(4)=   "textST36"
      Tab(1).Control(5)=   "textST35"
      Tab(1).Control(6)=   "textST38"
      Tab(1).Control(7)=   "textST34"
      Tab(1).Control(8)=   "LblBackDate"
      Tab(1).Control(9)=   "LblST40"
      Tab(1).Control(10)=   "Label25"
      Tab(1).Control(11)=   "Label1(33)"
      Tab(1).Control(12)=   "Label1(30)"
      Tab(1).Control(13)=   "Label1(11)"
      Tab(1).Control(14)=   "Label1(28)"
      Tab(1).Control(15)=   "Label1(27)"
      Tab(1).Control(16)=   "Label1(26)"
      Tab(1).Control(17)=   "Label1(25)"
      Tab(1).Control(18)=   "Label1(6)"
      Tab(1).Control(19)=   "Label11"
      Tab(1).ControlCount=   20
      TabCaption(2)   =   "³Ò°·°h²²ÄÝ"
      TabPicture(2)   =   "frm160001.frx":0038
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Label18"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Label1(42)"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Label1(44)"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Label20"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "Label1(34)"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "cboST50"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "txtSD16"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "txtSD17"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "cboST56"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).Control(9)=   "Frame1"
      Tab(2).Control(9).Enabled=   0   'False
      Tab(2).ControlCount=   10
      TabCaption(3)   =   "¾ú¦~¦ÒÁZ"
      TabPicture(3)   =   "frm160001.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "grd2"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "±Mªø¸ê®Æ"
      TabPicture(4)   =   "frm160001.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "textSS07"
      Tab(4).Control(1)=   "textSS06"
      Tab(4).Control(2)=   "textSS05"
      Tab(4).Control(3)=   "textSS04"
      Tab(4).Control(4)=   "textSS03"
      Tab(4).Control(5)=   "textSS02"
      Tab(4).Control(6)=   "Label1(43)"
      Tab(4).Control(7)=   "Label1(40)"
      Tab(4).Control(8)=   "Label1(39)"
      Tab(4).Control(9)=   "Label1(38)"
      Tab(4).Control(10)=   "Label1(37)"
      Tab(4).Control(11)=   "Label1(36)"
      Tab(4).Control(12)=   "Label1(35)"
      Tab(4).ControlCount=   13
      TabCaption(5)   =   "°·ÀË¸ê®Æ"
      TabPicture(5)   =   "frm160001.frx":008C
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "textST68"
      Tab(5).Control(1)=   "grd3"
      Tab(5).Control(2)=   "Label24"
      Tab(5).ControlCount=   3
      TabCaption(6)   =   "§ë«Oª÷ÃB"
      TabPicture(6)   =   "frm160001.frx":00A8
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "Label1(48)"
      Tab(6).Control(1)=   "Label1(45)"
      Tab(6).Control(2)=   "Label1(46)"
      Tab(6).Control(3)=   "lblSDdata(0)"
      Tab(6).Control(4)=   "lblSDdata(1)"
      Tab(6).Control(5)=   "lblSDdata(3)"
      Tab(6).Control(6)=   "lblSDdata(2)"
      Tab(6).Control(7)=   "Label1(47)"
      Tab(6).ControlCount=   8
      Begin VB.TextBox txtBackDate 
         Enabled         =   0   'False
         Height          =   270
         Left            =   -73200
         Locked          =   -1  'True
         MaxLength       =   7
         TabIndex        =   37
         Top             =   2520
         Width           =   945
      End
      Begin VB.TextBox textST68 
         Appearance      =   0  '¥­­±
         BackColor       =   &H8000000F&
         BorderStyle     =   0  '¨S¦³®Ø½u
         Height          =   270
         Left            =   -73080
         TabIndex        =   142
         Top             =   390
         Width           =   615
      End
      Begin VB.PictureBox tmpPic 
         AutoRedraw      =   -1  'True
         DragMode        =   1  '¦Û°Ê
         Height          =   2300
         Left            =   -67850
         ScaleHeight     =   2260
         ScaleWidth      =   1800
         TabIndex        =   130
         Top             =   360
         Width           =   1840
         Begin VB.Image tmpImg 
            BorderStyle     =   1  '³æ½u©T©w
            Height          =   960
            Left            =   420
            Stretch         =   -1  'True
            Top             =   390
            Visible         =   0   'False
            Width           =   1020
         End
      End
      Begin VB.CommandButton Command1 
         Caption         =   "·Ó¤ù¤W¶Ç(&U)"
         Height          =   285
         Left            =   -67320
         TabIndex        =   29
         Top             =   2700
         Width           =   1275
      End
      Begin VB.Frame Frame1 
         Caption         =   "²²ÄÝ¸ê®Æ"
         ForeColor       =   &H8000000D&
         Height          =   3375
         Left            =   90
         TabIndex        =   119
         Top             =   1080
         Width           =   8880
         Begin VB.CheckBox chkSR13 
            Caption         =   "ª\"
            Height          =   195
            Left            =   7785
            TabIndex        =   48
            Top             =   930
            Width           =   555
         End
         Begin VB.TextBox textSR10 
            Height          =   285
            Left            =   1035
            MaxLength       =   12
            TabIndex        =   49
            Top             =   1245
            Width           =   1095
         End
         Begin VB.TextBox textSR09 
            Height          =   270
            Left            =   5865
            MaxLength       =   20
            TabIndex        =   47
            Top             =   930
            Width           =   1440
         End
         Begin VB.CheckBox chkSR08 
            Caption         =   "°·«O²²ÄÝ"
            Height          =   285
            Left            =   45
            TabIndex        =   51
            Top             =   1560
            Width           =   1035
         End
         Begin VB.TextBox textSR07 
            Height          =   285
            Left            =   3585
            MaxLength       =   10
            TabIndex        =   46
            Top             =   930
            Width           =   1545
         End
         Begin VB.TextBox textSR06 
            Height          =   285
            Left            =   1035
            MaxLength       =   7
            TabIndex        =   45
            Top             =   930
            Width           =   1095
         End
         Begin VB.ComboBox textSR03 
            Height          =   300
            ItemData        =   "frm160001.frx":00C4
            Left            =   735
            List            =   "frm160001.frx":00D7
            TabIndex        =   42
            Top             =   600
            Width           =   1785
         End
         Begin VB.CommandButton cmdOK 
            Caption         =   "¨ú®ø"
            Enabled         =   0   'False
            Height          =   345
            Index           =   2
            Left            =   3420
            TabIndex        =   59
            Top             =   210
            Width           =   795
         End
         Begin VB.CommandButton cmdOK 
            Caption         =   "½T©w"
            Enabled         =   0   'False
            Height          =   345
            Index           =   1
            Left            =   2595
            TabIndex        =   58
            Top             =   210
            Width           =   795
         End
         Begin VB.CommandButton cmdOK 
            Caption         =   "·s¼W"
            Enabled         =   0   'False
            Height          =   345
            Index           =   0
            Left            =   60
            TabIndex        =   55
            Top             =   210
            Width           =   795
         End
         Begin VB.ComboBox textSR05 
            Height          =   300
            ItemData        =   "frm160001.frx":0108
            Left            =   5865
            List            =   "frm160001.frx":0112
            TabIndex        =   44
            Top             =   600
            Width           =   1125
         End
         Begin VB.CommandButton cmdLog 
            Caption         =   "°·«O²§°Ê¸ê®Æ"
            Height          =   315
            Left            =   7110
            TabIndex        =   60
            Top             =   210
            Width           =   1635
         End
         Begin VB.ComboBox cboHL05 
            Height          =   300
            ItemData        =   "frm160001.frx":0125
            Left            =   2520
            List            =   "frm160001.frx":0127
            Style           =   2  '³æ¯Â¤U©Ô¦¡
            TabIndex        =   52
            Top             =   1545
            Width           =   4005
         End
         Begin VB.TextBox textSR12 
            Height          =   285
            Left            =   7785
            MaxLength       =   7
            TabIndex        =   53
            Top             =   1560
            Width           =   1005
         End
         Begin VB.CommandButton cmdOK 
            Caption         =   "§R°£"
            Enabled         =   0   'False
            Height          =   345
            Index           =   3
            Left            =   1755
            TabIndex        =   57
            Top             =   210
            Width           =   795
         End
         Begin VB.CommandButton cmdOK 
            Caption         =   "­×§ï"
            Enabled         =   0   'False
            Height          =   345
            Index           =   4
            Left            =   900
            TabIndex        =   56
            Top             =   210
            Width           =   795
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid grd1 
            Height          =   1395
            Left            =   90
            TabIndex        =   54
            Top             =   1890
            Width           =   8715
            _ExtentX        =   15363
            _ExtentY        =   2469
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
               Name            =   "·s²Ó©úÅé-ExtB"
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
         Begin MSForms.TextBox textSR11 
            Height          =   285
            Left            =   2850
            TabIndex        =   50
            Top             =   1245
            Width           =   5925
            VariousPropertyBits=   679495707
            MaxLength       =   70
            Size            =   "10451;503"
            FontName        =   "·s²Ó©úÅé-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox textSR04 
            Height          =   285
            Left            =   3585
            TabIndex        =   43
            Top             =   600
            Width           =   1545
            VariousPropertyBits=   679495707
            MaxLength       =   12
            Size            =   "2725;503"
            FontName        =   "·s²Ó©úÅé-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "¦a§}¡G"
            Height          =   180
            Left            =   2295
            TabIndex        =   129
            Top             =   1290
            Width           =   540
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "¶l»¼°Ï¸¹¡G"
            Height          =   180
            Left            =   75
            TabIndex        =   128
            Top             =   1290
            Width           =   900
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "¹q¸Ü¡G"
            Height          =   180
            Left            =   5295
            TabIndex        =   127
            Top             =   975
            Width           =   540
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "¨­¤ÀÃÒ¦r¸¹¡G"
            Height          =   180
            Left            =   2475
            TabIndex        =   126
            Top             =   975
            Width           =   1080
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "¥X¥Í¤é´Á¡G"
            Height          =   180
            Left            =   75
            TabIndex        =   125
            Top             =   975
            Width           =   900
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "©Ê§O¡G"
            Height          =   180
            Left            =   5295
            TabIndex        =   124
            Top             =   660
            Width           =   540
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "©m¦W¡G"
            Height          =   180
            Left            =   2835
            TabIndex        =   123
            Top             =   660
            Width           =   540
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "ºÙ¿×¡G"
            Height          =   180
            Left            =   75
            TabIndex        =   122
            Top             =   660
            Width           =   540
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "°·«O¸É§UÃþ§O¡G"
            Height          =   180
            Left            =   1215
            TabIndex        =   121
            Top             =   1605
            Width           =   1260
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "§R°£¤é´Á¡G"
            Height          =   180
            Left            =   6795
            TabIndex        =   120
            Top             =   1605
            Width           =   900
         End
      End
      Begin VB.ComboBox cboST56 
         Height          =   260
         ItemData        =   "frm160001.frx":0129
         Left            =   4860
         List            =   "frm160001.frx":012B
         Style           =   2  '³æ¯Â¤U©Ô¦¡
         TabIndex        =   39
         Top             =   435
         Width           =   4005
      End
      Begin VB.TextBox txtSD17 
         Alignment       =   1  '¾a¥k¹ï»ô
         Height          =   270
         Left            =   1455
         MaxLength       =   2
         TabIndex        =   40
         Text            =   "99.99"
         Top             =   750
         Width           =   555
      End
      Begin VB.TextBox txtSD16 
         Height          =   270
         Left            =   1470
         MaxLength       =   1
         TabIndex        =   38
         Text            =   "Y"
         Top             =   450
         Width           =   285
      End
      Begin VB.ComboBox cboST50 
         Height          =   260
         ItemData        =   "frm160001.frx":012D
         Left            =   4860
         List            =   "frm160001.frx":012F
         Style           =   2  '³æ¯Â¤U©Ô¦¡
         TabIndex        =   41
         Top             =   735
         Width           =   4005
      End
      Begin VB.ComboBox textST27 
         Height          =   260
         Left            =   -73980
         TabIndex        =   14
         Top             =   1920
         Width           =   2205
      End
      Begin VB.TextBox textST51 
         Enabled         =   0   'False
         Height          =   270
         Left            =   -68940
         MaxLength       =   7
         TabIndex        =   16
         Top             =   1920
         Width           =   945
      End
      Begin VB.ComboBox textST25 
         Height          =   260
         ItemData        =   "frm160001.frx":0131
         Left            =   -73980
         List            =   "frm160001.frx":0141
         TabIndex        =   11
         Top             =   1610
         Width           =   1155
      End
      Begin VB.ComboBox textST22 
         Height          =   260
         ItemData        =   "frm160001.frx":0152
         Left            =   -73980
         List            =   "frm160001.frx":015C
         TabIndex        =   8
         Top             =   1290
         Width           =   1125
      End
      Begin VB.ComboBox textST24 
         Height          =   260
         ItemData        =   "frm160001.frx":016F
         Left            =   -70020
         List            =   "frm160001.frx":0179
         TabIndex        =   10
         Top             =   1290
         Width           =   2145
      End
      Begin VB.ComboBox textST21 
         Height          =   260
         ItemData        =   "frm160001.frx":018F
         Left            =   -70020
         List            =   "frm160001.frx":0191
         TabIndex        =   6
         Top             =   680
         Width           =   2145
      End
      Begin VB.ComboBox textST20 
         Height          =   260
         ItemData        =   "frm160001.frx":0193
         Left            =   -73980
         List            =   "frm160001.frx":0195
         TabIndex        =   5
         Top             =   680
         Width           =   2205
      End
      Begin VB.ComboBox textST06 
         Height          =   260
         ItemData        =   "frm160001.frx":0197
         Left            =   -70020
         List            =   "frm160001.frx":01AA
         TabIndex        =   4
         Top             =   360
         Width           =   2145
      End
      Begin VB.ComboBox textST03 
         Height          =   260
         Left            =   -73980
         TabIndex        =   3
         Top             =   360
         Width           =   2205
      End
      Begin VB.ComboBox textST37 
         Height          =   300
         ItemData        =   "frm160001.frx":01DB
         Left            =   -73950
         List            =   "frm160001.frx":01DD
         TabIndex        =   33
         Top             =   990
         Width           =   2985
      End
      Begin VB.TextBox textST49 
         Height          =   270
         Left            =   -73980
         MaxLength       =   80
         TabIndex        =   7
         Top             =   1000
         Width           =   6105
      End
      Begin VB.TextBox textST40 
         Height          =   270
         Left            =   -73950
         Locked          =   -1  'True
         MaxLength       =   4
         TabIndex        =   36
         Top             =   2190
         Width           =   615
      End
      Begin VB.TextBox textST39 
         Height          =   270
         Left            =   -73950
         MaxLength       =   40
         TabIndex        =   35
         Top             =   1620
         Width           =   5385
      End
      Begin VB.TextBox textST36 
         Height          =   270
         Left            =   -68940
         MaxLength       =   12
         TabIndex        =   31
         Top             =   390
         Width           =   1575
      End
      Begin VB.TextBox textST35 
         Height          =   270
         Left            =   -73530
         MaxLength       =   20
         TabIndex        =   30
         Top             =   390
         Width           =   1575
      End
      Begin VB.TextBox textST38 
         Height          =   270
         Left            =   -73950
         MaxLength       =   50
         TabIndex        =   34
         Top             =   1320
         Width           =   5385
      End
      Begin VB.TextBox textST42 
         Height          =   270
         Left            =   -73920
         MaxLength       =   50
         TabIndex        =   28
         Top             =   3960
         Width           =   4995
      End
      Begin VB.TextBox textST18 
         Height          =   270
         Left            =   -73920
         MaxLength       =   50
         TabIndex        =   27
         Top             =   3670
         Width           =   4995
      End
      Begin VB.TextBox textST41 
         Height          =   270
         Left            =   -70860
         MaxLength       =   7
         TabIndex        =   15
         Top             =   1920
         Width           =   945
      End
      Begin VB.TextBox textST19 
         Height          =   270
         Left            =   -68940
         MaxLength       =   20
         TabIndex        =   26
         Top             =   3380
         Width           =   2085
      End
      Begin VB.TextBox textST13 
         Height          =   270
         Left            =   -71940
         MaxLength       =   7
         TabIndex        =   12
         Top             =   1610
         Width           =   945
      End
      Begin VB.TextBox textST10 
         Height          =   270
         Left            =   -73530
         MaxLength       =   20
         TabIndex        =   25
         Top             =   3380
         Width           =   1575
      End
      Begin VB.TextBox textST09 
         Height          =   270
         Left            =   -73530
         MaxLength       =   20
         TabIndex        =   22
         Top             =   2800
         Width           =   1575
      End
      Begin VB.TextBox textST23 
         Height          =   270
         Left            =   -71940
         MaxLength       =   7
         TabIndex        =   9
         Top             =   1290
         Width           =   945
      End
      Begin VB.TextBox textST26 
         Height          =   270
         Left            =   -70020
         MaxLength       =   10
         TabIndex        =   13
         Top             =   1610
         Width           =   1305
      End
      Begin VB.TextBox textST28 
         Height          =   270
         Left            =   -73980
         MaxLength       =   7
         TabIndex        =   17
         Top             =   2220
         Width           =   945
      End
      Begin VB.TextBox textST29 
         Height          =   270
         Left            =   -72990
         MaxLength       =   7
         TabIndex        =   18
         Top             =   2220
         Width           =   945
      End
      Begin VB.TextBox textST30 
         Height          =   270
         Left            =   -68940
         MaxLength       =   6
         TabIndex        =   19
         Top             =   2220
         Width           =   945
      End
      Begin VB.TextBox textST31 
         Height          =   270
         Left            =   -73980
         MaxLength       =   7
         TabIndex        =   20
         Top             =   2510
         Width           =   945
      End
      Begin VB.TextBox textST32 
         Height          =   270
         Left            =   -68940
         MaxLength       =   7
         TabIndex        =   21
         Top             =   2510
         Width           =   945
      End
      Begin VB.TextBox textST33 
         Height          =   270
         Left            =   -68940
         MaxLength       =   12
         TabIndex        =   23
         Top             =   2800
         Width           =   1575
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grd2 
         Height          =   4095
         Left            =   -74910
         TabIndex        =   61
         Top             =   360
         Width           =   4035
         _ExtentX        =   7108
         _ExtentY        =   7214
         _Version        =   393216
         Cols            =   1
         FixedCols       =   0
         ScrollTrack     =   -1  'True
         HighLight       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "·s²Ó©úÅé-ExtB"
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
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grd3 
         Height          =   3855
         Left            =   -74850
         TabIndex        =   141
         Top             =   690
         Width           =   8775
         _ExtentX        =   15469
         _ExtentY        =   6809
         _Version        =   393216
         Cols            =   4
         FixedCols       =   0
         ScrollTrack     =   -1  'True
         HighLight       =   0
         FormatString    =   "°·ÀË¤é´Á|¸É§U¶O¥Î|Ãº¥æ¤é´Á|³Æ¡@¡@µù"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "·s²Ó©úÅé-ExtB"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   4
      End
      Begin MSForms.Label Label23 
         Height          =   195
         Left            =   -74250
         TabIndex        =   156
         Top             =   4320
         Width           =   7905
         VariousPropertyBits=   27
         Caption         =   "CREATE :                                                    UPDATE : "
         Size            =   "13944;344"
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textSS07 
         Height          =   696
         Left            =   -73716
         TabIndex        =   67
         Top             =   3900
         Width           =   7812
         VariousPropertyBits=   -1466939365
         MaxLength       =   2000
         ScrollBars      =   3
         Size            =   "13785;1217"
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textSS06 
         Height          =   720
         Left            =   -73716
         TabIndex        =   66
         Top             =   3180
         Width           =   7812
         VariousPropertyBits=   -1466939365
         MaxLength       =   2000
         ScrollBars      =   3
         Size            =   "13785;1270"
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textSS05 
         Height          =   720
         Left            =   -73716
         TabIndex        =   65
         Top             =   2460
         Width           =   7812
         VariousPropertyBits=   -1466939365
         MaxLength       =   2000
         ScrollBars      =   3
         Size            =   "13785;1270"
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textSS04 
         Height          =   720
         Left            =   -73716
         TabIndex        =   64
         Top             =   1740
         Width           =   7812
         VariousPropertyBits=   -1466939365
         MaxLength       =   2000
         ScrollBars      =   3
         Size            =   "13785;1270"
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textSS03 
         Height          =   720
         Left            =   -73716
         TabIndex        =   63
         Top             =   1020
         Width           =   7812
         VariousPropertyBits=   -1466939365
         MaxLength       =   2000
         ScrollBars      =   3
         Size            =   "13785;1270"
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textSS02 
         Height          =   696
         Left            =   -73716
         TabIndex        =   62
         Top             =   336
         Width           =   7812
         VariousPropertyBits=   -1466939365
         MaxLength       =   2000
         ScrollBars      =   3
         Size            =   "13785;1217"
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textST34 
         Height          =   285
         Left            =   -73530
         TabIndex        =   32
         Top             =   660
         Width           =   7485
         VariousPropertyBits=   679495707
         MaxLength       =   70
         Size            =   "13203;503"
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textST08 
         Height          =   285
         Left            =   -73530
         TabIndex        =   24
         Top             =   3080
         Width           =   7485
         VariousPropertyBits=   679495707
         MaxLength       =   70
         Size            =   "13203;503"
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label LblBackDate 
         AutoSize        =   -1  'True
         Caption         =   "¯dÂ¾°±Á~¯S¥ð°_ºâ¤é"
         Height          =   180
         Left            =   -74880
         TabIndex        =   154
         Top             =   2550
         Width           =   1620
      End
      Begin VB.Label LblST40 
         Caption         =   "99"
         ForeColor       =   &H000000C0&
         Height          =   180
         Left            =   -73890
         TabIndex        =   153
         Top             =   1950
         Width           =   360
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         Caption         =   "¯S§O°²¦~«×"
         ForeColor       =   &H000000C0&
         Height          =   180
         Left            =   -74880
         TabIndex        =   152
         Top             =   1950
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '³z©ú
         Caption         =   "³Ò°·«O¬O§_¥H¦X¹Ù¤H¨­¤À§ë«O¡G       (Y:¬O)"
         Height          =   180
         Index           =   47
         Left            =   -74340
         TabIndex        =   151
         Top             =   1470
         Width           =   3300
      End
      Begin VB.Label lblSDdata 
         AutoSize        =   -1  'True
         Caption         =   "Y"
         Height          =   180
         Index           =   2
         Left            =   -71740
         TabIndex        =   150
         Top             =   1470
         Width           =   120
      End
      Begin VB.Label lblSDdata 
         Alignment       =   1  '¾a¥k¹ï»ô
         AutoSize        =   -1  'True
         Caption         =   "99,999,999"
         Height          =   180
         Index           =   3
         Left            =   -72840
         TabIndex        =   149
         Top             =   1860
         Width           =   810
      End
      Begin VB.Label lblSDdata 
         Alignment       =   1  '¾a¥k¹ï»ô
         AutoSize        =   -1  'True
         Caption         =   "99,999,999"
         Height          =   180
         Index           =   1
         Left            =   -72840
         TabIndex        =   148
         Top             =   1080
         Width           =   810
      End
      Begin VB.Label lblSDdata 
         Alignment       =   1  '¾a¥k¹ï»ô
         AutoSize        =   -1  'True
         Caption         =   "99,999,999"
         Height          =   180
         Index           =   0
         Left            =   -72840
         TabIndex        =   147
         Top             =   690
         Width           =   810
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "°·«O§ë«Oª÷ÃB¡G"
         Height          =   180
         Index           =   46
         Left            =   -74150
         TabIndex        =   146
         Top             =   1080
         Width           =   1260
      End
      Begin VB.Label Label1 
         Alignment       =   1  '¾a¥k¹ï»ô
         AutoSize        =   -1  'True
         Caption         =   "³Ò«O§ë«Oª÷ÃB¡G"
         Height          =   180
         Index           =   45
         Left            =   -74150
         TabIndex        =   145
         Top             =   690
         Width           =   1260
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "°h¥ðª÷§ë«Oª÷ÃB¡G"
         Height          =   180
         Index           =   48
         Left            =   -74340
         TabIndex        =   144
         Top             =   1860
         Width           =   1440
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         Caption         =   "¤U¦¸À³Ãº¦~«×¡G"
         Height          =   180
         Left            =   -74340
         TabIndex        =   143
         Top             =   420
         Width           =   1260
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "·|­û"
         Height          =   180
         Index           =   43
         Left            =   -74670
         TabIndex        =   140
         Top             =   4200
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "±M·~/ªÀ·|¹ÎÅé"
         Height          =   180
         Index           =   40
         Left            =   -74940
         TabIndex        =   139
         Top             =   3990
         Width           =   1125
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "«D´¼°]±M·~ÃÒ·Ó"
         Height          =   180
         Index           =   39
         Left            =   -74964
         TabIndex        =   138
         Top             =   3300
         Width           =   1260
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "µÛ§@/µo©ú/³Ð§@"
         Height          =   180
         Index           =   38
         Left            =   -74970
         TabIndex        =   137
         Top             =   2550
         Width           =   1170
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "´¼°]±M·~ÃÒ·Ó"
         ForeColor       =   &H000000C0&
         Height          =   180
         Index           =   37
         Left            =   -74904
         TabIndex        =   136
         Top             =   1800
         Width           =   1080
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "»y¤å¯à¤OÀË©w"
         Height          =   180
         Index           =   36
         Left            =   -74910
         TabIndex        =   135
         Top             =   1110
         Width           =   1080
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "±Ð¨|»P°V½m"
         Height          =   180
         Index           =   35
         Left            =   -74910
         TabIndex        =   134
         Top             =   390
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "%"
         Height          =   180
         Index           =   34
         Left            =   2070
         TabIndex        =   118
         Top             =   810
         Width           =   135
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "­û¤u³Ò«O¸É§UÃþ§O¡G"
         Height          =   180
         Left            =   3225
         TabIndex        =   117
         Top             =   495
         Width           =   1620
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "³Ò°h¦Û´£¶O²v¡G"
         Height          =   180
         Index           =   44
         Left            =   180
         TabIndex        =   116
         Top             =   795
         Width           =   1260
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "¾A¥Î³Ò°h·s¨î¡G        ( Y:¾A¥Î )"
         Height          =   180
         Index           =   42
         Left            =   180
         TabIndex        =   115
         Top             =   495
         Width           =   2355
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "­û¤u°·«O¸É§UÃþ§O¡G"
         Height          =   180
         Left            =   3225
         TabIndex        =   114
         Top             =   795
         Width           =   1620
      End
      Begin VB.Label Label1 
         Caption         =   "PS¡G1. ¤áÄy¦a§}Äæ¥u¥i¿é¤J30­Ó¥þ§Î¦r¡A¶W¹LªÌ¥i§R°£¾F¨½¸ê®Æ          2. ¥~¹´µL¤¤¤å¦a§}®É½Ð¿é¤J¹µ¥D¦a§}"
         ForeColor       =   &H000000FF&
         Height          =   360
         Index           =   33
         Left            =   -74640
         TabIndex        =   113
         Top             =   3720
         Width           =   5100
      End
      Begin VB.Label LabelST30 
         AutoSize        =   -1  'True
         Height          =   180
         Left            =   -67980
         TabIndex        =   112
         Top             =   2280
         Width           =   735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Â÷Â¾¤é´Á"
         Height          =   180
         Index           =   32
         Left            =   -69690
         TabIndex        =   111
         Top             =   1975
         Width           =   720
      End
      Begin VB.Line Line1 
         X1              =   -73170
         X2              =   -72750
         Y1              =   2340
         Y2              =   2340
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "³Ì°ª¾Ç¾ú"
         Height          =   180
         Index           =   30
         Left            =   -74700
         TabIndex        =   109
         Top             =   1035
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Â¾ºÙ»¡©ú"
         Height          =   180
         Index           =   29
         Left            =   -74730
         TabIndex        =   108
         Top             =   1045
         Width           =   720
      End
      Begin VB.Label Label15 
         Height          =   255
         Left            =   -68520
         TabIndex        =   107
         Top             =   695
         Width           =   1605
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "¦a§}"
         Height          =   180
         Index           =   11
         Left            =   -73980
         TabIndex        =   106
         Top             =   735
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "¬ì¨t"
         Height          =   180
         Index           =   28
         Left            =   -74340
         TabIndex        =   105
         Top             =   1665
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "¶l»¼°Ï¸¹"
         Height          =   180
         Index           =   27
         Left            =   -69675
         TabIndex        =   104
         Top             =   435
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "¹q¸Ü"
         Height          =   180
         Index           =   26
         Left            =   -73980
         TabIndex        =   103
         Top             =   435
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "¤áÄy¸ê®Æ"
         Height          =   180
         Index           =   25
         Left            =   -74730
         TabIndex        =   102
         Top             =   435
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "²¦·~¾Ç®Õ"
         Height          =   180
         Index           =   6
         Left            =   -74700
         TabIndex        =   101
         Top             =   1365
         Width           =   720
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "¥i¥ð¯S§O°²"
         Height          =   180
         Left            =   -74880
         TabIndex        =   100
         Top             =   2235
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "¦a§}"
         Height          =   180
         Index           =   10
         Left            =   -73980
         TabIndex        =   99
         Top             =   3135
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "¥~³¡ E-mail 2"
         Height          =   180
         Index           =   4
         Left            =   -74940
         TabIndex        =   98
         Top             =   4005
         Width           =   1005
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "¥~³¡ E-mail 1"
         Height          =   180
         Index           =   3
         Left            =   -74940
         TabIndex        =   97
         Top             =   3715
         Width           =   1005
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "µ²±B¤é´Á"
         Height          =   180
         Left            =   -71670
         TabIndex        =   96
         Top             =   1975
         Width           =   720
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "¤â¾÷"
         Height          =   180
         Left            =   -69315
         TabIndex        =   95
         Top             =   3425
         Width           =   360
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "¤J©Ò¤é´Á"
         Height          =   180
         Left            =   -72750
         TabIndex        =   94
         Top             =   1670
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "¶Ç¯u"
         Height          =   180
         Index           =   9
         Left            =   -73980
         TabIndex        =   93
         Top             =   3425
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "¹q¸Ü"
         Height          =   180
         Index           =   8
         Left            =   -73980
         TabIndex        =   92
         Top             =   2845
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "³q°T¸ê®Æ"
         Height          =   180
         Index           =   7
         Left            =   -74730
         TabIndex        =   91
         Top             =   2845
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "­û¤u©ÒÄÝ©Ò§O"
         Height          =   180
         Index           =   5
         Left            =   -71160
         TabIndex        =   90
         Top             =   435
         Width           =   1080
      End
      Begin VB.Label Label1 
         Alignment       =   1  '¾a¥k¹ï»ô
         AutoSize        =   -1  'True
         Caption         =   "³¡ªù"
         Height          =   180
         Index           =   2
         Left            =   -74550
         TabIndex        =   89
         Top             =   410
         Width           =   540
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Height          =   180
         Left            =   -72930
         TabIndex        =   88
         Top             =   435
         Width           =   45
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Â¾ºÙ"
         Height          =   180
         Index           =   12
         Left            =   -74370
         TabIndex        =   87
         Top             =   740
         Width           =   360
      End
      Begin VB.Label Label6 
         Height          =   255
         Left            =   -73560
         TabIndex        =   86
         Top             =   695
         Width           =   1605
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Â¾¦ì"
         Height          =   180
         Index           =   13
         Left            =   -70410
         TabIndex        =   85
         Top             =   735
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "©Ê§O"
         Height          =   180
         Index           =   14
         Left            =   -74370
         TabIndex        =   84
         Top             =   1350
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "°êÄy          ( L. ¥»°ê  F. ¥~°ê )"
         Height          =   180
         Index           =   15
         Left            =   -70395
         TabIndex        =   83
         Top             =   1350
         Width           =   2205
      End
      Begin VB.Label Label17 
         Height          =   255
         Left            =   -68520
         TabIndex        =   82
         Top             =   1275
         Width           =   1605
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "¦å«¬"
         Height          =   180
         Index           =   16
         Left            =   -74370
         TabIndex        =   81
         Top             =   1670
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "¥X¥Í¤é´Á"
         Height          =   180
         Index           =   17
         Left            =   -72750
         TabIndex        =   80
         Top             =   1350
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "¨­¤ÀÃÒ¦r¸¹"
         Height          =   180
         Index           =   18
         Left            =   -70935
         TabIndex        =   79
         Top             =   1670
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "¥X¥Í¦a"
         Height          =   180
         Index           =   19
         Left            =   -74550
         TabIndex        =   78
         Top             =   1975
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "¸Õ¥Î´Á¶¡"
         Height          =   180
         Index           =   20
         Left            =   -74730
         TabIndex        =   77
         Top             =   2265
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Â¾°È¥N²z¤H"
         Height          =   180
         Index           =   21
         Left            =   -69855
         TabIndex        =   76
         Top             =   2265
         Width           =   900
      End
      Begin VB.Label Label21 
         Height          =   255
         Left            =   -67950
         TabIndex        =   75
         Top             =   2180
         Width           =   1605
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "¥[«O¤é´Á"
         Height          =   180
         Index           =   22
         Left            =   -74730
         TabIndex        =   74
         Top             =   2555
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "¤W¦¸Ã±«OÃÒ®Ñ¤é´Á"
         Height          =   180
         Index           =   23
         Left            =   -70395
         TabIndex        =   73
         Top             =   2555
         Width           =   1440
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "¶l»¼°Ï¸¹"
         Height          =   180
         Index           =   24
         Left            =   -69675
         TabIndex        =   72
         Top             =   2845
         Width           =   720
      End
   End
   Begin VB.TextBox textST01 
      Height          =   285
      Left            =   990
      MaxLength       =   5
      TabIndex        =   0
      Top             =   630
      Width           =   735
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   8190
      Top             =   0
      _ExtentX        =   988
      _ExtentY        =   988
      BackColor       =   -2147483643
      ImageWidth      =   20
      ImageHeight     =   20
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm160001.frx":01DF
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm160001.frx":04FB
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm160001.frx":0817
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm160001.frx":09F3
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm160001.frx":0D0F
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm160001.frx":102B
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm160001.frx":1347
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm160001.frx":1663
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm160001.frx":197F
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm160001.frx":1C9B
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm160001.frx":1FB7
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar TBar1 
      Align           =   1  '¹ï»ôªí³æ¤W¤è
      Height          =   520
      Left            =   0
      TabIndex        =   68
      Top             =   0
      Width           =   9170
      _ExtentX        =   16175
      _ExtentY        =   917
      ButtonWidth     =   1076
      ButtonHeight    =   882
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      DisabledImageList=   "ImageList1"
      HotImageList    =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   14
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "·s¼W"
            Key             =   "keyInsert"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "­×§ï"
            Key             =   "keyUpdate"
            ImageIndex      =   2
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   1
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "§R°£"
            Key             =   "keyDelete"
            ImageIndex      =   3
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   1
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "¬d¸ß"
            Key             =   "keyQuery"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   1
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "²Ä¤@µ§"
            Key             =   "keyFirst"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "«e¤@µ§"
            Key             =   "keyPrevious"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "«á¤@µ§"
            Key             =   "keyNext"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "³Ì«áµ§"
            Key             =   "keyLast"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   1
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "½T©w"
            Key             =   "keyOk"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "¨ú®ø"
            Key             =   "keyCancel"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   1
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "µ²§ô"
            Key             =   "keyExit"
            ImageIndex      =   11
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin VB.PictureBox G_SeekPicColor 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
         Enabled         =   0   'False
         Height          =   500
         Left            =   7650
         ScaleHeight     =   46
         ScaleMode       =   3  '¹³¯À
         ScaleWidth      =   46
         TabIndex        =   131
         Top             =   30
         Visible         =   0   'False
         Width           =   500
      End
   End
   Begin MSForms.TextBox textST12 
      Height          =   285
      Left            =   5100
      TabIndex        =   2
      Top             =   630
      Width           =   4005
      VariousPropertyBits=   679495707
      MaxLength       =   30
      Size            =   "7064;503"
      FontName        =   "·s²Ó©úÅé-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textST02 
      Height          =   285
      Left            =   2610
      TabIndex        =   1
      Top             =   630
      Width           =   1545
      VariousPropertyBits=   679495707
      MaxLength       =   12
      Size            =   "2725;503"
      FontName        =   "·s²Ó©úÅé-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      Caption         =   "PS¡G¿é¤J¸ê®Æ»Ý´«¦æ®É, ½Ð«ö Ctrl+Enter"
      ForeColor       =   &H000000FF&
      Height          =   210
      Index           =   49
      Left            =   120
      TabIndex        =   155
      Top             =   5610
      Width           =   5100
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "­^¤å©m¦W"
      Height          =   180
      Index           =   31
      Left            =   4350
      TabIndex        =   110
      Top             =   660
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "­û¤u¥N¸¹"
      Height          =   180
      Index           =   0
      Left            =   240
      TabIndex        =   70
      Top             =   675
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "¤¤¤å©m¦W"
      Height          =   180
      Index           =   1
      Left            =   1860
      TabIndex        =   69
      Top             =   660
      Width           =   720
   End
End
Attribute VB_Name = "frm160001"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/6/15 Form2.0¤w­×§ï
'Memo By Sindy 2012/12/3 ´¼Åv¤H­ûÄæ¤w­×§ï
'Memo By Sindy 2011/2/17 SQLDate¤wÀË¬d
'Memo By Sindy 2010/11/25 ­û¤u½s¸¹Äæ¤w­×§ï
'Modify By Sindy 2010/7/20 ¤é´ÁÄæ¤w­×§ï
'Create by nickc 2006/11/01 copy from frm140401
Option Explicit

Dim RcMain As New ADODB.Recordset, RsAdo As New ADODB.Recordset
' ÅÜ¼Æ«Å§i°Ï
Dim m_EditMode As Integer
Dim m_SubMode As Integer
'(°õ¦æ¦U¶µ¥\¯àªºÅv­­)
Dim m_bInsert As Boolean
Dim m_bUpdate As Boolean
Dim m_bDelete As Boolean
Dim m_bQuery As Boolean
Dim m_CU01 As String
' «Å§iÄæ¦ì¤º®eµ²ºc
Private Type FIELDITEM
   fiName As String
   fiOldData As String
   fiNewData As String
   fiType As Integer
End Type
Dim m_FieldList() As FIELDITEM
' ²Ä¤@µ§¸ê®Æªº¥»©Ò®×¸¹
Dim m_FirstKEY As String
' ³Ì«á¤@µ§¸ê®Æªº¥»©Ò®×¸¹
Dim m_LastKEY As String
' ¥Ø«e¥¿¦bÅã¥Üªº¥»©Ò®×¸¹
Dim m_CurrKEY As String
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
Dim tf_st As Integer
'Add by Morgan 2009/6/25
Dim iLstSelRow As Integer '«e¦¸ÂI¿ïªº²²ÄÝ
Dim m_EditMode2 As Integer '²²ÄÝ½s¿èª¬ºA
Public UpForm As Form 'Add By Sindy 2012/6/20


Private Sub chkSR08_Click()
   If chkSR08.Value = 1 Then
      'MsgBox chkSR08.Value
   End If
End Sub

Private Sub chkSR13_Click()
   If m_EditMode2 = 2 Then
      If chkSR13.Value = 1 And chkSR08.Value = 1 Then
         MsgBox "¡i" & textSR04 & "¡j¬OÂÂ¦³°·«O²²ÄÝ¡A½Ð¥ý¶i¡i°·«O²§°Ê¸ê®Æ¡j§@°·«O²¾¥X¡I"
         chkSR13.Value = 0
      End If
   End If
End Sub

Private Sub cmdok_Click(Index As Integer)
Dim i As Integer
   Select Case Index
      'Modify by Morgan 2009/6/10 ¨ú®ø­×§ï©M§R°£¥\¯à¡A§ï¬°¥t¥~¶}µe­±ºûÅ@
      'Case 0
      '        ' 2008/12/17 MODIFY BY SINDY
      '        'If textSR03.Text <> "" And textSR04.Text <> "" Then
      '        If textSR03.Text <> "" And textSR04.Text <> "" And textSR05.Text <> "" Then
      '        ' 2008/12/17 END
      '            If GRD1.TextMatrix(GRD1.Rows - 1, 0) <> "" Then
      '                GRD1.Rows = GRD1.Rows + 1
      '            End If
      '            GRD1.TextMatrix(GRD1.Rows - 1, 0) = textSR03.Text
      '            GRD1.TextMatrix(GRD1.Rows - 1, 1) = textSR04.Text
      '            GRD1.TextMatrix(GRD1.Rows - 1, 2) = textSR05.Text
      '            GRD1.TextMatrix(GRD1.Rows - 1, 3) = ChangeTStringToTDateString(textSR06.Text)
      '            GRD1.TextMatrix(GRD1.Rows - 1, 4) = textSR07.Text
      '            GRD1.TextMatrix(GRD1.Rows - 1, 5) = IIf(chkSR08.Value, "Y", "")
      '            GRD1.TextMatrix(GRD1.Rows - 1, 6) = IIf(chkSR13.Value, "Y", "")
      '            GRD1.TextMatrix(GRD1.Rows - 1, 7) = textSR09.Text
      '            GRD1.TextMatrix(GRD1.Rows - 1, 8) = textSR10.Text
      '            GRD1.TextMatrix(GRD1.Rows - 1, 9) = textSR11.Text
      '            GRD1.TextMatrix(GRD1.Rows - 1, 10) = ChangeTStringToTDateString(textSR12.Text)
      '            ClearSR
      '            GRD1.Refresh
      '        Else
      '            ' 2008/12/17 MODIFY BY SINDY
      '            If textSR03.Text = "" Then
      '               textSR03.SetFocus
      '            ElseIf textSR04.Text = "" Then
      '               textSR04.SetFocus
      '            ElseIf textSR05.Text = "" Then
      '               textSR05.SetFocus
      '            End If
      '            'MsgBox "ºÙ¿×¤Î©m¦W¤£¥iªÅ¥Õ¡I", vbExclamation, "µo¥Í¿ù»~¡I"
      '            MsgBox "ºÙ¿×¤Î©m¦W¤Î©Ê§O¤£¥iªÅ¥Õ¡I", vbExclamation, "µo¥Í¿ù»~¡I"
      '            ' 2008/12/17 END
      '        End If
      'Case 1
      '        ' 2008/12/17 MODIFY BY SINDY
      '        'If textSR03.Text <> "" And textSR04.Text <> "" Then
      '        If textSR03.Text <> "" And textSR04.Text <> "" And textSR05.Text <> "" Then
      '        ' 2008/12/17 END
      '            For i = 1 To GRD1.Rows - 1
      '                GRD1.row = i
      '                GRD1.col = 0
      '                If GRD1.CellBackColor = &HFFC0C0 Then
      '                    GRD1.TextMatrix(i, 0) = textSR03.Text
      '                    GRD1.TextMatrix(i, 1) = textSR04.Text
      '                    GRD1.TextMatrix(i, 2) = textSR05.Text
      '                    GRD1.TextMatrix(i, 3) = ChangeTStringToTDateString(textSR06.Text)
      '                    GRD1.TextMatrix(i, 4) = textSR07.Text
      '                    GRD1.TextMatrix(i, 5) = IIf(chkSR08.Value, "Y", "")
      '                    GRD1.TextMatrix(i, 6) = IIf(chkSR13.Value, "Y", "")
      '                    GRD1.TextMatrix(i, 7) = textSR09.Text
      '                    GRD1.TextMatrix(i, 8) = textSR10.Text
      '                    GRD1.TextMatrix(i, 9) = textSR11.Text
      '                    GRD1.TextMatrix(i, 10) = ChangeTStringToTDateString(textSR12.Text)
      '                    GRD1.Refresh
      '                    Exit For
      '                End If
      '            Next i
      '        Else
      '            ' 2008/12/17 MODIFY BY SINDY
      '            If textSR03.Text = "" Then
      '               textSR03.SetFocus
      '            ElseIf textSR04.Text = "" Then
      '               textSR04.SetFocus
      '            ElseIf textSR05.Text = "" Then
      '               textSR05.SetFocus
      '            End If
      '            'MsgBox "ºÙ¿×¤Î©m¦W¤£¥iªÅ¥Õ¡I", vbExclamation, "µo¥Í¿ù»~¡I"
      '            MsgBox "ºÙ¿×¤Î©m¦W¤Î©Ê§O¤£¥iªÅ¥Õ¡I", vbExclamation, "µo¥Í¿ù»~¡I"
      '            ' 2008/12/17 END
      '        End If
      'Case 2
      '        If textSR03.Text <> "" And textSR04.Text <> "" Then
      '            For i = 1 To grd1.Rows - 1
      '                grd1.row = i
      '                grd1.col = 0
      '                If grd1.CellBackColor = &HFFC0C0 Then
      '                    ' 2008/12/16 MODIFY BY SINDY
      '                    'grd1.RemoveItem i
      '                    If grd1.Rows = 2 Then
      '                        grd1.Clear
      '                    Else
      '                        grd1.RemoveItem i
      '                    End If
      '                    ' 2008/12/16 END
      '                    ClearSR
      '                    grd1.Refresh
      '                    Exit For
      '                End If
      '            Next i
      '        End If
      Case 0 '·s¼W
         iLstSelRow = -1
         For i = 1 To GRD1.Rows - 1
            GRD1.row = i
            GRD1.col = 0
            If GRD1.CellBackColor = &HFFC0C0 Then
               iLstSelRow = i
               Exit For
            End If
         Next
         ClearSR
         
         If GRD1.TextMatrix(GRD1.Rows - 1, 0) <> "" Then
            GRD1.Rows = GRD1.Rows + 1
         End If
         GRD1.row = GRD1.Rows - 1
         grd1_SelChange
         
         EnableSR True
         cmdOK(0).Enabled = False
         cmdOK(1).Enabled = True
         cmdOK(2).Enabled = True
         cmdOK(3).Enabled = False
         cmdOK(4).Enabled = False
         'cmdLog.Enabled = False
         GRD1.Enabled = False
         m_EditMode2 = 1
      Case 1 '½T©w
         For i = 1 To GRD1.Rows - 1
            GRD1.row = i
            GRD1.col = 0
            If GRD1.CellBackColor = &HFFC0C0 Then
               '2011/10/18 add by sonia
               If textSR11.Text <> "" Then
                  If CheckTaiwanAddr(textSR11.Text, "000", "²²ÄÝ¦a§}") = False Then
                     textSR11.SetFocus
                     textSR11_GotFocus
                     Exit For
                  End If
               End If
               '2011/10/18 end
               
               If textSR03.Text <> "" And textSR04.Text <> "" And textSR05.Text <> "" Then
                  GRD1.TextMatrix(i, 0) = textSR03.Text
                  If textSR12.Text <> "" Then
                     GRD1.TextMatrix(i, 2) = "§R"
                  Else
                     GRD1.TextMatrix(i, 2) = ""
                  End If
                  GRD1.TextMatrix(i, 1) = textSR04.Text
                  GRD1.TextMatrix(i, 3) = textSR05.Text
                  GRD1.TextMatrix(i, 4) = ChangeTStringToTDateString(textSR06.Text)
                  GRD1.TextMatrix(i, 5) = textSR07.Text
                  GRD1.TextMatrix(i, 6) = IIf(chkSR08.Value, "Y", "")
                  GRD1.TextMatrix(i, 7) = IIf(chkSR13.Value, "Y", "")
                  GRD1.TextMatrix(i, 8) = textSR09.Text
                  GRD1.TextMatrix(i, 9) = textSR10.Text
                  GRD1.TextMatrix(i, 10) = textSR11.Text
                  GRD1.TextMatrix(i, 11) = ChangeTStringToTDateString(textSR12.Text)
                  If chkSR08.Value Then
                     If cboHL05.ListIndex > 0 Then
                        GRD1.TextMatrix(i, 13) = Left(cboHL05, 2)
                     Else
                        GRD1.TextMatrix(i, 13) = ""
                     End If
                  End If
                  GRD1.Refresh
                  
                  EnableSR False
                  cmdOK(0).Enabled = True
                  cmdOK(1).Enabled = False
                  cmdOK(2).Enabled = False
                  cmdOK(3).Enabled = True
                  cmdOK(4).Enabled = True
                  'cmdLog.Enabled = True
                  m_EditMode2 = 0
               Else
                  If textSR03.Text = "" Then
                     textSR03.SetFocus
                  ElseIf textSR04.Text = "" Then
                     textSR04.SetFocus
                  ElseIf textSR05.Text = "" Then
                     textSR05.SetFocus
                  End If
                  MsgBox "ºÙ¿×¤Î©m¦W¤Î©Ê§O¤£¥iªÅ¥Õ¡I", vbExclamation, "µo¥Í¿ù»~¡I"
               End If
               Exit For
            End If
         Next
         GRD1.Enabled = True
         
      Case 2 '¨ú®ø
         If GRD1.TextMatrix(GRD1.Rows - 1, 0) = "" Then
            If GRD1.Rows = 2 Then
               GRD1.Clear
               SetGrd
            Else
               GRD1.RemoveItem GRD1.Rows - 1
            End If
         End If
         
         If iLstSelRow >= 0 Then
            GRD1.row = iLstSelRow
            GRD1.col = 0
            GRD1.CellBackColor = QBColor(15) 'ÃC¦âÁÙ­ì¥H«K­«·sÅª¨ú¸ê®Æ
            grd1_SelChange
         Else
            ClearSR
         End If
         
         EnableSR False
         cmdOK(0).Enabled = True
         cmdOK(1).Enabled = False
         cmdOK(2).Enabled = False
         cmdOK(3).Enabled = True
         cmdOK(4).Enabled = True
         'cmdLog.Enabled = True
         GRD1.Enabled = True
         m_EditMode2 = 0
         
      Case 3 '§R°£
         For i = 1 To GRD1.Rows - 1
             GRD1.row = i
             GRD1.col = 0
             If GRD1.CellBackColor = &HFFC0C0 Then
               If GRD1.TextMatrix(i, 12) <> "" And GRD1.TextMatrix(i, 6) = "Y" Then
                  MsgBox "¡i" & GRD1.TextMatrix(i, 1) & "¡j¬OÂÂ¦³°·«O²²ÄÝ¡A½Ð¥ý¶i¡i°·«O²§°Ê¸ê®Æ¡j§@°·«O²¾¥X¡I"
                  Exit For
               Else
                  If GRD1.TextMatrix(i, 12) <> "" Then
                     textSR12 = strSrvDate(2)
                     GRD1.TextMatrix(i, 11) = ChangeTStringToTDateString(textSR12)
                     GRD1.TextMatrix(i, 2) = "§R"
                     GRD1.Refresh
                  Else
                     If GRD1.Rows = 2 Then
                         GRD1.Clear
                         SetGrd
                     Else
                         GRD1.RemoveItem i
                     End If
                     ClearSR
                     GRD1.Refresh
                  End If
                  Exit For
               End If
             End If
         Next i
         m_EditMode2 = 0
         
      Case 4 '­×§ï
         For i = 1 To GRD1.Rows - 1
            GRD1.row = i
            GRD1.col = 0
            If GRD1.CellBackColor = &HFFC0C0 Then
               iLstSelRow = i
               EnableSR True, GRD1.TextMatrix(i, 12)
               cmdOK(0).Enabled = False
               cmdOK(1).Enabled = True
               cmdOK(2).Enabled = True
               cmdOK(3).Enabled = False
               cmdOK(4).Enabled = False
               'cmdLog.Enabled = False
               GRD1.Enabled = False
               m_EditMode2 = 2
               Exit For
            End If
         Next i
      'end 2009/6/10
      'Add By Sindy 2012/6/20
      Case 5 'µ²§ô
         Unload Me
         UpForm.Show
      Case 6 '¬d¸ß¤U¤@µ§
         Unload Me
         UpForm.cmdState = 0
         UpForm.PubShowNextData
      '2012/6/20 End
      Case Else
   End Select
End Sub

Private Sub cmdLog_Click()
Dim i As Integer, strSR08 As String, strHL05 As String
   
   For i = 1 To GRD1.Rows - 1
      GRD1.row = i
      GRD1.col = 0
      If GRD1.CellBackColor = &HFFC0C0 Then
         If GRD1.TextMatrix(i, 12) = "" Then
            MsgBox "¡i" & textSR04 & "¡j©|¥¼¦sÀÉ¡AµLªkºûÅ@°·«O²§°Ê¸ê®Æ¡I"
         Else
            With frm160001_1
               .strHL01 = textST01
               .strHL02 = GRD1.TextMatrix(i, 12)
               .InitForm
               .Show vbModal
               GetNewHiData textST01, GRD1.TextMatrix(i, 12), strSR08, strHL05
               GRD1.TextMatrix(i, 6) = strSR08
               If strSR08 = "Y" Then
                  chkSR08.Value = 1
               Else
                  chkSR08.Value = 0
               End If
               GRD1.TextMatrix(i, 13) = strHL05
               SelCombo cboHL05, strHL05
            End With
         End If
         Exit For
      End If
   Next
   If i = GRD1.Rows Then
      MsgBox "½Ð¥ýÂI¿ï²²ÄÝ¸ê®Æ¡I"
   End If
End Sub

'Add by Morgan 2009/6/11
'Åª¨ú³Ì·sªº°·«O¸ê®Æ
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

'Add By Sindy 2012/6/14
Private Sub Command1_Click()
   If textST01 = "" Then
      MsgBox "½Ð¿é¤J­û¤u¥N¸¹¡I"
      Exit Sub
   End If
   frmPic001.oCP01 = "000"
   frmPic001.oCP02 = textST01
   frmPic001.oCP03 = "0"
   frmPic001.oCP04 = "00"
   frmPic001.strWorkType = "1"
   frmPic001.Label11 = "­û¤u·Ó¤ù¤W¶Ç"
   If m_EditMode <> 1 And m_EditMode <> 2 Then
      frmPic001.bolQuery = True '¥u¬d¸ß
   Else
      frmPic001.bolQuery = False '¥i¦sÀÉ
   End If
   frmPic001.StrMenu
   'frmPic001.CanScan
   frmPic001.SetSeekCmdok 'Add by Amy 2018/07/19
   frmPic001.Show vbModal
   Call ReadPhoto '¸ü¤J·Ó¤ù
End Sub

Private Sub Form_Initialize()
   Set rsA = New ADODB.Recordset
   If rsA.State = 1 Then rsA.Close
   rsA.CursorLocation = adUseClient
   rsA.Open "select * from staff where rownum <2 ", cnnConnection, adOpenStatic, adLockReadOnly
   tf_st = rsA.Fields.Count
   InitialField rsA 'Added by Morgan 2023/12/15
   SetGrd
End Sub

' «ö¤U«öÁä
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   'Add By Sindy 2014/8/29 ·ífocus¦b¤U¦CÄæ¦ì®É,«öenterÁäºû«ù´«¦æ¥\¯à¦Ó¤£¬O¦sÀÉ¥\¯à
   If KeyCode = vbKeyReturn And _
      (UCase(Me.ActiveControl.Name) = UCase("textSS02") Or _
       UCase(Me.ActiveControl.Name) = UCase("textSS03") Or _
       UCase(Me.ActiveControl.Name) = UCase("textSS04") Or _
       UCase(Me.ActiveControl.Name) = UCase("textSS05") Or _
       UCase(Me.ActiveControl.Name) = UCase("textSS06") Or _
       UCase(Me.ActiveControl.Name) = UCase("textSS07")) Then
      Exit Sub
   End If
   '2014/8/29 END
   
   Select Case KeyCode
      ' ·s¼W
      Case vbKeyF2:
         If m_bInsert Then
            If m_EditMode = 0 Then
               OnAction KeyCode
               KeyCode = 0
            End If
         End If
      ' ­×§ï
      Case vbKeyF3:
         If m_bUpdate Then
            If m_EditMode = 0 Then
               OnAction KeyCode
               KeyCode = 0
            End If
         End If
      ' ¬d¸ß
      Case vbKeyF4:
         If m_bQuery Then
            If m_EditMode = 0 Then
               OnAction KeyCode
               KeyCode = 0
            End If
         End If
      ' §R°£
      Case vbKeyF5:
         If m_bDelete Then
            If m_EditMode = 0 Then
               OnAction KeyCode
               KeyCode = 0
            End If
         End If
      ' ²Ä¤@µ§, ¤W¤@µ§, ¤U¤@µ§, ³Ì«á¤@µ§
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
'edit by nickc 2006/11/13
'      Case vbKeyReturn:
'         If m_EditMode <> 0 Then
'            OnAction vbKeyF9
'         End If
      Case vbKeyEscape:
         If m_EditMode = 0 Then
            OnAction KeyCode
         Else
            OnAction vbKeyF10
         End If
   End Select
End Sub
'add by nickc 2006/11/13 Enter ¨Æ¥ó¡Aµ¥©ó¦sÀÉ¡A°µ§¹¨ú®ø¡A¤£µM form ¤º¨ä¥Lª«¥ó¦³¼g keycode ©Î¬O keyascii ¨Æ¥óªº¸Ü¡A¤]·|°µ¨ì
Private Sub Form_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
      Case vbKeyReturn:
         If m_EditMode <> 0 Then
            KeyAscii = 0
            OnAction vbKeyF9
         End If
    End Select
End Sub

Private Sub Form_Load()
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset

   'ReDim m_FieldList(tf_st) As FIELDITEM 'Removed by Morgan 2023/12/15

   If Pub_StrUserSt03 = "M31" Then
      m_bInsert = False
      m_bUpdate = False
      m_bDelete = False
   Else
      m_bInsert = IsUserHasRightOfFunction(Me.Name, strAdd, False)
      m_bUpdate = IsUserHasRightOfFunction(Me.Name, strEdit, False)
      m_bDelete = IsUserHasRightOfFunction(Me.Name, strDel, False)
   End If
   
   m_bQuery = IsUserHasRightOfFunction(Me.Name, strFind, False)
   
   
   textST01.BackColor = &H8000000F
   
   MoveFormToCenter Me
   
   'InitialField 'Removed by Morgan 2023/12/15
   InitialData
   RefreshRange
   ShowFirstRecord
   UpdateToolbarState
   SetCtrlReadOnly True
   OnAction vbKeyF4
   Me.SSTab1.Tab = 0
   'add by sonia 2016/5/3 ¤H¨Æ³B¤Î¹q¸£¤¤¤ß¤~¥i¥H¬Ý¨ì§ë«Oª÷ÃBªº­¶ÅÒ
   'Modify By Sindy 2019/10/28 ¼B¸g²z:¥u¶}©ñ¼B¸g²z¥i¥H¬Ý
   SSTab1.TabVisible(6) = False 'Add By Sindy 2019/10/28 ¼È®É¤£¶}©ñ
'   If (Pub_StrUserSt03 = "M21" And strUserNum = "68010") Or _
'      Pub_StrUserSt03 = "M51" Then
'      SSTab1.TabVisible(6) = True
'   Else
'      SSTab1.TabVisible(6) = False
'   End If
   'end 2016/5/3
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm160001 = Nothing
End Sub

Private Sub grd1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim nCol As Long, nRow As Long
   
   getGrdColRow GRD1, x, y, nCol, nRow
   GRD1.col = nCol
   GRD1.row = nRow
End Sub

Private Sub grd1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
'Debug.Print "UP"
End Sub

Private Sub grd1_SelChange()
'Debug.Print "change"
Dim tmpMouseRow
Dim i, j

   GRD1.Visible = False
   tmpMouseRow = GRD1.row
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
            If m_EditMode <> 0 Then
               'Remove by Morgan 2009/6/12
               'cmdOK(1).Enabled = True
               'cmdOK(2).Enabled = True
               'end 2009/6/12
            End If
            textSR03.Text = GRD1.TextMatrix(tmpMouseRow, 0)
            textSR04.Text = GRD1.TextMatrix(tmpMouseRow, 1)
            'Modify by Morgan 2009/6/10 +ª¬ºAÄæ¦ì,«á­±Äæ¦ì§Ç¦¸+1
            textSR05.Text = GRD1.TextMatrix(tmpMouseRow, 3)
            textSR06.Text = ChangeTDateStringToTString(GRD1.TextMatrix(tmpMouseRow, 4))
            textSR07.Text = GRD1.TextMatrix(tmpMouseRow, 5)
            chkSR08.Value = IIf(GRD1.TextMatrix(tmpMouseRow, 6) = "Y", vbChecked, vbUnchecked)
            chkSR13.Value = IIf(GRD1.TextMatrix(tmpMouseRow, 7) = "Y", vbChecked, vbUnchecked)
            textSR09.Text = GRD1.TextMatrix(tmpMouseRow, 8)
            textSR10.Text = GRD1.TextMatrix(tmpMouseRow, 9)
            textSR11.Text = GRD1.TextMatrix(tmpMouseRow, 10)
            textSR12.Text = ChangeTDateStringToTString(GRD1.TextMatrix(tmpMouseRow, 11))
            SelCombo cboHL05, GRD1.TextMatrix(tmpMouseRow, 13) 'Add by Morgan 2009/6/24
            GRD1.Visible = True
       End If
   End If
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
   Select Case SSTab1.Tab
      Case 0
        textST03.SelStart = 1
        textST03.SelLength = 0
        textST06.SelStart = 1
        textST06.SelLength = 0
        textST20.SelStart = 1
        textST20.SelLength = 0
        textST21.SelStart = 1
        textST21.SelLength = 0
        textST22.SelStart = 1
        textST22.SelLength = 0
        textST24.SelStart = 1
        textST24.SelLength = 0
        textST25.SelStart = 1
        textST25.SelLength = 0
        textST27.SelStart = 1
        textST27.SelLength = 0
      Case 1
        textST37.SelStart = 1
        textST37.SelLength = 0
      Case 2
        textSR03.SelStart = 1
        textSR03.SelLength = 0
        textSR05.SelStart = 1
        textSR05.SelLength = 0
      Case 3
      Case Else
   End Select
End Sub

Private Sub Tbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim strTit As String
Dim strMsg As String
Dim nResponse
   
   Select Case Button.Index
      ' ·s¼W
      Case 1: OnAction vbKeyF2
      ' ­×§ï
      Case 2: OnAction vbKeyF3
      ' §R°£
      Case 3: OnAction vbKeyF5
      ' ¬d¸ß
      Case 4: OnAction vbKeyF4
      ' ²Ä¤@µ§
      Case 6: OnAction vbKeyHome
      ' «e¤@µ§
      Case 7: OnAction vbKeyPageUp
      ' «á¤@µ§
      Case 8: OnAction vbKeyPageDown
      ' ³Ì«á¤@µ§
      Case 9: OnAction vbKeyEnd
      ' ½T©w
      Case 11: OnAction vbKeyF9
      ' ¨ú®ø
      Case 12: OnAction vbKeyF10
      ' Â÷¶}
      Case 14: OnAction vbKeyEscape
   End Select
End Sub

Private Sub ShowMsg(ByVal St As String)
   MsgBox St, vbInformation
End Sub

' §ó·s Create ¤Î Update ªº¤H
Private Sub UpdateCUID(ByRef rsSrcTmp As ADODB.Recordset)
Dim strTemp As String
Dim strCName As String
Dim strCDate As String
Dim strCTime As String
Dim strUName As String
Dim strUDate As String
Dim strUTime As String
   
   If IsNull(rsSrcTmp.Fields("st43")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("st43")) = False Then
         strCName = GetStaffName(rsSrcTmp.Fields("st43"), True)
      End If
   End If
   If IsNull(rsSrcTmp.Fields("st44")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("st44")) = False Then
         strTemp = TAIWANDATE(rsSrcTmp.Fields("st44"))
         strCDate = Format(strTemp, "###/##/##")
      End If
   End If
   If IsNull(rsSrcTmp.Fields("st45")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("st45")) = False Then
         strTemp = rsSrcTmp.Fields("st45")
         strCTime = Format(strTemp, "##:##")
      End If
   End If
   If IsNull(rsSrcTmp.Fields("st46")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("st46")) = False Then
         strUName = GetStaffName(rsSrcTmp.Fields("st46"), True)
      End If
   End If
   If IsNull(rsSrcTmp.Fields("st47")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("st47")) = False Then
         strTemp = TAIWANDATE(rsSrcTmp.Fields("st47"))
         strUDate = Format(strTemp, "###/##/##")
      End If
   End If
   If IsNull(rsSrcTmp.Fields("st48")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("st48")) = False Then
         strTemp = rsSrcTmp.Fields("st48")
         strUTime = Format(strTemp, "##:##")
      End If
   End If
   
   ' ³]©wCUID¤¤ªº¤å¦r
   Label23.Caption = "CREATE : " & strCName & " " & _
              " " & strCDate & " " & _
              " " & strCTime & String(10, " ") & _
              "UPDATE : " & strUName & " " & _
              " " & strUDate & " " & _
              " " & strUTime
              
End Sub

Private Function TxtValidate() As Boolean
Dim objTxt As Object
Dim ii As Integer
Dim Cancel As Boolean

   TxtValidate = False
   If Me.textST01.Enabled = True Then
      Cancel = False
      textST01_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If Me.textST02.Enabled = True Then
      Cancel = False
      textST02_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If Me.textST03.Enabled = True Then
      Cancel = False
      textST03_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If Me.textST06.Enabled = True Then
      Cancel = False
      textST06_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If Me.textST08.Enabled = True Then
      Cancel = False
      textST08_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
      '2011/10/18 add by sonia
      If CheckTaiwanAddr(textST08, "000", "³q°T¦a§}") = False Then
         Cancel = True
         Exit Function
      End If
      '2011/10/18 end
   End If
   If Me.textST09.Enabled = True Then
      Cancel = False
      textST09_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If Me.textST10.Enabled = True Then
      Cancel = False
      textST10_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If Me.textST13.Enabled = True Then
      Cancel = False
      textST13_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If Me.textST18.Enabled = True Then
      Cancel = False
      textST18_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If Me.textST19.Enabled = True Then
      Cancel = False
      textST19_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If Me.textST20.Enabled = True Then
      Cancel = False
      textST20_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If Me.textST21.Enabled = True Then
      Cancel = False
      textST21_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If Me.textST22.Enabled = True Then
      Cancel = False
      textST22_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If Me.textST23.Enabled = True Then
      Cancel = False
      textST23_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If Me.textST24.Enabled = True Then
      'Add by Morgan 2011/1/18
      If textST24 = "" Then
         MsgBox "°êÄy¤£¥iªÅ¥Õ¡I"
         Exit Function
      End If
      
      Cancel = False
      textST24_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If Me.textST25.Enabled = True Then
      Cancel = False
      textST25_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If Me.textST26.Enabled = True Then
      Cancel = False
      textST26_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If Me.textST27.Enabled = True Then
      Cancel = False
      textST27_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If Me.textST28.Enabled = True Then
      Cancel = False
      textST28_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If Me.textST29.Enabled = True Then
      Cancel = False
      textST29_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If Me.textST31.Enabled = True Then
      Cancel = False
      textST31_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If Me.textST32.Enabled = True Then
      Cancel = False
      textST32_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If Me.textST33.Enabled = True Then
      Cancel = False
      textST33_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If Me.textST34.Enabled = True Then
      Cancel = False
      textST34_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
      '2011/10/18 add by sonia
      If CheckTaiwanAddr(textST34, "000", "¤áÄy¦a§}") = False Then
         Cancel = True
         Exit Function
      End If
      '2011/10/18 end
   End If
   If Me.textST35.Enabled = True Then
      Cancel = False
      textST35_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If Me.textST36.Enabled = True Then
      Cancel = False
      textST36_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If Me.textST37.Enabled = True Then
      Cancel = False
      textST37_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If Me.textST38.Enabled = True Then
      Cancel = False
      textST38_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If Me.textST39.Enabled = True Then
      Cancel = False
      textST39_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If Me.textST40.Enabled = True Then
      Cancel = False
      textST40_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If Me.textST41.Enabled = True Then
      Cancel = False
      textST41_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If Me.textST42.Enabled = True Then
      Cancel = False
      textST42_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If Me.textST49.Enabled = True Then
      Cancel = False
      textST49_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   'Add By Sindy 2014/3/12
   If Me.textSS02.Enabled = True Then
      Cancel = False
      textSS02_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If Me.textSS03.Enabled = True Then
      Cancel = False
      textSS03_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If Me.textSS04.Enabled = True Then
      Cancel = False
      textSS04_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If Me.textSS05.Enabled = True Then
      Cancel = False
      textSS05_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If Me.textSS06.Enabled = True Then
      Cancel = False
      textSS06_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If Me.textSS07.Enabled = True Then
      Cancel = False
      textSS07_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   '2014/3/12 END
   If Me.textST12.Enabled = True Then
      Cancel = False
      textST12_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If Me.txtSD17.Enabled = True Then
      Cancel = False
      txtSD17_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   'Added by Morgan 2013/1/29
   '¤ºÂ½¤H­û¥²¶·¿é¤J¨­¤ÀÃÒ¸¹¥H«Kµ¹¥IÂ½Ä¶¶O®É­pºâ¸É¥R«O¶O
   If Left(textST03, 3) = "F52" And textST26 = "" Then
      MsgBox "¤ºÂ½¤H­û¥²¶·¿é¤J¨­¤ÀÃÒ¸¹¡I", vbExclamation
      textST26.SetFocus
      Exit Function
   End If
   'end 2013/1/29
   
   'Add by Sindy 2021/6/15 ÀË¬dµe­±ªº TextBox, ComboBox ¬O§_§t¦³Unicode¤å¦r
   If PUB_ChkUniText(Me) = False Then
      Exit Function
   End If
   '2021/6/15 END

   TxtValidate = True
End Function

'add by nickc 2006/10/24
' ³]©wÄæ¦ìªº¤º®e
Private Sub SetFieldNewData(ByVal strName As String, Optional ByVal strData As String = "#==#")
Dim nIndex As Integer
   
   For nIndex = 0 To tf_st - 1 'edit by nickc 2006/10/24  MAX_FIELD - 1
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

' ±q°O¿ý¤¤§ó·sÄæ¦ì¤º®e
Private Sub UpdateFieldOldData(ByRef rsTmp As ADODB.Recordset)
Dim nIndex As Integer
Dim strTmp As String
   
   For nIndex = 0 To tf_st - 1
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

' ·s¼W°O¿ý
Private Function AddRecord() As Boolean
Dim strSql As String
Dim strTmp As String
Dim strTit As String
Dim strMsg As String
Dim nResponse
Dim nIndex As Integer
Dim bDifference As Boolean
Dim bFirst As Boolean
Dim strST01 As String
Dim MyArr As Variant
'Dim m_SD09 As String     '2009/1/1 ¹w³]±B³à¤¬§U 'Remvoed by Morgan 2025/7/29 114/7/28°_¼o¤î±B³à¤¬§U¿ìªk
Dim oMailCount As String ' Add By Sindy 98/03/02
Dim strTemp As String
   
   AddRecord = False
   
   strST01 = textST01

   ' ÀË¬d°O¿ý¬O§_¤w¦s¦b
   If IsRecordExist(strST01) = True Then
      strTit = "·s¼W¸ê®Æ"
      strMsg = "¸Óµ§°O¿ý¤w¦s¦b"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      UpdateCtrlData
      Exit Function
   End If
   
   bFirst = True
   bDifference = False
   strSql = "INSERT INTO staff ("
   For nIndex = 0 To tf_st - 1
      strTmp = Empty
      If m_FieldList(nIndex).fiOldData <> m_FieldList(nIndex).fiNewData Then
         strTmp = m_FieldList(nIndex).fiName
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
   Next nIndex
   strSql = strSql & ") "
   strSql = strSql & "VALUES ("
   
   bFirst = True
   For nIndex = 0 To tf_st - 1
      strTmp = Empty
      If m_FieldList(nIndex).fiOldData <> m_FieldList(nIndex).fiNewData Then
         If m_FieldList(nIndex).fiType = 0 Then
            strTmp = "'" & ChgSQL(m_FieldList(nIndex).fiNewData) & "'"
         Else
            strTmp = m_FieldList(nIndex).fiNewData
         End If
      End If
      If strTmp <> Empty Then
         If bFirst = True Then
            strSql = strSql & strTmp
            bFirst = False
         Else
            strSql = strSql & "," & strTmp
         End If
      End If
   Next nIndex
   strSql = strSql & ")"
   
On Error GoTo ErrHand
   cnnConnection.BeginTrans
   
   Pub_SeekTbLog strSql
   cnnConnection.Execute strSql
   
'***********************************************************
'          ¿ËÄÝ¸ê®Æ²§°Ê
'***********************************************************
     Pub_SeekTbLog "delete from staff_relation where sr01='" & strST01 & "' "
     cnnConnection.Execute "delete from staff_relation where sr01='" & strST01 & "' "
        For nIndex = 1 To GRD1.Rows - 1
            If GRD1.TextMatrix(nIndex, 0) <> "" Then
                strSql = "insert into staff_relation (sr01,sr02,sr03,sr04,sr05,sr06,sr07,sr08,sr13,sr09,sr10,sr11,sr12) " & _
                              " select '" & strST01 & "',nvl(max(sr02),0)+1,"
                              
                MyArr = Split(GRD1.TextMatrix(nIndex, 0), " ")
                strSql = strSql & CNULL(Trim(MyArr(0))) & ","
                strSql = strSql & CNULL(ChgSQL(GRD1.TextMatrix(nIndex, 1))) & ","
                'Modify by Morgan 2009/6/10 +ª¬ºAÄæ,©Ò¦³«á­±ªº§Ç¦¸+1
                MyArr = Split(GRD1.TextMatrix(nIndex, 3), " ")
                strSql = strSql & CNULL(Trim(MyArr(0))) & ","
                strSql = strSql & CNULL(DBDATE(GRD1.TextMatrix(nIndex, 4))) & ","
                strSql = strSql & CNULL(ChgSQL(GRD1.TextMatrix(nIndex, 5))) & ","
                strSql = strSql & CNULL(ChgSQL(Trim(GRD1.TextMatrix(nIndex, 6)))) & ","
                strSql = strSql & CNULL(ChgSQL(Trim(GRD1.TextMatrix(nIndex, 7)))) & ","
                strSql = strSql & CNULL(ChgSQL(GRD1.TextMatrix(nIndex, 8))) & ","
                strSql = strSql & CNULL(ChgSQL(GRD1.TextMatrix(nIndex, 9))) & ","
                strSql = strSql & CNULL(ChgSQL(GRD1.TextMatrix(nIndex, 10))) & ","
                strSql = strSql & CNULL(DBDATE(GRD1.TextMatrix(nIndex, 11))) & _
                " from staff_relation where '" & strST01 & "'=sr01(+)  "
                Pub_SeekTbLog strSql
                cnnConnection.Execute strSql
                
                'Add by Morgan 2009/6/11
                '·s¼W°·«O²²ÄÝ®É¦P®É·s¼W¤@µ§²§°Ê¸ê®Æ
                If GRD1.TextMatrix(nIndex, 13) = "" And GRD1.TextMatrix(nIndex, 6) = "Y" Then
                  strSql = "INSERT INTO HiRelationLog (HL01,HL02,HL03,HL04,HL05)" & _
                     " VALUES ('" & strST01 & "'," & nIndex & "," & strSrvDate(1) & ",'1'," & CNULL(GRD1.TextMatrix(nIndex, 13)) & ")"
                  Pub_SeekTbLog strSql
                  cnnConnection.Execute strSql
                End If
            End If
        Next nIndex
        
   '2008/12/29 ADD BY SONIA
   '·s­û¤u¥B«DF½s¸¹ªÌ¼g¤H¨Æ²§°Ê¸ê®Æ01·s¶i
   If Mid(textST01, 1, 1) <> "F" Then
      'Modify By Sindy 2024/9/9 ¦]³æ¦ì§ïST93; CNULL(m_FieldList(2).fiNewData) => CNULL(Trim(Left(textST03.Text, 4)))
'      strSql = "insert into staff_change (sc01,sc02,sc03,sc04,sc05,sc06,sc07) " & _
'               "values (" & CNULL(m_FieldList(0).fiNewData) & "," & CNULL(m_FieldList(12).fiNewData) & ",'01'," & CNULL(m_FieldList(2).fiNewData) & "," & CNULL(m_FieldList(19).fiNewData) & "," & CNULL(m_FieldList(20).fiNewData) & "," & CNULL(m_FieldList(48).fiNewData) & ")"
      strSql = "insert into staff_change (sc01,sc02,sc03,sc04,sc05,sc06,sc07) " & _
               "values (" & CNULL(m_FieldList(0).fiNewData) & "," & CNULL(m_FieldList(12).fiNewData) & ",'01'," & CNULL(Trim(Left(textST03.Text, 4))) & "," & CNULL(m_FieldList(19).fiNewData) & "," & CNULL(m_FieldList(20).fiNewData) & "," & CNULL(m_FieldList(48).fiNewData) & ")"
      Pub_SeekTbLog strSql
      cnnConnection.Execute strSql
   End If
   '2008/12/29 END
   
   'Add By Sindy 2014/3/12 +­û¤u±Mªø¸ê®Æ
   If Trim(textSS02) <> "" Or Trim(textSS03) <> "" Or Trim(textSS04) <> "" Or _
      Trim(textSS05) <> "" Or Trim(textSS06) <> "" Or Trim(textSS07) <> "" Then
      strSql = "INSERT INTO staff_specialty(ss01,ss02,ss03,ss04,ss05,ss06,ss07)" & _
               " VALUES(" & CNULL(strST01) & "," & CNULL(textSS02) & "," & CNULL(textSS03) & _
               "," & CNULL(textSS04) & "," & CNULL(textSS05) & "," & CNULL(textSS06) & _
               "," & CNULL(textSS07) & ")"
      Pub_SeekTbLog strSql
      cnnConnection.Execute strSql
   End If
   '2014/3/12 END
   
   '2009/1/1 add by sonia ¦A¼gÁ~¸ê°ò¥»ÀÉ
   'Modified by Morgan 2024/4/30 +¦X¬ùÂ½Ä¶¤H­û(²Ä5½X>='A')°£¥~
   If Mid(textST01, 1, 1) <> "F" And Right(textST01, 1) < "A" Then
      'Modified by Morgan 2025/7/29 114/7/28°_¼o¤î±B³à¤¬§U¿ìªk
      'If PUB_GetHelpFee(m_FieldList(0).fiNewData, strExc(1)) = True Then
      '   m_SD09 = strExc(1) '±B³à¤¬§U
      'End If
      'strSql = "insert into salarydata (sd01,sd02,sd04,sd05,sd09,sd10,sd16,sd17) " & _
               "values (" & CNULL(m_FieldList(0).fiNewData) & ",'T','Y','1'," & CNULL(m_SD09) & "," & CNULL(m_SD09) & ",'Y'," & Val(txtSD17) & ")"
      strSql = "insert into salarydata (sd01,sd02,sd04,sd05,sd16,sd17) " & _
               "values (" & CNULL(m_FieldList(0).fiNewData) & ",'T','Y','1','Y'," & Val(txtSD17) & ")"
      'end 2025/7/29
   Else
      'Modified by Morgan 2013/3/7 ¨ú®øsd11(¸ÓÄæ¦ì¤w§ï¨ä¥L¥Î³~)
      'strSql = "insert into salarydata (sd01,sd02,sd04,sd05,sd11,sd19) " & _
               "values (" & CNULL(m_FieldList(0).fiNewData) & ",decode(" & CNULL(m_FieldList(2).fiNewData) & ",'F51','F','F52','P','T'),'Y',decode(" & CNULL(m_FieldList(2).fiNewData) & ",'F51','7','F52','1','7'),'N','2')"
      'Modify By Sindy 2024/9/9 ¦]³æ¦ì§ïST93; CNULL(m_FieldList(2).fiNewData) => CNULL(Trim(Left(textST03.Text, 4)))
'      strSql = "insert into salarydata (sd01,sd02,sd04,sd05,sd19) " & _
'               "values (" & CNULL(m_FieldList(0).fiNewData) & ",decode(" & CNULL(m_FieldList(2).fiNewData) & ",'F51','F','F52','P','T'),'Y',decode(" & CNULL(m_FieldList(2).fiNewData) & ",'F51','7','F52','1','7'),'2')"
      strSql = "insert into salarydata (sd01,sd02,sd04,sd05,sd19) " & _
               "values (" & CNULL(m_FieldList(0).fiNewData) & ",decode(" & CNULL(Trim(Left(textST03.Text, 4))) & ",'F51','F','F52','P','T'),'Y',decode(" & CNULL(Trim(Left(textST03.Text, 4))) & ",'F51','7','F52','1','7'),'2')"
   End If
   Pub_SeekTbLog strSql
   cnnConnection.Execute strSql
   '2009/1/1 end
   
   '2009/1/6 ADD BY SONIA
   '·s­û¤u¥B«DF½s¸¹ªÌ¼gÁ~¸ê²§°Ê¸ê®Æ
   If Mid(textST01, 1, 1) <> "F" Then
      'Modified by Morgan 2025/7/29 114/7/28°_¼o¤î±B³à¤¬§U¿ìªk
      'strSql = "insert into salarylog (sl01,sl02,sl03,sl05,sl06,sl35) " & _
               "values (" & CNULL(m_FieldList(0).fiNewData) & "," & CNULL(m_FieldList(12).fiNewData) & ",'T'," & CNULL(m_SD09) & "," & CNULL(m_SD09) & ",'N')"
      strSql = "insert into salarylog (sl01,sl02,sl03,sl35) " & _
               "values (" & CNULL(m_FieldList(0).fiNewData) & "," & CNULL(m_FieldList(12).fiNewData) & ",'T','N')"
      'end 2025/7/29
      Pub_SeekTbLog strSql
      cnnConnection.Execute strSql
   End If
   '2009/1/6 END
   
   'ADD BY Sindy 98/03/10
   '·s­û¤u¥B«DF½s¸¹ªÌ¼g­û¤u±K½XÀÉ
   Dim strPWD As String
   'Modified by Morgan 2024/4/30 +¦X¬ùÂ½Ä¶¤H­û(²Ä5½X>='A')°£¥~
   If Mid(textST01, 1, 1) <> "F" And Right(textST01, 1) < "A" Then
      '2014/4/21 MODIFY BY SONIA ¹w³]±K½X§ï¬°­û¤u½s¸¹«e¤G½X+¥½¤G½X
      'strPWD = Encrypt(m_FieldList(0).fiNewData, True)
      strPWD = Encrypt(Left(m_FieldList(0).fiNewData, 2) & Right(m_FieldList(0).fiNewData, 2), True)
      '2015/12/22 modify by sonia +sp02(¦Psp03)
      strSql = "insert into staff_pwd (sp01,sp02,sp03) " & _
               "values (" & CNULL(m_FieldList(0).fiNewData) & "," & CNULL(strPWD) & "," & CNULL(strPWD) & ")"
      'Pub_SeekTbLog strSQL '¤£¼gLog¬O¦]¬°¦bCall¨ç¼Æ®É,·|§PÅªSP01Äæ¦ì¦WºÙ,¾É­P¥X²{¡u¸ê®ÆÄæªº´¡¤J­È¹L¤j¡v
      cnnConnection.Execute strSql
      'add by sonia 2016/1/6 ¨C¦~10/2(§t)¥H«á¨ìÂ¾ªÌ¤£°Ñ¥[¦~²×¦ÒÁZ
      If Val(Right(DBDATE(textST13), 4)) >= 1002 Then
         strSql = "INSERT INTO yearmerit (ym01,ym02,ym03) VALUES(" & Val(Left(DBDATE(textST13), 4)) & ",'*','" & textST01 & "')"
         Pub_SeekTbLog strSql
         cnnConnection.Execute strSql
      End If
      'end 2016/1/6
   End If
   '98/03/10 END
   
   If ((strST01) < (m_FirstKEY)) Or ((strST01) > (m_LastKEY)) Then
      RefreshRange
   End If
   cnnConnection.CommitTrans
   
   'Add By Sindy 2021/2/18
   '·s­û¤u¥B«DF½s¸¹ªÌ¼g¤H¨Æ²§°Ê¸ê®Æ01·s¶i,
   '¦P¤¯²§°Ê®É,½Ð¨t²Îµo³qª¾µ¹¤À¾÷ºûÅ@¤H­û
   'Modified by Morgan 2024/4/30 +¦X¬ùÂ½Ä¶¤H­û(²Ä5½X>='A')°£¥~
   If Mid(textST01, 1, 1) <> "F" And Right(textST01, 1) < "A" Then
      Call PUB_CallScMailTOM13(textST01, m_FieldList(12).fiNewData)
   End If
   '2021/2/18 END
   
   ShowCurrRecord strST01
   AddRecord = True
   
   '2008/12/18 add BY SONIA µomailµ¹83002¸É­û¤uÀÉ¨ä¥L¸ê®Æ
   ' Modify By Sindy 98/03/02
   oMailCount = ""
   oMailCount = Pub_GetSpecMan("¤H¨Æ²§°Ê¶l¥ó³qª¾")
   'PUB_SendMail strUserNum, "83002", "", "·s­û¤u¨ìÂ¾³qª¾¡I", "­û¤u½s¸¹¡G" + textST01 + " " + textST02
   '2010/11/1 modify by sonia ¥[³qª¾«Ø¥ß²Õ§O
   'modify by sonia 2014/11/13 ¥[³qª¾´¼Åv³¡¤H­û«Ø¥ßsalesno¸ê®Æ,¥ý¨ú©m¦W²Ä¤G¦r,­Y­«ÂÐ«h¨ú²Ä¤T¦r,§_«h°Ý°]°È³B
   'Modify By Sindy 2020/7/6 ¼W¥[±±ºÞ½Ð°²®É,¤£µoÂ¾¥N
   'modify by sonia 2022/4/26 ¼W¥[³qª¾«Ø¥ß®×¥óªí³æÃ±®Ö¤H­û³]©w¸ê®Æªº¤å¦r
   PUB_SendMail strUserNum, oMailCount, "", "·s­û¤u¨ìÂ¾³qª¾¡I", "­û¤u½s¸¹¡G" + textST01 + " " + textST02 & vbCrLf & _
      "³¡¡@ªù¡G" + textST03 & vbCrLf & _
      "­û¤u©ÒÄÝ©Ò§O¡G" + textST06 & vbCrLf & _
      "Â¾¡@¦ì¡G" + textST21 & vbCrLf & vbCrLf & vbCrLf & _
      "¢Þ¢á¡G³¡ªù¬° F11,F21,F51,F52,F81,P10,P11,P14, ­n¿é¤J²Õ§O¡C" & vbCrLf & _
      "¡@¡@¡@µL«H½cªÌ¡A¥²¶·³]©w¡e¤º³¡¶l¥ó¦¬¥ó­û¤u½s¸¹¡f¬°99997(¤£±H«H)¡C" & vbCrLf & vbCrLf & _
      "¡@¡@¡@´¼Åv³¡¤H­û½Ð¹q¸£¤¤¤ß«Ø¥ßsalesno¸ê®Æ,¥ý¨ú©m¦W²Ä¤G¦r,­Y­«ÂÐ«h¨ú²Ä¤T¦r,§_«h°Ý°]°È³B¡C" & vbCrLf & vbCrLf & _
      "¡@¡@¡@´¼Åv³¡¤H­û½Ð¹q¸£¤¤¤ß«Ø¥ß®×¥óªí³æÃ±®Ö¤H­û³]©w¸ê®Æ¡C", , , , , , , , , , True
' 98/03/02 End
   
   Call Chk_Staff_Relation(textST01) 'Add By Sindy 2024/11/18
   
   Exit Function
ErrHand:
    cnnConnection.RollbackTrans
    MsgBox " ·s¼W¥¢±Ñ¡I" & vbCrLf & Err.Description
End Function

'Add By Sindy 2024/11/18 ¥Ñ¨t²ÎÀË¬d­Y»P¥»©Ò¦bÂ¾­û¤u©m¦W¬Û¦P®É(¤£¦Ò¼{µê½s¸¹¤Î¥~Â½½s¸¹)¡A
'¥u­n¦³¤@µ§©m¦W¬Û¦P§Y¸ß°Ý¾Þ§@¤H­û­Y¿ï¾Ü¬O¥»©Ò¦P¤¯®É¡A«h¥Ñ¨t²ÎµoEMAIL³qª¾¡u¸Õ¥Î´Áº¡°lÂÜÁ~¸ê¤H­û¡v
Private Sub Chk_Staff_Relation(strSR01 As String)
Dim nResponse
Dim strTo As String, strSR04 As String
   
   strExc(0) = "select * from staff,Staff_Relation" & _
               " where SR01='" & strSR01 & "' and SR03='3' and st02=SR04" & _
               " and st01<'F' and st04='1'" & _
               " and substr(st01,1,1)>='6'" & _
               " and substr(st01,4,1)<>'9'" & _
               " order by st03,st01"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      strSR04 = RsTemp.Fields("SR04")
      Me.SSTab1.Tab = 2
      nResponse = MsgBox("¦¹°t°¸¬O§_¬°¥»©Ò¦P¤¯¡H", vbYesNo + vbCritical + vbDefaultButton2, "¸ß°Ý")
      If nResponse = vbYes Then
         strTo = Pub_GetSpecMan("¸Õ¥Î´Áº¡°lÂÜÁ~¸ê¤H­û")
         PUB_SendMail strUserNum, strTo, "", GetPrjSalesNM(strSR01) & "¤§°t°¸¬°¥»©Ò¦P¤¯¡A½Ð¨Ì³W©w½Õ¾ã±B³à¦©´Úªº³]©w¡C", _
                     "­û¤u½s¸¹¡G" + strSR01 & vbCrLf & _
                     "­û¤u©m¦W¡G" + textST02 & vbCrLf & _
                     "°t°¸©m¦W¡G" + strSR04 & vbCrLf, , , , , , , , , , True
      End If
   End If
End Sub

' ­×§ï°O¿ý
Private Function ModRecord() As Boolean
Dim strSql As String
Dim strTmp As String
Dim strTit As String
Dim strMsg As String
Dim nResponse
Dim nIndex As Integer
Dim bDifference As Boolean
Dim bFirst As Boolean
Dim strST01 As String
Dim MyArr As Variant
Dim strTemp As String

   ModRecord = False
   
   strST01 = m_CurrKEY
   
   strSql = "begin user_data.user_enabled:=1; UPDATE staff SET "

   bFirst = True
   bDifference = False
   For nIndex = 0 To tf_st - 1
      strTmp = Empty
      'If nIndex < 42 Or nIndex > 47 Then
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
        'End If
   Next nIndex

   strSql = strSql & " " & _
                  "WHERE ST01 = '" & strST01 & "' ; end; "
On Error GoTo ErrHand
      cnnConnection.BeginTrans
   If bDifference = True Then
      Pub_SeekTbLog strSql
      cnnConnection.Execute strSql
      
      'Add By Sindy 2018/5/16 ¶}©ñ¥i¥H­×§ï,¦ý­n¦^¼gYEARVACATION
      If textST40.Tag <> textST40.Text And textST40.Locked = False And textST40.Enabled = True Then
         'modify by Sindy 2019/1/28 ¥[user_data.user_enabled:=1ºX¼Ð
         'Modify By Sindy 2019/7/16 C.¤H¨Æ³B½Õ¾ã
'         strSql = "begin user_data.user_enabled:=1; UPDATE YEARVACATION SET yv04=" & textST40.Text & ",yv11=" & strSrvDate(1) & ",yv12='C'" & _
'                  " WHERE yv01=" & Val(LblST40) + 1911 & " and yv02='" & strST01 & "'; end;"
         strSql = "UPDATE YEARVACATION SET yv04=" & textST40.Text & ",yv11=" & strSrvDate(1) & ",yv12='C'" & _
                  " WHERE yv01=" & Val(LblST40) + 1911 & " and yv02='" & strST01 & "'"
         Pub_SeekTbLog strSql
         cnnConnection.Execute strSql
      End If
      '2018/5/16 END
      
      'Add By Sindy 2024/10/17
      '¦³²§°Ê¨ìÂ¾¤é,ÀË¬d¬O§_­n§ï¬ÛÃö¤é´Á¸ê®Æ
      If textST13.Tag <> textST13.Text And Val(textST13.Tag) > 0 Then
         '·s­û¤u¥B«DF½s¸¹ªÌ
         If Mid(textST01, 1, 1) <> "F" Then
            '¤H¨Æ²§°Ê¸ê®Æ
            strSql = "UPDATE staff_change SET sc02=" & DBDATE(textST13.Text) & _
                     " WHERE sc01='" & textST01.Text & "' and sc02=" & DBDATE(textST13.Tag) & " and sc03='01'" '·s¶i
            Pub_SeekTbLog strSql
            cnnConnection.Execute strSql, intI
            If intI = 1 Then '¦³§ó·s¨ì"·s¶i"ªº¤H¨Æ²§°Ê¸ê®Æ
               'Á~¸ê²§°Ê¸ê®Æ
               strSql = "UPDATE salarylog SET SL02=" & DBDATE(textST13.Text) & _
                        " WHERE SL01='" & textST01.Text & "' and SL02=" & DBDATE(textST13.Tag)
               Pub_SeekTbLog strSql
               cnnConnection.Execute strSql, intI
               '­Y¦³¼g¤J­n§R°£(¨C¦~10/2(§t)¥H«á¨ìÂ¾ªÌ¤£°Ñ¥[¦~²×¦ÒÁZ)
               If Val(Right(DBDATE(textST13), 4)) < 1002 Then
                  strSql = "Delete From yearmerit Where ym01=" & Val(Left(DBDATE(textST13), 4)) & " and ym03='" & textST01 & "' and ym05=" & DBDATE(textST13.Tag)
                  Pub_SeekTbLog strSql
                  cnnConnection.Execute strSql, intI
               End If
            End If
         End If
      End If
      '2024/10/17 END
   End If
   
   'Add by Morgan 2009/6/15
   '­Y¾A¥Î³Ò°h·s¨î©Î¦Û´£¶O²v¦³²§°Ê®É§ó·sÁ~¸ê°ò¥»ÀÉ
   If txtSD16.Tag <> txtSD16 Or txtSD17.Tag <> txtSD17 Then
      strSql = "update salarydata set sd16='" & txtSD16 & "',sd17=" & Val(txtSD17) & " where sd01='" & strST01 & "'"
      Pub_SeekTbLog strSql
      cnnConnection.Execute strSql
   End If
   'end 2009/6/15
   
   'Add by Morgan 2010/7/14
   'Modified by Morgan 2024/4/30 +¦X¬ùÂ½Ä¶¤H­û(²Ä5½X>='A')°£¥~
   'Removed by Morgan 2025/7/29 114/7/28°_¼o¤î±B³à¤¬§U¿ìªk
   'If Mid(textST01, 1, 1) <> "F" And Right(textST01, 1) < "A" Then
   '   If m_FieldList(20).fiOldData <> m_FieldList(20).fiNewData Then
   '      '±B³à¤¬§U
   '      If PUB_GetHelpFee(textST01, strExc(1)) = True Then
   '         strSql = "update salarydata set sd09=decode(nvl(sd09,0),0,sd09," & Val(strExc(1)) & "),sd10=decode(nvl(sd10,0),0,sd10," & Val(strExc(1)) & ") where sd01='" & strST01 & "'"
   '         Pub_SeekTbLog strSql
   '         cnnConnection.Execute strSql
   '      End If
   '   End If
   'End If
   'end 2025/7/29
   
'***********************************************************
'          ¿ËÄÝ¸ê®Æ²§°Ê
'***********************************************************
     Pub_SeekTbLog "delete from staff_relation where sr01='" & strST01 & "' "
     cnnConnection.Execute "delete from staff_relation where sr01='" & strST01 & "' "
        For nIndex = 1 To GRD1.Rows - 1
            If GRD1.TextMatrix(nIndex, 0) <> "" Then
                strSql = "insert into staff_relation (sr01,sr02,sr03,sr04,sr05,sr06,sr07,sr08,sr13,sr09,sr10,sr11,sr12) " & _
                              " select '" & strST01 & "',nvl(max(sr02),0)+1,"
                MyArr = Split(GRD1.TextMatrix(nIndex, 0), " ")
                strSql = strSql & CNULL(Trim(MyArr(0))) & ","
                strSql = strSql & CNULL(ChgSQL(GRD1.TextMatrix(nIndex, 1))) & ","
                'Modify by Morgan 2009/6/10 +ª¬ºAÄæ,©Ò¦³«á­±ªº§Ç¦¸+1
                MyArr = Split(GRD1.TextMatrix(nIndex, 3), " ")
                strSql = strSql & CNULL(Trim(MyArr(0))) & ","
                strSql = strSql & CNULL(DBDATE(GRD1.TextMatrix(nIndex, 4))) & ","
                strSql = strSql & CNULL(ChgSQL(GRD1.TextMatrix(nIndex, 5))) & ","
                'Modify by Morgan 2009/6/22 SR08,SR13 §ï©ñ Y
                'If grd1.TextMatrix(nIndex, 6) = "Y" Then
                '   strTemp = ""
                'Else
                '   strTemp = "N"
                'End If
                'strSQL = strSQL & CNULL(ChgSQL(strTemp)) & ","
                'If grd1.TextMatrix(nIndex, 7) = "Y" Then
                '   strTemp = ""
                'Else
                '   strTemp = "N"
                'End If
                'strSQL = strSQL & CNULL(ChgSQL(strTemp)) & ","
                strSql = strSql & CNULL(ChgSQL(GRD1.TextMatrix(nIndex, 6))) & ","
                strSql = strSql & CNULL(ChgSQL(GRD1.TextMatrix(nIndex, 7))) & ","
                'end 2009/6/22
                strSql = strSql & CNULL(ChgSQL(GRD1.TextMatrix(nIndex, 8))) & ","
                strSql = strSql & CNULL(ChgSQL(GRD1.TextMatrix(nIndex, 9))) & ","
                strSql = strSql & CNULL(ChgSQL(GRD1.TextMatrix(nIndex, 10))) & ","
                strSql = strSql & CNULL(DBDATE(GRD1.TextMatrix(nIndex, 11))) & " from staff_relation where '" & strST01 & "'=sr01(+)  "
                Pub_SeekTbLog strSql
                cnnConnection.Execute strSql
                
                'Add by Morgan 2009/6/11
                '·s¼W°·«O²²ÄÝ®É¦P®É·s¼W¤@µ§²§°Ê¸ê®Æ
                If GRD1.TextMatrix(nIndex, 12) = "" And GRD1.TextMatrix(nIndex, 6) = "Y" Then
                  strSql = "insert into HiRelationLog (HL01,HL02,HL03,HL04,HL05,HL07,HL08,HL09)" & _
                     " values ('" & strST01 & "'," & nIndex & "," & strSrvDate(1) & ",'1','" & GRD1.TextMatrix(nIndex, 13) & "','" & strUserNum & "',to_char(sysdate,'yyyymmdd'),to_char(sysdate,'hh24miss'))"
                  Pub_SeekTbLog strSql
                  cnnConnection.Execute strSql
                End If
            End If
        Next nIndex
   
   'Modify By Sindy 2010/5/4
   '­Y¸Ó­û¤uªº¤H¨Æ²§°ÊÀÉ¥u¦³¤@µ§·s¶i¸ê®Æ®É, ­×§ï®É¦P®É§ó·s¤H¨Æ²§°ÊÀÉ.¨Ò: 98003Â¾¦ì­×§ï
   strExc(0) = "select count(*) from staff_change where sc01='" & strST01 & "' "
   intI = 1
   Set adoRecordset = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      If Val(adoRecordset.Fields(0)) = 1 Then
         strExc(0) = "select count(*) from staff_change where sc01='" & strST01 & "' and sc03='01' "
         intI = 1
         Set adoRecordset = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            If Val(adoRecordset.Fields(0)) = 1 Then
               'Modify By Sindy 2024/9/9 ¦]³æ¦ì§ïST93; CNULL(m_FieldList(2).fiNewData) => CNULL(Trim(Left(textST03.Text, 4)))
               strSql = "update staff_change " & _
                               "set sc04=" & CNULL(Trim(Left(textST03.Text, 4))) & "," & _
                                     "sc05=" & CNULL(m_FieldList(19).fiNewData) & "," & _
                                     "sc06=" & CNULL(m_FieldList(20).fiNewData) & "," & _
                                     "sc07=" & CNULL(m_FieldList(48).fiNewData) & " " & _
                               "where sc01='" & strST01 & "' and sc03='01' "
               Pub_SeekTbLog strSql
               cnnConnection.Execute strSql
            End If
         End If
      End If
   End If
   '2010/5/4 End
   
   'Add By Sindy 2014/3/12 +­û¤u±Mªø¸ê®Æ
   If Trim(textSS02) <> "" Or Trim(textSS03) <> "" Or Trim(textSS04) <> "" Or _
      Trim(textSS05) <> "" Or Trim(textSS06) <> "" Or Trim(textSS07) <> "" Then
      strExc(0) = "select * from staff_specialty where ss01='" & strST01 & "'"
      intI = 1
      Set adoRecordset = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         strSql = "update staff_specialty set" & _
                   " ss02=" & CNULL(textSS02) & _
                  " ,ss03=" & CNULL(textSS03) & _
                  " ,ss04=" & CNULL(textSS04) & _
                  " ,ss05=" & CNULL(textSS05) & _
                  " ,ss06=" & CNULL(textSS06) & _
                  " ,ss07=" & CNULL(textSS07) & _
                  " where ss01='" & strST01 & "'"
      Else
         strSql = "INSERT INTO staff_specialty(ss01,ss02,ss03,ss04,ss05,ss06,ss07)" & _
                  " VALUES(" & CNULL(strST01) & "," & CNULL(textSS02) & "," & CNULL(textSS03) & _
                  "," & CNULL(textSS04) & "," & CNULL(textSS05) & "," & CNULL(textSS06) & _
                  "," & CNULL(textSS07) & ")"
      End If
      Pub_SeekTbLog strSql
      cnnConnection.Execute strSql
   'Add By Sindy 2017/12/22
   Else
      strSql = "delete from staff_specialty where ss01='" & strST01 & "'"
      Pub_SeekTbLog strSql
      cnnConnection.Execute strSql
   '2017/12/22 END
   End If
   '2014/3/12 END
   
   cnnConnection.CommitTrans
   
   'Add By Sindy 2023/8/17
   If textSS04.Tag <> textSS04.Text Then
      PUB_SendMail strUserNum, Pub_GetSpecMan("¸Õ¥Î´Áº¡³qª¾"), "", "¤H­û´¼°]±M·~ÃÒ·Ó²§°Ê³qª¾¡I", "­û¤u½s¸¹¡G" + textST01 + " " + textST02 & vbCrLf & _
         "³¡¡@ªù¡G" + textST03 & vbCrLf & _
         "­û¤u©ÒÄÝ©Ò§O¡G" + textST06 & vbCrLf & _
         "Â¾¡@¦ì¡G" + textST21 & vbCrLf & vbCrLf & vbCrLf & _
         "ÃÒ·Ó²§°Ê«e¡G" & textSS04.Tag & vbCrLf & _
         "ÃÒ·Ó²§°Ê«á¡G" & textSS04.Text & vbCrLf, , , , , , , , , , True
   End If
   '2023/8/17 END
    
   Call Chk_Staff_Relation(textST01) 'Add By Sindy 2024/11/18
   
'cancel by sonia 2021/12/30 ¤w¦Û°Ê§ó·sYearVacation·í¦~«×ªº¸ê®Æ¬G¤£¥²¦A´£¿ô
'   'add by sonia 2019/1/28
'   If textST40.Tag <> textST40.Text And textST40.Locked = False And textST40.Enabled = True Then
'      MsgBox "­Y¬°·s¤H¥i¥ð°²¤é¼Æ²¾¦Ü¦¸¦~¡A½Ð³qª¾¹q¸£¤¤¤ß±N«e¤@¦~¥i¥ð°²¤é¼Æ¦©°£ !!"
'   End If
'   'end 2019/1/28

   ShowCurrRecord strST01
      
   ModRecord = True
   Exit Function
ErrHand:
    cnnConnection.RollbackTrans
    MsgBox (Err.Description)
End Function

' §R°£°O¿ý
Private Function DelRecord() As Boolean
Dim strSql As String
Dim strST01 As String
   
   DelRecord = False
   
On Error GoTo ErrHand
   
   cnnConnection.BeginTrans
   
   strST01 = m_CurrKEY

   strSql = "DELETE FROM staff WHERE st01 = '" & strST01 & "'"
   Pub_SeekTbLog strSql
   cnnConnection.Execute strSql

'***********************************************************
'          ¿ËÄÝ¸ê®Æ²§°Ê
'***********************************************************
   strSql = "delete from staff_relation where sr01='" & strST01 & "'"
   Pub_SeekTbLog strSql
   cnnConnection.Execute strSql

   'Add by Morgan 2009/7/7
   strSql = "DELETE HiRelationLog WHERE HL01='" & strST01 & "'"
   Pub_SeekTbLog strSql
   cnnConnection.Execute strSql
   'END 2009/7/7
   
   'Add by Sindy 2012/6/14
   strSql = "DELETE imgbytefile WHERE ibf01='000' and ibf02='" & strST01 & "' and ibf03='0' and ibf04='00' and ibf05='3'"
   cnnConnection.Execute strSql
   '2012/6/14 End
   
   'Add By Sindy 2014/3/12 +­û¤u±Mªø¸ê®Æ
   strSql = "DELETE FROM staff_specialty WHERE ss01 = '" & strST01 & "'"
   Pub_SeekTbLog strSql
   cnnConnection.Execute strSql
   '2014/3/12 END
   
   If (strST01 = m_LastKEY) Or (strST01 = m_FirstKEY) Then
      RefreshRange
   End If
   ShowCurrRecord strST01
   DelRecord = True
   cnnConnection.CommitTrans
   
   Exit Function
ErrHand:
    cnnConnection.RollbackTrans
    MsgBox "§R°£¥¢±Ñ¡I" & vbCrLf & Err.Description
End Function

' ¬d¸ß°O¿ý
'Modify By Sindy 2012/6/14 Mark
'Private Function QueryRecord() As Boolean
Public Function QueryRecord(strST01) As Boolean
'Dim strST01 As String
   
   QueryRecord = False
'   strST01 = textST01
   
   If IsRecordExist(strST01) = True Then
      m_CurrKEY = strST01
      QueryRecord = True
      UpdateCtrlData
   Else
      QueryRecord = False
   End If

   UpdateToolbarState
End Function

' ¨Ï¥ÎªÌ«ö¤U½T©wªº«ö¯Ã
Private Function OnWork() As Boolean
Dim strMsg As String
Dim strTit As String
Dim nResponse
   
   OnWork = False
   Select Case m_EditMode
      Case 1: '·s¼W
         If CheckDataValid() = True Then
            '­«·sÀË¬dÄæ¦ì¦³®Ä©Ê
            If TxtValidate = False Then Exit Function
            UpdateFieldNewData
            If AddRecord = True Then
                RefreshRange
            Else
                Exit Function
            End If
         Else
            GoTo EXITSUB
         End If
      Case 2: '­×§ï
         If CheckDataValid() = True Then
            'Add By Cheng 2002/05/22
            '­«·sÀË¬dÄæ¦ì¦³®Ä©Ê
            If TxtValidate = False Then Exit Function
            UpdateFieldNewData
            If ModRecord = False Then Exit Function
         Else
            GoTo EXITSUB
         End If
      Case 3: '§R°£
         If DelRecord = True Then
            RefreshRange
            ClearField
            ShowCurrRecord m_CurrKEY
         Else
            Exit Function
         End If
      Case 4: '¬d¸ß
         If textST01 <> "" Then
            ' 2008/12/16 ADD BY SINDY
            ' ÀË¬d­û¤u½s¸¹³W«h
            If ChkStaffID(textST01) Then
               'Call textST01_GotFocus
               textST01.SetFocus
               nResponse = 1
               UpdateCtrlData
            ' 2008/12/16 END
            Else
               If QueryRecord(textST01) = False Then
                  strMsg = "µL¦¹¸ê®Æ"
                  strTit = "¬d¸ß¸ê®Æ"
                  nResponse = MsgBox(strMsg, vbOKOnly, strTit)
                  UpdateCtrlData
               End If
            End If
         Else
            GoTo EXITSUB
         End If
   End Select
   m_EditMode = 0
   SetCtrlReadOnly True
   OnWork = True
EXITSUB:
End Function

' ¶}©l¿é¤J¸ê®Æ
Private Sub SetInputEntry()
   Select Case m_EditMode
      Case 1: If Me.Visible = True Then textST01.SetFocus
      Case 2: If Me.Visible = True Then textST02.SetFocus
      Case 4: If Me.Visible = True Then textST01.SetFocus
   End Select
End Sub

' ÀË¬d°O¿ý¬O§_¤w¸g¦s¦b
Private Function IsRecordExist(ByVal strKEY01 As String) As Boolean
Dim rsTmp As New ADODB.Recordset
Dim strSql As String
   
   IsRecordExist = False
      
   strSql = "SELECT * FROM staff " & _
            "WHERE st01 = '" & strKEY01 & "'  "
                  
   ' Åª¨ú¸ê®Æ®w
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   ' ÀË¬dÅª¨úªº¸ê®Æµ§¼Æ
   If rsTmp.RecordCount > 0 Then
      IsRecordExist = True
   Else
      IsRecordExist = False
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Function


' Åã¥Ü¸ê®Æ
Private Sub ShowCurrRecord(ByVal strKEY01 As String)
Dim strSql As String
Dim rsTmp As New ADODB.Recordset

   If IsRecordExist(strKEY01) = True Then
      m_CurrKEY = strKEY01
   Else
      strSql = "SELECT st01 FROM staff " & _
               "WHERE st01 = '" & m_CurrKEY & "' "
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         If IsNull(rsTmp.Fields("st01")) = False Then: m_CurrKEY = rsTmp.Fields("st01")
         rsTmp.Close
         UpdateCtrlData
         GoTo EXITSUB
      End If
      rsTmp.Close
      
      ' 2008/12/16 MODIFY BY SINDY
      'strSQL = "SELECT st01 FROM staff " & _
      '         "WHERE st01 = (SELECT MIN(st01) FROM staff ) "
      '2011/4/1 MODIFY BY SONIA
      'strSQL = "SELECT st01 FROM staff " & _
               "WHERE st01 = (SELECT MIN(st01) FROM staff where (substr(st01,1,1)>='6' and substr(st01,1,1)<='9') or (substr(st01,1,1)='F')) "
      strSql = "SELECT st01 FROM staff " & _
               "WHERE st01 = (SELECT MIN(st01) FROM staff where (substr(st01,1,1)>='6' and substr(st01,1,1)<='F')) "
      ' 2008/12/16 END
      
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         If IsNull(rsTmp.Fields("st01")) = False Then: m_CurrKEY = rsTmp.Fields("st01")
      Else
         ShowLastRecord
         GoTo EXITSUB
      End If
      rsTmp.Close
   End If
   UpdateCtrlData
EXITSUB:
End Sub

' Åã¥Ü²Ä¤@µ§¸ê®Æ
Private Sub ShowFirstRecord()
   m_CurrKEY = m_FirstKEY
  
   UpdateCtrlData
End Sub

' Åã¥Ü¤W¤@µ§¸ê®Æ
Private Sub ShowPrevRecord()
Dim strSql As String
Dim rsTmp As New ADODB.Recordset
   
   If m_CurrKEY = m_FirstKEY Then
      ShowMsg MsgText(9008)
      GoTo EXITSUB
   End If
   
   strSql = "SELECT st01 FROM staff " & _
            "WHERE st01 = (SELECT MAX(st01) FROM staff " & _
                          "WHERE st01 < '" & m_CurrKEY & "' )"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("st01")) = False Then: m_CurrKEY = rsTmp.Fields("st01")
      rsTmp.Close
      UpdateCtrlData
      GoTo EXITSUB
   End If
   rsTmp.Close
   
   strSql = "SELECT st01 FROM staff " & _
            "WHERE st01 = (SELECT Min(st01) FROM staff ) "

   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("st01")) = False Then: m_CurrKEY = rsTmp.Fields("st01")
   End If
   rsTmp.Close
   UpdateCtrlData
   
EXITSUB:
   Set rsTmp = Nothing
End Sub

' Åã¥Ü¤U¤@µ§¸ê®Æ
Private Sub ShowNextRecord()
Dim strSql As String
Dim rsTmp As New ADODB.Recordset
   
   If m_CurrKEY = m_LastKEY Then
      ShowMsg MsgText(9009)
      GoTo EXITSUB
   End If
   
   strSql = "SELECT st01 FROM staff " & _
            "WHERE st01 = (SELECT MIN(st01) FROM staff " & _
                          "WHERE st01  > '" & m_CurrKEY & "' )"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("st01")) = False Then: m_CurrKEY = rsTmp.Fields("st01")
      rsTmp.Close
      UpdateCtrlData
      GoTo EXITSUB
   End If
   rsTmp.Close
   
   strSql = "SELECT st01 FROM staff " & _
            "WHERE st01 = (SELECT max(st01) FROM staff ) "

   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("st01")) = False Then: m_CurrKEY = rsTmp.Fields("st01")
   End If
   rsTmp.Close
   
   UpdateCtrlData
   
EXITSUB:
   Set rsTmp = Nothing
End Sub

' Åã¥Ü³Ì«á¤@µ§¸ê®Æ
Private Sub ShowLastRecord()
   m_CurrKEY = m_LastKEY
   UpdateCtrlData
End Sub

' °õ¦æ«ü¥O
Private Sub OnAction(ByVal KeyCode As Integer)
Dim strTit As String
Dim strMsg As String
Dim nResponse
   
   m_SubMode = 0
   Select Case KeyCode
      ' ·s¼W
      Case vbKeyF2:
         m_EditMode = 1
         ClearField
         SetCtrlReadOnly False
         UpdateToolbarState
         SetInputEntry
         Me.SSTab1.Tab = 0
      ' ­×§ï
      Case vbKeyF3:
         m_EditMode = 2
         SetCtrlReadOnly False
         SetKeyReadOnly True
         UpdateToolbarState
         SetInputEntry
         'Add By Sindy 2010/7/7 ÀË¬d¸Ó­û¤u¬O§_¦³¤@µ§¥H¤Wªº¤H¨Æ²§°Ê¸ê®Æ,­Y¦³,«h±N³¡ªù,Â¾ºÙ,Â¾¦ì¤T­ÓÄæ¦ìÂê¦í
         strExc(0) = "select count(*) from Staff_Change where sc01='" & textST01 & "' "
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            If RsTemp.Fields(0) > 1 Then
               textST03.Enabled = False
               textST20.Enabled = False
               textST21.Enabled = False
            End If
         End If
         '2010/7/7 End
      ' §R°£
      Case vbKeyF5:
         strTit = "¸ß°Ý"
         strMsg = "¬O§_­n§R°£¦¹µ§¸ê®Æ?"
         nResponse = MsgBox(strMsg, vbYesNo + vbCritical + vbDefaultButton2, strTit)
         If nResponse = vbYes Then
            m_EditMode = 3
            If OnWork = True Then
                UpdateToolbarState
            Else
                Exit Sub
            End If
         End If
      ' ¬d¸ß
      Case vbKeyF4:
         m_EditMode = 4
         SetCtrlReadOnly True
         SetKeyReadOnly False
         ClearField
         UpdateToolbarState
         SetInputEntry
      ' ²Ä¤@µ§
      Case vbKeyHome:
         ShowFirstRecord
      ' «e¤@µ§
      Case vbKeyPageUp:
         ShowPrevRecord
      ' «á¤@µ§
      Case vbKeyPageDown:
         ShowNextRecord
      ' ³Ì«á¤@µ§
      Case vbKeyEnd:
         ShowLastRecord
      ' ½T©w
      Case vbKeyF9:
         'Add by Morgan 2009/6/15
         If cmdOK(1).Enabled = True Then
            If MsgBox("¿ËÄÝ¸ê®Æ©|¥¼§@·~§¹²¦¡A½T©w©ñ±ó½s¿è¶Ü?", vbYesNo + vbDefaultButton2) = vbNo Then
               Exit Sub
            End If
         End If
         'end 2009/6/15
         ' ±N©Ò¦³Äæ¦ìªº¤º®e§ó·s¨ìÄæ¦ì¦ê¦C¤¤ªºÄæ¦ì¤º®e¶µ¥Ø
         UpdateFieldNewData
         If OnWork = True Then
            UpdateToolbarState
         Else
            Exit Sub
         End If
      ' ¨ú®ø
      Case vbKeyF10:
         Select Case m_EditMode
            Case 1, 2:
               strTit = "¸ß°Ý"
               strMsg = "§A¨Ã¥¼¦sÀÉ, ½T©wÂ÷¶}¶Ü?"
               nResponse = MsgBox(strMsg, vbYesNo + vbCritical + vbDefaultButton2, strTit)
               If nResponse = vbYes Then
                  m_EditMode = 0
                  UpdateCtrlData
                  SetCtrlReadOnly True
                  UpdateToolbarState
               End If
            Case Else
               m_EditMode = 0
               UpdateCtrlData
               SetCtrlReadOnly True
               UpdateToolbarState
         End Select
      ' Â÷¶}
      Case vbKeyEscape:
         Unload Me
   End Select
   If KeyCode <> vbKeyEscape And KeyCode <> vbKeyF3 Then
'      tabCustomer.Tab = 0
   End If
End Sub

Private Sub RefreshRange()
Dim strSql As String
Dim rsTmp As New ADODB.Recordset
   
   ' 2008/12/16 MODIFY BY SINDY
   'strSQL = "SELECT st01 FROM staff " & _
   '         "WHERE st01 = (SELECT MIN(st01) FROM staff) "
   '2011/4/1 MODIFY BY SONIA
   'strSQL = "SELECT st01 FROM staff " & _
            "WHERE st01 = (SELECT MIN(st01) FROM staff where (substr(st01,1,1)>='6' and substr(st01,1,1)<='9') or (substr(st01,1,1)='F')) "
   strSql = "SELECT st01 FROM staff " & _
            "WHERE st01 = (SELECT MIN(st01) FROM staff where (substr(st01,1,1)>='6' and substr(st01,1,1)<='F'))"
   ' 2008/12/16 END
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("st01")) = False Then: m_FirstKEY = rsTmp.Fields("st01")
   End If
   rsTmp.Close
   
   ' 2008/12/16 MODIFY BY SINDY
   'strSQL = "SELECT st01 FROM staff " & _
   '         "WHERE st01 = (SELECT MAX(st01) FROM staff)  "
   '2011/4/1 MODIFY BY SONIA
   'strSQL = "SELECT st01 FROM staff " & _
            "WHERE st01 = (SELECT MAX(st01) FROM staff  where (substr(st01,1,1)>='6' and substr(st01,1,1)<='9') or (substr(st01,1,1)='F')) "
   strSql = "SELECT st01 FROM staff " & _
            "WHERE st01 = (SELECT MAX(st01) FROM staff  where (substr(st01,1,1)>='6' and substr(st01,1,1)<='F'))"
   ' 2008/12/16 END
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("st01")) = False Then: m_LastKEY = rsTmp.Fields("st01")
   End If
   rsTmp.Close
   
   Set rsTmp = Nothing
End Sub

' ±N¸ê®Æ®w¤¤ªº¸ê®Æ§ó·s¨ì©Ò¦³Äæ¦ì¤¤
Private Sub UpdateCtrlData()
Dim rsTmp As New ADODB.Recordset
Dim strSql As String
Dim lngInsureSalary As Long '§ë«OÁ~¸ê  'add by sonia 2016/5/5
Dim strBackTaieDate As String
   
   'Modify By Sindy 2014/3/12 +staff_specialty ­û¤u±Mªø¸ê®Æ
   strSql = "SELECT * FROM staff,staff_specialty " & _
            "WHERE st01='" & m_CurrKEY & "' and st01=ss01(+)"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      ClearField
      If IsNull(rsTmp.Fields("st01")) = False Then: textST01 = rsTmp.Fields("st01")
      If IsNull(rsTmp.Fields("st02")) = False Then: textST02 = rsTmp.Fields("st02")
      'Modify By Sindy 2023/12/20
'      If strSrvDate(1) >= ·s³¡ªù±Ò¥Î¤é Then
         If IsNull(rsTmp.Fields("st93")) = False Then
            textST03 = rsTmp.Fields("st93")
         ElseIf IsNull(rsTmp.Fields("st03")) = False Then
            textST03 = rsTmp.Fields("st03")
            textST03.Tag = "ST03" 'Add By Sindy 2024/1/9
         End If
'      Else
'      '2023/12/20 END
'         If IsNull(rsTmp.Fields("st03")) = False Then
'            textST03 = rsTmp.Fields("st03")
'            textST03.Tag = "ST03" 'Add By Sindy 2024/3/13
'         End If
'      End If
      If IsNull(rsTmp.Fields("st06")) = False Then: textST06 = rsTmp.Fields("st06")
      If IsNull(rsTmp.Fields("st08")) = False Then: textST08 = rsTmp.Fields("st08")
      If IsNull(rsTmp.Fields("st09")) = False Then: textST09 = rsTmp.Fields("st09")
      If IsNull(rsTmp.Fields("st10")) = False Then: textST10 = rsTmp.Fields("st10")
      '¨ìÂ¾¤é
      If IsNull(rsTmp.Fields("st13")) = False Then
         textST13 = TAIWANDATE(rsTmp.Fields("st13"))
         'Add By Sindy 2024/10/17
         textST13.Tag = textST13.Text
         '2024/10/17 END
      End If
      
      If IsNull(rsTmp.Fields("st18")) = False Then: textST18 = rsTmp.Fields("st18")
      If IsNull(rsTmp.Fields("st19")) = False Then: textST19 = rsTmp.Fields("st19")
      If IsNull(rsTmp.Fields("st20")) = False Then: textST20 = rsTmp.Fields("st20")
      If IsNull(rsTmp.Fields("st21")) = False Then: textST21 = rsTmp.Fields("st21")
      textST21.Tag = textST21 'Add by Morgan 2010/7/14
      If IsNull(rsTmp.Fields("st22")) = False Then: textST22 = rsTmp.Fields("st22")
      If IsNull(rsTmp.Fields("st23")) = False Then: textST23 = TAIWANDATE(rsTmp.Fields("st23"))
      If IsNull(rsTmp.Fields("st24")) = False Then: textST24 = rsTmp.Fields("st24")
      If IsNull(rsTmp.Fields("st25")) = False Then: textST25 = rsTmp.Fields("st25")
      If IsNull(rsTmp.Fields("st26")) = False Then: textST26 = rsTmp.Fields("st26")
      If IsNull(rsTmp.Fields("st27")) = False Then: textST27 = rsTmp.Fields("st27")
      If IsNull(rsTmp.Fields("st28")) = False Then: textST28 = TAIWANDATE(rsTmp.Fields("st28"))
      If IsNull(rsTmp.Fields("st29")) = False Then: textST29 = TAIWANDATE(rsTmp.Fields("st29"))
      If IsNull(rsTmp.Fields("st30")) = False Then: textST30 = rsTmp.Fields("st30")
      If IsNull(rsTmp.Fields("st31")) = False Then: textST31 = TAIWANDATE(rsTmp.Fields("st31"))
      If IsNull(rsTmp.Fields("st32")) = False Then: textST32 = TAIWANDATE(rsTmp.Fields("st32"))
      If IsNull(rsTmp.Fields("st33")) = False Then: textST33 = rsTmp.Fields("st33")
      If IsNull(rsTmp.Fields("st34")) = False Then: textST34 = rsTmp.Fields("st34")
      If IsNull(rsTmp.Fields("st35")) = False Then: textST35 = rsTmp.Fields("st35")
      If IsNull(rsTmp.Fields("st36")) = False Then: textST36 = rsTmp.Fields("st36")
      If IsNull(rsTmp.Fields("st37")) = False Then: textST37 = rsTmp.Fields("st37")
      If IsNull(rsTmp.Fields("st38")) = False Then: textST38 = rsTmp.Fields("st38")
      If IsNull(rsTmp.Fields("st39")) = False Then: textST39 = rsTmp.Fields("st39")
      If IsNull(rsTmp.Fields("st40")) = False Then
         textST40 = rsTmp.Fields("st40")
      'Modify By Sindy 2018/5/16 ¶}©ñ¥i¥H­×§ï,¦ý­n¦^¼gYEARVACATION
         textST40.Tag = textST40.Text
'         strSql = "select yv01 from YEARVACATION where yv02='" & m_CurrKEY & "'" & _
'                  " and yv11=(select nvl(max(yv11),0) from YEARVACATION where yv02='" & m_CurrKEY & "')" & _
'                  " order by yv01 desc"
         'Modify By Sindy 2019/1/2
         'Modify By Sindy 2021/2/18 + and yv11>0 :¥Nªí¤w§ó·s¯S§O°²
         'modify by sonia 2021/12/30 +¤º¼h¤]­n¥[yv01<=" & Left(strSrvDate(1), 4)±ø¥ó¡A§_«h¤U¤@¦~«×YEARVACATION¤w²£¥Í¦ý¥¼§ó·s«e·|§ì¤£¨ì¦~¸ê®Æ
         strSql = "select yv01 from YEARVACATION where yv02='" & m_CurrKEY & "'" & _
                  " and yv01=(select nvl(max(yv01),0) from YEARVACATION where yv02='" & m_CurrKEY & "' and yv01<=" & Left(strSrvDate(1), 4) & ")  and yv01<=" & Left(strSrvDate(1), 4) & " and yv11>0" & _
                  " order by yv01 desc"
         '2019/1/2 END
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strSql)
         If intI = 1 Then
            LblST40 = Val(RsTemp.Fields("yv01")) - 1911
         End If
      End If
      '2018/5/16 END
      
      If IsNull(rsTmp.Fields("st41")) = False Then: textST41 = TAIWANDATE(rsTmp.Fields("st41"))
      If IsNull(rsTmp.Fields("st42")) = False Then: textST42 = rsTmp.Fields("st42")
      If IsNull(rsTmp.Fields("st49")) = False Then: textST49 = rsTmp.Fields("st49")
      If IsNull(rsTmp.Fields("st12")) = False Then: textST12 = rsTmp.Fields("st12")
      If IsNull(rsTmp.Fields("st51")) = False Then: textST51 = TAIWANDATE(rsTmp.Fields("st51"))
      If IsNull(rsTmp.Fields("st68")) = False Then: textST68 = Val(rsTmp.Fields("ST68")) - 1911 'Add By Sindy 2015/8/13
      
      'Modify By Sindy 2014/3/12 +­û¤u±Mªø¸ê®Æ
      If IsNull(rsTmp.Fields("ss02")) = False Then: textSS02 = rsTmp.Fields("ss02")
      If IsNull(rsTmp.Fields("ss03")) = False Then: textSS03 = rsTmp.Fields("ss03")
      textSS04.Tag = "" 'Add By Sindy 2023/8/17
      If IsNull(rsTmp.Fields("ss04")) = False Then
         textSS04 = rsTmp.Fields("ss04")
         textSS04.Tag = textSS04.Text 'Add By Sindy 2023/8/17
      End If
      If IsNull(rsTmp.Fields("ss05")) = False Then: textSS05 = rsTmp.Fields("ss05")
      If IsNull(rsTmp.Fields("ss06")) = False Then: textSS06 = rsTmp.Fields("ss06")
      If IsNull(rsTmp.Fields("ss07")) = False Then: textSS07 = rsTmp.Fields("ss07")
      '2014/3/12 END
      
      'Add by Morgan 2009/6/24 ³Ò°·«O¸É§UÃþ§O
      SelCombo cboST50, "" & rsTmp.Fields("st50")
      SelCombo cboST56, "" & rsTmp.Fields("st56")
      'end 2009/6/24
      
      ' §ó·sCUID
      UpdateCUID rsTmp
      
      ' 2008/12/17 ADD BY SINDY
      If IsNull(rsTmp.Fields("st30")) = False Then
         If IsEmptyText(rsTmp.Fields("st30")) = False Then
            LabelST30.Caption = GetStaffName(rsTmp.Fields("st30"), False)
         End If
      End If
      ' 2008/12/17 END
      
      'Add By Sindy 2019/6/25 ¯dÂ¾°±Á~¯S¥ð°_ºâ¤é
      strBackTaieDate = Pub_BackTaieToDate(m_CurrKEY, Left(strSrvDate(1), 4))
      If Val(strBackTaieDate) > 0 Then
         LblBackDate.Visible = True
         txtBackDate.Visible = True
         txtBackDate.Text = TAIWANDATE(strBackTaieDate)
      Else
         LblBackDate.Visible = False
         txtBackDate.Visible = False
      End If
      '2019/6/25 END
      
      Call ReadPhoto 'Add By Sindy 2012/6/20 ¸ü¤J·Ó¤ù
      
      ' §ó·s¼È¦s°Ïªº¸ê®Æ
      UpdateFieldOldData rsTmp
      
      textST03_Validate False
      textST06_Validate False
      textST20_Validate False
      textST21_Validate False
      textST22_Validate False
      textST24_Validate False
      textST25_Validate False
      textST27_Validate False
      textST37_Validate False
      
      'Add by Morgan 2009/6/15
      'Åª¨ú¾A¥Î³Ò°h·s¨îÄæ¦ì
      'modify by sonia 2016/5/3 ¥[¤J§ë«Oª÷ÃB¬ÛÃöÄæ¦ìsd12,sd13,sd27,sd43,sd45,sd47
      strSql = "select sd16,sd17,sd11,sd12,sd45,sd27,sd43,sd47 from salarydata where sd01='" & m_CurrKEY & "'"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strSql)
      If intI = 1 Then
         txtSD16 = "" & RsTemp(0)
         txtSD16.Tag = txtSD16
         txtSD17 = "" & RsTemp(1)
         txtSD17.Tag = txtSD17
         'add by sonia 2016/5/3
         '³Ò«O§ë«Oª÷ÃB
         If Not IsNull(RsTemp("sd12")) Then
            lngInsureSalary = Val("" & RsTemp("sd12"))
         Else
         'end 2016/3/31
            lngInsureSalary = Val("" & RsTemp("sd45"))
         End If
         lblSDdata(0) = PUB_GetInsureBase(lngInsureSalary, "L")
         '°·«O§ë«Oª÷ÃB
         lblSDdata(1) = Val("" & RsTemp("sd47"))
         '³Ò°·«O¬O§_¥H¦X¹Ù¤H¨­¤À§ë«O
         lblSDdata(2) = "" & RsTemp("sd11")
         '°h¥ðª÷§ë«Oª÷ÃB
         If Not IsNull(RsTemp("sd27")) Then
            lngInsureSalary = Val("" & RsTemp("sd27"))
         Else
            lngInsureSalary = Val("" & RsTemp("sd43"))
         End If
         lblSDdata(3) = PUB_GetInsureBase(lngInsureSalary, "R")
         '³Ò°h·s¨î¤H­û¤~Åã¥Ü°h¥ðª÷§ë«Oª÷ÃBÄæ
         If txtSD16 = "" Then
            Label1(48).Visible = False
            lblSDdata(3).Visible = False
         Else
            Label1(48).Visible = True
            lblSDdata(3).Visible = True
         End If
      Else
         txtSD16 = ""
         txtSD17 = ""
         lblSDdata(0) = ""
         lblSDdata(1) = ""
         lblSDdata(2) = ""
         lblSDdata(3) = ""
         If txtSD16 = "" Then
            Label1(48).Visible = False
            lblSDdata(3).Visible = False
         Else
            Label1(48).Visible = True
            lblSDdata(3).Visible = True
         End If
         'end 2016/5/3
      End If
      'end 2009/6/15
      
      '§ì¨ú²²ÄÝ
      'Modify by Morgan 2009/6/22 +SR02,HL05;SR08,SR13 §ï©ñ Y
      'strSQL = "SELECT sr03||' '||decode(sr03,'1','¤÷¿Ë','2','¥À¿Ë','3','°t°¸','4','¤l¤k','¨ä¥L'),Sr04,DECODE(SR12,NULL,NULL,'§R') STATUS,sr05||' '||decode(sr05,'M','¨k','F','¤k','¤£¸Ô'),sqldatet(sr06),sr07,decode(sr08,'','Y',' '),decode(sr13,'','Y',' '),sr09,sr10,sr11,sqldatet(sr12),sr02  FROM staff_relation " & _
              "WHERE SR01 = '" & m_CurrKEY & "'  order by sr02 "
      strSql = "SELECT sr03||' '||decode(sr03,'1','¤÷¿Ë','2','¥À¿Ë','3','°t°¸','4','¤l¤k','¨ä¥L')" & _
         ",Sr04,DECODE(SR12,NULL,NULL,'§R') STATUS,sr05||' '||decode(sr05,'M','¨k','F','¤k','¤£¸Ô')" & _
         ",sqldatet(sr06),sr07,sr08,sr13,sr09,sr10,sr11,sqldatet(sr12),sr02,hl05" & _
         " FROM staff_relation,(select hl02,hl05 from HIrelationlog a" & _
         " where hl01='" & m_CurrKEY & "'" & _
         " and (hl02,hl03) in (select b.hl02,max(b.hl03) from hirelationlog b where b.hl01=a.hl01" & _
         " and b.hl03<=to_char(sysdate,'yyyymmdd') group by hl02)" & _
         ") X WHERE SR01 = '" & m_CurrKEY & "' and hl02(+)=SR02 order by sr02 "
      
      If rsTmp.State = 1 Then rsTmp.Close
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      Set GRD1.Recordset = rsTmp
      'Add by Morgan 2009/6/10
      cmdOK(1).Enabled = False
      cmdOK(2).Enabled = False
      'end 2009/6/10
      
      ' 2008/12/17 MODIFY BY SINDY
      'strSQL = "SELECT myyear,decode(ym02,'1','Àu','3','¤A','4','¤þ','¥Ò')  FROM yearmerit,(select distinct to_number(substr(st01,1,2))+1911 as MyYear from staff where st01>='" & m_CurrKEY & "' and st01<rtrim(ltrim(to_char(to_number(to_char(sysdate,'YYYY'))-1911)))) AA where AA.MyYear=ym01(+) and '" & m_CurrKEY & "'=ym02(+)  "
      'Modify by Morgan 2009/6/19 ­ì¨Ó´_Â¾ªº¸ê®Æ·|¦³°ÝÃD¡A§ï¥Òµ¥¤]¦sÀÉ¥iª½±µ§ì¦ÒÁZÀÉ
      'strSQL = "SELECT myyear,decode(ym02,'1','Àu','3','¤A','4','¤þ','¥Ò')  FROM yearmerit,(select distinct to_number(substr(st01,1,2))+1911 as MyYear from staff where st01>='" & m_CurrKEY & "' and st01<rtrim(ltrim(to_char(to_number(to_char(sysdate,'YYYY'))-1911)))) AA where AA.MyYear=ym01(+) and '" & m_CurrKEY & "'=ym03(+)  "
      'modify by sonia 2016/1/5 ym02¥[*¤£°Ñ¥[¦Ò®Ö
      'strSql = "SELECT ym01-1911,decode(ym02,'1','Àu','3','¤A','4','¤þ','¥Ò')  FROM yearmerit where ym03='" & m_CurrKEY & "' order by 1 desc"
      strSql = "SELECT ym01-1911,decode(ym02,'1','Àu','3','¤A','4','¤þ','*','¤£°Ñ¥[¦Ò®Ö','¥Ò')  FROM yearmerit where ym03='" & m_CurrKEY & "' order by 1 desc"
      'end 2016/1/5
      'end 2009/6/19
      ' 2008/12/17 END
      If rsTmp.State = 1 Then rsTmp.Close
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      Set GRD2.Recordset = rsTmp
      'Add by Morgan 2009/6/19
      GRD2.ColAlignment(0) = flexAlignCenterCenter
      GRD2.ColAlignment(1) = flexAlignCenterCenter
      
      'Add By Sindy 2015/8/13 °·ÀË¸ê®Æ
      strSql = "SELECT sqldatet(SH02),SH03,sqldatet(SH06),SH04 FROM staff_health where SH01='" & m_CurrKEY & "' order by 1 desc"
      If rsTmp.State = 1 Then rsTmp.Close
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      Set grd3.Recordset = rsTmp
      grd3.ColAlignment(0) = flexAlignCenterCenter
      grd3.ColAlignment(1) = flexAlignCenterCenter
      grd3.ColAlignment(2) = flexAlignCenterCenter
      '2015/8/13 END
   End If
   SetGrd
   rsTmp.Close
   
EXITSUB:
   Set rsTmp = Nothing
End Sub

'Add By Sindy 2012/6/20 ¸ü¤J·Ó¤ù
Private Function ReadPhoto() As Boolean
Dim PicRs As New ADODB.Recordset
Dim file_num As Integer
Dim bytes() As Byte
Dim IsWmf As Boolean
Dim pWidth As Integer '¹Ï¤ù¼e«×
Dim pHeight As Integer '¹Ï¤ù°ª«×
Dim dblTmp As Double
Dim sW As Integer, sH As Integer
Dim stAttachFile As String
   
   Screen.MousePointer = vbHourglass
   
   ReadPhoto = False
   
   '²M¹Ï¤ù
   tmpPic.Picture = LoadPicture()
   tmpImg.Picture = LoadPicture()
   G_SeekPicColor.Picture = LoadPicture()
   G_SeekPicColor.Width = 0
   G_SeekPicColor.Height = 0
   
   DoEvents
   Set PicRs = New ADODB.Recordset
   PicRs.CursorLocation = adUseClient
   PicRs.Open "select ImgByteFile.*,S1.st02 as Cst02,s2.st02 as Ust02 from ImgByteFile,staff S1,staff S2 where ibf05='3' and ibf01='000' and ibf02='" & textST01 & "' and ibf03='0' and ibf04='00' and ibf07=s1.st01(+) and ibf10=s2.st01(+) ", cnnConnection, adOpenStatic, adLockOptimistic
   If PicRs.RecordCount <> 0 Then
      ReadPhoto = True
      
      PicRs.MoveFirst
      If CheckStr(PicRs.Fields("ibf06")) = "3" Or CheckStr(PicRs.Fields("ibf06")) = "4" Or CheckStr(PicRs.Fields("ibf06")) = "6" Then
         IsWmf = True
         stAttachFile = App.path & "\NowPic.wmf"
      Else
         IsWmf = False
         stAttachFile = App.path & "\NowPic.jpg"
      End If
      
      'Add By Sindy 2017/8/10
'      If "" & PicRs.Fields("IBF15") <> "" Then
         If PUB_GetFtpFile(PicRs.Fields("IBF15"), stAttachFile, UCase("ImgByteFile")) = False Then
            Screen.MousePointer = vbDefault
            Exit Function
         End If
'      Else
'      '2017/8/10 END
'         ReDim bytes(Val(PicRs.Fields("ibf13").Value))
'         bytes() = PicRs.Fields("ibf14").GetChunk(Val(PicRs.Fields("ibf13").Value))
'         file_num = FreeFile
'         If IsWmf = False Then
'            Open App.path & "\NowPic.jpg" For Binary Access Write As #file_num
'         Else
'            Open App.path & "\NowPic.wmf" For Binary Access Write As #file_num
'         End If
'         Put #file_num, , bytes()
'         Close #file_num
'      End If
      
      G_SeekPicColor.Picture = LoadPicture(App.path & "\NowPic.jpg")
      pWidth = G_SeekPicColor.Width
      pHeight = G_SeekPicColor.Height
      sH = 0: sW = 0
      If pWidth < pHeight Then '¥H°ªªº¤ñ¨Ò
         dblTmp = pHeight / tmpPic.Height
         sH = tmpPic.Height
      Else '¥H¼eªº¤ñ¨Ò
         dblTmp = pWidth / tmpPic.Width
         sW = tmpPic.Width
      End If
      If sW = 0 Then
         sW = pWidth / dblTmp
         'Add By Sindy 2012/7/27
         If sW > tmpPic.Width Then
            '¼e«×µ¥¤ñ¨ÒÁY¤p«áÁÙ¬O¤j©ó¹Ï®Ø¼e,¦A¥H¼eªº¤ñ¨ÒÁY©ñ
            dblTmp = sW / tmpPic.Width
            sW = tmpPic.Width
            sH = sH / dblTmp
         End If
         '2012/7/27 End
      ElseIf sH = 0 Then
         sH = pHeight / dblTmp
         'Add By Sindy 2012/7/27
         If sH > tmpPic.Height Then
            '°ª«×µ¥¤ñ¨ÒÁY¤p«áÁÙ¬O¤j©ó¹Ï®Ø°ª,¦A¥H°ªªº¤ñ¨ÒÁY©ñ
            dblTmp = sH / tmpPic.Height
            sH = tmpPic.Height
            sW = sW / dblTmp
         End If
         '2012/7/27 End
      End If
      tmpImg.Width = sW: tmpImg.Height = sH
      Set tmpImg.Picture = G_SeekPicColor
      'tmpPic.PaintPicture G_SeekPicColor, ((tmpPic.Width - sW) / 2) / 2, ((tmpPic.Height - sH) / 2), sW, sH
      tmpPic.PaintPicture G_SeekPicColor, IIf(tmpPic.ScaleWidth / 2 - (sW / 2) < 0, 0, tmpPic.ScaleWidth / 2 - (sW / 2)), IIf(tmpPic.ScaleHeight / 2 - (sH / 2) < 0, 0, tmpPic.ScaleHeight / 2 - (sH / 2)), sW, sH
      Set tmpPic.Picture = tmpPic.Image
      
      If Dir(App.path & "\NowPic.jpg") <> "" Then
         Kill App.path & "\NowPic.jpg"
      End If
      If Dir(App.path & "\NowPic.wmf") <> "" Then
         Kill App.path & "\NowPic.wmf"
      End If
   End If
   
   Screen.MousePointer = vbDefault
End Function

' §ó·stoolbar¤W«ö¯Ãªºª¬ºA
Private Sub UpdateToolbarState()
   Me.Enabled = False
   Select Case m_EditMode
      ' µL¥ô¦ó°Ê§@
      Case 0:
         If m_bInsert Then
            TBar1.Buttons(1).Enabled = True
         Else
            TBar1.Buttons(1).Enabled = False
         End If
         If m_bUpdate Then
            TBar1.Buttons(2).Enabled = True
         Else
            TBar1.Buttons(2).Enabled = False
         End If
         If m_bDelete Then
            TBar1.Buttons(3).Enabled = True
         Else
            TBar1.Buttons(3).Enabled = False
         End If
         If m_bQuery Then
            TBar1.Buttons(4).Enabled = True
         Else
            TBar1.Buttons(4).Enabled = False
         End If
         If m_bQuery Then
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
         ' ·s¼W
      Case 1, 2, 3, 4:
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
   Me.Enabled = True
End Sub

Private Function CheckDataValid() As Boolean
Dim nResponse As Boolean
Dim strTit As String
Dim strMsg As String
   
   CheckDataValid = False
   
   '2008/12/15 ADD BY SINDY
   ' ­û¤u¥N¸¹¤£¥i¬°ªÅ¥Õ
   If IsEmptyText(textST01) = True Then
      strTit = "ÀË®Ö¸ê®Æ"
      strMsg = "­û¤u¥N¸¹¤£¥i¬°ªÅ¥Õ"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textST01.SetFocus
      GoTo EXITSUB
   End If
   ' ¤¤¤å©m¦W¤£¥i¬°ªÅ¥Õ
   If IsEmptyText(textST02) = True Then
      strTit = "ÀË®Ö¸ê®Æ"
      strMsg = "¤¤¤å©m¦W¤£¥i¬°ªÅ¥Õ"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textST02.SetFocus
      GoTo EXITSUB
   End If
   ' ³¡ªù¤£¥i¬°ªÅ¥Õ
   If IsEmptyText(textST03) = True Then
      strTit = "ÀË®Ö¸ê®Æ"
      strMsg = "³¡ªù¤£¥i¬°ªÅ¥Õ"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textST03.SetFocus
      GoTo EXITSUB
   End If
   ' ­û¤u©ÒÄÝ©Ò§O¤£¥i¬°ªÅ¥Õ
   If IsEmptyText(textST06) = True Then
      strTit = "ÀË®Ö¸ê®Æ"
      strMsg = "­û¤u©ÒÄÝ©Ò§O¤£¥i¬°ªÅ¥Õ"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textST06.SetFocus
      GoTo EXITSUB
   End If
   ' ©Ê§O¤£¥i¬°ªÅ¥Õ
   If IsEmptyText(textST22) = True Then
      'Modify By Sindy 2025/9/19
      '­û¤u½s¸¹¬°FXX®É,©Ê§OÄæªÅ¥Õ®É´£¿ô´N¦n¡A¤£¥²­­¨î¤@©w­n¿é¡F¦ý¦³10½Xªº¨­¥÷ÃÒ¦r¸¹«h¤@©w­n¿é¤J
      If Left(textST01, 1) = "F" And GetTextLength(textSR07.Text) <> 10 Then
         If MsgBox("©Ê§O©|¥¼¿é¤J¡A½T©wÄ~Äò¶Ü¡H", vbYesNo + vbCritical + vbDefaultButton2) = vbNo Then
            textST22.SetFocus
            GoTo EXITSUB
         End If
      Else
      '2025/9/19 END
         strTit = "ÀË®Ö¸ê®Æ"
         strMsg = "©Ê§O¤£¥i¬°ªÅ¥Õ"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textST22.SetFocus
         GoTo EXITSUB
      End If
   End If
   ' ¤J©Ò¤é´Á¤£¥i¬°ªÅ¥Õ
   If Mid(textST01, 1, 1) <> "F" Then
      If IsEmptyText(textST13) = True Then
         strTit = "ÀË®Ö¸ê®Æ"
         strMsg = "¤J©Ò¤é´Á¤£¥i¬°ªÅ¥Õ"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textST13.SetFocus
         GoTo EXITSUB
      End If
   End If
   
   'Add By Sindy 2021/7/19 ±Æ°£ textST01 <= "99053" ¦]ÂÂ­û¤u³£¨S¦³¸Õ¥Î´Á¶¡ªº¸ê®Æ
   If Not (textST01 <= "99053") Then
      'Modify By Sindy 2025/9/19
      '­û¤u½s¸¹¬°FXX®É,¸Õ¥Î´Á¶¡¥i¥H¤£¥Î¿é¤J
      If Left(textST01, 1) <> "F" Then
      '2025/9/19 END
         If IsEmptyText(textST28) = True Then
            strTit = "ÀË®Ö¸ê®Æ"
            strMsg = "¸Õ¥Î°_©l¤é´Á¤£¥i¬°ªÅ¥Õ"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textST28.SetFocus
            GoTo EXITSUB
         End If
         If IsEmptyText(textST29) = True Then
            strTit = "ÀË®Ö¸ê®Æ"
            strMsg = "¸Õ¥ÎºI¤î¤é´Á¤£¥i¬°ªÅ¥Õ"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textST29.SetFocus
            GoTo EXITSUB
         End If
      End If
   End If
   '2021/7/19 END
   
   ' Â¾ºÙ©ÎÂ¾¦ì¦Ü¤Ö­n¿é¤J¤@¶µ!!!
   If IsEmptyText(textST20) = True And IsEmptyText(textST21) = True Then
      If Mid(textST01, 1, 1) <> "F" Then '2009/1/1 ADD BY SONIA  F½s¸¹¤£ÀË¬d
         strTit = "ÀË®Ö¸ê®Æ"
         strMsg = "Â¾ºÙ©ÎÂ¾¦ì¦Ü¤Ö­n¿é¤J¤@¶µ!!!"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textST21.SetFocus
         GoTo EXITSUB
      End If       '2009/1/1 ADD BY SONIA
   End If
   '2008/12/15 END
   
   'Add by Morgan 2009/6/19
   If Trim(txtSD16.Text) <> "" And Trim(txtSD17) = "" Then
      MsgBox "³Ò°h·s¨î¦P¤¯¶·¿é¤J¦Û´£¶O²v¡A­Y¨S¦³¦Û´£½Ð¿é 0 !!"
      Me.SSTab1.Tab = 2
      txtSD17.SetFocus
      GoTo EXITSUB
   End If
   
   nResponse = False
   textST01_Validate nResponse
   If nResponse = True Then GoTo EXITSUB
   nResponse = False
   textST02_Validate nResponse
   If nResponse = True Then GoTo EXITSUB
   nResponse = False
   textST03_Validate nResponse
   If nResponse = True Then GoTo EXITSUB
   nResponse = False
   textST06_Validate nResponse
   If nResponse = True Then GoTo EXITSUB
   nResponse = False
   textST08_Validate nResponse
   If nResponse = True Then GoTo EXITSUB
   nResponse = False
   textST09_Validate nResponse
   If nResponse = True Then GoTo EXITSUB
   nResponse = False
   textST10_Validate nResponse
   If nResponse = True Then GoTo EXITSUB
   nResponse = False
   textST13_Validate nResponse
   If nResponse = True Then GoTo EXITSUB
   nResponse = False
   textST18_Validate nResponse
   If nResponse = True Then GoTo EXITSUB
   nResponse = False
   textST19_Validate nResponse
   If nResponse = True Then GoTo EXITSUB
   nResponse = False
   textST20_Validate nResponse
   If nResponse = True Then GoTo EXITSUB
   nResponse = False
   textST21_Validate nResponse
   If nResponse = True Then GoTo EXITSUB
   nResponse = False
   textST22_Validate nResponse
   If nResponse = True Then GoTo EXITSUB
   nResponse = False
   textST23_Validate nResponse
   If nResponse = True Then GoTo EXITSUB
   nResponse = False
   textST24_Validate nResponse
   If nResponse = True Then GoTo EXITSUB
   nResponse = False
   textST25_Validate nResponse
   If nResponse = True Then GoTo EXITSUB
   nResponse = False
   textST26_Validate nResponse
   If nResponse = True Then GoTo EXITSUB
   nResponse = False
   textST27_Validate nResponse
   If nResponse = True Then GoTo EXITSUB
   nResponse = False
   textST28_Validate nResponse
   If nResponse = True Then GoTo EXITSUB
   nResponse = False
   textST29_Validate nResponse
   If nResponse = True Then GoTo EXITSUB
   ' 2008/12/17 ADD BY SINDY
   nResponse = False
   textST30_Validate nResponse
   If nResponse = True Then GoTo EXITSUB
   ' 2008/12/17 END
   nResponse = False
   textST31_Validate nResponse
   If nResponse = True Then GoTo EXITSUB
   nResponse = False
   textST32_Validate nResponse
   If nResponse = True Then GoTo EXITSUB
   nResponse = False
   textST33_Validate nResponse
   If nResponse = True Then GoTo EXITSUB
   nResponse = False
   textST34_Validate nResponse
   If nResponse = True Then GoTo EXITSUB
   nResponse = False
   textST35_Validate nResponse
   If nResponse = True Then GoTo EXITSUB
   nResponse = False
   textST36_Validate nResponse
   If nResponse = True Then GoTo EXITSUB
   nResponse = False
   textST37_Validate nResponse
   If nResponse = True Then GoTo EXITSUB
   nResponse = False
   textST38_Validate nResponse
   If nResponse = True Then GoTo EXITSUB
   nResponse = False
   textST39_Validate nResponse
   If nResponse = True Then GoTo EXITSUB
   nResponse = False
   textST40_Validate nResponse
   If nResponse = True Then GoTo EXITSUB
   nResponse = False
   textST41_Validate nResponse
   If nResponse = True Then GoTo EXITSUB
   nResponse = False
   textST42_Validate nResponse
   If nResponse = True Then GoTo EXITSUB
   
   nResponse = False
   textST49_Validate nResponse
   If nResponse = True Then GoTo EXITSUB
   
   nResponse = False
   textST12_Validate nResponse
   If nResponse = True Then GoTo EXITSUB
   
   'Add By Sindy 2014/3/12
   nResponse = False
   textSS02_Validate nResponse
   If nResponse = True Then GoTo EXITSUB
   nResponse = False
   textSS03_Validate nResponse
   If nResponse = True Then GoTo EXITSUB
   nResponse = False
   textSS04_Validate nResponse
   If nResponse = True Then GoTo EXITSUB
   nResponse = False
   textSS05_Validate nResponse
   If nResponse = True Then GoTo EXITSUB
   nResponse = False
   textSS06_Validate nResponse
   If nResponse = True Then GoTo EXITSUB
   nResponse = False
   textSS07_Validate nResponse
   If nResponse = True Then GoTo EXITSUB
   '2014/3/12 END
   
   CheckDataValid = True
EXITSUB:
End Function

' §ó·sKeyªºª¬ºA
Private Sub SetKeyReadOnly(ByVal bEnable As Boolean)
   textST01.Locked = bEnable
   If bEnable Then textST01.BackColor = &H8000000F Else textST01.BackColor = &H80000005
End Sub

' §ó·s¦U±±¨î¶µªºª¬ºA
Private Sub SetCtrlReadOnly(ByVal bEnable As Boolean)
Dim i As Integer
   
   textST01.Locked = bEnable
   If bEnable Then textST01.BackColor = &H8000000F Else textST01.BackColor = &H80000005
   textST02.Locked = bEnable
   textST03.Enabled = Not bEnable
   textST06.Enabled = Not bEnable
   textST08.Locked = bEnable
   textST09.Locked = bEnable
   textST10.Locked = bEnable
   textST13.Locked = bEnable
   textST18.Locked = bEnable
   textST19.Locked = bEnable
   textST20.Enabled = Not bEnable
   textST21.Enabled = Not bEnable
   textST22.Enabled = Not bEnable
   textST23.Locked = bEnable
   textST24.Enabled = Not bEnable
   textST25.Enabled = Not bEnable
   textST26.Locked = bEnable
   textST27.Enabled = Not bEnable
   textST28.Locked = bEnable
   textST29.Locked = bEnable
   textST30.Locked = bEnable
   textST31.Locked = bEnable
   textST32.Locked = bEnable
   textST33.Locked = bEnable
   textST34.Locked = bEnable
   textST35.Locked = bEnable
   textST36.Locked = bEnable
   textST37.Enabled = Not bEnable
   textST38.Locked = bEnable
   textST39.Locked = bEnable
   'Modify By Sindy 2018/5/16 ¶}©ñ¥i¥H­×§ï,¦ý­n¦^¼gYEARVACATION
   '2010/1/11 modify by sonia ¤£¥i§ï¥i¥ð°²¤Ñ¼Æ,¦]¬°¦~²×¼úª÷¬O­pºâ¥H¯S§O°²°O¿ýÀÉYEARVACATION­pºâ,­Y¦¹³B­×§ï«h·|¤£¤@­P
   If m_EditMode = 1 Or textST40 = "" Or LblST40 = "" Then '1.·s¼W
      textST40.Locked = True
   Else
      textST40.Locked = bEnable
   End If
'   textST40.Locked = True
'   textST40.Enabled = False
   '2010/1/11 end
   '2018/5/16 END
   textST41.Locked = bEnable
   textST42.Locked = bEnable
   textST49.Locked = bEnable
   textST12.Locked = bEnable
   
   'Add by Morgan 2009/6/26 --³Ò«Ø«O¸É§U¤H¨Æ¥u¯à¬Ý,°]°È¤~¯à§ï
   cboST50.Enabled = False
   cboST56.Enabled = False
   txtSD16.Locked = bEnable
   txtSD17.Locked = bEnable
   'end 2009/6/24
   
   cmdOK(0).Enabled = Not bEnable
'Modify by Morgan 2009/6/11
'   If cmdok(0).Enabled = True Then
'        For i = 1 To GRD1.Rows - 1
'            GRD1.row = i
'            GRD1.col = 0
'            If GRD1.CellBackColor = &HFFC0C0 Then
'               cmdok(1).Enabled = Not bEnable
'               cmdok(2).Enabled = Not bEnable
'               Exit For ' 2009/01/05 Add BY Sindy
'
'            Else
'               cmdok(1).Enabled = bEnable
'               cmdok(2).Enabled = bEnable
'            End If
'            'Exit For
'        Next i
'   Else
'        cmdok(1).Enabled = Not bEnable
'        cmdok(2).Enabled = Not bEnable
'   End If
'   textSR03.Enabled = Not bEnable
'   textSR04.Locked = bEnable
'   textSR05.Enabled = Not bEnable
'   textSR06.Locked = bEnable
'   textSR07.Locked = bEnable
'   textSR09.Locked = bEnable
'   textSR10.Locked = bEnable
'   textSR11.Locked = bEnable
'   textSR12.Locked = bEnable
   
   EnableSR False
   cmdOK(1).Enabled = False
   cmdOK(2).Enabled = False
   cmdOK(3).Enabled = Not bEnable
   cmdOK(4).Enabled = Not bEnable
   'cmdLog.Enabled = Not bEnable
   GRD1.Enabled = True
   
   'Add By Sindy 2014/3/12
   textSS02.Locked = bEnable
   textSS03.Locked = bEnable
   textSS04.Locked = bEnable
   textSS05.Locked = bEnable
   textSS06.Locked = bEnable
   textSS07.Locked = bEnable
   '2014/3/12 END
   
   Command1.Enabled = Not bEnable 'Add By Sindy 2012/6/20
'end 2009/6/11
End Sub

Private Sub ClearField()
Dim nIndex As Integer
   
   textST01 = Empty
   textST02 = Empty
   textST03 = Empty
   textST06 = Empty
   textST08 = Empty
   textST09 = Empty
   textST10 = Empty
   textST13 = Empty
   textST18 = Empty
   textST19 = Empty
   textST20 = Empty
   textST21 = Empty
   textST22 = Empty
   textST23 = Empty
   textST24 = Empty
   textST25 = Empty
   textST26 = Empty
   textST27 = Empty
   textST28 = Empty
   textST29 = Empty
   textST30 = Empty
   textST31 = Empty
   textST32 = Empty
   textST33 = Empty
   textST34 = Empty
   textST35 = Empty
   textST36 = Empty
   textST37 = Empty
   textST38 = Empty
   textST39 = Empty
   textST40 = Empty: LblST40 = Empty
   textST41 = Empty
   textST42 = Empty
   textST49 = Empty
   textST12 = Empty
   textST51 = Empty
   textST68 = Empty 'Add By Sindy 2015/8/13
   Label23 = Empty
   ' 2008/12/17 ADD BY SINDY
   LabelST30 = Empty
   ' 2008/12/17 END
   GRD1.Clear
   GRD1.Rows = 2
   GRD2.Clear
   GRD2.Rows = 2
   ClearSR
   SetGrd
   For nIndex = 0 To tf_st - 1
      m_FieldList(nIndex).fiOldData = Empty
      m_FieldList(nIndex).fiNewData = Empty
   Next nIndex
   
   'Add by Morgan 2009/6/24
   cboST50.ListIndex = 0
   cboST56.ListIndex = 0
   txtSD16 = "Y"
   txtSD17 = Empty
   'end 2009/6/24
   
   'Add By Sindy 2014/3/12
   textSS02 = Empty
   textSS03 = Empty
   textSS04 = Empty
   textSS05 = Empty
   textSS06 = Empty
   textSS07 = Empty
   '2014/3/12 END
   
   'Add By Sindy 2019/6/25
   txtBackDate = Empty
   '2019/6/25 END
   
   tmpPic.Picture = LoadPicture() 'Add By Sindy 2012/6/20 ²M¹Ï¤ù
   'Add By Sindy 2024/3/19
   textST03.Tag = "ST93" 'Add By Sindy 2024/3/13 ¹w³]­È
   Label1(2).Caption = "³¡ªù"
   '2024/3/19 END
End Sub
Sub ClearSR()
   textSR03 = Empty
   textSR04 = Empty
   textSR05 = Empty
   textSR06 = Empty
   textSR07 = Empty
   textSR09 = Empty
   textSR10 = Empty
   textSR11 = Empty
   textSR12 = Empty
   chkSR08.Value = vbUnchecked
   chkSR13.Value = vbUnchecked
   'Add by Morgan 2009/6/10
   cboHL05.ListIndex = 0
End Sub

Private Sub UpdateFieldNewData()
Dim MyArr As Variant
   
   '­Y·s¼W¸ê®Æ
   If m_EditMode = 1 Then
      SetFieldNewData "ST01", textST01
      SetFieldNewData "ST04", "1"           '2008/12/18 ADD BY SONIA
      'Modify By Sindy 2023/12/20
      If strSrvDate(1) < ·s³¡ªù±Ò¥Î¤é Then
      '2023/12/20 END
         '2012/10/2 ADD BY SONIA ¥[¦¬¤å·~°È°Ï
         If textST03.Text <> "" Then
              MyArr = Split(textST03, " ")
              SetFieldNewData "ST15", MyArr(0)
         Else
              SetFieldNewData "ST15", Empty
         End If
         '2012/10/2 END
      End If
   End If
   'Modify By Sindy 2014/3/3 ex.A3007³¯¬¸¡¨Ž«¡¨¬O³y¦r,«á­±­n¥[¡¨ªÅ¥Õ¡¨¤~¯à¦s¤J,©Ò¥H¤£¯àtrim±¼
   'Modify By Sindy 2014/3/3 ¸g²z»¡¤£­n¥[ªÅ¥Õ,¦³¯S®í³y¦r¦s¤£°_¨Ó®É,³qª¾¥L³B²z
   SetFieldNewData "ST02", Trim(textST02)
   'SetFieldNewData "ST02", textST02
   '2014/3/3 END
   'Modify By Sindy 2023/12/20
   If strSrvDate(1) >= ·s³¡ªù±Ò¥Î¤é Then
      If textST03.Text <> "" Then
           MyArr = Split(textST03, " ")
           SetFieldNewData "ST93", MyArr(0)
      Else
           SetFieldNewData "ST93", Empty
      End If
   Else
   '2023/12/20 END
      If textST03.Text <> "" Then
           MyArr = Split(textST03, " ")
           SetFieldNewData "ST03", MyArr(0)
      Else
           SetFieldNewData "ST03", Empty
      End If
   End If
   
   If textST06.Text <> "" Then
        MyArr = Split(textST06, " ")
        SetFieldNewData "ST06", MyArr(0)
   Else
        SetFieldNewData "ST06", Empty
   End If
   SetFieldNewData "ST08", ChgSQL(textST08)
   SetFieldNewData "ST09", textST09
   SetFieldNewData "ST10", textST10
   SetFieldNewData "ST13", DBDATE(textST13)
   SetFieldNewData "ST18", textST18
   SetFieldNewData "ST19", textST19
   If textST20.Text <> "" Then
        MyArr = Split(textST20, " ")
        SetFieldNewData "ST20", MyArr(0)
   Else
        SetFieldNewData "ST20", Empty
   End If
   If textST21.Text <> "" Then
        MyArr = Split(textST21, " ")
        SetFieldNewData "ST21", MyArr(0)
   Else
        SetFieldNewData "ST21", Empty
   End If
   If textST22.Text <> "" Then
        MyArr = Split(textST22, " ")
        SetFieldNewData "ST22", MyArr(0)
   Else
        SetFieldNewData "ST22", Empty
   End If
   SetFieldNewData "ST23", DBDATE(textST23)
   If textST24.Text <> "" Then
        MyArr = Split(textST24, " ")
        SetFieldNewData "ST24", MyArr(0)
   Else
        SetFieldNewData "ST24", Empty
   End If
   If textST25.Text <> "" Then
        MyArr = Split(textST25, " ")
        SetFieldNewData "ST25", MyArr(0)
   Else
        SetFieldNewData "ST25", Empty
   End If
   SetFieldNewData "ST26", textST26
   If textST27.Text <> "" Then
        MyArr = Split(textST27, " ")
        SetFieldNewData "ST27", MyArr(0)
   Else
        SetFieldNewData "ST27", Empty
   End If
   SetFieldNewData "ST28", DBDATE(textST28)
   SetFieldNewData "ST29", DBDATE(textST29)
   SetFieldNewData "ST30", textST30
   SetFieldNewData "ST31", DBDATE(textST31)
   SetFieldNewData "ST32", DBDATE(textST32)
   SetFieldNewData "ST33", textST33
   SetFieldNewData "ST34", ChgSQL(textST34)
   SetFieldNewData "ST35", textST35
   SetFieldNewData "ST36", textST36
   If textST37.Text <> "" Then
        MyArr = Split(textST37, " ")
        SetFieldNewData "ST37", MyArr(0)
   Else
        SetFieldNewData "ST37", Empty
   End If
   SetFieldNewData "ST38", ChgSQL(textST38)
   SetFieldNewData "ST39", ChgSQL(textST39)
   SetFieldNewData "ST40", textST40
   SetFieldNewData "ST41", DBDATE(textST41)
   SetFieldNewData "ST42", textST42
   SetFieldNewData "ST49", ChgSQL(textST49)
   SetFieldNewData "ST12", textST12
   
   'Add by Morgan 2009/6/15 ­û¤u³Ò°·«O¸É§UÃþ§O
   If cboST50.ListIndex > 0 Then
      SetFieldNewData "ST50", Left(cboST50, 2)
   Else
      SetFieldNewData "ST50", Empty
   End If
   If cboST56.ListIndex > 0 Then
      SetFieldNewData "ST56", Left(cboST56, 2)
   Else
      SetFieldNewData "ST56", Empty
   End If
   'end 2009/6/24
   If m_EditMode = 1 Then 'Add by Morgan 2011/3/9 ·s¼W¤~³]
      SetFieldNewData "ST58", "N" 'Add by Morgan 2011/1/10 ¬O§_¦Û°Ê¦¬¤å¹w³] N,µe­±¤£©ñ
      SetFieldNewData "ST64", "N" 'Add by Sindy 2015/4/23 ³]©w­^®Öªí¹w³] N,µe­±¤£©ñ
   End If
End Sub

' ªì©l¤ÆÄæ¦ì°}¦C
'Modified by Morgan 2023/12/15
'Private Sub InitialField()
Private Sub InitialField(pRst As ADODB.Recordset)
   Dim nIndex As Integer
   Dim strTmp As String
   ReDim m_FieldList(tf_st) As FIELDITEM 'Added by Morgan 2023/12/15
   
   ' ªì©l¤ÆÄæ¦ì°}¦C
   For nIndex = 1 To tf_st
      'Modified by Morgan 2023/12/15 ¦]±qST73,ª½±µ·s¼WST93¤§¬G
      'strTmp = Format(nIndex, "00")
      'm_FieldList(nIndex - 1).fiName = "ST" & strTmp
      m_FieldList(nIndex - 1).fiName = pRst.Fields(nIndex - 1).Name
      'end 2023/12/15
      m_FieldList(nIndex - 1).fiOldData = Empty
      m_FieldList(nIndex - 1).fiNewData = Empty
      m_FieldList(nIndex - 1).fiType = 0 '¤å¦r«¬ºA
      'Modified by Morgan 2023/12/15
      'Select Case nIndex
      Select Case Val(Right(m_FieldList(nIndex - 1).fiName, 2))
      'end 2023/12/15
         Case 13, 23, 28, 29, 31, 32, 40, 41, 44, 45, 47, 48:
            m_FieldList(nIndex - 1).fiType = 1 '¼Æ­È«¬ºA
      End Select
   Next nIndex
End Sub

'±a¹w³]¸ê®Æ
Private Sub InitialData()
Dim MyRs As New ADODB.Recordset
   
   textST03.Clear
   Set MyRs = New ADODB.Recordset
   If MyRs.State = 1 Then MyRs.Close
   '2009/3/2 modify by sonia
   'strSQL = "select a0901||' '||a0902 from acc090 order by a0901"
   'Modify By Sindy 2023/12/20
   If strSrvDate(1) >= ·s³¡ªù±Ò¥Î¤é Then
      strSql = "select a0921||' '||a0922 from acc090NEW order by a0921"
   Else
   '2023/12/20 END
      strSql = "select a0901||' '||a0902 from acc090 where a0904<>'Y' and a0901<>'CFL' order by a0901"
   End If
   MyRs.CursorLocation = adUseClient
   MyRs.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If MyRs.RecordCount <> 0 Then
       While Not MyRs.EOF
           textST03.AddItem "" & MyRs.Fields(0).Value
           MyRs.MoveNext
       Wend
   End If
   textST20.Clear
   Set MyRs = New ADODB.Recordset
   If MyRs.State = 1 Then MyRs.Close
   strSql = "select ac02||' '||ac03 from allcode where ac01='01' order by ac02"
   MyRs.CursorLocation = adUseClient
   MyRs.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If MyRs.RecordCount <> 0 Then
       While Not MyRs.EOF
           textST20.AddItem "" & MyRs.Fields(0).Value
           MyRs.MoveNext
       Wend
   End If
   textST21.Clear
   Set MyRs = New ADODB.Recordset
   If MyRs.State = 1 Then MyRs.Close
   strSql = "select ac02||' '||ac03 from allcode where ac01='02' order by ac02"
   MyRs.CursorLocation = adUseClient
   MyRs.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If MyRs.RecordCount <> 0 Then
       While Not MyRs.EOF
           textST21.AddItem "" & MyRs.Fields(0).Value
           MyRs.MoveNext
       Wend
   End If
   textST27.Clear
   Set MyRs = New ADODB.Recordset
   If MyRs.State = 1 Then MyRs.Close
   strSql = "select ac02||' '||ac03 from allcode where ac01='06' order by ac02"
   MyRs.CursorLocation = adUseClient
   MyRs.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If MyRs.RecordCount <> 0 Then
       While Not MyRs.EOF
           textST27.AddItem "" & MyRs.Fields(0).Value
           MyRs.MoveNext
       Wend
   End If
   textST37.Clear
   Set MyRs = New ADODB.Recordset
   If MyRs.State = 1 Then MyRs.Close
   strSql = "select ac02||' '||ac03 from allcode where ac01='03' order by ac02"
   MyRs.CursorLocation = adUseClient
   MyRs.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If MyRs.RecordCount <> 0 Then
       While Not MyRs.EOF
           textST37.AddItem "" & MyRs.Fields(0).Value
           MyRs.MoveNext
       Wend
   End If
   
   'Add by Morgan 2009/6/24
   cboHL05.Clear
   cboHL05.AddItem "µL"
   cboST50.Clear
   cboST50.AddItem "µL"
   cboST56.Clear
   cboST56.AddItem "µL"
   Set MyRs = New ADODB.Recordset
   If MyRs.State = 1 Then MyRs.Close
   strSql = "select HR01||' '||HR04 from HiReduce order by 1"
   MyRs.CursorLocation = adUseClient
   MyRs.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If MyRs.RecordCount <> 0 Then
       While Not MyRs.EOF
           cboHL05.AddItem "" & MyRs.Fields(0).Value
           cboST50.AddItem "" & MyRs.Fields(0).Value
           MyRs.MoveNext
       Wend
   End If
   Set MyRs = New ADODB.Recordset
   If MyRs.State = 1 Then MyRs.Close
   strSql = "select LR01||' '||LR04 from LiReduce order by 1"
   MyRs.CursorLocation = adUseClient
   MyRs.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If MyRs.RecordCount <> 0 Then
       While Not MyRs.EOF
           cboST56.AddItem "" & MyRs.Fields(0).Value
           MyRs.MoveNext
       Wend
   End If
   'end 2009/6/24
   SetGrd
End Sub

Private Sub textSR03_GotFocus()
   If m_EditMode <> 0 Then
       textSR03.SetFocus
       InverseTextBox textSR03
   End If
End Sub

Private Sub textSR03_Validate(Cancel As Boolean)
Dim MyArr As Variant
Dim MyArr2 As Variant
Dim Myi As Integer
   
   If m_EditMode <> 0 And textSR03.Text <> "" Then
       MyArr = Split(textSR03, " ")
       For Myi = 0 To textSR03.ListCount - 1
           MyArr2 = Split(textSR03.List(Myi), " ")
           If MyArr(0) = MyArr2(0) Then
               textSR03.Text = textSR03.List(Myi)
               Exit Sub
           End If
       Next Myi
       If m_EditMode <> 0 Then
           MsgBox "ºÙ¿×¥N¸¹¿é¤J¿ù»~!!!", vbExclamation + vbOKOnly
           ' 2008/12/17 ADD BY SINDY
           Call textSR03_GotFocus
           ' 2008/12/17 END
           Cancel = True
           Exit Sub
       End If
   End If
End Sub

Private Sub textSR04_GotFocus()
   If m_EditMode <> 0 Then
       InverseTextBox textSR04
       OpenIme
   End If
End Sub

Private Sub textSR04_Validate(Cancel As Boolean)
   If m_EditMode <> 0 And textSR04 <> "" Then
       If CheckLengthIsOK(textSR04, textSR04.MaxLength) = False Then
           ' 2008/12/17 ADD BY SINDY
           Call textSR04_GotFocus
           ' 2008/12/17 END
           Cancel = True
           Exit Sub
       End If
   End If
   CloseIme
End Sub

Private Sub textSR05_GotFocus()
   If m_EditMode <> 0 Then
       InverseTextBox textSR05
   End If
End Sub

Private Sub textSR05_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textSR05_Validate(Cancel As Boolean)
Dim MyArr As Variant
Dim MyArr2 As Variant
Dim Myi As Integer
   
   If textSR05.Text <> "" And m_EditMode <> 0 Then
       MyArr = Split(textSR05, " ")
       For Myi = 0 To textSR05.ListCount - 1
           MyArr2 = Split(textSR05.List(Myi), " ")
           If MyArr(0) = MyArr2(0) Then
               textSR05.Text = textSR05.List(Myi)
               Exit Sub
           End If
       Next Myi
       If m_EditMode <> 0 Then
           MsgBox "©Ê§O¥N¸¹¿é¤J¿ù»~!!!", vbExclamation + vbOKOnly
           Call textSR05_GotFocus
           Cancel = True
           Exit Sub
       End If
   End If
End Sub

Private Sub textSR06_GotFocus()
   If m_EditMode <> 0 Then
       InverseTextBox textSR06
   End If
End Sub

Private Sub textSR06_KeyPress(KeyAscii As Integer)
   KeyAscii = Pub_NumAscii(KeyAscii)
End Sub

Private Sub textSR06_Validate(Cancel As Boolean)
   If m_EditMode <> 0 And textSR06 <> "" Then
      If textSR06 = "" Then Exit Sub
      ' 2008/12/17 MODIFY BY SINDY
      'Cancel = Not ChkDate(textSR06)
      If ChkDate(textSR06) = False Then
         Call textSR06_GotFocus
         Cancel = True
         Exit Sub
      End If
      ' 2008/12/17 END
   End If
End Sub

Private Sub textSR07_GotFocus()
   If m_EditMode <> 0 Then
       InverseTextBox textSR07
   End If
End Sub

Private Sub textSR07_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textSR07_Validate(Cancel As Boolean)
   If m_EditMode <> 0 And textSR07 <> "" Then
       If CheckLengthIsOK(textSR07, textSR07.MaxLength) = False Then
           ' 2008/12/17 ADD BY SINDY
           Call textSR07_GotFocus
           ' 2008/12/17 END
           Cancel = True
           Exit Sub
       End If
       
       ' 2008/12/17 ADD BY SINDY
      If m_EditMode <> 1 And m_EditMode <> 2 Then Exit Sub
      If textSR07.Text = "" Then Exit Sub
      Dim strTmp As String
      If GetTextLength(textSR07.Text) <> 10 Then
         Call textSR07_GotFocus
         strTmp = "¨­¥÷ÃÒ¥²¶·¬O10½X ! ½Ð½T©w ?"
         If MsgBox(strTmp, vbYesNo + vbCritical) = vbNo Then
            Cancel = True
            Exit Sub
         End If
      End If
      If CheckID(0, textSR07.Text) = False Then
         Call textSR07_GotFocus
         strTmp = "¨­¤ÀÃÒ¦r¸¹¿ù»~¡A¬O§_½T©w ?"
         If MsgBox(strTmp, vbYesNo + vbCritical) = vbNo Then
            Cancel = True
         End If
      End If
      ' 2008/12/17 END
   End If
   CloseIme
End Sub

Private Sub textSR09_GotFocus()
   If m_EditMode <> 0 Then
       InverseTextBox textSR09
   End If
End Sub

Private Sub textSR09_Validate(Cancel As Boolean)
   If m_EditMode <> 0 And textSR09 <> "" Then
       If CheckLengthIsOK(textSR09, textSR09.MaxLength) = False Then
           ' 2008/12/17 ADD BY SINDY
           Call textSR09_GotFocus
           ' 2008/12/17 END
           Cancel = True
           Exit Sub
       End If
   End If
   CloseIme
End Sub

Private Sub textSR10_GotFocus()
   If m_EditMode <> 0 Then
       InverseTextBox textSR10
       ' 2008/12/17 ADD BY SINDY
       OpenIme
       ' 2008/12/17 END
   End If
End Sub

' 2008/12/17 ADD BY SINDY
Private Sub textSR10_KeyPress(KeyAscii As Integer)
   KeyAscii = ChangeZIP(KeyAscii)
End Sub
' 2008/12/17 END

Private Sub textSR10_Validate(Cancel As Boolean)
   If m_EditMode <> 0 And textSR10 <> "" Then
       If CheckLengthIsOK(textSR10, textSR10.MaxLength) = False Then
           ' 2008/12/17 ADD BY SINDY
           Call textSR10_GotFocus
           ' 2008/12/17 END
           Cancel = True
           Exit Sub
       End If
   End If
   CloseIme
End Sub

Private Sub textSR11_GotFocus()
   If m_EditMode <> 0 Then
       InverseTextBox textSR11
       OpenIme
   End If
End Sub

' 2008/12/17 ADD BY SINDY
Private Sub textSR11_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = ChangeZIP(KeyAscii)
End Sub
' 2008/12/17 END

Private Sub textSR11_Validate(Cancel As Boolean)
   If m_EditMode <> 0 And textSR11 <> "" Then
       If CheckLengthIsOK(textSR11, textSR11.MaxLength) = False Then
           ' 2008/12/17 ADD BY SINDY
           Call textSR11_GotFocus
           ' 2008/12/17 END
           Cancel = True
           Exit Sub
       End If
   End If
   CloseIme
End Sub

Private Sub textSR12_GotFocus()
   If m_EditMode <> 0 Then
       InverseTextBox textSR12
   End If
End Sub

Private Sub textSR12_KeyPress(KeyAscii As Integer)
   KeyAscii = 0
End Sub

Private Sub textSR12_Validate(Cancel As Boolean)
   If m_EditMode <> 0 And textSR12 <> "" Then
      If textSR12 = "" Then Exit Sub
      If ChkDate(textSR12) = False Then
         Call textSR12_GotFocus
         Cancel = True
         Exit Sub
      End If
   End If
End Sub

'Add By Sindy 2014/3/12
Private Sub textSS02_GotFocus()
   If m_EditMode <> 0 Then
       InverseTextBox textSS02
       OpenIme
   End If
End Sub
Private Sub textSS02_Validate(Cancel As Boolean)
   If m_EditMode <> 0 And textSS02 <> "" Then
       If CheckLengthIsOK(textSS02, textSS02.MaxLength) = False Then
           Call textSS02_GotFocus
           Cancel = True
           Exit Sub
       End If
   End If
   CloseIme
End Sub
Private Sub textSS03_GotFocus()
   If m_EditMode <> 0 Then
       InverseTextBox textSS03
       OpenIme
   End If
End Sub
Private Sub textSS03_Validate(Cancel As Boolean)
   If m_EditMode <> 0 And textSS03 <> "" Then
       If CheckLengthIsOK(textSS03, textSS03.MaxLength) = False Then
           Call textSS03_GotFocus
           Cancel = True
           Exit Sub
       End If
   End If
   CloseIme
End Sub
Private Sub textSS04_GotFocus()
   If m_EditMode <> 0 Then
       InverseTextBox textSS04
       OpenIme
   End If
End Sub
Private Sub textSS04_Validate(Cancel As Boolean)
   If m_EditMode <> 0 And textSS04 <> "" Then
       If CheckLengthIsOK(textSS04, textSS04.MaxLength) = False Then
           Call textSS04_GotFocus
           Cancel = True
           Exit Sub
       End If
   End If
   CloseIme
End Sub
Private Sub textSS05_GotFocus()
   If m_EditMode <> 0 Then
       InverseTextBox textSS05
       OpenIme
   End If
End Sub
Private Sub textSS05_Validate(Cancel As Boolean)
   If m_EditMode <> 0 And textSS05 <> "" Then
       If CheckLengthIsOK(textSS05, textSS05.MaxLength) = False Then
           Call textSS05_GotFocus
           Cancel = True
           Exit Sub
       End If
   End If
   CloseIme
End Sub
Private Sub textSS06_GotFocus()
   If m_EditMode <> 0 Then
       InverseTextBox textSS06
       OpenIme
   End If
End Sub
Private Sub textSS06_Validate(Cancel As Boolean)
   If m_EditMode <> 0 And textSS06 <> "" Then
       If CheckLengthIsOK(textSS06, textSS06.MaxLength) = False Then
           Call textSS06_GotFocus
           Cancel = True
           Exit Sub
       End If
   End If
   CloseIme
End Sub
Private Sub textSS07_GotFocus()
   If m_EditMode <> 0 Then
       InverseTextBox textSS07
       OpenIme
   End If
End Sub
Private Sub textSS07_Validate(Cancel As Boolean)
   If m_EditMode <> 0 And textSS07 <> "" Then
       If CheckLengthIsOK(textSS07, textSS07.MaxLength) = False Then
           Call textSS07_GotFocus
           Cancel = True
           Exit Sub
       End If
   End If
   CloseIme
End Sub
'2014/3/12 END

Private Sub textST01_GotFocus()
   If m_EditMode <> 0 Then
       InverseTextBox textST01
   End If
End Sub

Private Sub textST01_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textST01_Validate(Cancel As Boolean)
   If m_EditMode = 1 And textST01 <> "" Then
       If IsRecordExist(textST01) = True And textST01.Enabled = True And textST01.Locked = False Then
           MsgBox "¦¹½s¸¹¤w¸g¦s¦b¡A½Ð½T»{¡I", vbInformation
           ' 2008/12/17 ADD BY SINDY
           Call textST01_GotFocus
           ' 2008/12/17 END
           Cancel = True
           Exit Sub
       Else
            ' 2008/12/16 ADD BY SINDY
            ' ÀË¬d­û¤u½s¸¹³W«h
            If ChkStaffID(textST01) Then
               Call textST01_GotFocus
               Cancel = True
               Exit Sub
            End If
            ' 2008/12/16 END
       End If
   End If
End Sub

Private Sub textST02_GotFocus()
   If m_EditMode <> 0 Then
       InverseTextBox textST02
       OpenIme
   End If
End Sub

Private Sub textST02_Validate(Cancel As Boolean)
   If m_EditMode <> 0 And textST02 <> "" Then
       If CheckLengthIsOK(textST02, textST02.MaxLength) = False Then
           ' 2008/12/17 ADD BY SINDY
           Call textST02_GotFocus
           ' 2008/12/17 END
           Cancel = True
           Exit Sub
       End If
   End If
   CloseIme
End Sub

Private Sub textST03_GotFocus()
   If m_EditMode <> 0 Then
       InverseTextBox textST03
   End If
End Sub

Private Sub textST03_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textST03_Validate(Cancel As Boolean)
Dim MyRs As New ADODB.Recordset
Dim MyArr As Variant
   
   If textST03.Text <> "" Then
       MyArr = Split(textST03, " ")
       Set MyRs = New ADODB.Recordset
       If MyRs.State = 1 Then MyRs.Close
       'Modify By Sindy 2023/12/20
       'Modify By Sindy 2024/1/9 + And Not (textST03.Tag = "ST03" And textST03.Tag <> "")
       If strSrvDate(1) >= ·s³¡ªù±Ò¥Î¤é And textST03.Tag <> "ST03" Then
         strSql = "select a0921||' '||a0922 from acc090NEW where a0921='" & MyArr(0) & "' order by a0921"
         Label1(2).Caption = "³¡ªù" 'Add By Sindy 2024/1/11
       Else
       '2023/12/20 END
         strSql = "select a0901||' '||a0902 from acc090 where a0901='" & MyArr(0) & "' order by a0901"
         Label1(2).Caption = "ÂÂ³¡ªù" 'Add By Sindy 2024/1/11
       End If
       MyRs.CursorLocation = adUseClient
       MyRs.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
       If MyRs.RecordCount <> 0 Then
           textST03.Text = "" & MyRs.Fields(0).Value
       Else
           If m_EditMode <> 0 Then
               MsgBox "³¡ªù¥N¸¹¿é¤J¿ù»~!!!", vbExclamation + vbOKOnly
               ' 2008/12/17 ADD BY SINDY
               Call textST03_GotFocus
               ' 2008/12/17 END
               Cancel = True
               Exit Sub
           End If
       End If
   End If
End Sub

Private Sub textST06_GotFocus()
   If m_EditMode <> 0 Then
       InverseTextBox textST06
   End If
End Sub

Private Sub textST06_Validate(Cancel As Boolean)
Dim MyArr As Variant
Dim MyArr2 As Variant
Dim Myi As Integer

   If textST06.Text <> "" Then
       MyArr = Split(textST06, " ")
       For Myi = 0 To textST06.ListCount - 1
           MyArr2 = Split(textST06.List(Myi), " ")
           If MyArr(0) = MyArr2(0) Then
               textST06.Text = textST06.List(Myi)
               Exit Sub
           End If
       Next Myi
       If m_EditMode <> 0 Then
           MsgBox "©Ò§O¥N¸¹¿é¤J¿ù»~!!!", vbExclamation + vbOKOnly
           ' 2008/12/17 ADD BY SINDY
           Call textST06_GotFocus
           ' 2008/12/17 END
           Cancel = True
           Exit Sub
       End If
   End If
End Sub

Private Sub textST08_GotFocus()
   If m_EditMode <> 0 Then
       InverseTextBox textST08
       OpenIme
   End If
End Sub

' 2008/12/17 ADD BY SINDY
Private Sub textST08_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = ChangeZIP(KeyAscii)
End Sub
' 2008/12/17 END

Private Sub textST08_Validate(Cancel As Boolean)
   If m_EditMode <> 0 And textST08 <> "" Then
       If CheckLengthIsOK(textST08, textST08.MaxLength) = False Then
           ' 2008/12/17 ADD BY SINDY
           Call textST08_GotFocus
           ' 2008/12/17 END
           Cancel = True
           Exit Sub
       End If
   End If
   CloseIme
End Sub

Private Sub textST09_GotFocus()
   If m_EditMode <> 0 Then
       InverseTextBox textST09
   End If
End Sub

Private Sub textST09_Validate(Cancel As Boolean)
   If m_EditMode <> 0 And textST09 <> "" Then
       If CheckLengthIsOK(textST09, textST09.MaxLength) = False Then
           ' 2008/12/17 ADD BY SINDY
           Call textST09_GotFocus
           ' 2008/12/17 END
           Cancel = True
           Exit Sub
       End If
   End If
   CloseIme
End Sub

Private Sub textST10_GotFocus()
   If m_EditMode <> 0 Then
       InverseTextBox textST10
   End If
End Sub

Private Sub textST10_Validate(Cancel As Boolean)
   If m_EditMode <> 0 And textST10 <> "" Then
       If CheckLengthIsOK(textST10, textST10.MaxLength) = False Then
           ' 2008/12/17 ADD BY SINDY
           Call textST10_GotFocus
           ' 2008/12/17 END
           Cancel = True
           Exit Sub
       End If
   End If
   CloseIme
End Sub

Private Sub textST13_GotFocus()
   If m_EditMode <> 0 Then
       InverseTextBox textST13
   End If
End Sub

Private Sub textST13_KeyPress(KeyAscii As Integer)
   KeyAscii = Pub_NumAscii(KeyAscii)
End Sub

Private Sub textST13_Validate(Cancel As Boolean)
   If m_EditMode <> 0 And textST13 <> "" Then
      If textST13 = "" Then Exit Sub
      ' 2008/12/17 MODIFY BY SINDY
      'Cancel = Not ChkDate(textST13)
      If ChkDate(textST13) = False Then
         Call textST13_GotFocus
         Cancel = True
         Exit Sub
      End If
      ' 2008/12/17 END
   End If
End Sub

Private Sub textST18_GotFocus()
   If m_EditMode <> 0 Then
       InverseTextBox textST18
   End If
End Sub

Private Sub textST18_Validate(Cancel As Boolean)
   If m_EditMode <> 0 And textST18 <> "" Then
       If CheckLengthIsOK(textST18, textST18.MaxLength) = False Then
           Call textST18_GotFocus
           Cancel = True
           Exit Sub
       End If
       ' 2008/12/17 ADD BY SINDY
       If PUB_CheckMail(textST18.Text) = False Then
          Call textST18_GotFocus
          Cancel = True
          Exit Sub
       End If
       ' 2008/12/17 END
   End If
   CloseIme
End Sub

Private Sub textST19_GotFocus()
   If m_EditMode <> 0 Then
       InverseTextBox textST19
   End If
End Sub

Private Sub textST19_Validate(Cancel As Boolean)
   If m_EditMode <> 0 And textST19 <> "" Then
       If CheckLengthIsOK(textST19, textST19.MaxLength) = False Then
           ' 2008/12/17 ADD BY SINDY
           Call textST19_GotFocus
           ' 2008/12/17 END
           Cancel = True
           Exit Sub
       End If
   End If
   CloseIme
End Sub

Private Sub textST20_GotFocus()
   If m_EditMode <> 0 Then
       InverseTextBox textST20
   End If
End Sub

Private Sub textST20_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textST20_Validate(Cancel As Boolean)
Dim MyRs As New ADODB.Recordset
Dim MyArr As Variant

   If textST20.Text <> "" Then
       MyArr = Split(textST20, " ")
       Set MyRs = New ADODB.Recordset
       If MyRs.State = 1 Then MyRs.Close
       strSql = "select ac02||' '||ac03 from allcode where ac02='" & MyArr(0) & "' and ac01='01' order by ac02"
       MyRs.CursorLocation = adUseClient
       MyRs.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
       If MyRs.RecordCount <> 0 Then
               textST20.Text = "" & MyRs.Fields(0).Value
       Else
           If m_EditMode <> 0 Then
               MsgBox "Â¾ºÙ¥N¸¹¿é¤J¿ù»~!!!", vbExclamation + vbOKOnly
               ' 2008/12/17 ADD BY SINDY
               Call textST20_GotFocus
               ' 2008/12/17 END
               Cancel = True
               Exit Sub
           End If
       End If
   End If
End Sub

Private Sub textST21_GotFocus()
   If m_EditMode <> 0 Then
       InverseTextBox textST21
   End If
End Sub

Private Sub textST21_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textST21_Validate(Cancel As Boolean)
Dim MyRs As New ADODB.Recordset
Dim MyArr As Variant

   If textST21.Text <> "" Then
       MyArr = Split(textST21, " ")
       Set MyRs = New ADODB.Recordset
       If MyRs.State = 1 Then MyRs.Close
       strSql = "select ac02||' '||ac03 from allcode where ac02='" & MyArr(0) & "' and ac01='02' order by ac02"
       MyRs.CursorLocation = adUseClient
       MyRs.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
       If MyRs.RecordCount <> 0 Then
               textST21.Text = "" & MyRs.Fields(0).Value
       Else
           If m_EditMode <> 0 Then
               MsgBox "Â¾¦ì¥N¸¹¿é¤J¿ù»~!!!", vbExclamation + vbOKOnly
               ' 2008/12/17 ADD BY SINDY
               Call textST21_GotFocus
               ' 2008/12/17 END
               Cancel = True
               Exit Sub
           End If
       End If
   End If
End Sub

Private Sub textST22_GotFocus()
   If m_EditMode <> 0 Then
       InverseTextBox textST22
   End If
End Sub

Private Sub textST22_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textST22_Validate(Cancel As Boolean)
Dim MyArr As Variant
Dim MyArr2 As Variant
Dim Myi As Integer
   
   If textST22.Text <> "" Then
       MyArr = Split(textST22, " ")
       For Myi = 0 To textST22.ListCount - 1
           MyArr2 = Split(textST22.List(Myi), " ")
           If MyArr(0) = MyArr2(0) Then
               textST22.Text = textST22.List(Myi)
               Exit Sub
           End If
       Next Myi
       If m_EditMode <> 0 Then
           MsgBox "©Ê§O¥N¸¹¿é¤J¿ù»~!!!", vbExclamation + vbOKOnly
           ' 2008/12/17 ADD BY SINDY
           Call textST22_GotFocus
           ' 2008/12/17 END
           Cancel = True
           Exit Sub
       End If
   End If
End Sub

Private Sub textST23_GotFocus()
   If m_EditMode <> 0 Then
       InverseTextBox textST23
   End If
End Sub

Private Sub textST23_KeyPress(KeyAscii As Integer)
Dim MyArr As Variant
Dim MyArr2 As Variant
Dim Myi As Integer
End Sub

Private Sub textST23_Validate(Cancel As Boolean)
   If m_EditMode <> 0 And textST23 <> "" Then
       If textST23 = "" Then Exit Sub
      ' 2008/12/17 MODIFY BY SINDY
      'Cancel = Not ChkDate(textST23)
      If ChkDate(textST23) = False Then
         Call textST23_GotFocus
         Cancel = True
         Exit Sub
      End If
      ' 2008/12/17 END
   End If
End Sub

Private Sub textST24_GotFocus()
   If m_EditMode <> 0 Then
       InverseTextBox textST24
   End If
End Sub

Private Sub textST24_Validate(Cancel As Boolean)
Dim MyArr As Variant
Dim MyArr2 As Variant
Dim Myi As Integer

   If textST24.Text <> "" Then
       MyArr = Split(textST24, " ")
       For Myi = 0 To textST24.ListCount - 1
           MyArr2 = Split(textST24.List(Myi), " ")
           If MyArr(0) = MyArr2(0) Then
               textST24.Text = textST24.List(Myi)
               Exit Sub
           End If
       Next Myi
       If m_EditMode <> 0 Then
           MsgBox "°êÄy¥N¸¹¿é¤J¿ù»~!!!", vbExclamation + vbOKOnly
           ' 2008/12/17 ADD BY SINDY
           Call textST24_GotFocus
           ' 2008/12/17 END
           Cancel = True
           Exit Sub
       End If
   End If
End Sub

Private Sub textST25_GotFocus()
   If m_EditMode <> 0 Then
       InverseTextBox textST25
   End If
End Sub

Private Sub textST25_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textST25_Validate(Cancel As Boolean)
Dim MyArr As Variant
Dim MyArr2 As Variant
Dim Myi As Integer

   If textST25.Text <> "" Then
       MyArr = Split(textST25, " ")
       For Myi = 0 To textST25.ListCount - 1
           MyArr2 = Split(textST25.List(Myi), " ")
           If MyArr(0) = MyArr2(0) Then
               textST25.Text = textST25.List(Myi)
               Exit Sub
           End If
       Next Myi
       If m_EditMode <> 0 Then
           MsgBox "¦å«¬¥N¸¹¿é¤J¿ù»~!!!", vbExclamation + vbOKOnly
           ' 2008/12/17 ADD BY SINDY
           Call textST25_GotFocus
           ' 2008/12/17 END
           Cancel = True
           Exit Sub
       End If
   End If
End Sub

Private Sub textST26_GotFocus()
   If m_EditMode <> 0 Then
       InverseTextBox textST26
   End If
End Sub

Private Sub textST26_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textST26_Validate(Cancel As Boolean)
   'If m_EditMode <> 0 And textST26 <> "" Then
   '    If (Mid(textST26, 2, 1) = "1" And Mid(textST22, 1, 1) <> "M") Or (Mid(textST26, 2, 1) = "2" And Mid(textST22, 1, 1) <> "F") Then
   '        MsgBox "¨­¤ÀÃÒ¸¹©Ê§OÀË¬d¿ù»~!!!", vbExclamation + vbOKOnly
   '        Cancel = True
   '        Exit Sub
   '    End If
   '    If GetTextLength(textST26.Text) <> 10 Then
   '       Call textST26_GotFocus
   '       If MsgBox("¨­¥÷ÃÒ¥²¶·¬O10½X ! ½Ð½T©w ?", vbYesNo + vbCritical) = vbNo Then
   '          Cancel = True
   '          Exit Sub
   '       End If
   '    End If
   'End If
   ' 2008/12/17 ADD BY SINDY
   If m_EditMode <> 1 And m_EditMode <> 2 Then Exit Sub
   If textST26.Text = "" Then Exit Sub
   
Dim strTmp As String
   If GetTextLength(textST26.Text) <> 10 Then
      Call textST26_GotFocus
      strTmp = "¨­¥÷ÃÒ¥²¶·¬O10½X ! ½Ð½T©w ?"
      If MsgBox(strTmp, vbYesNo + vbCritical) = vbNo Then
         Cancel = True
         Exit Sub
      End If
   End If
   If CheckID(0, textST26.Text) = False Then
      Call textST26_GotFocus
      strTmp = "¨­¤ÀÃÒ¦r¸¹¿ù»~¡A¬O§_½T©w ?"
      If MsgBox(strTmp, vbYesNo + vbCritical) = vbNo Then
         Cancel = True
      End If
   End If
   ' 2008/12/17 END
End Sub

Private Sub textST27_GotFocus()
   If m_EditMode <> 0 Then
       InverseTextBox textST27
       OpenIme
   End If
End Sub

Private Sub textST27_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textST27_Validate(Cancel As Boolean)
Dim MyRs As New ADODB.Recordset
Dim MyArr As Variant

   If textST27.Text <> "" Then
       MyArr = Split(textST27, " ")
       Set MyRs = New ADODB.Recordset
       If MyRs.State = 1 Then MyRs.Close
       strSql = "select ac02||' '||ac03 from allcode where ac02='" & MyArr(0) & "' and ac01='06' order by ac02"
       MyRs.CursorLocation = adUseClient
       MyRs.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
       If MyRs.RecordCount <> 0 Then
               textST27.Text = "" & MyRs.Fields(0).Value
       Else
           If m_EditMode <> 0 Then
               MsgBox "¥X¥Í¦a¥N¸¹¿é¤J¿ù»~!!!", vbExclamation + vbOKOnly
               ' 2008/12/17 ADD BY SINDY
               Call textST27_GotFocus
               ' 2008/12/17 END
               Cancel = True
               Exit Sub
           End If
       End If
   End If
   CloseIme
End Sub

Private Sub textST28_GotFocus()
   If m_EditMode <> 0 Then
       InverseTextBox textST28
   End If
End Sub

Private Sub textST28_KeyPress(KeyAscii As Integer)
   KeyAscii = Pub_NumAscii(KeyAscii)
End Sub

Private Sub textST28_Validate(Cancel As Boolean)
   If m_EditMode <> 0 And textST28 <> "" Then
      ' 2008/12/17 MODIFY BY SINDY
      'Cancel = Not ChkDate(textST28)
      If ChkDate(textST28) = False Then
         Call textST28_GotFocus
         Cancel = True
         Exit Sub
      End If
      ' 2008/12/17 END
      
      'Add By Sindy 2021/7/20 ¤£¥i¤p©ó¤J©Ò¤é´Á
      If Val(textST28) < Val(textST13) Then
         MsgBox "¸Õ¥Î°_©l¤é´Á¤£¥i¤p©ó¤J©Ò¤é´Á¡I", vbExclamation + vbOKOnly
         Call textST28_GotFocus
         Cancel = True
         Exit Sub
      End If
      '2021/7/20 END
   End If
End Sub

Private Sub textST29_GotFocus()
   If m_EditMode <> 0 Then
       InverseTextBox textST29
   End If
End Sub

Private Sub textST29_KeyPress(KeyAscii As Integer)
   KeyAscii = Pub_NumAscii(KeyAscii)
End Sub

Private Sub textST29_Validate(Cancel As Boolean)
   If m_EditMode <> 0 And textST29 <> "" Then
      ' 2008/12/17 MODIFY BY SINDY
      'Cancel = Not ChkDate(textST29)
      If ChkDate(textST29) = False Then
         Call textST29_GotFocus
         Cancel = True
         Exit Sub
      End If
      ' 2008/12/17 END
      
      'Add By Sindy 2021/7/20 ¤£¥i¤p©ó¤J©Ò¤é´Á
      If Val(textST29) < Val(textST13) Then
         MsgBox "¸Õ¥ÎºI¤î¤é´Á¤£¥i¤p©ó¤J©Ò¤é´Á¡I", vbExclamation + vbOKOnly
         Call textST29_GotFocus
         Cancel = True
         Exit Sub
      End If
      '2021/7/20 END
      
      ' 2008/12/17 ADD BY SINDY
      If RunNick2(textST28, textST29) Then
         Call textST29_GotFocus
         Cancel = True
         Exit Sub
      End If
      ' 2008/12/17 END
   End If
End Sub

Private Sub textST30_GotFocus()
   If m_EditMode <> 0 Then
       InverseTextBox textST30
   End If
End Sub

'Add By Sindy 2010/11/25
Private Sub textST30_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textST30_Validate(Cancel As Boolean)
Dim strTemp As String

   If m_EditMode <> 0 And textST30 <> "" Then
      ' 2008/12/17 ADD BY SINDY
      LabelST30.Caption = ""
      If ClsPDGetStaff(textST30, strTemp) Then
         LabelST30.Caption = strTemp
      Else
         'Call textST30_GotFocus
         Cancel = True
         Exit Sub
      End If
      ' 2008/12/17 END
   End If
   If Cancel Then TextInverse textST30
End Sub

Private Sub textST31_GotFocus()
   If m_EditMode <> 0 Then
       InverseTextBox textST31
   End If
End Sub

Private Sub textST31_KeyPress(KeyAscii As Integer)
   KeyAscii = Pub_NumAscii(KeyAscii)
End Sub

Private Sub textST31_Validate(Cancel As Boolean)
   If m_EditMode <> 0 And textST31 <> "" Then
      ' 2008/12/17 MODIFY BY SINDY
      'Cancel = Not ChkDate(textST31)
      If ChkDate(textST31) = False Then
         Call textST31_GotFocus
         Cancel = True
         Exit Sub
      End If
      ' 2008/12/17 END
   End If
End Sub

Private Sub textST32_GotFocus()
   If m_EditMode <> 0 Then
       InverseTextBox textST32
   End If
End Sub

Private Sub textST32_KeyPress(KeyAscii As Integer)
   KeyAscii = Pub_NumAscii(KeyAscii)
End Sub

Private Sub textST32_Validate(Cancel As Boolean)
   If m_EditMode <> 0 And textST32 <> "" Then
      ' 2008/12/17 MODIFY BY SINDY
      'Cancel = Not ChkDate(textST32)
      If ChkDate(textST32) = False Then
         Call textST32_GotFocus
         Cancel = True
         Exit Sub
      End If
      ' 2008/12/17 END
   End If
End Sub

Private Sub textST33_GotFocus()
   If m_EditMode <> 0 Then
       InverseTextBox textST33
       ' 2008/12/17 ADD BY SINDY
       OpenIme
       ' 2008/12/17 END
   End If
End Sub

' 2008/12/17 ADD BY SINDY
Private Sub textST33_KeyPress(KeyAscii As Integer)
   KeyAscii = ChangeZIP(KeyAscii)
End Sub
' 2008/12/17 END

Private Sub textST33_Validate(Cancel As Boolean)
   If m_EditMode <> 0 And textST33 <> "" Then
       If CheckLengthIsOK(textST33, textST33.MaxLength) = False Then
           ' 2008/12/17 ADD BY SINDY
           Call textST33_GotFocus
           ' 2008/12/17 END
           Cancel = True
           Exit Sub
       End If
   End If
   CloseIme
End Sub

Private Sub textST34_GotFocus()
   If m_EditMode <> 0 Then
       InverseTextBox textST34
       OpenIme
   End If
End Sub

' 2008/12/17 ADD BY SINDY
Private Sub textST34_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = ChangeZIP(KeyAscii)
End Sub
' 2008/12/17 END

Private Sub textST34_Validate(Cancel As Boolean)
   If m_EditMode <> 0 And textST34 <> "" Then
       If CheckLengthIsOK(textST34, textST34.MaxLength) = False Then
           ' 2008/12/17 ADD BY SINDY
           Call textST34_GotFocus
           ' 2008/12/17 END
           Cancel = True
           Exit Sub
       End If
   End If
   CloseIme
End Sub

Private Sub textST35_GotFocus()
   If m_EditMode <> 0 Then
       InverseTextBox textST35
   End If
End Sub

Private Sub textST35_Validate(Cancel As Boolean)
   If m_EditMode <> 0 And textST35 <> "" Then
       If CheckLengthIsOK(textST35, textST35.MaxLength) = False Then
           ' 2008/12/17 ADD BY SINDY
           Call textST35_GotFocus
           ' 2008/12/17 END
           Cancel = True
           Exit Sub
       End If
   End If
   CloseIme
End Sub

Private Sub textST36_GotFocus()
   If m_EditMode <> 0 Then
       InverseTextBox textST36
       ' 2008/12/17 ADD BY SINDY
       OpenIme
       ' 2008/12/17 END
   End If
End Sub

' 2008/12/17 ADD BY SINDY
Private Sub textST36_KeyPress(KeyAscii As Integer)
   KeyAscii = ChangeZIP(KeyAscii)
End Sub
' 2008/12/17 END

Private Sub textST36_Validate(Cancel As Boolean)
   If m_EditMode <> 0 And textST36 <> "" Then
       If CheckLengthIsOK(textST36, textST36.MaxLength) = False Then
           ' 2008/12/17 ADD BY SINDY
           Call textST36_GotFocus
           ' 2008/12/17 END
           Cancel = True
           Exit Sub
       End If
   End If
   CloseIme
End Sub

Private Sub textST37_GotFocus()
   If m_EditMode <> 0 Then
       InverseTextBox textST37
   End If
End Sub

Private Sub textST37_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textST37_Validate(Cancel As Boolean)
Dim MyRs As New ADODB.Recordset
Dim MyArr As Variant

   If textST37.Text <> "" Then
       MyArr = Split(textST37, " ")
       Set MyRs = New ADODB.Recordset
       If MyRs.State = 1 Then MyRs.Close
       strSql = "select ac02||' '||ac03 from allcode where ac02='" & MyArr(0) & "' and ac01='03' order by ac02"
       MyRs.CursorLocation = adUseClient
       MyRs.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
       If MyRs.RecordCount <> 0 Then
           textST37.Text = "" & MyRs.Fields(0).Value
       Else
           If m_EditMode <> 0 Then
               MsgBox "¾Ç¾ú¥N¸¹¿é¤J¿ù»~!!!", vbExclamation + vbOKOnly
               ' 2008/12/17 ADD BY SINDY
               Call textST37_GotFocus
               ' 2008/12/17 END
               Cancel = True
               Exit Sub
           End If
       End If
   End If
End Sub

Private Sub textST38_GotFocus()
   If m_EditMode <> 0 Then
       InverseTextBox textST38
       OpenIme
   End If
End Sub

Private Sub textST38_Validate(Cancel As Boolean)
   If m_EditMode <> 0 And textST38 <> "" Then
       If CheckLengthIsOK(textST38, textST38.MaxLength) = False Then
           ' 2008/12/17 ADD BY SINDY
           Call textST38_GotFocus
           ' 2008/12/17 END
           Cancel = True
           Exit Sub
       End If
   End If
   CloseIme
End Sub

Private Sub textST39_GotFocus()
   If m_EditMode <> 0 Then
       InverseTextBox textST39
       OpenIme
   End If
End Sub

Private Sub textST39_Validate(Cancel As Boolean)
   If m_EditMode <> 0 And textST39 <> "" Then
       If CheckLengthIsOK(textST39, textST39.MaxLength) = False Then
           ' 2008/12/17 ADD BY SINDY
           Call textST39_GotFocus
           ' 2008/12/17 END
           Cancel = True
           Exit Sub
       End If
   End If
   CloseIme
End Sub

Private Sub textST40_GotFocus()
   If m_EditMode <> 0 Then
       InverseTextBox textST40
   End If
End Sub

Private Sub textST40_KeyPress(KeyAscii As Integer)
   KeyAscii = Pub_NumAscii(KeyAscii, True)
End Sub

Private Sub textST40_Validate(Cancel As Boolean)
   If m_EditMode <> 0 And textST40 <> "" Then
       If CheckLengthIsOK(textST40, textST40.MaxLength) = False Then
           ' 2008/12/17 ADD BY SINDY
           Call textST40_GotFocus
           ' 2008/12/17 END
           Cancel = True
           Exit Sub
       End If
       If Val(textST40) > 30 Then
           MsgBox "¦~°²¤Ó¦h¡A½Ð½T©w¬O§_¥¿½T!!", vbInformation, "¾Þ§@¿ù»~!!"
           ' 2008/12/17 ADD BY SINDY
           Call textST40_GotFocus
           ' 2008/12/17 END
           Cancel = True
           Exit Sub
       End If
   End If
   CloseIme
End Sub

Private Sub textST41_GotFocus()
   If m_EditMode <> 0 Then
       InverseTextBox textST41
   End If
End Sub

Private Sub textST41_KeyPress(KeyAscii As Integer)
   KeyAscii = Pub_NumAscii(KeyAscii)
End Sub

Private Sub textST41_Validate(Cancel As Boolean)
   If m_EditMode <> 0 And textST41 <> "" Then
      ' 2008/12/17 MODIFY BY SINDY
      'Cancel = Not ChkDate(textST41)
      If ChkDate(textST41) = False Then
         Call textST41_GotFocus
         Cancel = True
         Exit Sub
      End If
      ' 2008/12/17 END
   End If
End Sub

Private Sub textST42_GotFocus()
   If m_EditMode <> 0 Then
       InverseTextBox textST42
   End If
End Sub

Private Sub textST42_Validate(Cancel As Boolean)
   If m_EditMode <> 0 And textST42 <> "" Then
       If CheckLengthIsOK(textST42, textST42.MaxLength) = False Then
           Call textST42_GotFocus
           Cancel = True
           Exit Sub
       End If
       ' 2008/12/17 ADD BY SINDY
       If PUB_CheckMail(textST42.Text) = False Then
          Call textST42_GotFocus
          Cancel = True
          Exit Sub
       End If
       ' 2008/12/17 END
   End If
   CloseIme
End Sub

Private Sub textST49_GotFocus()
   If m_EditMode <> 0 Then
       InverseTextBox textST49
       OpenIme
   End If
End Sub

Private Sub textST49_Validate(Cancel As Boolean)
   If m_EditMode <> 0 And textST49 <> "" Then
       If CheckLengthIsOK(textST49, textST49.MaxLength) = False Then
           ' 2008/12/17 ADD BY SINDY
           Call textST49_GotFocus
           ' 2008/12/17 END
           Cancel = True
           Exit Sub
       End If
   End If
   CloseIme
End Sub

Private Sub textST12_GotFocus()
   If m_EditMode <> 0 Then
       InverseTextBox textST12
   End If
End Sub

Private Sub textST12_Validate(Cancel As Boolean)
   If m_EditMode <> 0 And textST12 <> "" Then
       If CheckLengthIsOK(textST12, textST12.MaxLength) = False Then
           ' 2008/12/17 ADD BY SINDY
           Call textST12_GotFocus
           ' 2008/12/17 END
           Cancel = True
           Exit Sub
       End If
   End If
End Sub
Private Sub SetGrd()
Dim arrGridHeadText, arrGridHeadWidth
Dim iRow As Integer
   
   'Modify by Morgan 2009/6/9
   'arrGridHeadText = Array("ºÙ¿×", "©m¦W", "©Ê§O", "¥X¥Í¤é´Á", "¨­¤ÀÃÒ¦r¸¹", "°·«O²²ÄÝ", "ª\", "¹q¸Ü", "¶l»¼°Ï¸¹", "¦a§}", "°·«O«O¶O")
   'arrGridHeadWidth = Array(600, 1200, 600, 800, 1000, 800, 400, 1000, 600, 2000, 2000)
   arrGridHeadText = Array("ºÙ¿×", "©m¦W", "ª¬ºA", "©Ê§O", "¥X¥Í¤é´Á", "¨­¤ÀÃÒ¦r¸¹", "°·«O²²ÄÝ", "ª\", "¹q¸Ü", "¶l»¼°Ï¸¹", "¦a§}", "§R°£¤é´Á", "§Ç¸¹", "°·«O¸É§UÃþ§O")
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
   arrGridHeadText = Array("¦~«×", "µûµ¥")
   arrGridHeadWidth = Array(1000, 1000)
   GRD2.Visible = False
   GRD2.Cols = UBound(arrGridHeadText) + 1
   For iRow = 0 To GRD2.Cols - 1
      GRD2.row = 0
      GRD2.col = iRow
      GRD2.Text = arrGridHeadText(iRow)
      GRD2.ColWidth(iRow) = arrGridHeadWidth(iRow)
      GRD2.CellAlignment = flexAlignCenterCenter
   Next
   GRD2.Visible = True
   'Add By Sindy 2015/8/13
   arrGridHeadText = Array("°·ÀË¤é´Á", "¸É§U¶O¥Î", "Ãº¥æ¤é´Á", "³Æ¡@¡@µù")
   arrGridHeadWidth = Array(900, 900, 900, 5500)
   grd3.Visible = False
   grd3.Cols = UBound(arrGridHeadText) + 1
   For iRow = 0 To grd3.Cols - 1
      grd3.row = 0
      grd3.col = iRow
      grd3.Text = arrGridHeadText(iRow)
      grd3.ColWidth(iRow) = arrGridHeadWidth(iRow)
      grd3.CellAlignment = flexAlignCenterCenter
   Next
   grd3.Visible = True
   '2015/8/13 END
End Sub
'Add by Morgan 2009/6/11
Private Sub EnableSR(ByVal bEnable As Boolean, Optional ByVal SR02 As String)
   textSR03.Locked = Not bEnable
   textSR04.Locked = Not bEnable
   textSR05.Locked = Not bEnable
   textSR06.Locked = Not bEnable
   textSR07.Locked = Not bEnable
   textSR09.Locked = Not bEnable
   textSR10.Locked = Not bEnable
   textSR11.Locked = Not bEnable
   chkSR13.Enabled = bEnable
   If SR02 = "" Then
      textSR12.Locked = True
      chkSR08.Enabled = bEnable
      cboHL05.Enabled = bEnable
   Else
      textSR12.Locked = Not bEnable
      chkSR08.Enabled = False
      cboHL05.Enabled = False
   End If
End Sub

Private Sub txtSD16_GotFocus()
   TextInverse txtSD16
End Sub

Private Sub txtSD16_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 8 And KeyAscii <> Asc("Y") Then
      KeyAscii = 0
      Beep
   End If
End Sub
Private Sub txtSD17_GotFocus()
   TextInverse txtSD17
End Sub

Private Sub txtSD17_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 8 And Not IsNumeric(Chr(KeyAscii)) Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub txtSD17_Validate(Cancel As Boolean)
   If Val(txtSD17) > 6 Then
      Cancel = True
      MsgBox "³Ò°h¦Û´£¶O²v¤£¥i¤j©ó 6¡I"
   End If
End Sub
'Add by Morgan 2009/6/24
'¿ï¨ú¿ï³æ
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
