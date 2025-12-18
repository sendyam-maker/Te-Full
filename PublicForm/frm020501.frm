VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm020501 
   BorderStyle     =   1  '單線固定
   Caption         =   "商標基本資料維護"
   ClientHeight    =   6060
   ClientLeft      =   168
   ClientTop       =   996
   ClientWidth     =   9216
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6060
   ScaleWidth      =   9216
   Begin VB.CommandButton cmdIns 
      Caption         =   "各項指示"
      Height          =   285
      Left            =   1800
      TabIndex        =   188
      Top             =   570
      Visible         =   0   'False
      Width           =   1215
   End
   Begin TabDlg.SSTab tabCtrl 
      CausesValidation=   0   'False
      Height          =   5124
      Left            =   0
      TabIndex        =   90
      Top             =   900
      Width           =   9180
      _ExtentX        =   16193
      _ExtentY        =   9038
      _Version        =   393216
      Style           =   1
      Tabs            =   8
      TabsPerRow      =   8
      TabHeight       =   420
      TabCaption(0)   =   "基本資料1"
      TabPicture(0)   =   "frm020501.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1(1)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1(3)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label1(4)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label1(5)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label1(6)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label1(8)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label1(10)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label1(12)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label1(14)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label1(16)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label1(18)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label1(20)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Line1"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Label1(2)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Label1(7)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Label1(9)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Label1(11)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Label1(13)"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Label1(15)"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Label1(17)"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "Label1(19)"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "Label1(21)"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "Label1(22)"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "Label1(23)"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "Label1(111)"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "textCUID_1"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "textCUID_2"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "textTM01"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "textTM02_1"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "textTM02_2"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "textTM03"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "textTM28"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "textTM10_1"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "textTM04"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "textTM10_2"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "textTM09"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "textTM11"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "textTM14"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).Control(39)=   "textTM20"
      Tab(0).Control(39).Enabled=   0   'False
      Tab(0).Control(40)=   "textTM21"
      Tab(0).Control(40).Enabled=   0   'False
      Tab(0).Control(41)=   "textTM22"
      Tab(0).Control(41).Enabled=   0   'False
      Tab(0).Control(42)=   "textTM16"
      Tab(0).Control(42).Enabled=   0   'False
      Tab(0).Control(43)=   "textTM18"
      Tab(0).Control(43).Enabled=   0   'False
      Tab(0).Control(44)=   "textTM53"
      Tab(0).Control(44).Enabled=   0   'False
      Tab(0).Control(45)=   "textTM30"
      Tab(0).Control(45).Enabled=   0   'False
      Tab(0).Control(46)=   "textTM08_1"
      Tab(0).Control(46).Enabled=   0   'False
      Tab(0).Control(47)=   "textTM08_2"
      Tab(0).Control(47).Enabled=   0   'False
      Tab(0).Control(48)=   "textTM12"
      Tab(0).Control(48).Enabled=   0   'False
      Tab(0).Control(49)=   "textTM15"
      Tab(0).Control(49).Enabled=   0   'False
      Tab(0).Control(50)=   "textTM13"
      Tab(0).Control(50).Enabled=   0   'False
      Tab(0).Control(51)=   "textTM27"
      Tab(0).Control(51).Enabled=   0   'False
      Tab(0).Control(52)=   "textTM17"
      Tab(0).Control(52).Enabled=   0   'False
      Tab(0).Control(53)=   "textTM19"
      Tab(0).Control(53).Enabled=   0   'False
      Tab(0).Control(54)=   "textTM29"
      Tab(0).Control(54).Enabled=   0   'False
      Tab(0).Control(55)=   "textTM31_1"
      Tab(0).Control(55).Enabled=   0   'False
      Tab(0).Control(56)=   "textTM31_2"
      Tab(0).Control(56).Enabled=   0   'False
      Tab(0).Control(57)=   "textTM72_1"
      Tab(0).Control(57).Enabled=   0   'False
      Tab(0).Control(58)=   "textTM72_2"
      Tab(0).Control(58).Enabled=   0   'False
      Tab(0).Control(59)=   "textTM118"
      Tab(0).Control(59).Enabled=   0   'False
      Tab(0).Control(60)=   "textTM77"
      Tab(0).Control(60).Enabled=   0   'False
      Tab(0).Control(61)=   "textTM05"
      Tab(0).Control(61).Enabled=   0   'False
      Tab(0).Control(62)=   "Label1(87)"
      Tab(0).Control(62).Enabled=   0   'False
      Tab(0).Control(63)=   "textTM136"
      Tab(0).Control(63).Enabled=   0   'False
      Tab(0).Control(64)=   "cboTM08"
      Tab(0).Control(64).Enabled=   0   'False
      Tab(0).Control(65)=   "cboTM72"
      Tab(0).Control(65).Enabled=   0   'False
      Tab(0).ControlCount=   66
      TabCaption(1)   =   "基本資料2"
      TabPicture(1)   =   "frm020501.frx":005E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label1(24)"
      Tab(1).Control(1)=   "Label1(25)"
      Tab(1).Control(2)=   "Label1(27)"
      Tab(1).Control(3)=   "Label1(29)"
      Tab(1).Control(4)=   "Label1(31)"
      Tab(1).Control(5)=   "Label1(26)"
      Tab(1).Control(6)=   "Label1(28)"
      Tab(1).Control(7)=   "Label1(30)"
      Tab(1).Control(8)=   "Label1(35)"
      Tab(1).Control(9)=   "Label1(36)"
      Tab(1).Control(10)=   "Label1(37)"
      Tab(1).Control(11)=   "Label1(38)"
      Tab(1).Control(12)=   "Label1(160)"
      Tab(1).Control(13)=   "Label1(115)"
      Tab(1).Control(14)=   "textTM32"
      Tab(1).Control(15)=   "textTM34"
      Tab(1).Control(16)=   "textTM36"
      Tab(1).Control(17)=   "textTM54_1"
      Tab(1).Control(18)=   "textSP32"
      Tab(1).Control(19)=   "textTM35"
      Tab(1).Control(20)=   "textTM37"
      Tab(1).Control(21)=   "textTM55"
      Tab(1).Control(22)=   "textTM54_2"
      Tab(1).Control(23)=   "textTM67"
      Tab(1).Control(24)=   "textTM58"
      Tab(1).Control(25)=   "textFA29"
      Tab(1).Control(26)=   "textFA39"
      Tab(1).Control(27)=   "textTM122"
      Tab(1).Control(28)=   "cboContact"
      Tab(1).Control(29)=   "textCU72"
      Tab(1).Control(30)=   "Label1(68)"
      Tab(1).Control(31)=   "Label1(67)"
      Tab(1).Control(32)=   "textCU79"
      Tab(1).Control(33)=   "Label1(88)"
      Tab(1).Control(34)=   "textTM140"
      Tab(1).Control(35)=   "Label1(89)"
      Tab(1).Control(36)=   "textTM141"
      Tab(1).ControlCount=   37
      TabCaption(2)   =   "申請人1-3"
      TabPicture(2)   =   "frm020501.frx":007A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "textTM91"
      Tab(2).Control(1)=   "textTM87"
      Tab(2).Control(2)=   "textTM83"
      Tab(2).Control(3)=   "textTM90"
      Tab(2).Control(4)=   "textTM86"
      Tab(2).Control(5)=   "textTM82"
      Tab(2).Control(6)=   "textTM26"
      Tab(2).Control(7)=   "textTM25"
      Tab(2).Control(8)=   "textTM24"
      Tab(2).Control(9)=   "textTM79_2"
      Tab(2).Control(10)=   "textTM78_2"
      Tab(2).Control(11)=   "textTM23_2"
      Tab(2).Control(12)=   "textTM79_1"
      Tab(2).Control(13)=   "textTM78_1"
      Tab(2).Control(14)=   "textTM23_1"
      Tab(2).Control(15)=   "Label1(56)"
      Tab(2).Control(16)=   "Label1(57)"
      Tab(2).Control(17)=   "Label1(58)"
      Tab(2).Control(18)=   "Label1(53)"
      Tab(2).Control(19)=   "Label1(54)"
      Tab(2).Control(20)=   "Label1(55)"
      Tab(2).Control(21)=   "Label1(47)"
      Tab(2).Control(22)=   "Label1(46)"
      Tab(2).Control(23)=   "Label1(52)"
      Tab(2).Control(24)=   "Label1(51)"
      Tab(2).Control(25)=   "Label1(50)"
      Tab(2).Control(26)=   "Label1(45)"
      Tab(2).ControlCount=   27
      TabCaption(3)   =   "申請人4-5 / 延展"
      TabPicture(3)   =   "frm020501.frx":0096
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "textTM93"
      Tab(3).Control(1)=   "textTM89"
      Tab(3).Control(2)=   "textTM85"
      Tab(3).Control(3)=   "textTM92"
      Tab(3).Control(4)=   "textTM88"
      Tab(3).Control(5)=   "textTM84"
      Tab(3).Control(6)=   "textTM81_2"
      Tab(3).Control(7)=   "textTM80_2"
      Tab(3).Control(8)=   "textTM81_1"
      Tab(3).Control(9)=   "textTM80_1"
      Tab(3).Control(10)=   "Label1(62)"
      Tab(3).Control(11)=   "Label1(63)"
      Tab(3).Control(12)=   "Label1(64)"
      Tab(3).Control(13)=   "Label1(59)"
      Tab(3).Control(14)=   "Label1(60)"
      Tab(3).Control(15)=   "Label1(61)"
      Tab(3).Control(16)=   "Label1(49)"
      Tab(3).Control(17)=   "Label1(48)"
      Tab(3).Control(18)=   "textTM66_2"
      Tab(3).Control(19)=   "textTM33_2"
      Tab(3).Control(20)=   "textTM65"
      Tab(3).Control(21)=   "textTM70_2"
      Tab(3).Control(22)=   "textTM70_1"
      Tab(3).Control(23)=   "textTM33_1"
      Tab(3).Control(24)=   "Label1(41)"
      Tab(3).Control(25)=   "Label1(39)"
      Tab(3).Control(26)=   "Label1(40)"
      Tab(3).Control(27)=   "Label1(42)"
      Tab(3).Control(28)=   "Label1(43)"
      Tab(3).Control(29)=   "Label1(44)"
      Tab(3).Control(30)=   "textTM66_1"
      Tab(3).Control(31)=   "textTM71_1"
      Tab(3).Control(32)=   "textTM68"
      Tab(3).Control(33)=   "Frame1"
      Tab(3).ControlCount=   34
      TabCaption(4)   =   "代理人/聯絡人"
      TabPicture(4)   =   "frm020501.frx":00B2
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "textTM127"
      Tab(4).Control(1)=   "textTM38"
      Tab(4).Control(2)=   "textTM39"
      Tab(4).Control(3)=   "textTM40"
      Tab(4).Control(4)=   "textTM41"
      Tab(4).Control(5)=   "textTM42"
      Tab(4).Control(6)=   "textTM43"
      Tab(4).Control(7)=   "textTM76"
      Tab(4).Control(8)=   "textTM44_2"
      Tab(4).Control(9)=   "textTM44_1"
      Tab(4).Control(10)=   "textTM45"
      Tab(4).Control(11)=   "Label1(169)"
      Tab(4).Control(12)=   "Label1(69)"
      Tab(4).Control(13)=   "Label1(70)"
      Tab(4).Control(14)=   "Label1(71)"
      Tab(4).Control(15)=   "Label1(72)"
      Tab(4).Control(16)=   "Label1(73)"
      Tab(4).Control(17)=   "Label1(74)"
      Tab(4).Control(18)=   "Label1(75)"
      Tab(4).Control(19)=   "Label1(65)"
      Tab(4).Control(20)=   "Label1(66)"
      Tab(4).Control(21)=   "textTM69_2"
      Tab(4).Control(22)=   "textTM69_1"
      Tab(4).Control(23)=   "textTM56_2"
      Tab(4).Control(24)=   "Label1(33)"
      Tab(4).Control(25)=   "Label1(34)"
      Tab(4).Control(26)=   "Label1(32)"
      Tab(4).Control(27)=   "textTM46"
      Tab(4).Control(28)=   "Label5"
      Tab(4).Control(29)=   "Label11(1)"
      Tab(4).Control(30)=   "textTM56_1"
      Tab(4).Control(31)=   "Combo4"
      Tab(4).Control(32)=   "Combo5"
      Tab(4).ControlCount=   33
      TabCaption(5)   =   "代表人"
      TabPicture(5)   =   "frm020501.frx":00CE
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "textTM117"
      Tab(5).Control(1)=   "textTM116"
      Tab(5).Control(2)=   "textTM115"
      Tab(5).Control(3)=   "textTM114"
      Tab(5).Control(4)=   "textTM113"
      Tab(5).Control(5)=   "textTM112"
      Tab(5).Control(6)=   "textTM111"
      Tab(5).Control(7)=   "textTM110"
      Tab(5).Control(8)=   "textTM109"
      Tab(5).Control(9)=   "textTM108"
      Tab(5).Control(10)=   "textTM107"
      Tab(5).Control(11)=   "textTM106"
      Tab(5).Control(12)=   "textTM105"
      Tab(5).Control(13)=   "textTM104"
      Tab(5).Control(14)=   "textTM103"
      Tab(5).Control(15)=   "textTM102"
      Tab(5).Control(16)=   "textTM101"
      Tab(5).Control(17)=   "textTM100"
      Tab(5).Control(18)=   "textTM99"
      Tab(5).Control(19)=   "textTM98"
      Tab(5).Control(20)=   "textTM97"
      Tab(5).Control(21)=   "textTM96"
      Tab(5).Control(22)=   "textTM95"
      Tab(5).Control(23)=   "textTM94"
      Tab(5).Control(24)=   "textTM47"
      Tab(5).Control(25)=   "textTM48"
      Tab(5).Control(26)=   "textTM49"
      Tab(5).Control(27)=   "textTM50"
      Tab(5).Control(28)=   "textTM51"
      Tab(5).Control(29)=   "textTM52"
      Tab(5).Control(30)=   "Label1(83)"
      Tab(5).Control(31)=   "Label1(84)"
      Tab(5).Control(32)=   "Label1(77)"
      Tab(5).Control(33)=   "Label1(78)"
      Tab(5).Control(34)=   "Label1(79)"
      Tab(5).Control(35)=   "Label1(80)"
      Tab(5).Control(36)=   "Label1(81)"
      Tab(5).Control(37)=   "Label1(82)"
      Tab(5).Control(38)=   "Label1(85)"
      Tab(5).Control(39)=   "Label1(86)"
      Tab(5).ControlCount=   40
      TabCaption(6)   =   "銷卷資料"
      TabPicture(6)   =   "frm020501.frx":00EA
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "textTM57"
      Tab(6).Control(1)=   "textTM73"
      Tab(6).Control(2)=   "textTM74"
      Tab(6).Control(3)=   "textTM75"
      Tab(6).Control(4)=   "Label1(107)"
      Tab(6).Control(5)=   "Label1(108)"
      Tab(6).Control(6)=   "Label1(109)"
      Tab(6).Control(7)=   "Label1(110)"
      Tab(6).Control(8)=   "cmdTFBaseNo"
      Tab(6).ControlCount=   9
      TabCaption(7)   =   "其他/商標描述"
      TabPicture(7)   =   "frm020501.frx":0106
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "textTM131"
      Tab(7).Control(1)=   "textTM125"
      Tab(7).Control(2)=   "textTM124"
      Tab(7).Control(3)=   "Label1(118)"
      Tab(7).Control(4)=   "Label1(113)"
      Tab(7).Control(5)=   "Label1(114)"
      Tab(7).Control(6)=   "Label1(112)"
      Tab(7).Control(7)=   "Label1(76)"
      Tab(7).Control(8)=   "Label1(117)"
      Tab(7).Control(9)=   "textTM121"
      Tab(7).Control(10)=   "textTM126"
      Tab(7).Control(11)=   "textTM130"
      Tab(7).Control(12)=   "lblTM137"
      Tab(7).Control(13)=   "textTM137"
      Tab(7).Control(14)=   "lblTM138"
      Tab(7).Control(15)=   "textTM138"
      Tab(7).Control(16)=   "textTM139"
      Tab(7).Control(17)=   "lblTM139"
      Tab(7).ControlCount=   18
      Begin VB.CommandButton cmdTFBaseNo 
         Caption         =   "TF基礎案號數"
         Height          =   285
         Left            =   -74928
         Style           =   1  '圖片外觀
         TabIndex        =   272
         Top             =   1656
         Width           =   1404
      End
      Begin VB.Frame Frame1 
         Height          =   280
         Left            =   -70620
         TabIndex        =   267
         Top             =   3660
         Width           =   2050
         Begin MSForms.TextBox textTM129 
            Height          =   290
            Left            =   930
            TabIndex        =   268
            Top             =   0
            Width           =   330
            VariousPropertyBits=   671105051
            MaxLength       =   1
            Size            =   "582;503"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "不催延展：        (Y:不催)"
            Height          =   180
            Index           =   116
            Left            =   30
            TabIndex        =   269
            Top             =   30
            Width           =   2060
         End
      End
      Begin VB.ComboBox Combo5 
         Height          =   276
         ItemData        =   "frm020501.frx":0122
         Left            =   -69360
         List            =   "frm020501.frx":0135
         Style           =   2  '單純下拉式
         TabIndex        =   82
         Top             =   1560
         Width           =   1470
      End
      Begin VB.ComboBox Combo4 
         Height          =   276
         ItemData        =   "frm020501.frx":0169
         Left            =   -73720
         List            =   "frm020501.frx":016B
         Style           =   2  '單純下拉式
         TabIndex        =   81
         Top             =   1560
         Width           =   1320
      End
      Begin MSForms.TextBox textTM141 
         Height          =   290
         Left            =   -70860
         TabIndex        =   39
         Top             =   1320
         Width           =   410
         VariousPropertyBits=   671105051
         MaxLength       =   2
         Size            =   "723;512"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "延展折扣 :                  ( % )"
         Height          =   180
         Index           =   89
         Left            =   -72120
         TabIndex        =   271
         Top             =   1370
         Width           =   2080
      End
      Begin MSForms.TextBox textTM140 
         Height          =   290
         Left            =   -73710
         TabIndex        =   38
         Top             =   1320
         Width           =   410
         VariousPropertyBits=   671105051
         MaxLength       =   2
         Size            =   "723;512"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "繳註冊費折扣:          ( % )"
         Height          =   180
         Index           =   88
         Left            =   -74880
         TabIndex        =   270
         Top             =   1370
         Width           =   1990
      End
      Begin VB.Label lblTM139 
         AutoSize        =   -1  'True
         Caption         =   "商標描述日文:"
         Height          =   180
         Left            =   -74850
         TabIndex        =   266
         Top             =   3990
         Width           =   1130
      End
      Begin MSForms.TextBox textTM139 
         Height          =   1010
         Left            =   -73620
         TabIndex        =   134
         Top             =   3930
         Width           =   7490
         VariousPropertyBits=   -1467989989
         MaxLength       =   1000
         ScrollBars      =   2
         Size            =   "13212;1782"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM138 
         Height          =   1010
         Left            =   -73620
         TabIndex        =   133
         Top             =   2910
         Width           =   7490
         VariousPropertyBits=   -1467989989
         MaxLength       =   1000
         ScrollBars      =   2
         Size            =   "13212;1782"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label lblTM138 
         AutoSize        =   -1  'True
         Caption         =   "商標描述英文:"
         Height          =   180
         Left            =   -74850
         TabIndex        =   265
         Top             =   2970
         Width           =   1130
      End
      Begin MSForms.TextBox textTM137 
         Height          =   1010
         Left            =   -73620
         TabIndex        =   132
         Top             =   1890
         Width           =   7490
         VariousPropertyBits=   -1467989989
         MaxLength       =   1000
         ScrollBars      =   2
         Size            =   "13212;1782"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label lblTM137 
         AutoSize        =   -1  'True
         Caption         =   "商標描述中文:"
         Height          =   180
         Left            =   -74850
         TabIndex        =   264
         Top             =   1950
         Width           =   1130
      End
      Begin MSForms.ComboBox cboTM72 
         Height          =   300
         Left            =   1560
         TabIndex        =   28
         Top             =   4344
         Width           =   2124
         VariousPropertyBits=   679495707
         DisplayStyle    =   3
         Size            =   "3746;529"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "新細明體"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.ComboBox cboTM08 
         Height          =   300
         Left            =   5220
         TabIndex        =   6
         Top             =   792
         Width           =   2220
         VariousPropertyBits=   679495707
         DisplayStyle    =   3
         Size            =   "3916;529"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "新細明體"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM136 
         Height          =   285
         Left            =   7470
         TabIndex        =   263
         Top             =   3240
         Width           =   375
         VariousPropertyBits=   671105051
         MaxLength       =   1
         Size            =   "656;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "註冊證形式 :            (1:電子 2:紙本)"
         Height          =   180
         Index           =   87
         Left            =   6390
         TabIndex        =   21
         Top             =   3300
         Width           =   2685
      End
      Begin MSForms.TextBox textTM56_1 
         Height          =   290
         Left            =   -73720
         TabIndex        =   78
         Top             =   930
         Width           =   1130
         VariousPropertyBits=   671105051
         MaxLength       =   8
         Size            =   "1993;512"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "代表人9(中,英,日)："
         Height          =   180
         Index           =   86
         Left            =   -74940
         TabIndex        =   262
         Top             =   4148
         Width           =   1560
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "代表人10(中,英,日)："
         Height          =   180
         Index           =   85
         Left            =   -74940
         TabIndex        =   261
         Top             =   4620
         Width           =   1650
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "代表人6(中,英,日)："
         Height          =   180
         Index           =   82
         Left            =   -74940
         TabIndex        =   260
         Top             =   2750
         Width           =   1560
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "代表人7(中,英,日)："
         Height          =   180
         Index           =   81
         Left            =   -74940
         TabIndex        =   259
         Top             =   3216
         Width           =   1560
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "代表人8(中,英,日)："
         Height          =   180
         Index           =   80
         Left            =   -74940
         TabIndex        =   258
         Top             =   3682
         Width           =   1560
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "請款幣別："
         Height          =   180
         Index           =   1
         Left            =   -74940
         TabIndex        =   257
         Top             =   1620
         Width           =   900
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "請款單列印幣別格式："
         Height          =   180
         Left            =   -71220
         TabIndex        =   256
         Top             =   1620
         Width           =   1800
      End
      Begin MSForms.TextBox textCU79 
         Height          =   650
         Left            =   -73710
         TabIndex        =   46
         Top             =   3990
         Width           =   7500
         VariousPropertyBits=   -1466941409
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "13229;1138"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         Caption         =   "客戶備註 :"
         Height          =   260
         Index           =   67
         Left            =   -74700
         TabIndex        =   255
         Top             =   3990
         Width           =   1100
      End
      Begin VB.Label Label1 
         Caption         =   "客戶收款後辦案 :"
         Height          =   260
         Index           =   68
         Left            =   -74880
         TabIndex        =   254
         Top             =   4650
         Width           =   1490
      End
      Begin MSForms.TextBox textCU72 
         Height          =   290
         Left            =   -73350
         TabIndex        =   47
         Top             =   4650
         Width           =   410
         VariousPropertyBits=   671105055
         Size            =   "714;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM46 
         Height          =   285
         Left            =   -73170
         TabIndex        =   79
         Top             =   1230
         Width           =   375
         VariousPropertyBits=   671105051
         MaxLength       =   1
         Size            =   "661;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "D/N是否列印申請人 :             (  Y:印 )"
         Height          =   180
         Index           =   32
         Left            =   -74940
         TabIndex        =   253
         Top             =   1260
         Width           =   2820
      End
      Begin VB.Label Label1 
         Caption         =   "固定請款對象 :"
         Height          =   255
         Index           =   34
         Left            =   -74940
         TabIndex        =   252
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "D/N固定列印對象 :"
         Height          =   276
         Index           =   33
         Left            =   -70932
         TabIndex        =   251
         Top             =   1236
         Width           =   1572
      End
      Begin MSForms.TextBox textTM56_2 
         Height          =   290
         Left            =   -72570
         TabIndex        =   250
         TabStop         =   0   'False
         Top             =   930
         Width           =   6140
         VariousPropertyBits=   671105055
         Size            =   "10830;512"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM69_1 
         Height          =   285
         Left            =   -69360
         TabIndex        =   80
         Top             =   1230
         Width           =   1130
         VariousPropertyBits=   671105051
         MaxLength       =   8
         Size            =   "1993;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM69_2 
         Height          =   290
         Left            =   -68220
         TabIndex        =   249
         TabStop         =   0   'False
         Top             =   1230
         Width           =   1920
         VariousPropertyBits=   671105055
         Size            =   "3387;503"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM05 
         Height          =   620
         Left            =   1560
         TabIndex        =   8
         Top             =   1140
         Width           =   7215
         VariousPropertyBits=   -1467989989
         ScrollBars      =   2
         Size            =   "12726;1094"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.ComboBox cboContact 
         Height          =   320
         Left            =   -68010
         TabIndex        =   37
         Top             =   1030
         Width           =   1770
         VariousPropertyBits=   679495707
         DisplayStyle    =   7
         Size            =   "3122;564"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM68 
         Height          =   290
         Left            =   -67140
         TabIndex        =   71
         Top             =   3660
         Visible         =   0   'False
         Width           =   380
         VariousPropertyBits=   671105051
         MaxLength       =   1
         Size            =   "661;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM71_1 
         Height          =   285
         Left            =   -73830
         TabIndex        =   74
         Top             =   4590
         Width           =   3045
         VariousPropertyBits=   671105051
         MaxLength       =   35
         Size            =   "5371;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM66_1 
         Height          =   285
         Left            =   -73560
         TabIndex        =   72
         Top             =   3960
         Width           =   1095
         VariousPropertyBits=   671105051
         MaxLength       =   8
         Size            =   "1931;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         Caption         =   "延展聯絡人 :"
         Height          =   210
         Index           =   44
         Left            =   -74880
         TabIndex        =   248
         Top             =   4620
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "延展D/N列印對象 :"
         Height          =   276
         Index           =   43
         Left            =   -74880
         TabIndex        =   247
         Top             =   4272
         Width           =   1548
      End
      Begin VB.Label Label1 
         Caption         =   "延展請款對象 :"
         Height          =   270
         Index           =   42
         Left            =   -74880
         TabIndex        =   246
         Top             =   3960
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "延展彼所案號 :"
         Height          =   270
         Index           =   40
         Left            =   -74880
         TabIndex        =   245
         Top             =   3645
         Width           =   1305
      End
      Begin VB.Label Label1 
         Caption         =   "延展代理人 :"
         Height          =   270
         Index           =   39
         Left            =   -74880
         TabIndex        =   244
         Top             =   3330
         Width           =   1095
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "延展單筆不跑 :           Y:不跑"
         Height          =   180
         Index           =   41
         Left            =   -68370
         TabIndex        =   243
         Top             =   3710
         Visible         =   0   'False
         Width           =   2340
      End
      Begin MSForms.TextBox textTM33_1 
         Height          =   285
         Left            =   -73560
         TabIndex        =   69
         Top             =   3330
         Width           =   1095
         VariousPropertyBits=   671105051
         MaxLength       =   8
         Size            =   "1931;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM70_1 
         Height          =   285
         Left            =   -73350
         TabIndex        =   73
         Top             =   4275
         Width           =   1095
         VariousPropertyBits=   671105051
         MaxLength       =   8
         Size            =   "1931;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM70_2 
         Height          =   285
         Left            =   -72180
         TabIndex        =   242
         TabStop         =   0   'False
         Top             =   4275
         Width           =   5910
         VariousPropertyBits=   671105055
         Size            =   "10425;503"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM65 
         Height          =   285
         Left            =   -73560
         TabIndex        =   70
         Top             =   3645
         Width           =   2775
         VariousPropertyBits=   671105051
         MaxLength       =   30
         Size            =   "4895;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM33_2 
         Height          =   285
         Left            =   -72420
         TabIndex        =   241
         TabStop         =   0   'False
         Top             =   3330
         Width           =   6135
         VariousPropertyBits=   671105055
         Size            =   "10821;503"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM66_2 
         Height          =   285
         Left            =   -72420
         TabIndex        =   240
         TabStop         =   0   'False
         Top             =   3960
         Width           =   6135
         VariousPropertyBits=   671105055
         Size            =   "10821;503"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "申請人4 :"
         Height          =   180
         Index           =   48
         Left            =   -74880
         TabIndex        =   239
         Top             =   382
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "申請人5 :"
         Height          =   180
         Index           =   49
         Left            =   -74880
         TabIndex        =   238
         Top             =   699
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "申請地址4(日) :"
         Height          =   180
         Index           =   61
         Left            =   -74880
         TabIndex        =   237
         Top             =   1820
         Width           =   1200
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "申請地址4(英) :"
         Height          =   180
         Index           =   60
         Left            =   -74880
         TabIndex        =   236
         Top             =   1437
         Width           =   1200
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "申請地址4(中) :"
         Height          =   180
         Index           =   59
         Left            =   -74880
         TabIndex        =   235
         Top             =   1054
         Width           =   1200
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "申請地址5(日) :"
         Height          =   180
         Index           =   64
         Left            =   -74880
         TabIndex        =   234
         Top             =   2970
         Width           =   1200
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "申請地址5(英) :"
         Height          =   180
         Index           =   63
         Left            =   -74880
         TabIndex        =   233
         Top             =   2586
         Width           =   1200
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "申請地址5(中) :"
         Height          =   180
         Index           =   62
         Left            =   -74880
         TabIndex        =   232
         Top             =   2203
         Width           =   1200
      End
      Begin MSForms.TextBox textTM80_1 
         Height          =   285
         Left            =   -74130
         TabIndex        =   61
         Top             =   330
         Width           =   1095
         VariousPropertyBits=   671105051
         MaxLength       =   9
         Size            =   "1926;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM81_1 
         Height          =   285
         Left            =   -74130
         TabIndex        =   62
         Top             =   647
         Width           =   1095
         VariousPropertyBits=   671105051
         MaxLength       =   9
         Size            =   "1931;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM80_2 
         Height          =   285
         Left            =   -73020
         TabIndex        =   231
         TabStop         =   0   'False
         Top             =   330
         Width           =   6885
         VariousPropertyBits=   671105055
         Size            =   "12144;503"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM81_2 
         Height          =   285
         Left            =   -73020
         TabIndex        =   230
         TabStop         =   0   'False
         Top             =   647
         Width           =   6885
         VariousPropertyBits=   671105055
         Size            =   "12144;503"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM84 
         Height          =   360
         Left            =   -73560
         TabIndex        =   63
         Top             =   964
         Width           =   7335
         VariousPropertyBits=   671105051
         MaxLength       =   80
         Size            =   "12938;635"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM88 
         Height          =   360
         Left            =   -73560
         TabIndex        =   64
         Top             =   1347
         Width           =   7335
         VariousPropertyBits=   671105051
         MaxLength       =   185
         Size            =   "12938;635"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM92 
         Height          =   360
         Left            =   -73560
         TabIndex        =   65
         Top             =   1730
         Width           =   7335
         VariousPropertyBits=   671105051
         MaxLength       =   80
         Size            =   "12938;635"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM85 
         Height          =   360
         Left            =   -73560
         TabIndex        =   66
         Top             =   2113
         Width           =   7335
         VariousPropertyBits=   671105051
         MaxLength       =   80
         Size            =   "12938;635"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM89 
         Height          =   360
         Left            =   -73560
         TabIndex        =   67
         Top             =   2496
         Width           =   7335
         VariousPropertyBits=   671105051
         MaxLength       =   185
         Size            =   "12938;635"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM93 
         Height          =   360
         Left            =   -73560
         TabIndex        =   68
         Top             =   2880
         Width           =   7335
         VariousPropertyBits=   671105051
         MaxLength       =   80
         Size            =   "12938;635"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         Caption         =   "彼所案號 :"
         Height          =   270
         Index           =   66
         Left            =   -74940
         TabIndex        =   229
         Top             =   644
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "FC代理人 :"
         Height          =   270
         Index           =   65
         Left            =   -74940
         TabIndex        =   228
         Top             =   337
         Width           =   1215
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "聯絡人部門(日) :"
         Height          =   180
         Index           =   75
         Left            =   -74940
         TabIndex        =   227
         Top             =   3795
         Width           =   1290
      End
      Begin VB.Label Label1 
         Caption         =   "聯絡人2(日) :"
         Height          =   255
         Index           =   74
         Left            =   -74940
         TabIndex        =   226
         Top             =   3495
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "聯絡人2(英) :"
         Height          =   255
         Index           =   73
         Left            =   -74940
         TabIndex        =   225
         Top             =   3180
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "聯絡人2(中) :"
         Height          =   255
         Index           =   72
         Left            =   -74940
         TabIndex        =   224
         Top             =   2865
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "聯絡人1(日) :"
         Height          =   255
         Index           =   71
         Left            =   -74940
         TabIndex        =   223
         Top             =   2550
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "聯絡人1(英) :"
         Height          =   255
         Index           =   70
         Left            =   -74940
         TabIndex        =   222
         Top             =   2235
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "聯絡人1(中) :"
         Height          =   255
         Index           =   69
         Left            =   -74940
         TabIndex        =   221
         Top             =   1920
         Width           =   1125
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "CLIENT_MATTER_ID:"
         Height          =   180
         Index           =   169
         Left            =   -70660
         TabIndex        =   220
         Top             =   690
         Width           =   1730
      End
      Begin MSForms.TextBox textTM45 
         Height          =   290
         Left            =   -73720
         TabIndex        =   76
         Top             =   640
         Width           =   3050
         VariousPropertyBits=   671105051
         MaxLength       =   100
         Size            =   "5380;512"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM44_1 
         Height          =   290
         Left            =   -73720
         TabIndex        =   75
         Top             =   330
         Width           =   1130
         VariousPropertyBits=   671105051
         MaxLength       =   8
         Size            =   "1993;512"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM44_2 
         Height          =   290
         Left            =   -72570
         TabIndex        =   219
         TabStop         =   0   'False
         Top             =   330
         Width           =   6140
         VariousPropertyBits=   671105055
         Size            =   "10821;503"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM76 
         Height          =   290
         Left            =   -73590
         TabIndex        =   89
         Top             =   3750
         Width           =   7320
         VariousPropertyBits=   671105051
         Size            =   "12912;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM43 
         Height          =   290
         Left            =   -73720
         TabIndex        =   88
         Top             =   3440
         Width           =   7460
         VariousPropertyBits=   671105051
         Size            =   "13150;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM42 
         Height          =   290
         Left            =   -73720
         TabIndex        =   87
         Top             =   3140
         Width           =   7460
         VariousPropertyBits=   671105051
         Size            =   "13150;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM41 
         Height          =   290
         Left            =   -73720
         TabIndex        =   86
         Top             =   2820
         Width           =   7460
         VariousPropertyBits=   671105051
         Size            =   "13150;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM40 
         Height          =   290
         Left            =   -73720
         TabIndex        =   85
         Top             =   2520
         Width           =   7460
         VariousPropertyBits=   671105051
         Size            =   "13150;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM39 
         Height          =   290
         Left            =   -73720
         TabIndex        =   84
         Top             =   2210
         Width           =   7460
         VariousPropertyBits=   671105051
         Size            =   "13150;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM38 
         Height          =   290
         Left            =   -73720
         TabIndex        =   83
         Top             =   1910
         Width           =   7460
         VariousPropertyBits=   671105051
         Size            =   "13150;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM127 
         Height          =   285
         Left            =   -68910
         TabIndex        =   77
         Top             =   637
         Width           =   2700
         VariousPropertyBits=   671105051
         MaxLength       =   20
         Size            =   "4762;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "代表人3(中,英,日)："
         Height          =   180
         Index           =   79
         Left            =   -74940
         TabIndex        =   218
         Top             =   1352
         Width           =   1560
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "代表人2(中,英,日)："
         Height          =   180
         Index           =   78
         Left            =   -74940
         TabIndex        =   217
         Top             =   886
         Width           =   1560
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "代表人1(中,英,日)："
         Height          =   180
         Index           =   77
         Left            =   -74940
         TabIndex        =   216
         Top             =   420
         Width           =   1560
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "代表人5(中,英,日)："
         Height          =   180
         Index           =   84
         Left            =   -74940
         TabIndex        =   215
         Top             =   2284
         Width           =   1560
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "代表人4(中,英,日)："
         Height          =   180
         Index           =   83
         Left            =   -74940
         TabIndex        =   214
         Top             =   1818
         Width           =   1560
      End
      Begin MSForms.TextBox textTM52 
         Height          =   420
         Left            =   -68310
         TabIndex        =   99
         Top             =   766
         Width           =   2445
         VariousPropertyBits=   -1467989989
         MaxLength       =   40
         ScrollBars      =   2
         Size            =   "4313;741"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM51 
         Height          =   420
         Left            =   -70800
         TabIndex        =   98
         Top             =   766
         Width           =   2445
         VariousPropertyBits=   -1467989989
         MaxLength       =   80
         ScrollBars      =   2
         Size            =   "4313;741"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM50 
         Height          =   420
         Left            =   -73260
         TabIndex        =   97
         Top             =   766
         Width           =   2445
         VariousPropertyBits=   -1467989989
         MaxLength       =   50
         ScrollBars      =   2
         Size            =   "4313;741"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM49 
         Height          =   420
         Left            =   -68310
         TabIndex        =   96
         Top             =   300
         Width           =   2445
         VariousPropertyBits=   -1467989989
         MaxLength       =   40
         ScrollBars      =   2
         Size            =   "4313;741"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM48 
         Height          =   420
         Left            =   -70800
         TabIndex        =   95
         Top             =   300
         Width           =   2445
         VariousPropertyBits=   -1467989989
         MaxLength       =   80
         ScrollBars      =   2
         Size            =   "4313;741"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM47 
         Height          =   420
         Left            =   -73260
         TabIndex        =   94
         Top             =   300
         Width           =   2445
         VariousPropertyBits=   -1467989989
         MaxLength       =   50
         ScrollBars      =   2
         Size            =   "4313;741"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM94 
         Height          =   420
         Left            =   -73260
         TabIndex        =   100
         Top             =   1232
         Width           =   2445
         VariousPropertyBits=   -1467989989
         MaxLength       =   50
         ScrollBars      =   2
         Size            =   "4313;741"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM95 
         Height          =   420
         Left            =   -70800
         TabIndex        =   101
         Top             =   1232
         Width           =   2445
         VariousPropertyBits=   -1467989989
         MaxLength       =   80
         ScrollBars      =   2
         Size            =   "4313;741"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM96 
         Height          =   420
         Left            =   -68310
         TabIndex        =   102
         Top             =   1232
         Width           =   2445
         VariousPropertyBits=   -1467989989
         MaxLength       =   40
         ScrollBars      =   2
         Size            =   "4313;741"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM97 
         Height          =   420
         Left            =   -73260
         TabIndex        =   103
         Top             =   1698
         Width           =   2445
         VariousPropertyBits=   -1467989989
         MaxLength       =   50
         ScrollBars      =   2
         Size            =   "4313;741"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM98 
         Height          =   420
         Left            =   -70800
         TabIndex        =   104
         Top             =   1698
         Width           =   2445
         VariousPropertyBits=   -1467989989
         MaxLength       =   80
         ScrollBars      =   2
         Size            =   "4313;741"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM99 
         Height          =   420
         Left            =   -68310
         TabIndex        =   105
         Top             =   1698
         Width           =   2445
         VariousPropertyBits=   -1467989989
         MaxLength       =   40
         ScrollBars      =   2
         Size            =   "4313;741"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM100 
         Height          =   420
         Left            =   -73260
         TabIndex        =   106
         Top             =   2164
         Width           =   2445
         VariousPropertyBits=   -1467989989
         MaxLength       =   50
         ScrollBars      =   2
         Size            =   "4313;741"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM101 
         Height          =   420
         Left            =   -70800
         TabIndex        =   107
         Top             =   2164
         Width           =   2445
         VariousPropertyBits=   -1467989989
         MaxLength       =   80
         ScrollBars      =   2
         Size            =   "4313;741"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM102 
         Height          =   420
         Left            =   -68310
         TabIndex        =   108
         Top             =   2164
         Width           =   2445
         VariousPropertyBits=   -1467989989
         MaxLength       =   40
         ScrollBars      =   2
         Size            =   "4313;741"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM103 
         Height          =   420
         Left            =   -73260
         TabIndex        =   109
         Top             =   2630
         Width           =   2445
         VariousPropertyBits=   -1467989989
         MaxLength       =   50
         ScrollBars      =   2
         Size            =   "4313;741"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM104 
         Height          =   420
         Left            =   -70800
         TabIndex        =   110
         Top             =   2630
         Width           =   2445
         VariousPropertyBits=   -1467989989
         MaxLength       =   80
         ScrollBars      =   2
         Size            =   "4313;741"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM105 
         Height          =   420
         Left            =   -68310
         TabIndex        =   111
         Top             =   2630
         Width           =   2445
         VariousPropertyBits=   -1467989989
         MaxLength       =   40
         ScrollBars      =   2
         Size            =   "4313;741"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM106 
         Height          =   420
         Left            =   -73260
         TabIndex        =   112
         Top             =   3096
         Width           =   2445
         VariousPropertyBits=   -1467989989
         MaxLength       =   50
         ScrollBars      =   2
         Size            =   "4313;741"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM107 
         Height          =   420
         Left            =   -70800
         TabIndex        =   113
         Top             =   3096
         Width           =   2445
         VariousPropertyBits=   -1467989989
         MaxLength       =   80
         ScrollBars      =   2
         Size            =   "4313;741"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM108 
         Height          =   420
         Left            =   -68310
         TabIndex        =   114
         Top             =   3096
         Width           =   2445
         VariousPropertyBits=   -1467989989
         MaxLength       =   40
         ScrollBars      =   2
         Size            =   "4313;741"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM109 
         Height          =   420
         Left            =   -73260
         TabIndex        =   115
         Top             =   3562
         Width           =   2445
         VariousPropertyBits=   -1467989989
         MaxLength       =   50
         ScrollBars      =   2
         Size            =   "4313;741"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM110 
         Height          =   420
         Left            =   -70800
         TabIndex        =   116
         Top             =   3562
         Width           =   2445
         VariousPropertyBits=   -1467989989
         MaxLength       =   80
         ScrollBars      =   2
         Size            =   "4313;741"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM111 
         Height          =   420
         Left            =   -68310
         TabIndex        =   117
         Top             =   3562
         Width           =   2445
         VariousPropertyBits=   -1467989989
         MaxLength       =   40
         ScrollBars      =   2
         Size            =   "4313;741"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM112 
         Height          =   420
         Left            =   -73260
         TabIndex        =   118
         Top             =   4028
         Width           =   2445
         VariousPropertyBits=   -1467989989
         MaxLength       =   50
         ScrollBars      =   2
         Size            =   "4313;741"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM113 
         Height          =   420
         Left            =   -70800
         TabIndex        =   119
         Top             =   4028
         Width           =   2445
         VariousPropertyBits=   -1467989989
         MaxLength       =   80
         ScrollBars      =   2
         Size            =   "4313;741"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM114 
         Height          =   420
         Left            =   -68310
         TabIndex        =   120
         Top             =   4028
         Width           =   2445
         VariousPropertyBits=   -1467989989
         MaxLength       =   40
         ScrollBars      =   2
         Size            =   "4313;741"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM115 
         Height          =   420
         Left            =   -73260
         TabIndex        =   121
         Top             =   4500
         Width           =   2445
         VariousPropertyBits=   -1467989989
         MaxLength       =   50
         ScrollBars      =   2
         Size            =   "4313;741"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM116 
         Height          =   420
         Left            =   -70800
         TabIndex        =   122
         Top             =   4500
         Width           =   2445
         VariousPropertyBits=   -1467989989
         MaxLength       =   80
         ScrollBars      =   2
         Size            =   "4313;741"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM117 
         Height          =   420
         Left            =   -68310
         TabIndex        =   123
         Top             =   4500
         Width           =   2445
         VariousPropertyBits=   -1467989989
         MaxLength       =   40
         ScrollBars      =   2
         Size            =   "4313;741"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "分所銷卷備註："
         Height          =   180
         Index           =   110
         Left            =   -74910
         TabIndex        =   213
         Top             =   1290
         Width           =   1260
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "分所銷卷員："
         Height          =   180
         Index           =   109
         Left            =   -74910
         TabIndex        =   212
         Top             =   990
         Width           =   1080
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "分所銷卷日："
         Height          =   180
         Index           =   108
         Left            =   -74910
         TabIndex        =   211
         Top             =   720
         Width           =   1080
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "北所銷卷日："
         Height          =   180
         Index           =   107
         Left            =   -74910
         TabIndex        =   210
         Top             =   450
         Width           =   1080
      End
      Begin MSForms.TextBox textTM75 
         Height          =   285
         Left            =   -73680
         TabIndex        =   209
         Top             =   1290
         Width           =   7605
         VariousPropertyBits=   671105055
         BackColor       =   16777215
         Size            =   "13414;503"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM74 
         Height          =   285
         Left            =   -73830
         TabIndex        =   208
         Top             =   990
         Width           =   1305
         VariousPropertyBits=   671105055
         BackColor       =   16777215
         Size            =   "2302;503"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM73 
         Height          =   285
         Left            =   -73830
         TabIndex        =   207
         Top             =   720
         Width           =   1305
         VariousPropertyBits=   671105055
         BackColor       =   16777215
         Size            =   "2302;503"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM57 
         Height          =   285
         Left            =   -73830
         TabIndex        =   206
         Top             =   450
         Width           =   1305
         VariousPropertyBits=   671105055
         BackColor       =   16777215
         Size            =   "2302;503"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM130 
         Height          =   290
         Left            =   -73620
         TabIndex        =   130
         Top             =   960
         Width           =   320
         VariousPropertyBits=   671105051
         MaxLength       =   1
         Size            =   "564;512"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM126 
         Height          =   290
         Left            =   -69870
         TabIndex        =   129
         Top             =   650
         Width           =   300
         VariousPropertyBits=   671105051
         MaxLength       =   1
         Size            =   "529;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM121 
         Height          =   290
         Left            =   -73620
         TabIndex        =   128
         Top             =   650
         Width           =   300
         VariousPropertyBits=   671105051
         MaxLength       =   1
         Size            =   "529;512"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "特殊出名公司:          ( J:智權公司 空白:系統預設)"
         Height          =   180
         Index           =   117
         Left            =   -74850
         TabIndex        =   205
         Top             =   1010
         Width           =   3740
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "以EMail通知:        (Y:是   D:僅D/N）"
         Height          =   180
         Index           =   76
         Left            =   -74730
         TabIndex        =   204
         Top             =   700
         Width           =   2720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "EMail同時寄紙本 :         ( Y:是)"
         Height          =   180
         Index           =   112
         Left            =   -71340
         TabIndex        =   203
         Top             =   700
         Width           =   2330
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "定稿份數 : "
         Height          =   180
         Index           =   114
         Left            =   -74580
         TabIndex        =   202
         Top             =   380
         Width           =   860
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "請款單份數 : "
         Height          =   180
         Index           =   113
         Left            =   -70950
         TabIndex        =   201
         Top             =   380
         Width           =   1040
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "定稿商標名稱:"
         Height          =   180
         Index           =   118
         Left            =   -74850
         TabIndex        =   200
         Top             =   1310
         Width           =   1130
      End
      Begin MSForms.TextBox textTM124 
         Height          =   290
         Left            =   -73620
         TabIndex        =   126
         Top             =   330
         Width           =   380
         VariousPropertyBits=   671105051
         MaxLength       =   1
         Size            =   "670;512"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM125 
         Height          =   290
         Left            =   -69870
         TabIndex        =   127
         Top             =   330
         Width           =   380
         VariousPropertyBits=   671105051
         MaxLength       =   1
         Size            =   "661;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM131 
         Height          =   620
         Left            =   -73620
         TabIndex        =   131
         Top             =   1250
         Width           =   7490
         VariousPropertyBits=   -1467989989
         MaxLength       =   140
         ScrollBars      =   2
         Size            =   "13212;1094"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM77 
         Height          =   285
         Left            =   5220
         TabIndex        =   25
         Top             =   3795
         Width           =   375
         VariousPropertyBits=   671105051
         MaxLength       =   1
         Size            =   "656;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM122 
         Height          =   290
         Left            =   -71100
         TabIndex        =   45
         Top             =   3690
         Width           =   320
         VariousPropertyBits=   671105051
         MaxLength       =   1
         Size            =   "556;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM118 
         Height          =   360
         Left            =   1560
         TabIndex        =   31
         Top             =   4650
         Width           =   7455
         VariousPropertyBits=   -1467989989
         MaxLength       =   150
         ScrollBars      =   2
         Size            =   "13150;626"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM91 
         Height          =   360
         Left            =   -73560
         TabIndex        =   60
         Top             =   4110
         Width           =   7335
         VariousPropertyBits=   671105051
         MaxLength       =   80
         Size            =   "12938;635"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM87 
         Height          =   360
         Left            =   -73560
         TabIndex        =   59
         Top             =   3761
         Width           =   7335
         VariousPropertyBits=   671105051
         MaxLength       =   185
         Size            =   "12938;635"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM83 
         Height          =   360
         Left            =   -73560
         TabIndex        =   58
         Top             =   3408
         Width           =   7335
         VariousPropertyBits=   671105051
         MaxLength       =   80
         Size            =   "12938;635"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM90 
         Height          =   360
         Left            =   -73560
         TabIndex        =   57
         Top             =   3055
         Width           =   7335
         VariousPropertyBits=   671105051
         MaxLength       =   80
         Size            =   "12938;635"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM86 
         Height          =   360
         Left            =   -73560
         TabIndex        =   55
         Top             =   2702
         Width           =   7335
         VariousPropertyBits=   671105051
         MaxLength       =   185
         Size            =   "12938;635"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM82 
         Height          =   360
         Left            =   -73560
         TabIndex        =   56
         Top             =   2349
         Width           =   7335
         VariousPropertyBits=   671105051
         MaxLength       =   80
         Size            =   "12938;635"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM26 
         Height          =   360
         Left            =   -73560
         TabIndex        =   54
         Top             =   1996
         Width           =   7335
         VariousPropertyBits=   671105051
         MaxLength       =   80
         Size            =   "12938;635"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM25 
         Height          =   360
         Left            =   -73560
         TabIndex        =   53
         Top             =   1643
         Width           =   7335
         VariousPropertyBits=   671105051
         MaxLength       =   185
         Size            =   "12938;635"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM24 
         Height          =   360
         Left            =   -73560
         TabIndex        =   52
         Top             =   1290
         Width           =   7335
         VariousPropertyBits=   671105051
         MaxLength       =   80
         Size            =   "12938;635"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM79_2 
         Height          =   285
         Left            =   -72960
         TabIndex        =   180
         TabStop         =   0   'False
         Top             =   975
         Width           =   6885
         VariousPropertyBits=   671105055
         Size            =   "12144;503"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM78_2 
         Height          =   285
         Left            =   -72960
         TabIndex        =   178
         TabStop         =   0   'False
         Top             =   645
         Width           =   6885
         VariousPropertyBits=   671105055
         Size            =   "12144;503"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM23_2 
         Height          =   285
         Left            =   -72960
         TabIndex        =   174
         TabStop         =   0   'False
         Top             =   330
         Width           =   6885
         VariousPropertyBits=   671105055
         Size            =   "12144;503"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM79_1 
         Height          =   285
         Left            =   -74070
         TabIndex        =   51
         Top             =   975
         Width           =   1095
         VariousPropertyBits=   671105051
         MaxLength       =   9
         Size            =   "1926;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM78_1 
         Height          =   285
         Left            =   -74070
         TabIndex        =   50
         Top             =   645
         Width           =   1095
         VariousPropertyBits=   671105051
         MaxLength       =   9
         Size            =   "1926;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM72_2 
         Height          =   288
         Left            =   2904
         TabIndex        =   177
         TabStop         =   0   'False
         Top             =   4056
         Visible         =   0   'False
         Width           =   612
         VariousPropertyBits=   671105055
         BackColor       =   16777152
         Size            =   "1080;508"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM72_1 
         Height          =   288
         Left            =   2544
         TabIndex        =   29
         Top             =   4056
         Visible         =   0   'False
         Width           =   372
         VariousPropertyBits=   671105051
         BackColor       =   16777152
         MaxLength       =   1
         Size            =   "656;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM23_1 
         Height          =   285
         Left            =   -74070
         TabIndex        =   49
         Top             =   330
         Width           =   1095
         VariousPropertyBits=   671105051
         MaxLength       =   9
         Size            =   "1926;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textFA39 
         Height          =   290
         Left            =   -73260
         TabIndex        =   44
         TabStop         =   0   'False
         Top             =   3690
         Width           =   320
         VariousPropertyBits=   671105055
         Size            =   "556;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textFA29 
         Height          =   650
         Left            =   -73710
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   3060
         Width           =   7500
         VariousPropertyBits=   -1466941409
         ScrollBars      =   2
         Size            =   "13229;1138"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM58 
         Height          =   650
         Left            =   -73710
         TabIndex        =   42
         Top             =   2190
         Width           =   7500
         VariousPropertyBits=   -1466941413
         MaxLength       =   2000
         ScrollBars      =   2
         Size            =   "13229;1138"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM67 
         Height          =   420
         Left            =   -71250
         TabIndex        =   48
         Top             =   4650
         Width           =   5030
         VariousPropertyBits=   -1476378597
         MaxLength       =   200
         Size            =   "8864;741"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM54_2 
         Height          =   290
         Left            =   -72330
         TabIndex        =   165
         TabStop         =   0   'False
         Top             =   1620
         Width           =   6020
         VariousPropertyBits=   671105055
         Size            =   "10610;503"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM55 
         Height          =   290
         Left            =   -73710
         TabIndex        =   41
         Top             =   1920
         Width           =   2720
         VariousPropertyBits=   671105051
         MaxLength       =   35
         Size            =   "4789;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM37 
         Height          =   290
         Left            =   -70860
         TabIndex        =   36
         Top             =   1030
         Width           =   410
         VariousPropertyBits=   671105051
         MaxLength       =   2
         Size            =   "723;512"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM35 
         Height          =   290
         Left            =   -69600
         TabIndex        =   34
         Top             =   740
         Width           =   3380
         VariousPropertyBits=   671105051
         MaxLength       =   50
         Size            =   "5953;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textSP32 
         Height          =   290
         Left            =   -69600
         TabIndex        =   161
         TabStop         =   0   'False
         Top             =   1920
         Width           =   3380
         VariousPropertyBits=   671105055
         Size            =   "5953;503"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM54_1 
         Height          =   290
         Left            =   -73710
         TabIndex        =   40
         Top             =   1620
         Width           =   1340
         VariousPropertyBits=   671105051
         MaxLength       =   8
         Size            =   "2355;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM36 
         Height          =   290
         Left            =   -73710
         TabIndex        =   35
         Top             =   1030
         Width           =   410
         VariousPropertyBits=   671105051
         MaxLength       =   2
         Size            =   "723;512"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM34 
         Height          =   290
         Left            =   -73710
         TabIndex        =   33
         Top             =   740
         Width           =   2750
         VariousPropertyBits=   671105051
         MaxLength       =   50
         Size            =   "4842;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM32 
         Height          =   460
         Left            =   -73710
         TabIndex        =   32
         Top             =   270
         Width           =   7460
         VariousPropertyBits=   -1467989989
         MaxLength       =   1500
         ScrollBars      =   2
         Size            =   "13150;811"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM31_2 
         Height          =   285
         Left            =   5640
         TabIndex        =   155
         TabStop         =   0   'False
         Top             =   4350
         Width           =   2475
         VariousPropertyBits=   671105055
         Size            =   "4360;503"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM31_1 
         Height          =   285
         Left            =   5220
         TabIndex        =   30
         Top             =   4350
         Width           =   375
         VariousPropertyBits=   671105051
         MaxLength       =   2
         Size            =   "656;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM29 
         Height          =   285
         Left            =   5220
         TabIndex        =   27
         Top             =   4065
         Width           =   375
         VariousPropertyBits=   671105051
         MaxLength       =   1
         Size            =   "656;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM19 
         Height          =   285
         Left            =   5220
         TabIndex        =   23
         Top             =   3510
         Width           =   375
         VariousPropertyBits=   671105051
         MaxLength       =   1
         Size            =   "656;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM17 
         Height          =   285
         Left            =   5220
         TabIndex        =   20
         Top             =   3225
         Width           =   375
         VariousPropertyBits=   671105051
         MaxLength       =   1
         Size            =   "656;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM27 
         Height          =   285
         Left            =   5730
         TabIndex        =   18
         Top             =   2955
         Width           =   2955
         VariousPropertyBits=   671105051
         MaxLength       =   20
         Size            =   "5207;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM13 
         Height          =   285
         Left            =   5730
         TabIndex        =   15
         Top             =   2670
         Width           =   975
         VariousPropertyBits=   671105051
         MaxLength       =   7
         Size            =   "1714;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM15 
         Height          =   285
         Left            =   5730
         TabIndex        =   13
         Top             =   2400
         Width           =   2415
         VariousPropertyBits=   671105051
         MaxLength       =   20
         Size            =   "4254;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM12 
         Height          =   285
         Left            =   5730
         TabIndex        =   11
         Top             =   2115
         Width           =   2415
         VariousPropertyBits=   671105051
         MaxLength       =   30
         Size            =   "4254;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM08_2 
         Height          =   288
         Left            =   4608
         TabIndex        =   154
         TabStop         =   0   'False
         Top             =   504
         Visible         =   0   'False
         Width           =   1932
         VariousPropertyBits=   671105055
         BackColor       =   16777152
         Size            =   "3408;503"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM08_1 
         Height          =   288
         Left            =   4200
         TabIndex        =   7
         Top             =   528
         Visible         =   0   'False
         Width           =   372
         VariousPropertyBits=   671105051
         BackColor       =   16777152
         MaxLength       =   1
         Size            =   "656;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM30 
         Height          =   285
         Left            =   1560
         TabIndex        =   26
         Top             =   4066
         Width           =   972
         VariousPropertyBits=   671105051
         MaxLength       =   7
         Size            =   "1714;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM53 
         Height          =   285
         Left            =   1560
         TabIndex        =   24
         Top             =   3788
         Width           =   372
         VariousPropertyBits=   671105051
         MaxLength       =   1
         Size            =   "656;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM18 
         Height          =   285
         Left            =   1560
         TabIndex        =   22
         Top             =   3510
         Width           =   372
         VariousPropertyBits=   671105051
         MaxLength       =   1
         Size            =   "656;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM16 
         Height          =   285
         Left            =   1560
         TabIndex        =   19
         Top             =   3232
         Width           =   372
         VariousPropertyBits=   671105051
         MaxLength       =   1
         Size            =   "656;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM22 
         Height          =   285
         Left            =   2856
         TabIndex        =   17
         Top             =   2954
         Width           =   972
         VariousPropertyBits=   671105051
         MaxLength       =   8
         Size            =   "1714;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM21 
         Height          =   285
         Left            =   1560
         TabIndex        =   16
         Top             =   2954
         Width           =   972
         VariousPropertyBits=   671105051
         MaxLength       =   8
         Size            =   "1714;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM20 
         Height          =   285
         Left            =   1560
         TabIndex        =   14
         Top             =   2676
         Width           =   972
         VariousPropertyBits=   671105051
         MaxLength       =   7
         Size            =   "1714;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM14 
         Height          =   285
         Left            =   1560
         TabIndex        =   12
         Top             =   2398
         Width           =   972
         VariousPropertyBits=   671105051
         MaxLength       =   7
         Size            =   "1714;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM11 
         Height          =   285
         Left            =   1560
         TabIndex        =   10
         Top             =   2120
         Width           =   972
         VariousPropertyBits=   671105051
         MaxLength       =   7
         Size            =   "1714;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM09 
         Height          =   355
         Left            =   1560
         TabIndex        =   9
         Top             =   1760
         Width           =   6612
         VariousPropertyBits=   -1467989989
         MaxLength       =   395
         ScrollBars      =   2
         Size            =   "11663;626"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM10_2 
         Height          =   285
         Left            =   2220
         TabIndex        =   144
         TabStop         =   0   'False
         Top             =   856
         Width           =   1812
         VariousPropertyBits=   671105055
         Size            =   "3196;503"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM04 
         Height          =   285
         Left            =   3360
         TabIndex        =   4
         Top             =   300
         Width           =   372
         VariousPropertyBits=   671105051
         MaxLength       =   2
         Size            =   "656;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM10_1 
         Height          =   285
         Left            =   1560
         TabIndex        =   5
         Top             =   856
         Width           =   612
         VariousPropertyBits=   671105051
         MaxLength       =   3
         Size            =   "1080;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM28 
         Height          =   285
         Left            =   1560
         TabIndex        =   143
         TabStop         =   0   'False
         Top             =   578
         Width           =   2532
         VariousPropertyBits=   671105055
         Size            =   "4466;503"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM03 
         Height          =   285
         Left            =   3120
         TabIndex        =   3
         Top             =   300
         Width           =   252
         VariousPropertyBits=   671105051
         MaxLength       =   1
         Size            =   "444;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM02_2 
         Height          =   285
         Left            =   2880
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   300
         Width           =   252
         VariousPropertyBits=   671105051
         MaxLength       =   1
         Size            =   "444;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM02_1 
         Height          =   285
         Left            =   2040
         TabIndex        =   1
         Top             =   300
         Width           =   1092
         VariousPropertyBits=   671105051
         MaxLength       =   5
         Size            =   "1926;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM01 
         Height          =   285
         Left            =   1560
         TabIndex        =   0
         Top             =   300
         Width           =   495
         VariousPropertyBits=   671105051
         MaxLength       =   3
         Size            =   "873;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCUID_2 
         Height          =   285
         Left            =   5820
         TabIndex        =   199
         Top             =   578
         Width           =   3135
         VariousPropertyBits=   671105055
         Size            =   "5530;503"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCUID_1 
         Height          =   285
         Left            =   5820
         TabIndex        =   198
         Top             =   300
         Width           =   3135
         VariousPropertyBits=   671105055
         Size            =   "5530;503"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "案件備註欄：不可銷卷案請加註 ""不銷卷"" 字樣！  與他案合併計算結餘請註明""與某案號合併計算結餘""！"
         ForeColor       =   &H000000FF&
         Height          =   180
         Index           =   115
         Left            =   -74430
         TabIndex        =   197
         Top             =   2850
         Width           =   8220
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "畫面上定稿語文 :           (N:不印 1:台->各國 2:外->台 3:英文)"
         Height          =   180
         Index           =   111
         Left            =   3780
         TabIndex        =   196
         Top             =   3840
         Width           =   4590
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "接洽人:"
         Height          =   180
         Index           =   160
         Left            =   -68670
         TabIndex        =   195
         Top             =   1080
         Width           =   590
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "同意書商標號數 :"
         Height          =   180
         Index           =   23
         Left            =   120
         TabIndex        =   194
         Top             =   4650
         Width           =   1350
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "申請地址3(中) :"
         Height          =   180
         Index           =   56
         Left            =   -74910
         TabIndex        =   187
         Top             =   3498
         Width           =   1200
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "申請地址3(英) :"
         Height          =   180
         Index           =   57
         Left            =   -74910
         TabIndex        =   186
         Top             =   3851
         Width           =   1200
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "申請地址3(日) :"
         Height          =   180
         Index           =   58
         Left            =   -74910
         TabIndex        =   185
         Top             =   4200
         Width           =   1200
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "申請地址2(中) :"
         Height          =   180
         Index           =   53
         Left            =   -74910
         TabIndex        =   184
         Top             =   2439
         Width           =   1200
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "申請地址2(英) :"
         Height          =   180
         Index           =   54
         Left            =   -74910
         TabIndex        =   183
         Top             =   2792
         Width           =   1200
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "申請地址2(日) :"
         Height          =   180
         Index           =   55
         Left            =   -74910
         TabIndex        =   182
         Top             =   3145
         Width           =   1200
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "申請人3 :"
         Height          =   180
         Index           =   47
         Left            =   -74910
         TabIndex        =   181
         Top             =   1020
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "申請人2 :"
         Height          =   180
         Index           =   46
         Left            =   -74910
         TabIndex        =   179
         Top             =   705
         Width           =   720
      End
      Begin VB.Label Label1 
         Caption         =   "特殊商標 :"
         Height          =   255
         Index           =   22
         Left            =   120
         TabIndex        =   176
         Top             =   4359
         Width           =   1335
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "申請地址1(日) :"
         Height          =   180
         Index           =   52
         Left            =   -74910
         TabIndex        =   173
         Top             =   2086
         Width           =   1200
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "申請地址1(英) :"
         Height          =   180
         Index           =   51
         Left            =   -74910
         TabIndex        =   172
         Top             =   1733
         Width           =   1200
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "申請地址1(中) :"
         Height          =   180
         Index           =   50
         Left            =   -74910
         TabIndex        =   171
         Top             =   1380
         Width           =   1200
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "申請人1 :"
         Height          =   180
         Index           =   45
         Left            =   -74910
         TabIndex        =   170
         Top             =   375
         Width           =   720
      End
      Begin VB.Label Label1 
         Caption         =   "代理人收款後辦案 :            FCT註冊費自動代繳:         (Y:自動代繳)"
         Height          =   260
         Index           =   38
         Left            =   -74910
         TabIndex        =   169
         Top             =   3720
         Width           =   5990
      End
      Begin VB.Label Label1 
         Caption         =   "代理人備註 :"
         Height          =   260
         Index           =   37
         Left            =   -74910
         TabIndex        =   168
         Top             =   3060
         Width           =   1100
      End
      Begin VB.Label Label1 
         Caption         =   "案件備註 :"
         Height          =   260
         Index           =   36
         Left            =   -74880
         TabIndex        =   167
         Top             =   2190
         Width           =   1100
      End
      Begin VB.Label Label1 
         Caption         =   "放棄專用權 :"
         Height          =   260
         Index           =   35
         Left            =   -72330
         TabIndex        =   166
         Top             =   4650
         Width           =   1220
      End
      Begin VB.Label Label1 
         Caption         =   "副本聯絡人 :"
         Height          =   260
         Index           =   30
         Left            =   -74880
         TabIndex        =   164
         Top             =   1940
         Width           =   1100
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "申請/翻譯折扣 :          ( % )"
         Height          =   180
         Index           =   28
         Left            =   -72120
         TabIndex        =   163
         Top             =   1080
         Width           =   2090
      End
      Begin VB.Label Label1 
         Caption         =   "客戶案件案號 :"
         Height          =   240
         Index           =   26
         Left            =   -70890
         TabIndex        =   162
         Top             =   760
         Width           =   1460
      End
      Begin VB.Label Label1 
         Caption         =   "監視系統案號 :"
         Height          =   260
         Index           =   31
         Left            =   -70870
         TabIndex        =   160
         Top             =   1940
         Width           =   1460
      End
      Begin VB.Label Label1 
         Caption         =   "副本收受人 :"
         Height          =   260
         Index           =   29
         Left            =   -74880
         TabIndex        =   159
         Top             =   1620
         Width           =   1100
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "全部折扣 :                ( % )"
         Height          =   180
         Index           =   27
         Left            =   -74880
         TabIndex        =   158
         Top             =   1080
         Width           =   1980
      End
      Begin VB.Label Label1 
         Caption         =   "分所案號 :"
         Height          =   260
         Index           =   25
         Left            =   -74880
         TabIndex        =   157
         Top             =   750
         Width           =   860
      End
      Begin VB.Label Label1 
         Caption         =   "商品組群 :"
         Height          =   260
         Index           =   24
         Left            =   -74880
         TabIndex        =   156
         Top             =   270
         Width           =   1100
      End
      Begin VB.Label Label1 
         Caption         =   "閉卷原因 :"
         Height          =   255
         Index           =   21
         Left            =   3780
         TabIndex        =   153
         Top             =   4380
         Width           =   975
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "是否閉卷 :                       ( Y:閉卷 )"
         Height          =   180
         Index           =   19
         Left            =   3780
         TabIndex        =   152
         Top             =   4125
         Width           =   2580
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "是否有爭議程序 :           ( Y:有 )"
         Height          =   180
         Index           =   17
         Left            =   3780
         TabIndex        =   151
         Top             =   3555
         Width           =   2400
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "專用權是否存在 :           ( Y/N)"
         Height          =   180
         Index           =   15
         Left            =   3780
         TabIndex        =   150
         Top             =   3285
         Width           =   2295
      End
      Begin VB.Label Label1 
         Caption         =   "正商標號數 :"
         Height          =   255
         Index           =   13
         Left            =   4290
         TabIndex        =   149
         Top             =   2970
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "審定來函日 :"
         Height          =   255
         Index           =   11
         Left            =   4290
         TabIndex        =   148
         Top             =   2685
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "審定號 :"
         Height          =   255
         Index           =   9
         Left            =   4290
         TabIndex        =   147
         Top             =   2415
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "申請案號 :"
         Height          =   255
         Index           =   7
         Left            =   4290
         TabIndex        =   146
         Top             =   2130
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "商標種類 :"
         Height          =   255
         Index           =   2
         Left            =   4350
         TabIndex        =   145
         Top             =   856
         Width           =   975
      End
      Begin VB.Line Line1 
         X1              =   2610
         X2              =   2730
         Y1              =   3060
         Y2              =   3060
      End
      Begin VB.Label Label1 
         Caption         =   "閉卷日期 :"
         Height          =   255
         Index           =   20
         Left            =   120
         TabIndex        =   142
         Top             =   4081
         Width           =   1335
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "定稿語文 :                        ( 1:中 2:英 3:日 )"
         Height          =   180
         Index           =   18
         Left            =   120
         TabIndex        =   141
         Top             =   3840
         Width           =   3135
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "是否有救濟程序 :             ( Y:有 )"
         Height          =   180
         Index           =   16
         Left            =   120
         TabIndex        =   140
         Top             =   3562
         Width           =   2490
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "目前准駁 :                         ( 1:准  2:駁 )"
         Height          =   180
         Index           =   14
         Left            =   120
         TabIndex        =   139
         Top             =   3284
         Width           =   2865
      End
      Begin VB.Label Label1 
         Caption         =   "專用期限 :"
         Height          =   252
         Index           =   12
         Left            =   120
         TabIndex        =   138
         Top             =   2970
         Width           =   852
      End
      Begin VB.Label Label1 
         Caption         =   "發證日 :"
         Height          =   252
         Index           =   10
         Left            =   120
         TabIndex        =   137
         Top             =   2692
         Width           =   732
      End
      Begin VB.Label Label1 
         Caption         =   "公告日 :"
         Height          =   252
         Index           =   8
         Left            =   120
         TabIndex        =   136
         Top             =   2414
         Width           =   972
      End
      Begin VB.Label Label1 
         Caption         =   "申請日 :"
         Height          =   252
         Index           =   6
         Left            =   120
         TabIndex        =   135
         Top             =   2136
         Width           =   1092
      End
      Begin VB.Label Label1 
         Caption         =   "商品類別 :"
         Height          =   252
         Index           =   5
         Left            =   120
         TabIndex        =   125
         Top             =   1823
         Width           =   1092
      End
      Begin VB.Label Label1 
         Caption         =   "案件名稱 :"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   124
         Top             =   1190
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "申請國家 :"
         Height          =   252
         Index           =   3
         Left            =   120
         TabIndex        =   93
         Top             =   856
         Width           =   972
      End
      Begin VB.Label Label1 
         Caption         =   "卷宗性質 :"
         Height          =   252
         Index           =   1
         Left            =   120
         TabIndex        =   92
         Top             =   594
         Width           =   972
      End
      Begin VB.Label Label1 
         Caption         =   "本所案號 :"
         Height          =   252
         Index           =   0
         Left            =   120
         TabIndex        =   91
         Top             =   316
         Width           =   972
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "商品及服務"
      Height          =   285
      Left            =   4560
      TabIndex        =   190
      Top             =   570
      Width           =   1395
   End
   Begin VB.CommandButton Command3 
      Caption         =   "已設定代表圖"
      Enabled         =   0   'False
      Height          =   285
      Left            =   3030
      Style           =   1  '圖片外觀
      TabIndex        =   189
      Top             =   570
      Width           =   1515
   End
   Begin VB.CommandButton Command1 
      Caption         =   "分割案"
      Height          =   285
      Left            =   5970
      TabIndex        =   191
      Top             =   570
      Width           =   945
   End
   Begin VB.CommandButton ButtonPrior 
      Caption         =   "優先權"
      Height          =   285
      Left            =   6930
      TabIndex        =   192
      Top             =   570
      Width           =   945
   End
   Begin VB.CommandButton ButtonRelation 
      Caption         =   "相關卷號"
      Height          =   285
      Left            =   7896
      TabIndex        =   193
      Top             =   570
      Width           =   1212
   End
   Begin MSComctlLib.ImageList ImgList 
      Left            =   7860
      Top             =   30
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
            Picture         =   "frm020501.frx":016D
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm020501.frx":0489
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm020501.frx":07A5
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm020501.frx":0981
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm020501.frx":0C9D
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm020501.frx":0FB9
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm020501.frx":12D5
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm020501.frx":15F1
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm020501.frx":190D
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm020501.frx":1C29
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm020501.frx":1F45
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbar 
      Align           =   1  '對齊表單上方
      Height          =   576
      Left            =   0
      TabIndex        =   175
      Top             =   0
      Width           =   9216
      _ExtentX        =   16256
      _ExtentY        =   1016
      ButtonWidth     =   1101
      ButtonHeight    =   974
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImgList"
      DisabledImageList=   "ImgList"
      HotImageList    =   "ImgList"
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
            Enabled         =   0   'False
            Caption         =   "確定"
            Key             =   "keyOk"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
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
   End
End
Attribute VB_Name = "frm020501"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2025/10/23 TF基礎案號(TM06,TM07)改成可以輸入多筆(Table: TFBaseNo)，原本的輸入欄位直接刪除改成按鈕呼叫其他表單，若已有設定則按鈕設為綠色。
'Memo by Lydia 2021/11/29 改成Form2.0 ; 所有TextBox、cboContact
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/25 員工編號欄已修改
'Modify By Sindy 2010/7/26 日期欄已修改
Option Explicit

'Modify By Cheng 2003/09/23
'Const MAX_FIELD = 68
'Const MAX_FIELD = 70
'Const MAX_FIELD = 72
'edit by nickc 2006/07/12
'Const MAX_FIELD = T_TM
Dim MAX_FIELD
' 宣告欄位內容結構
Private Type FIELDITEM
   fiName As String
   fiOldData As String
   fiNewData As String
   fiType As Integer
End Type
'edit by nickc 2006/07/12
'Dim m_FieldList(MAX_FIELD) As FIELDITEM
Dim m_FieldList() As FIELDITEM
' 變數宣告區
'edit by nickc 2006/12/06 改可以外部傳
'Dim m_EditMode As Integer
Public m_EditMode As Integer
Dim m_SubMode As Integer
' 辦識其為外商還是內商的程式
' 0 表內商
' 1 表外商
Dim m_SysKind As Integer
' 第一筆資料的本所案號
Dim m_FirstTM(4) As String
' 最後一筆資料的本所案號
Dim m_LastTM(4) As String
' 目前正在顯示的本所案號
Dim m_CurrTM(4) As String
' 90.07.13 modify by louis (執行各項功能的權限)
Dim m_bInsert As Boolean
Dim m_bUpdate As Boolean
Dim m_bDelete As Boolean
Dim m_bQuery As Boolean
'Add By Cheng 2003/08/07
Dim m_TM23 As String '記錄原申請人
'add by nick 2004/10/05 檢查是否已經有商品及服務
Public ChkTG As Boolean
'add by nickc 206/12/11
Dim m_TM78 As String
Dim m_TM79 As String
Dim m_TM80 As String
Dim m_TM81 As String
Dim m_TM44 As String 'Added by Lydia 2024/06/13
'add by nickc 2008/02/05
Dim IsCreate716 As Boolean
Dim IsCreate102 As Boolean
Dim IsCreate105 As Boolean
Dim IsCreate105Before As Boolean  'add by sonia 2021/9/22發證前之使用宣誓期限
Dim Is105OK As Boolean            'add by sonia 2023/5/31使用宣誓國家檢查發證日
Dim m_716CP06 As String
Dim m_716CP07 As String
Dim m_716NP08 As String
Dim m_716NP09 As String
Dim m_102CP06 As String
Dim m_102CP07 As String
Dim m_102NP08 As String
Dim m_102NP09 As String
Dim m_105CP06 As String
Dim m_105CP07 As String
Dim m_105NP08 As String
Dim m_105NP09 As String
'add by sonia 2021/9/24發證前之使用宣誓期限
Dim m_105CP06Before As String
Dim m_105CP07Before As String
Dim m_105NP08Before As String
Dim m_105NP09Before As String
'end 2021/9/24
Dim m_716Key As String
Dim m_102Key As String
Dim m_105Key As String
Dim m_105KeyBefore As String   'add by sonia 2021/9/24發證前之使用宣誓期限
Dim m_716tmpDate1 As String
Dim m_716tmpDate2 As String
Dim m_102tmpDate1 As String
Dim m_102tmpDate2 As String
Dim m_105tmpDate1 As String
Dim m_105tmpDate2 As String
Dim m_105tmpDate1Before As String   'add by sonia 2021/9/24發證前之使用宣誓期限
Dim m_105tmpDate2Before As String   'add by sonia 2021/9/24發證前之使用宣誓期限
Dim strNP22 As String
Dim strCP09 As String
Dim m_form As Form 'Add By Sindy 2012/6/1
Dim strChkCuAreaMail As String, strChkCuAreaMailTo As String 'Added by Lydia 2017/06/19 檢查收文智權人員和客戶智權人員不同業務區要發mail的內文和通知人員
Dim m_ContactList As String 'Added by Lydia 2021/11/29
Dim bolUpdClose As Boolean 'Added by Lydia 2022/09/13 是否可以更新閉卷欄位
Dim strPTM As String, strSPT As String 'Added by Lydia 2023/11/16 暫存商標種類及特殊商標的Combo.ItemData

' 設定顯示的本所案號
'Modify By Sindy 2012/6/1 +strForm
Public Sub SetCurrKey(Optional ByVal strKEY01 As String = Empty, Optional ByVal strKEY02 As String = Empty, Optional ByVal strKEY03 As String = Empty, Optional ByVal strKEY04 As String = Empty, Optional ByRef strForm As Form)
   If IsEmptyText(strKEY01) Or IsEmptyText(strKEY02) Then
      m_CurrTM(0) = Empty
      m_CurrTM(1) = Empty
      m_CurrTM(2) = Empty
      m_CurrTM(3) = Empty
      Exit Sub
   End If
   m_CurrTM(0) = strKEY01
   m_CurrTM(1) = strKEY02
   m_CurrTM(2) = strKEY03
   If IsEmptyText(m_CurrTM(2)) Then
      m_CurrTM(2) = "0"
   End If
   m_CurrTM(3) = strKEY04
   If IsEmptyText(m_CurrTM(3)) Then
      m_CurrTM(3) = "00"
   End If
   'Add By Sindy 2012/6/1
   If strForm Is Nothing = False Then
      Set m_form = strForm
      m_form.Enabled = False
   End If
   '2012/6/1 End
End Sub

Private Sub RefreshRange()
Dim strSql As String
Dim rsTmp As New ADODB.Recordset
   
   ' 設定 Query 的命令
   strSql = "SELECT TM01,TM02,TM03,TM04 FROM TRADEMARK " & _
            "WHERE TM01 = '" & textTM01 & "' AND " & _
                  "TM02 = (SELECT MIN(TM02) FROM TRADEMARK WHERE TM01 = '" & textTM01 & "') AND " & _
                  "TM03 = (SELECT MIN(TM03) FROM TRADEMARK WHERE TM01 = '" & textTM01 & "' AND TM02 = (SELECT MIN(TM02) FROM TRADEMARK WHERE TM01 = '" & textTM01 & "' )) AND " & _
                  "TM04 = (SELECT MIN(TM04) FROM TRADEMARK WHERE TM01 = '" & textTM01 & "' AND TM02 = (SELECT MIN(TM02) FROM TRADEMARK WHERE TM01 = '" & textTM01 & "' ) AND TM03 = (SELECT MIN(TM03) FROM TRADEMARK WHERE TM01 = '" & textTM01 & "' AND TM02 = (SELECT MIN(TM02) FROM TRADEMARK WHERE TM01 = '" & textTM01 & "' ))) "
   
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("TM01")) = False Then: m_FirstTM(0) = rsTmp.Fields("TM01")
      If IsNull(rsTmp.Fields("TM02")) = False Then: m_FirstTM(1) = rsTmp.Fields("TM02")
      If IsNull(rsTmp.Fields("TM03")) = False Then: m_FirstTM(2) = rsTmp.Fields("TM03")
      If IsNull(rsTmp.Fields("TM04")) = False Then: m_FirstTM(3) = rsTmp.Fields("TM04")
   End If
   rsTmp.Close

   ' 設定 Query 的命令
   strSql = "SELECT TM01,TM02,TM03,TM04 FROM TRADEMARK " & _
            "WHERE TM01 = '" & textTM01 & "' AND " & _
                  "TM02 = (SELECT MAX(TM02) FROM TRADEMARK WHERE TM01 = '" & textTM01 & "') AND " & _
                  "TM03 = (SELECT MAX(TM03) FROM TRADEMARK WHERE TM01 = '" & textTM01 & "' AND TM02 = (SELECT MAX(TM02) FROM TRADEMARK WHERE TM01 = '" & textTM01 & "' )) AND " & _
                  "TM04 = (SELECT MAX(TM04) FROM TRADEMARK WHERE TM01 = '" & textTM01 & "' AND TM02 = (SELECT MAX(TM02) FROM TRADEMARK WHERE TM01 = '" & textTM01 & "' ) AND TM03 = (SELECT MAX(TM03) FROM TRADEMARK WHERE TM01 = '" & textTM01 & "' AND TM02 = (SELECT MAX(TM02) FROM TRADEMARK WHERE TM01 = '" & textTM01 & "' ))) "
                  
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("TM01")) = False Then: m_LastTM(0) = rsTmp.Fields("TM01")
      If IsNull(rsTmp.Fields("TM02")) = False Then: m_LastTM(1) = rsTmp.Fields("TM02")
      If IsNull(rsTmp.Fields("TM03")) = False Then: m_LastTM(2) = rsTmp.Fields("TM03")
      If IsNull(rsTmp.Fields("TM04")) = False Then: m_LastTM(3) = rsTmp.Fields("TM04")
   End If
   rsTmp.Close
  
   Set rsTmp = Nothing
End Sub

' 設定其為外商還是內商的系統
' Input : nSys 系統類別
'         0 ==> 內商
'         1 ==> 外商
Public Sub SetSystem(ByVal nSys As Integer)
   If nSys = 1 Then
      m_SysKind = 1
   Else
      m_SysKind = 0
   End If
End Sub

' 優先權資料
Private Sub ButtonPrior_Click()
' 優先權畫面所使用的變數
Dim strPA(1 To 4) As String '本所案號
'Dim strPriority(1 To 5) As String 'Modify by Amy 2014/04/08 +pd08,pd09
Dim strPriority(1 To 6) As String 'Modify By Sindy 2017/9/29 +pd10
Dim intOrgPWhere As Integer
    
    'Add By Cheng 2003/04/10
    '外專系統優先權日為西元格式
    intOrgPWhere = intPWhere
    If m_SysKind = 1 Then
        intPWhere = 1
    Else
        intPWhere = 0
    End If
    ' 讀取優先權資料
    strPA(1) = m_CurrTM(0)
    strPA(2) = m_CurrTM(1)
    strPA(3) = m_CurrTM(2)
    strPA(4) = m_CurrTM(3)
    
    'Modify by Amy 2014/04/08 +pd08,pd09
    'edit by nickc 2007/02/02 不用 dll 了
    'objPublicData.ReadPriority strPA, strPriority(1), strPriority(2), strPriority(3)
    'ClsPDReadPriority strPA, strPriority(1), strPriority(2), strPriority(3), strPriority(4), strPriority(5)
    'Modify By Sindy 2017/10/12 + , strPriority(6)
    ClsPDReadPriority strPA, strPriority(1), strPriority(2), strPriority(3), strPriority(4), strPriority(5), strPriority(6)
    
    ' 修改優先權資料
    'ModifyPriority strPriority(1), strPriority(2), strPriority(3), , , , , , strPriority(4), strPriority(5)
    'Modify By Sindy 2017/10/12 + , strPriority(6)
    'Modify by Sindy 2019/1/23 + strPA(1) & strPA(2) & strPA(3) & strPA(4)
    ModifyPriority strPriority(1), strPriority(2), strPriority(3), , , strPA(1) & strPA(2) & strPA(3) & strPA(4), , , strPriority(4), strPriority(5), strPriority(6)
    
    ' 儲存優先權資料
    'edit by nickc 2007/02/02 不用 dll 了
    'objPublicData.SavePriority strPA, strPriority(1), strPriority(2), strPriority(3)
    'ClsPDSavePriority strPA, strPriority(1), strPriority(2), strPriority(3), strPriority(4), strPriority(5)
    'Modify By Sindy 2017/10/12 + , strPriority(6)
    ClsPDSavePriority strPA, strPriority(1), strPriority(2), strPriority(3), strPriority(4), strPriority(5), strPriority(6)
    'end 2014/04/08
    
    'Add By Cheng 2003/04/10
    '還原由何系統進入變數值
    intPWhere = intOrgPWhere
End Sub

' 相關卷號按紐
Private Sub ButtonRelation_Click()
Dim strTM01 As String
Dim strTM02 As String
Dim strTM03 As String
Dim strTM04 As String
Dim strTit As String
Dim strMsg As String
Dim nResponse
   
   strTM01 = textTM01
   strTM02 = textTM02_1
   If textTM01 = "TF" Then: strTM02 = strTM02 & textTM02_2
   strTM03 = textTM03
   If IsEmptyText(strTM03) = True Then: strTM03 = "0"
   strTM04 = textTM04
   If IsEmptyText(strTM04) = True Then: strTM04 = "00"
   
   If IsEmptyText(textTM01) = True Or IsEmptyText(strTM02) = True Then
      strTit = "檢核資料"
      strMsg = "請先輸入本所案號"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
   Else
      Where1103ComeFrom Me, strTM01, strTM02, strTM03, strTM04
   End If
End Sub

'Added by Lydia 2016/11/24 各項指示
Private Sub cmdIns_Click()
   If m_CurrTM(0) = "" Or m_CurrTM(1) = "" Then
      MsgBox "請輸入本所案號", vbInformation
      Exit Sub
   End If
   'Added by Lydia 2020/05/05
   If m_EditMode <> 0 And m_EditMode <> 4 Then
      MsgBox IIf(m_EditMode = 1, "新增中", "修改中") & "不可執行！", vbInformation
      Exit Sub
   End If
   'end 2020/05/05
   'Added by Lydia 2020/05/05 各項指示：檢查表單是否開啟中
   If PUB_CheckFormExist("frm12040159") Then
       MsgBox "請先關閉〔申請人/代理人/案件各項指示資料〕的畫面！", vbInformation
       Exit Sub
   End If
   'end 2020/05/05
   
   frm12040159.SetParent "E", Trim(m_CurrTM(0) & m_CurrTM(1) & m_CurrTM(2) & m_CurrTM(3)), Me
   frm12040159.Show
End Sub

'Add By Sindy 2016/11/23
Private Sub Combo4_Click()
   Call GetCurrType
End Sub
Private Sub Combo4_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub
Private Sub Combo4_Validate(Cancel As Boolean)
   If Combo4 = MsgText(601) Then
      Combo4.Tag = Combo4.Text
      Combo5.ListIndex = 0
      Combo5.Enabled = False
      Exit Sub
   End If
   If ExistCheck("acc1y0", "a1y01", Combo4, Label11(1)) = False Then
      Cancel = True
      Combo4.SetFocus
   End If
   If Combo4 <> "USD" Then
      If ExistCheck("DebitNoteRate", "DNR01", Combo4, Label11(1) & "匯率") = False Then
         Cancel = True
         Combo4.SetFocus
         Exit Sub
      End If
   End If
   Call GetCurrType
End Sub
Private Sub GetCurrType()
Dim intType As Integer
   
   If Combo4 = MsgText(601) Then
      Combo4.Tag = Combo4.Text
      Combo5.ListIndex = 0
      Combo5.Enabled = False
      Exit Sub
   End If
   '若更改請款幣別
   If Me.Combo4.Text <> Me.Combo4.Tag Then
      Me.Combo4.Tag = Me.Combo4.Text
      '請款幣別變更要重新預設列印幣別
      '台幣
      If Me.Combo4.Text = "NTD" Then
         intType = 1 '純台幣
      '人民幣
      ElseIf Me.Combo4.Text = "RMB" Then
         intType = 4 '外幣+美金合計
      '其他幣別
      Else
         intType = 2 '台幣+外幣合計
      End If
      Combo5.ListIndex = intType
      '若為台幣時則格式欄位鎖住不可修改
      If Me.Combo4.Text = "NTD" Then
         Combo5.Enabled = False
      Else
         Combo5.Enabled = True
      End If
   End If
End Sub
'2016/11/23 END

Private Sub Command1_Click()
    frm02010604_3.m_CP01 = Me.textTM01.Text
    frm02010604_3.m_CP02 = Me.textTM02_1.Text & IIf(Me.textTM01.Text = "TF", Me.textTM02_2.Text, "")
    frm02010604_3.m_CP03 = IIf(Me.textTM03.Text = "", "0", Me.textTM03.Text)
    frm02010604_3.m_CP04 = IIf(Me.textTM04.Text = "", "00", Me.textTM04.Text)
    frm02010604_3.intWhereToGo = 1
    frm02010604_3.Show
End Sub

Private Sub Command2_Click()
   frm03010303_04.Hide
   Set frm03010303_04.UpForm = Me
   frm03010303_04.TGKey = textTM01 & "-" & textTM02_1 & textTM02_2 & "-" & textTM03 & "-" & textTM04
   textTM08_1_Validate False 'Add By Sindy 2015/6/30
   frm03010303_04.AllClass = textTM09.Text
   frm03010303_04.cmdOK(2).Visible = True
   
   If m_EditMode <> 1 And m_EditMode <> 2 Then
       frm03010303_04.Label2.Visible = False
       frm03010303_04.cmdOK(0).Visible = False
       frm03010303_04.cmdOK(2).Visible = False
       frm03010303_04.cmd.Visible = False
       frm03010303_04.cmd2.Visible = False
       frm03010303_04.txt2(0).Visible = False
       frm03010303_04.txt2(1).Visible = False
       frm03010303_04.txt2(2).Visible = False
       frm03010303_04.txt2(3).Visible = False
       frm03010303_04.Line1.Visible = False
   End If
   If textTM09 <> "" Then  '2010/5/10 MODIFY BY SONIA 有商品類別才可進入 T-113511團體標章
      Me.Hide
      frm03010303_04.QueryData
      frm03010303_04.Show vbModal 'Modify By Sindy 2009/09/17 改為強制回應表單
   '2010/5/10 ADD BY SONIA
   Else
      MsgBox ("無商品類別，不可使用此按鈕 !")
   End If
   '2010/5/10 END
End Sub

'Add By Sindy 2012/6/15
Private Sub Command3_Click()
   frmPic001.oCP01 = textTM01
   frmPic001.oCP02 = textTM02_1 & textTM02_2
   frmPic001.oCP03 = textTM03
   frmPic001.oCP04 = textTM04
   frmPic001.StrMenu
   frmPic001.CanScan
   frmPic001.SetSeekCmdok 'Add by Amy 2018/07/16
   frmPic001.Show vbModal
   '檢查有無代表圖
   'Modify by Amy 2018/07/16  改寫至function
'   strSql = "SELECT ibf01,ibf02 FROM imgbytefile WHERE ibf01='" & textTM01 & "' and ibf02='" & textTM02_1 & textTM02_2 & "' and ibf03='" & textTM03 & "' and ibf04='" & textTM04 & "' and ibf05='1'"
'   CheckOC2
'   adoRecordset1.CursorLocation = adUseClient
'   adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
'   If adoRecordset1.RecordCount <> 0 And adoRecordset1.RecordCount > 0 Then
   If ChkImgByteFile(textTM01, textTM02_1 & textTM02_2, textTM03, textTM04) = True Then
       'Modified by Lydia 2021/12/16 拿掉快速鍵(&I)
       Command3.Caption = "已設定代表圖"
       Command3.BackColor = &HC0FFC0
   Else
       'Modified by Lydia 2021/12/16 拿掉快速鍵(&I)
       Command3.Caption = "未設定代表圖"
       Command3.BackColor = &HC0C0FF
   End If
'   CheckOC2
   'end 2018/07/16
End Sub

Private Sub Form_Initialize()
   'add by nickc 2006/07/12
   MAX_FIELD = TF_TM
   ReDim m_FieldList(MAX_FIELD) As FIELDITEM
End Sub

'add by nickc 2006/11/10 Enter 事件，等於存檔，做完取消，不然 form 內其他物件有寫 keycode 或是 keyascii 事件的話，也會做到
Private Sub Form_KeyPress(KeyAscii As Integer)
'Remove by Lydia 2021/11/29 取消以ENTER控制為換行的功能 (Form2.0修改之維護資料功能Toolbar之修改統一)
'   Select Case KeyAscii
'      Case 13:
'         If m_EditMode <> 0 Then
'            KeyAscii = 0
'            OnAction vbKeyF9
'         End If
'   End Select
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
'Memo by Lydia 2021/11/29 從Form_KeyDown搬來
   Select Case KeyCode
      ' 90.07.13 modify by louis
      ' 新增
      'Case vbKeyF2, vbKeyF3, vbKeyF4, vbKeyF5, vbKeyHome, vbKeyPageUp, vbKeyPageDown, vbKeyEnd:
      '   If m_EditMode = 0 Then
      '      OnAction KeyCode
      '      KeyCode = 0
      '   End If
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
      ' 確定, 取消
      Case vbKeyF9, vbKeyF10:
         If m_EditMode <> 0 Then
            OnAction KeyCode
            KeyCode = 0
         End If
'edit by nickc 2006/11/10
'      ' 確定
'      Case vbKeyReturn:
'         If m_EditMode <> 0 Then
'            OnAction vbKeyF9
'         End If
      ' 取消或離開
      Case vbKeyEscape:
         If m_EditMode = 0 Then
            OnAction KeyCode
         Else
            OnAction vbKeyF10
         End If
   End Select
   
End Sub

' Load Form
Private Sub Form_Load()
   ' 90.07.13 modify by louis (取得使用者執行各項功能的權限)
   m_bInsert = IsUserHasRightOfFunction("frm020501", strAdd, False)
   m_bUpdate = IsUserHasRightOfFunction("frm020501", strEdit, False)
   m_bDelete = IsUserHasRightOfFunction("frm020501", strDel, False)
   m_bQuery = IsUserHasRightOfFunction("frm020501", strFind, False)
   
   textTM08_2.BackColor = &H8000000F
   textTM10_2.BackColor = &H8000000F
   textTM28.BackColor = &H8000000F
   textTM31_2.BackColor = &H8000000F
   textSP32.BackColor = &H8000000F
   textTM54_2.BackColor = &H8000000F
   textTM56_2.BackColor = &H8000000F
   textFA29.BackColor = &H8000000F
   textFA39.BackColor = &H8000000F
   textCU72.BackColor = &H8000000F
   textCU79.BackColor = &H8000000F
   textTM23_2.BackColor = &H8000000F
   textTM44_2.BackColor = &H8000000F
   textTM33_2.BackColor = &H8000000F
   textTM66_2.BackColor = &H8000000F
   textTM69_2.BackColor = &H8000000F
   textTM70_2.BackColor = &H8000000F
   'Added by Lydia 2021/12/15
   textTM57.BackColor = &H8000000F
   textTM73.BackColor = &H8000000F
   textTM74.BackColor = &H8000000F
   textTM75.BackColor = &H8000000F
   
   'add by nickc 2006/12/08
   textTM78_2.BackColor = &H8000000F
   textTM79_2.BackColor = &H8000000F
   textTM80_2.BackColor = &H8000000F
   textTM81_2.BackColor = &H8000000F
   
   textCUID_1.BackColor = &H8000000F
   textCUID_2.BackColor = &H8000000F
   
   'Added by Lydia 2016/09/10 設定代表人中文名稱和英文名稱長度
    textTM47.MaxLength = Pub_MaxCEL10
    textTM48.MaxLength = Pub_MaxCEL11
    textTM50.MaxLength = Pub_MaxCEL10
    textTM51.MaxLength = Pub_MaxCEL11
    textTM94.MaxLength = Pub_MaxCEL10
    textTM95.MaxLength = Pub_MaxCEL11
    textTM97.MaxLength = Pub_MaxCEL10
    textTM98.MaxLength = Pub_MaxCEL11
    textTM100.MaxLength = Pub_MaxCEL10
    textTM101.MaxLength = Pub_MaxCEL11
    textTM103.MaxLength = Pub_MaxCEL10
    textTM104.MaxLength = Pub_MaxCEL11
    textTM106.MaxLength = Pub_MaxCEL10
    textTM107.MaxLength = Pub_MaxCEL11
    textTM109.MaxLength = Pub_MaxCEL10
    textTM110.MaxLength = Pub_MaxCEL11
    textTM112.MaxLength = Pub_MaxCEL10
    textTM113.MaxLength = Pub_MaxCEL11
    textTM115.MaxLength = Pub_MaxCEL10
    textTM116.MaxLength = Pub_MaxCEL11
   'end 2016/09/10
   
   m_EditMode = 0
   m_SubMode = 0
   MoveFormToCenter Me
   
   InitialField
   
   If m_SysKind = 0 Then
      textTM01 = "T"
   Else
      '92.4.25 modify by sonia
      'textTM01 = "FCT"
      textTM01 = ""
   End If
   
   ' 90.10.18 modify by louis (程式進入時不去找第一筆及最後一筆的Key)
   'RefreshRange
   'ShowFirstRecord
   'UpdateToolbarState
   'SetCtrlReadOnly True
   
   If Not IsEmptyText(m_CurrTM(0)) And Not IsEmptyText(m_CurrTM(1)) And Not IsEmptyText(m_CurrTM(2)) And Not IsEmptyText(m_CurrTM(3)) Then
      ShowCurrRecord m_CurrTM(0), m_CurrTM(1), m_CurrTM(2), m_CurrTM(3)
      UpdateToolbarState
      SetCtrlReadOnly True
   Else
      m_EditMode = 4
      SetCtrlReadOnly True
      SetKeyReadOnly False
      UpdateToolbarState
   End If
   
   'Add By Sindy 2016/11/23
   '抓有輸入過匯率的請款幣別
   Combo4.Clear
   Combo4.AddItem ""
   Combo4.AddItem "USD"
   If RsTemp.State <> adStateClosed Then RsTemp.Close
   RsTemp.CursorLocation = adUseClient
   RsTemp.Open "select distinct DNR01 from DebitNoteRate order by DNR01 asc", adoTaie, adOpenStatic, adLockReadOnly
   Do While RsTemp.EOF = False
      Combo4.AddItem RsTemp.Fields("DNR01").Value
      RsTemp.MoveNext
   Loop
   RsTemp.Close
   '2016/11/23 End
   
   'Added by Lydia 2020/05/05 各項指示：顯示按鈕
   If strSrvDate(1) >= 各項指示啟用日 Then
      cmdIns.Visible = True
   Else
      cmdIns.Visible = False
   End If
   'end 2020/05/05
       
   'Added by Lydia 2022/09/13 外商承辦F11判斷人員職稱等級決定是否鎖住「閉卷」
   bolUpdClose = True
   If Pub_StrUserSt03 = "F11" Then
       strExc(0) = "select nvl(st20,'99') st20 from staff where st01='" & strUserNum & "' "
       intI = 1
       Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
       If intI = 1 Then
           If "" & RsTemp.Fields("st20") > "52" Then
               bolUpdClose = False
           End If
       End If
   End If
   'end 2022/09/13
   
   'Added by Lydia 2023/11/16 內外商之分案及商標基本資料維護之商標種類、特殊商標欄位增加下拉功能
   Pub_SetTMcombo "1", cboTM08, , , strPTM '商標種類
   Pub_SetTMcombo "2", cboTM72, , , strSPT '特殊商標種類
   'end 2023/11/16
'   'Add by Amy 2024/03/08 隱藏延展單筆不跑,將不催延展移位
'   'Modify By Sindy 2024/6/18 改用Frame因物件(Label1(116)、textTM129)會出現在其他頁籤中
''   Label1(116).Left = 4560
''   textTM129.Left = 5490
'   Frame1.Left = 4560
   Frame1.BorderStyle = 0 '無線
'   '2024/6/18 END
    
  
   tabCtrl.Tab = 0
End Sub

' 初始化欄位陣列
Private Sub InitialField()
Dim nIndex As Integer
Dim strTmp As String
   
   ' 初始化欄位陣列
   For nIndex = 1 To MAX_FIELD
      strTmp = Format(nIndex, "00")
      m_FieldList(nIndex - 1).fiName = "TM" & strTmp
      m_FieldList(nIndex - 1).fiOldData = Empty
      m_FieldList(nIndex - 1).fiNewData = Empty
      m_FieldList(nIndex - 1).fiType = 0
      Select Case nIndex
         Case 11, 13, 14, 20, 21, 22, 30, 36, 37, 60, 61, 63, 64:
            m_FieldList(nIndex - 1).fiType = 1
      End Select
   Next nIndex
End Sub

' 設定欄位的內容
Private Sub SetFieldNewData(ByVal strName As String, Optional ByVal strData As String = "#==#")
Dim nIndex As Integer
   
   For nIndex = 0 To MAX_FIELD - 1
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

' 更新欄位的內容
Private Sub UpdateFieldNewData()
   SetFieldNewData "TM01", textTM01
   SetFieldNewData "TM02", textTM02_1 & textTM02_2
   SetFieldNewData "TM03", textTM03 & String(1 - Len(textTM03), "0")
   SetFieldNewData "TM04", textTM04 & String(2 - Len(textTM04), "0")
   SetFieldNewData "TM05", textTM05
   SetFieldNewData "TM131", textTM131 'Add By Sindy 2015/6/30
   'Add By Sindy 2024/6/14
   SetFieldNewData "TM137", textTM137
   SetFieldNewData "TM138", textTM138
   SetFieldNewData "TM139", textTM139
   '2024/6/14 END
   SetFieldNewData "TM08", textTM08_1: SetFieldNewData "TM09", textTM09: SetFieldNewData "TM10", textTM10_1
   If IsEmptyText(textTM11) = False Then
      SetFieldNewData "TM11", DBDATE(textTM11)
   Else
      SetFieldNewData "TM11", textTM11
   End If
   SetFieldNewData "TM12", textTM12
   If IsEmptyText(textTM13) = False Then
      SetFieldNewData "TM13", DBDATE(textTM13)
   Else
      SetFieldNewData "TM13", textTM13
   End If
   If IsEmptyText(textTM14) = False Then
      SetFieldNewData "TM14", DBDATE(textTM14)
   Else
      SetFieldNewData "TM14", textTM14
   End If
   SetFieldNewData "TM15", textTM15
   SetFieldNewData "TM16", textTM16: SetFieldNewData "TM17", textTM17: SetFieldNewData "TM18", textTM18: SetFieldNewData "TM19", textTM19
   If IsEmptyText(textTM20) = False Then
      SetFieldNewData "TM20", DBDATE(textTM20)
   Else
      SetFieldNewData "TM20", textTM20
   End If
   If IsEmptyText(textTM21) = False Then
      SetFieldNewData "TM21", DBDATE(textTM21)
   Else
      SetFieldNewData "TM21", textTM21
   End If
   If IsEmptyText(textTM22) = False Then
      SetFieldNewData "TM22", DBDATE(textTM22)
   Else
      SetFieldNewData "TM22", textTM22
   End If
   If IsEmptyText(textTM23_1) = False Then
      SetFieldNewData "TM23", textTM23_1 & String(9 - Len(textTM23_1), "0")
   Else
      SetFieldNewData "TM23", textTM23_1
   End If
   SetFieldNewData "TM24", textTM24: SetFieldNewData "TM25", textTM25
   SetFieldNewData "TM26", textTM26: SetFieldNewData "TM27", textTM27: SetFieldNewData "TM29", textTM29
   If IsEmptyText(textTM30) = False Then
      SetFieldNewData "TM30", DBDATE(textTM30)
   Else
      SetFieldNewData "TM30", textTM30
   End If
   SetFieldNewData "TM31", textTM31_1
   SetFieldNewData "TM32", textTM32
   If IsEmptyText(textTM33_1) = False Then
      SetFieldNewData "TM33", textTM33_1 & String(9 - Len(textTM33_1), "0")
   Else
      SetFieldNewData "TM33", textTM33_1
   End If
   SetFieldNewData "TM34", textTM34: SetFieldNewData "TM35", textTM35
   SetFieldNewData "TM36", textTM36: SetFieldNewData "TM37", textTM37: SetFieldNewData "TM38", textTM38: SetFieldNewData "TM39", textTM39: SetFieldNewData "TM40", textTM40
   
   SetFieldNewData "TM140", textTM140: SetFieldNewData "TM141", textTM141 'Add By Sindy 2025/3/6
   
   SetFieldNewData "TM41", textTM41: SetFieldNewData "TM42", textTM42: SetFieldNewData "TM43", textTM43
   If IsEmptyText(textTM44_1) = False Then
      SetFieldNewData "TM44", textTM44_1 & String(9 - Len(textTM44_1), "0")
   Else
      SetFieldNewData "TM44", textTM44_1
   End If
   SetFieldNewData "TM45", textTM45
   SetFieldNewData "TM46", textTM46: SetFieldNewData "TM47", textTM47: SetFieldNewData "TM48", textTM48: SetFieldNewData "TM49", textTM49: SetFieldNewData "TM50", textTM50
   SetFieldNewData "TM130", textTM130 'Add By Sindy 2013/12/13
   SetFieldNewData "TM51", textTM51: SetFieldNewData "TM52", textTM52: SetFieldNewData "TM53", textTM53
   If IsEmptyText(textTM54_1) = False Then
      SetFieldNewData "TM54", textTM54_1 & String(9 - Len(textTM54_1), "0")
   Else
      SetFieldNewData "TM54", textTM54_1
   End If
   SetFieldNewData "TM55", textTM55
   If IsEmptyText(textTM56_1) = False Then
      SetFieldNewData "TM56", textTM56_1 & String(9 - Len(textTM56_1), "0")
   Else
      SetFieldNewData "TM56", textTM56_1
   End If
   SetFieldNewData "TM58", textTM58
   SetFieldNewData "TM65", textTM65
   If IsEmptyText(textTM66_1) = False Then
      SetFieldNewData "TM66", textTM66_1 & String(9 - Len(textTM66_1), "0")
   Else
      SetFieldNewData "TM66", textTM66_1
   End If
   SetFieldNewData "TM67", textTM67: SetFieldNewData "TM68", textTM68
   ' 設定資料同舊的
   'Modify By Sindy 2011/2/23 新增時預設卷宗性質TM28='1'
   If m_EditMode = 1 Then
      SetFieldNewData "TM28", "1"
   '2011/2/23 End
   Else
      SetFieldNewData "TM28"
   End If
   If IsEmptyText(textTM69_1) = False Then
      SetFieldNewData "TM69", textTM69_1 & String(9 - Len(textTM69_1), "0")
   Else
      SetFieldNewData "TM69", textTM69_1
   End If
   If IsEmptyText(textTM70_1) = False Then
      SetFieldNewData "TM70", textTM70_1 & String(9 - Len(textTM70_1), "0")
   Else
      SetFieldNewData "TM70", textTM70_1
   End If
   If IsEmptyText(textTM71_1) = False Then
      SetFieldNewData "TM71", textTM71_1
   Else
      SetFieldNewData "TM71", textTM71_1
   End If
   SetFieldNewData "TM72", textTM72_1
   
   'add by nickc 2006/12/08
   If IsEmptyText(textTM78_1) = False Then
      SetFieldNewData "TM78", textTM78_1 & String(9 - Len(textTM78_1), "0")
   Else
      SetFieldNewData "TM78", textTM78_1
   End If
   If IsEmptyText(textTM79_1) = False Then
      SetFieldNewData "TM79", textTM79_1 & String(9 - Len(textTM79_1), "0")
   Else
      SetFieldNewData "TM79", textTM79_1
   End If
   If IsEmptyText(textTM80_1) = False Then
      SetFieldNewData "TM80", textTM80_1 & String(9 - Len(textTM80_1), "0")
   Else
      SetFieldNewData "TM80", textTM80_1
   End If
   If IsEmptyText(textTM81_1) = False Then
      SetFieldNewData "TM81", textTM81_1 & String(9 - Len(textTM81_1), "0")
   Else
      SetFieldNewData "TM81", textTM81_1
   End If
   SetFieldNewData "TM82", textTM82
   SetFieldNewData "TM83", textTM83
   SetFieldNewData "TM84", textTM84
   SetFieldNewData "TM85", textTM85
   SetFieldNewData "TM86", textTM86
   SetFieldNewData "TM87", textTM87
   SetFieldNewData "TM88", textTM88
   SetFieldNewData "TM89", textTM89
   SetFieldNewData "TM90", textTM90
   SetFieldNewData "TM91", textTM91
   SetFieldNewData "TM92", textTM92
   SetFieldNewData "TM93", textTM93
   SetFieldNewData "TM94", textTM94
   SetFieldNewData "TM95", textTM95
   SetFieldNewData "TM96", textTM96
   SetFieldNewData "TM97", textTM97
   SetFieldNewData "TM98", textTM98
   SetFieldNewData "TM99", textTM99
   SetFieldNewData "TM100", textTM100
   SetFieldNewData "TM101", textTM101
   SetFieldNewData "TM102", textTM102
   SetFieldNewData "TM103", textTM103
   SetFieldNewData "TM104", textTM104
   SetFieldNewData "TM105", textTM105
   SetFieldNewData "TM106", textTM106
   SetFieldNewData "TM107", textTM107
   SetFieldNewData "TM108", textTM108
   SetFieldNewData "TM109", textTM109
   SetFieldNewData "TM110", textTM110
   SetFieldNewData "TM111", textTM111
   SetFieldNewData "TM112", textTM112
   SetFieldNewData "TM113", textTM113
   SetFieldNewData "TM114", textTM114
   SetFieldNewData "TM115", textTM115
   SetFieldNewData "TM116", textTM116
   SetFieldNewData "TM117", textTM117
   'add by nickc 2007/03/08
   SetFieldNewData "TM118", textTM118
   'add by nickc 2007/01/02
   SetFieldNewData "TM76", textTM76
   SetFieldNewData "TM121", textTM121 'Add by Morgan 2008/5/23
   'add by Toni 2008/10/21
   SetFieldNewData "TM122", textTM122
   'end 2008/10/21
   
   'Add By Sindy 2009/09/09
   SetFieldNewData "TM77", textTM77
   SetFieldNewData "TM124", textTM124
   SetFieldNewData "TM125", textTM125
   SetFieldNewData "TM126", textTM126
   '2009/09/09 End
   SetFieldNewData "TM127", textTM127 'Add by Morgan 2010/11/5
   SetFieldNewData "TM129", textTM129 'Add by Sindy 2013/8/15
   'Add By Sindy 2016/11/23
   SetFieldNewData "TM134", Combo4.Text
   SetFieldNewData "TM135", IIf(Combo5.Text <> "", Combo5.ListIndex, "")
   '2016/11/23 END
   
    SetFieldNewData "TM136", textTM136 'Added by Morgan 2022/12/1
End Sub

' 從記錄中更新欄位內容
Private Sub UpdateFieldOldData(ByRef rsTmp As ADODB.Recordset)
Dim nIndex As Integer
Dim strTmp As String
   
   For nIndex = 0 To MAX_FIELD - 1
      If m_FieldList(nIndex).fiName <> Empty Then
         If IsNull(rsTmp.Fields(m_FieldList(nIndex).fiName)) = False Then
            m_FieldList(nIndex).fiOldData = rsTmp.Fields(m_FieldList(nIndex).fiName)
            'add by nickc 2007/03/03
            m_FieldList(nIndex).fiNewData = rsTmp.Fields(m_FieldList(nIndex).fiName)
         Else
            m_FieldList(nIndex).fiOldData = Empty
            'add by nickc 2007/03/03
            m_FieldList(nIndex).fiNewData = Empty
         End If
      End If
   Next nIndex
EXITSUB:
End Sub

' 讀取資料庫所有的資料
Private Sub QueryDB()
   RefreshRange
End Sub

' 清除欄位內的資料內容
Private Sub ClearField()
Dim nIndex As Integer
   
   textTM01 = Empty: textTM02_1 = Empty: textTM02_2 = Empty: textTM03 = Empty: textTM04 = Empty: textTM05 = Empty: textTM131 = Empty
   'Add By Sindy 2024/6/14
   textTM137 = Empty: textTM138 = Empty: textTM139 = Empty
   '2024/6/14 END
   textTM08_1 = Empty: textTM08_2 = Empty:  textTM09 = Empty: textTM10_1 = Empty: textTM10_2 = Empty
   textTM11 = Empty: textTM12 = Empty: textTM13 = Empty: textTM14 = Empty: textTM15 = Empty
   textTM16 = Empty: textTM17 = Empty: textTM18 = Empty: textTM19 = Empty: textTM20 = Empty
   textTM21 = Empty: textTM22 = Empty: textTM23_1 = Empty: textTM23_2 = Empty: textTM24 = Empty: textTM25 = Empty
   textTM26 = Empty: textTM27 = Empty: textTM28 = Empty: textTM29 = Empty: textTM30 = Empty
   textTM31_1 = Empty: textTM31_2 = Empty: textTM32 = Empty: textTM33_1 = Empty: textTM33_2 = Empty: textTM34 = Empty: textTM35 = Empty
   textTM36 = Empty: textTM37 = Empty: textTM38 = Empty: textTM39 = Empty: textTM40 = Empty
   
   textTM140 = Empty: textTM141 = Empty 'Add By Sindy 2025/3/6
   
   textTM41 = Empty: textTM42 = Empty: textTM43 = Empty: textTM44_1 = Empty: textTM44_2 = Empty: textTM45 = Empty
   textTM46 = Empty: textTM47 = Empty: textTM48 = Empty: textTM49 = Empty: textTM50 = Empty
   textTM51 = Empty: textTM52 = Empty: textTM53 = Empty: textTM54_1 = Empty: textTM54_2 = Empty: textTM55 = Empty
   textTM56_1 = Empty: textTM56_2 = Empty: textTM57 = Empty: textTM58 = Empty
   textTM65 = Empty: textTM66_1 = Empty: textTM66_2 = Empty: textTM67 = Empty: textTM68 = Empty
   textCU72 = Empty: textCU79 = Empty: textFA29 = Empty: textFA39 = Empty: textTM69_1 = Empty: textTM69_2 = Empty: textTM70_1 = Empty: textTM70_2 = Empty
   textCUID_1 = Empty: textCUID_2 = Empty: textTM71_1 = Empty: textTM72_1 = Empty: textTM72_2 = Empty
   textTM130 = Empty 'Add By Sindy 2013/12/13
   cboTM08.Text = Empty: cboTM72.Text = "": cboTM08.Tag = Empty: cboTM72.Tag = "" 'Added by Lydia 2023/11/16
   
   'add by nickc 2006/07/12
   textTM73 = Empty: textTM74 = Empty: textTM75 = Empty
   
   'add by nickc 2006/12/08
   textTM78_1 = Empty: textTM78_2 = Empty: textTM79_1 = Empty: textTM79_2 = Empty
   textTM80_1 = Empty: textTM80_2 = Empty: textTM81_1 = Empty: textTM81_2 = Empty
   textTM82 = Empty: textTM83 = Empty: textTM84 = Empty: textTM85 = Empty
   textTM86 = Empty: textTM87 = Empty: textTM88 = Empty: textTM89 = Empty
   textTM90 = Empty: textTM91 = Empty: textTM92 = Empty: textTM93 = Empty
   textTM94 = Empty: textTM95 = Empty: textTM96 = Empty: textTM97 = Empty
   textTM98 = Empty: textTM99 = Empty: textTM100 = Empty: textTM101 = Empty
   textTM102 = Empty: textTM103 = Empty: textTM104 = Empty: textTM105 = Empty
   textTM106 = Empty: textTM107 = Empty: textTM108 = Empty: textTM109 = Empty
   textTM110 = Empty: textTM111 = Empty: textTM112 = Empty: textTM113 = Empty
   textTM114 = Empty: textTM115 = Empty: textTM116 = Empty: textTM117 = Empty
   'add by nickc 2007/03/08
   textTM118 = Empty
   'add by nickc 2007/01/02
   textTM76 = Empty
   textTM121 = Empty 'Add by Morgan 2008/5/23
   
   textTM122 = Empty 'add by Toni 2008/10/21
   
   'Add By Sindy 2009/09/09
   textTM77 = Empty
   textTM124 = Empty
   textTM125 = Empty
   textTM126 = Empty
   '2009/09/09 End
   textTM127 = Empty 'Add by Morgan 2010/11/5
   textTM129 = Empty 'Add by Sindy 2013/8/15
   textTM136 = Empty 'Add by Morgan 2022/12/1
   
   For nIndex = 0 To MAX_FIELD - 1
      m_FieldList(nIndex).fiOldData = Empty
      m_FieldList(nIndex).fiNewData = Empty
   Next nIndex
   
   If IsEmpty(m_CurrTM(0)) = False Then
      textTM01 = m_CurrTM(0)
   End If
   cboContact.Clear 'Add by Morgan 2008/8/4
   'Added by Lydia 2021/11/29
   cboContact.Tag = ""
   m_ContactList = ""
   
   'Add By Sindy 2016/11/23
   If Combo4.Visible = True Then
      Me.Combo4.ListIndex = 0
      Me.Combo5.ListIndex = 0
   End If
   '2016/11/23 End
End Sub

' 更新各控制項的狀態
Private Sub SetCtrlReadOnly(ByVal bEnable As Boolean)
   textTM01.Locked = bEnable
   textTM02_1.Locked = bEnable: textTM02_2.Locked = bEnable: textTM03.Locked = bEnable: textTM04.Locked = bEnable: textTM05.Locked = bEnable: textTM131.Locked = bEnable
   'Add By Sindy 2024/6/14
   textTM137.Locked = bEnable
   textTM138.Locked = bEnable
   textTM139.Locked = bEnable
   '2024/6/14 END
   textTM08_1.Locked = bEnable: textTM08_2.Locked = bEnable:  textTM09.Locked = bEnable: textTM10_1.Locked = bEnable: textTM10_2.Locked = bEnable
   textTM11.Locked = bEnable: textTM12.Locked = bEnable: textTM13.Locked = bEnable: textTM14.Locked = bEnable: textTM15.Locked = bEnable
   textTM16.Locked = bEnable: textTM17.Locked = bEnable: textTM18.Locked = bEnable: textTM19.Locked = bEnable: textTM20.Locked = bEnable
   textTM21.Locked = bEnable: textTM22.Locked = bEnable: textTM23_1.Locked = bEnable: textTM23_2.Locked = bEnable: textTM24.Locked = bEnable: textTM25.Locked = bEnable
   textTM26.Locked = bEnable: textTM27.Locked = bEnable: textTM28.Locked = bEnable: textTM29.Locked = bEnable: textTM30.Locked = bEnable
   textTM31_1.Locked = bEnable: textTM31_2.Locked = bEnable: textTM32.Locked = bEnable: textTM33_1.Locked = bEnable: textTM33_2.Locked = bEnable: textTM34.Locked = bEnable: textTM35.Locked = bEnable
   textTM36.Locked = bEnable: textTM37.Locked = bEnable: textTM38.Locked = bEnable: textTM39.Locked = bEnable: textTM40.Locked = bEnable
   
   textTM140.Locked = bEnable: textTM141.Locked = bEnable 'Add By Sindy 2025/3/6
   
   textTM41.Locked = bEnable: textTM42.Locked = bEnable: textTM43.Locked = bEnable: textTM44_1.Locked = bEnable: textTM44_2.Locked = bEnable: textTM45.Locked = bEnable
   textTM46.Locked = bEnable: textTM47.Locked = bEnable: textTM48.Locked = bEnable: textTM49.Locked = bEnable: textTM50.Locked = bEnable
   textTM51.Locked = bEnable: textTM52.Locked = bEnable: textTM53.Locked = bEnable: textTM54_1.Locked = bEnable: textTM54_2.Locked = bEnable: textTM55.Locked = bEnable
   textTM56_1.Locked = bEnable: textTM56_2.Locked = bEnable: textTM57.Locked = bEnable: textTM58.Locked = bEnable
   textTM65.Locked = bEnable: textTM66_1.Locked = bEnable: textTM66_2.Locked = bEnable: textTM67.Locked = bEnable: textTM68.Locked = bEnable
   textCU72.Locked = bEnable: textCU79.Locked = bEnable: textFA29.Locked = bEnable: textFA39.Locked = bEnable: textTM69_1.Locked = bEnable: textTM69_2.Locked = bEnable: textTM70_1.Locked = bEnable: textTM70_2.Locked = bEnable
   textTM71_1.Locked = bEnable: textTM72_1.Locked = bEnable: textTM72_2.Locked = bEnable
   cboTM08.Locked = bEnable: cboTM72.Locked = bEnable  'Added by Lydia 2023/11/16
   
   'Modify by Amy 2018/07/03 只有電腦中心才可改 特殊出名公司
   textTM130.Locked = True
   If Pub_StrUserSt03 = "M51" Then
      textTM130.Locked = bEnable 'Add By Sindy 2013/12/13
   End If
   
   'add by nickc 2006/12/08
   textTM78_1.Locked = bEnable: textTM78_2.Locked = bEnable: textTM79_1.Locked = bEnable: textTM79_2.Locked = bEnable: textTM80_1.Locked = bEnable
   textTM80_2.Locked = bEnable: textTM81_1.Locked = bEnable: textTM81_2.Locked = bEnable: textTM82.Locked = bEnable: textTM83.Locked = bEnable
   textTM84.Locked = bEnable: textTM85.Locked = bEnable: textTM86.Locked = bEnable: textTM87.Locked = bEnable: textTM88.Locked = bEnable
   textTM89.Locked = bEnable: textTM90.Locked = bEnable: textTM91.Locked = bEnable: textTM92.Locked = bEnable: textTM93.Locked = bEnable
   textTM94.Locked = bEnable: textTM95.Locked = bEnable: textTM96.Locked = bEnable: textTM97.Locked = bEnable: textTM98.Locked = bEnable
   textTM99.Locked = bEnable: textTM100.Locked = bEnable: textTM101.Locked = bEnable: textTM102.Locked = bEnable: textTM103.Locked = bEnable
   textTM104.Locked = bEnable: textTM105.Locked = bEnable: textTM106.Locked = bEnable: textTM107.Locked = bEnable: textTM108.Locked = bEnable
   textTM109.Locked = bEnable: textTM110.Locked = bEnable: textTM111.Locked = bEnable: textTM112.Locked = bEnable: textTM113.Locked = bEnable
   textTM114.Locked = bEnable: textTM115.Locked = bEnable: textTM116.Locked = bEnable: textTM117.Locked = bEnable
   'add by nickc 2007/03/08
   textTM118.Locked = bEnable
   'add by nickc 2007/01/02
   textTM76.Locked = bEnable
   textTM121.Locked = bEnable 'Add by Morgan 2008/5/23
   
   textTM122.Locked = bEnable 'add by Toni 2008/10/21
   
   'Add By Sindy 2009/09/09
   textTM77.Locked = bEnable
   textTM124.Locked = bEnable
   textTM125.Locked = bEnable
   textTM126.Locked = bEnable
   '2009/09/09 End
   textTM127.Locked = bEnable 'Add by Morgan 2010/11/5
   textTM129.Locked = bEnable 'Add by Sindy 2013/8/15
   textTM136.Locked = bEnable 'Add by Morgan 2022/12/1
   
   'Add By Sindy 2016/11/23
   Combo4.Locked = bEnable
   Combo5.Locked = bEnable
   '2016/11/23 End
End Sub

' 更新Key的狀態
Private Sub SetKeyReadOnly(ByVal bEnable As Boolean)
   textTM01.Locked = bEnable: textTM02_1.Locked = bEnable: textTM02_2.Locked = bEnable: textTM03.Locked = bEnable: textTM04.Locked = bEnable
End Sub

Private Sub SetTM72forCol(strTM72)
   If Trim(strTM72) = "" Then
      lblTM137.Visible = False
      textTM137.Visible = False
      lblTM138.Visible = False
      textTM138.Visible = False
      lblTM139.Visible = False
      textTM139.Visible = False
   Else
      lblTM137.Visible = True
      textTM137.Visible = True
      lblTM138.Visible = True
      textTM138.Visible = True
      lblTM139.Visible = True
      textTM139.Visible = True
   End If
End Sub

' 將資料庫中的資料更新到所有欄位中
Private Sub UpdateCtrlData()
Dim rsTmp As New ADODB.Recordset
Dim strSql As String
Dim i As Integer 'Add By Sindy 2016/11/23
Dim bolTmp As Boolean 'Added by Lydia 2025/09/12

   Command3.Enabled = False 'Add By Sindy 2014/2/18
   strSql = "SELECT * FROM TradeMark " & _
            "WHERE TM01 = '" & m_CurrTM(0) & "' AND " & _
                  "TM02 = '" & m_CurrTM(1) & "' AND " & _
                  "TM03 = '" & m_CurrTM(2) & "' AND " & _
                  "TM04 = '" & m_CurrTM(3) & "' "
                  
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
   Command3.Enabled = True 'Add By Sindy 2014/2/18
   ClearField
   If rsTmp.Fields("TM01") = "TF" Then
      textTM02_1.MaxLength = 5
      textTM02_2.Visible = True
      textTM02_2.TabStop = True
   Else
      textTM02_1.MaxLength = 6
      textTM02_2.Visible = False
      textTM02_2.TabStop = False
   End If
   
   textTM01 = rsTmp.Fields("TM01")
   If rsTmp.Fields("TM01") = "TF" Then
      textTM02_1 = Mid(rsTmp.Fields("TM02"), 1, 5)
      textTM02_2 = Mid(rsTmp.Fields("TM02"), 6, 1)
   Else
      textTM02_1 = rsTmp.Fields("TM02")
      textTM02_2 = Empty
   End If
   textTM03 = rsTmp.Fields("TM03")
   textTM04 = rsTmp.Fields("TM04")
   If Not IsNull(rsTmp.Fields("TM05")) Then: textTM05 = rsTmp.Fields("TM05"): 'End If
   If Not IsNull(rsTmp.Fields("TM131")) Then: textTM131 = "" & rsTmp.Fields("TM131") 'Add By Sindy 2015/6/30
   'Added by Lydia 2025/09/12 TF基礎案號設定：卷宗性質為申請之母案案號
   'Modified by Lydia 2025/10/23 TF基礎案號(TM06,TM07)改成可以輸入多筆(Table: TFBaseNo)，原本的輸入欄位直接刪除改成按鈕呼叫其他表單，若已有設定則按鈕設為綠色。
   tabCtrl.TabCaption(6) = "銷卷資料"
   cmdTFBaseNo.Visible = False
   cmdTFBaseNo.BackColor = &H8000000F
   If rsTmp.Fields("TM01") = "TF" And Mid(rsTmp.Fields("TM02"), 6, 1) = "0" And rsTmp.Fields("TM03") = "0" And rsTmp.Fields("TM04") = "00" And rsTmp.Fields("TM28") = "1" Then
      tabCtrl.TabCaption(6) = "銷卷/TF基礎案號數"
      cmdTFBaseNo.Visible = True
      strExc(0) = Pub_GetField("TFBaseNo", "TFBN01='" & textTM01 & "' AND TFBN02='" & textTM02_1 & textTM02_2 & "' AND TFBN03='" & textTM03 & "' AND TFBN04='" & textTM04 & "'", "TFBN05")
      If strExc(0) <> "" Then
          cmdTFBaseNo.BackColor = &HC0FFC0
      End If
   End If
   'end 2025/09/12
   
   'Add By Sindy 2024/6/14
   Call SetTM72forCol("" & rsTmp.Fields("TM72"))
   If Not IsNull(rsTmp.Fields("TM137")) Then: textTM137 = "" & rsTmp.Fields("TM137")
   If Not IsNull(rsTmp.Fields("TM138")) Then: textTM138 = "" & rsTmp.Fields("TM138")
   If Not IsNull(rsTmp.Fields("TM139")) Then: textTM139 = "" & rsTmp.Fields("TM139")
   '2024/6/14 END
   If Not IsNull(rsTmp.Fields("TM08")) Then: textTM08_1 = rsTmp.Fields("TM08"): 'End If
   If Not IsNull(rsTmp.Fields("TM09")) Then: textTM09 = rsTmp.Fields("TM09"): 'End If
   If Not IsNull(rsTmp.Fields("TM10")) Then: textTM10_1 = rsTmp.Fields("TM10"): 'End If
   If Not IsNull(rsTmp.Fields("TM11")) Then
      If rsTmp.Fields("TM11") <> "0" Then
         textTM11 = TAIWANDATE(rsTmp.Fields("TM11"))
      End If
   End If
   If Not IsNull(rsTmp.Fields("TM12")) Then: textTM12 = rsTmp.Fields("TM12")
   If Not IsNull(rsTmp.Fields("TM13")) Then
      If rsTmp.Fields("TM13") <> "0" Then
         textTM13 = TAIWANDATE(rsTmp.Fields("TM13"))
      End If
   End If
   If Not IsNull(rsTmp.Fields("TM14")) Then
      If rsTmp.Fields("TM14") <> "0" Then
         textTM14 = TAIWANDATE(rsTmp.Fields("TM14"))
      End If
   End If
   If Not IsNull(rsTmp.Fields("TM15")) Then: textTM15 = rsTmp.Fields("TM15"): 'End If
   'add by sonia 2022/10/12
   textTM15.Tag = "": textTM15.Tag = textTM15
   'end 2022/10/12

   If Not IsNull(rsTmp.Fields("TM16")) Then: textTM16 = rsTmp.Fields("TM16"): 'End If
   If Not IsNull(rsTmp.Fields("TM17")) Then: textTM17 = rsTmp.Fields("TM17"): 'End If
   If Not IsNull(rsTmp.Fields("TM18")) Then: textTM18 = rsTmp.Fields("TM18"): 'End If
   If Not IsNull(rsTmp.Fields("TM19")) Then: textTM19 = rsTmp.Fields("TM19"): 'End If
   If Not IsNull(rsTmp.Fields("TM20")) Then
      If rsTmp.Fields("TM20") <> "0" Then
         textTM20 = TAIWANDATE(rsTmp.Fields("TM20"))
      End If
   End If
   If Not IsNull(rsTmp.Fields("TM21")) Then: textTM21 = rsTmp.Fields("TM21"): 'End If
   If Not IsNull(rsTmp.Fields("TM22")) Then: textTM22 = rsTmp.Fields("TM22"): 'End If
   'add by sonia 2022/10/12
   textTM22.Tag = "": textTM22.Tag = textTM22
   'end 2022/10/12
   If Not IsNull(rsTmp.Fields("TM23")) Then: textTM23_1 = rsTmp.Fields("TM23"): 'End If
    'Add By Cheng 2003/08/07
    '記錄原申請人
    If "" & rsTmp.Fields("TM23").Value <> "" Then
        m_TM23 = Left("" & rsTmp.Fields("TM23").Value & "000000000", 9)
    Else
        m_TM23 = ""
    End If
   If Not IsNull(rsTmp.Fields("TM24")) Then: textTM24 = rsTmp.Fields("TM24"): 'End If
   If Not IsNull(rsTmp.Fields("TM25")) Then: textTM25 = rsTmp.Fields("TM25"): 'End If
   If Not IsNull(rsTmp.Fields("TM26")) Then: textTM26 = rsTmp.Fields("TM26"): 'End If
   If Not IsNull(rsTmp.Fields("TM27")) Then: textTM27 = rsTmp.Fields("TM27"): 'End If
   If Not IsNull(rsTmp.Fields("TM28")) Then
      Select Case rsTmp.Fields("TM28")
         Case 1: textTM28 = "申請"
         Case 2: textTM28 = "異議"
         Case 3: textTM28 = "評定"
         Case 4: textTM28 = "廢止"
      End Select
   End If
   If Not IsNull(rsTmp.Fields("TM29")) Then: textTM29 = rsTmp.Fields("TM29"): 'End If
   If Not IsNull(rsTmp.Fields("TM30")) Then
      If rsTmp.Fields("TM30") <> "0" Then
         textTM30 = TAIWANDATE(rsTmp.Fields("TM30"))
      End If
   End If
   If Not IsNull(rsTmp.Fields("TM31")) Then: textTM31_1 = rsTmp.Fields("TM31"): 'End If
   If Not IsNull(rsTmp.Fields("TM32")) Then: textTM32 = rsTmp.Fields("TM32"): 'End If
   If Not IsNull(rsTmp.Fields("TM33")) Then: textTM33_1 = rsTmp.Fields("TM33"): 'End If
   If Not IsNull(rsTmp.Fields("TM34")) Then: textTM34 = rsTmp.Fields("TM34"): 'End If
   If Not IsNull(rsTmp.Fields("TM35")) Then: textTM35 = rsTmp.Fields("TM35"): 'End If
   If Not IsNull(rsTmp.Fields("TM36")) Then: textTM36 = rsTmp.Fields("TM36"): 'End If
   'Add By Sindy 2025/3/6
   If Not IsNull(rsTmp.Fields("TM140")) Then: textTM140 = rsTmp.Fields("TM140"):
   If Not IsNull(rsTmp.Fields("TM141")) Then: textTM141 = rsTmp.Fields("TM141"):
   '2025/3/6 END
   If Not IsNull(rsTmp.Fields("TM37")) Then: textTM37 = rsTmp.Fields("TM37"): 'End If
   If Not IsNull(rsTmp.Fields("TM38")) Then: textTM38 = rsTmp.Fields("TM38"): 'End If
   If Not IsNull(rsTmp.Fields("TM39")) Then: textTM39 = rsTmp.Fields("TM39"): 'End If
   If Not IsNull(rsTmp.Fields("TM40")) Then: textTM40 = rsTmp.Fields("TM40"): 'End If
   If Not IsNull(rsTmp.Fields("TM41")) Then: textTM41 = rsTmp.Fields("TM41"): 'End If
   If Not IsNull(rsTmp.Fields("TM42")) Then: textTM42 = rsTmp.Fields("TM42"): 'End If
   If Not IsNull(rsTmp.Fields("TM43")) Then: textTM43 = rsTmp.Fields("TM43"): 'End If
   If Not IsNull(rsTmp.Fields("TM44")) Then: textTM44_1 = rsTmp.Fields("TM44"): 'End If
   m_TM44 = "" & rsTmp.Fields("TM44") 'Added by Lydia 2024/06/13
   If Not IsNull(rsTmp.Fields("TM45")) Then: textTM45 = rsTmp.Fields("TM45"): 'End If
   If Not IsNull(rsTmp.Fields("TM46")) Then: textTM46 = rsTmp.Fields("TM46"): 'End If
   If Not IsNull(rsTmp.Fields("TM47")) Then: textTM47 = rsTmp.Fields("TM47"): 'End If
   If Not IsNull(rsTmp.Fields("TM48")) Then: textTM48 = rsTmp.Fields("TM48"): 'End If
   If Not IsNull(rsTmp.Fields("TM49")) Then: textTM49 = rsTmp.Fields("TM49"): 'End If
   If Not IsNull(rsTmp.Fields("TM50")) Then: textTM50 = rsTmp.Fields("TM50"): 'End If
   If Not IsNull(rsTmp.Fields("TM51")) Then: textTM51 = rsTmp.Fields("TM51"): 'End If
   If Not IsNull(rsTmp.Fields("TM52")) Then: textTM52 = rsTmp.Fields("TM52"): 'End If
   If Not IsNull(rsTmp.Fields("TM53")) Then: textTM53 = rsTmp.Fields("TM53"): 'End If
   If Not IsNull(rsTmp.Fields("TM54")) Then: textTM54_1 = rsTmp.Fields("TM54"): 'End If
   If Not IsNull(rsTmp.Fields("TM55")) Then: textTM55 = rsTmp.Fields("TM55"): 'End If
   If Not IsNull(rsTmp.Fields("TM56")) Then: textTM56_1 = rsTmp.Fields("TM56"): 'End If
   If Not IsNull(rsTmp.Fields("TM130")) Then: textTM130 = rsTmp.Fields("TM130") 'Add By Sindy 2013/12/13
'edit by nickc 2006/07/12 已經不允許修改了
'   If Not IsNull(rsTmp.Fields("TM57")) Then: textTM57 = rsTmp.Fields("TM57"): 'End If
   If Not IsNull(rsTmp.Fields("TM58")) Then: textTM58 = rsTmp.Fields("TM58"): 'End If
   If Not IsNull(rsTmp.Fields("TM65")) Then: textTM65 = rsTmp.Fields("TM65"): 'End If
   If Not IsNull(rsTmp.Fields("TM66")) Then: textTM66_1 = rsTmp.Fields("TM66"): 'End If
   If Not IsNull(rsTmp.Fields("TM67")) Then: textTM67 = rsTmp.Fields("TM67"): 'End If
   If Not IsNull(rsTmp.Fields("TM68")) Then: textTM68 = rsTmp.Fields("TM68"): 'End If
   If Not IsNull(rsTmp.Fields("TM69")) Then: textTM69_1 = rsTmp.Fields("TM69"): 'End If
   If Not IsNull(rsTmp.Fields("TM70")) Then: textTM70_1 = rsTmp.Fields("TM70"): 'End If
   If Not IsNull(rsTmp.Fields("TM71")) Then: textTM71_1 = rsTmp.Fields("TM71"): 'End If
   If Not IsNull(rsTmp.Fields("TM72")) Then: textTM72_1 = rsTmp.Fields("TM72"): 'End If
   
   'add by Toni 2008/10/21
   If Not IsNull(rsTmp.Fields("TM122")) Then: textTM122 = rsTmp.Fields("TM122"): 'End If
   'end 2008/10/21
   
   'Add By Sindy 2009/09/09
   If Not IsNull(rsTmp.Fields("TM77")) Then: textTM77 = rsTmp.Fields("TM77"): 'End If
   If Not IsNull(rsTmp.Fields("TM124")) Then: textTM124 = rsTmp.Fields("TM124"): 'End If
   If Not IsNull(rsTmp.Fields("TM125")) Then: textTM125 = rsTmp.Fields("TM125"): 'End If
   If Not IsNull(rsTmp.Fields("TM126")) Then: textTM126 = rsTmp.Fields("TM126"): 'End If
   '2009/09/09 End
   
   textTM127 = "" & rsTmp.Fields("TM127") 'Add by Morgan 2010/11/5
   textTM129 = "" & rsTmp.Fields("TM129") 'Add by Sindy 2013/8/15
   textTM136 = "" & rsTmp.Fields("TM136") 'Add by Morgan 2022/12/1
   
   'add by nickc 2006/12/08
   textTM78_1 = CheckStr(rsTmp.Fields("tm78"))
   textTM79_1 = CheckStr(rsTmp.Fields("tm79"))
   textTM80_1 = CheckStr(rsTmp.Fields("tm80"))
   textTM81_1 = CheckStr(rsTmp.Fields("tm81"))
    '記錄原申請人2
    If "" & rsTmp.Fields("TM78").Value <> "" Then
        m_TM78 = Left("" & rsTmp.Fields("TM78").Value & "000000000", 9)
    Else
        m_TM78 = ""
    End If
    '記錄原申請人3
    If "" & rsTmp.Fields("TM79").Value <> "" Then
        m_TM79 = Left("" & rsTmp.Fields("TM79").Value & "000000000", 9)
    Else
        m_TM79 = ""
    End If
    '記錄原申請人4
    If "" & rsTmp.Fields("TM80").Value <> "" Then
        m_TM80 = Left("" & rsTmp.Fields("TM80").Value & "000000000", 9)
    Else
        m_TM80 = ""
    End If
    '記錄原申請人5
    If "" & rsTmp.Fields("TM81").Value <> "" Then
        m_TM81 = Left("" & rsTmp.Fields("TM81").Value & "000000000", 9)
    Else
        m_TM81 = ""
    End If
   textTM82 = CheckStr(rsTmp.Fields("tm82"))
   textTM83 = CheckStr(rsTmp.Fields("tm83"))
   textTM84 = CheckStr(rsTmp.Fields("tm84"))
   textTM85 = CheckStr(rsTmp.Fields("tm85"))
   textTM86 = CheckStr(rsTmp.Fields("tm86"))
   textTM87 = CheckStr(rsTmp.Fields("tm87"))
   textTM88 = CheckStr(rsTmp.Fields("tm88"))
   textTM89 = CheckStr(rsTmp.Fields("tm89"))
   textTM90 = CheckStr(rsTmp.Fields("tm90"))
   textTM91 = CheckStr(rsTmp.Fields("tm91"))
   textTM92 = CheckStr(rsTmp.Fields("tm92"))
   textTM93 = CheckStr(rsTmp.Fields("tm93"))
   textTM94 = CheckStr(rsTmp.Fields("tm94"))
   textTM95 = CheckStr(rsTmp.Fields("tm95"))
   textTM96 = CheckStr(rsTmp.Fields("tm96"))
   textTM97 = CheckStr(rsTmp.Fields("tm97"))
   textTM98 = CheckStr(rsTmp.Fields("tm98"))
   textTM99 = CheckStr(rsTmp.Fields("tm99"))
   textTM100 = CheckStr(rsTmp.Fields("tm100"))
   textTM101 = CheckStr(rsTmp.Fields("tm101"))
   textTM102 = CheckStr(rsTmp.Fields("tm102"))
   textTM103 = CheckStr(rsTmp.Fields("tm103"))
   textTM104 = CheckStr(rsTmp.Fields("tm104"))
   textTM105 = CheckStr(rsTmp.Fields("tm105"))
   textTM106 = CheckStr(rsTmp.Fields("tm106"))
   textTM107 = CheckStr(rsTmp.Fields("tm107"))
   textTM108 = CheckStr(rsTmp.Fields("tm108"))
   textTM109 = CheckStr(rsTmp.Fields("tm109"))
   textTM110 = CheckStr(rsTmp.Fields("tm110"))
   textTM111 = CheckStr(rsTmp.Fields("tm111"))
   textTM112 = CheckStr(rsTmp.Fields("tm112"))
   textTM113 = CheckStr(rsTmp.Fields("tm113"))
   textTM114 = CheckStr(rsTmp.Fields("tm114"))
   textTM115 = CheckStr(rsTmp.Fields("tm115"))
   textTM116 = CheckStr(rsTmp.Fields("tm116"))
   textTM117 = CheckStr(rsTmp.Fields("tm117"))
   'add by nickc 2007/03/08
   textTM118 = CheckStr(rsTmp.Fields("tm118"))
   
   'add by nickc 2007/01/02
   textTM76 = CheckStr(rsTmp.Fields("tm76"))
   textTM121 = CheckStr(rsTmp.Fields("tm121")) 'Add by Morgan 2008/5/23
   
   'Added by Lydia 2025/09/12
   textTM12.Tag = "": textTM12.Tag = textTM12
   textTM10_1.Tag = "": textTM10_1.Tag = textTM10_1
   textTM29.Tag = "": textTM29.Tag = textTM29
   textTM16.Tag = "": textTM16.Tag = textTM16
   'end 2025/09/12
   
   'Add By Sindy 2016/11/23
   If Combo4.Visible = True Then
      If IsNull(rsTmp.Fields("tm134")) = False Then
         For i = 0 To Combo4.ListCount - 1
            Combo4.ListIndex = i
            If InStr(Combo4.Text, rsTmp.Fields("tm134")) > 0 Then
               Exit For
            End If
         Next
      Else
         Combo4.ListIndex = 0
      End If
      If IsNull(rsTmp.Fields("tm135")) = False Then
         Combo5.ListIndex = Val(rsTmp.Fields("tm135"))
      Else
         Combo5.ListIndex = 0
      End If
   End If
   '2016/11/23 End
   
   'Modified by Lydia 2021/11/29 改成Form 2.0
   'PUB_AddContact "" & rsTmp.Fields("tm23"), cboContact, "" & rsTmp.Fields("tm123") 'Add by Morgan 2008/8/4
   m_ContactList = cboContact.Tag
   PUB_AddContact "" & rsTmp.Fields("tm23"), cboContact, "" & rsTmp.Fields("tm123"), , True, m_ContactList
   cboContact.Tag = m_ContactList
   'end 2021/11/29
   UpdateFieldOldData rsTmp
   
   'add by nickc 2006/07/12
   Dim strTemp As String
   If IsNull(rsTmp.Fields("TM57")) = False Then
      If IsEmptyText(rsTmp.Fields("TM57")) = False Then
         strTemp = TAIWANDATE(rsTmp.Fields("TM57"))
         textTM57 = Format(strTemp, "###/##/##")
      End If
   End If
   If IsNull(rsTmp.Fields("TM73")) = False Then
      If IsEmptyText(rsTmp.Fields("TM73")) = False Then
         strTemp = TAIWANDATE(rsTmp.Fields("TM73"))
         textTM73 = Format(strTemp, "###/##/##")
      End If
   End If
   If IsNull(rsTmp.Fields("TM74")) = False Then
      If IsEmptyText(rsTmp.Fields("TM74")) = False Then
         textTM74 = GetStaffName(rsTmp.Fields("TM74"), True)
      End If
   End If
   If Not IsNull(rsTmp.Fields("TM75")) Then: textTM75 = rsTmp.Fields("TM75")
   ' 更新顯示 Create 及 Update 的人
   UpdateCUID rsTmp
   
   textTM08_1_Validate False
   textTM08_2 = GetTradeMarkName(textTM08_1, IIf(textTM10_1 = "020", 1, 0)) 'Add By Sindy 2015/8/13\
   Pub_SetTMcombo "1", cboTM08, textTM08_1, IIf(textTM10_1 <> "000", True, False), strPTM 'Added by Lydia 2023/11/16 內外商之分案及商標基本資料維護之商標種類、特殊商標欄位增加下拉功能: 商標種類
   
   textTM10_1_Validate False
   textTM15_Validate False
    'Modify By Cheng 2003/08/07
    '讀取資料時, 不重抓申請人的相關資料
'   textTM23_1_Validate False
    Me.textTM23_2.Text = GetTM23Name(Me.textTM23_1.Text)
    
    Me.textTM78_2.Text = GetTM23Name(Me.textTM78_1.Text)
    Me.textTM79_2.Text = GetTM23Name(Me.textTM79_1.Text)
    Me.textTM80_2.Text = GetTM23Name(Me.textTM80_1.Text)
    Me.textTM81_2.Text = GetTM23Name(Me.textTM81_1.Text)
    
   textTM31_1_Validate False
   textTM44_1_Validate False
   textTM33_1_Validate False
   textTM54_1_Validate False
   textTM56_1_Validate False
   textTM66_1_Validate False
   textTM69_1_Validate False
   textTM70_1_Validate False
   textTM72_1_Validate False
   Pub_SetTMcombo "2", cboTM72, textTM72_1, IIf(textTM10_1 <> "000", True, False), strSPT 'Added by Lydia 2023/11/16 內外商之分案及商標基本資料維護之商標種類、特殊商標欄位增加下拉功能: 特殊商標種類
  
   'Add By Sindy 2012/6/15 檢查有無代表圖
   'Modify by Amy 2018/07/16  改寫至function
'   strSql = "SELECT ibf01,ibf02 FROM imgbytefile WHERE ibf01='" & m_CurrTM(0) & "' and ibf02='" & m_CurrTM(1) & "' and ibf03='" & m_CurrTM(2) & "' and ibf04='" & m_CurrTM(3) & "' and ibf05='1'"
'   CheckOC2
'   adoRecordset1.CursorLocation = adUseClient
'   adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
'   If adoRecordset1.RecordCount <> 0 And adoRecordset1.RecordCount > 0 Then
   If ChkImgByteFile(m_CurrTM(0), m_CurrTM(1), m_CurrTM(2), m_CurrTM(3)) = True Then
       'Modified by Lydia 2021/12/16 拿掉快速鍵(&I)
       Command3.Caption = "已設定代表圖"
       Command3.BackColor = &HC0FFC0
   Else
       'Modified by Lydia 2021/12/16 拿掉快速鍵(&I)
       Command3.Caption = "未設定代表圖"
       Command3.BackColor = &HC0C0FF
   End If
'   CheckOC2
   'end 2018/07/16
   '2012/6/15 End
   End If
   rsTmp.Close
   
   If textTM02_1.Enabled = True And Me.Visible = True Then textTM02_1.SetFocus 'Add By Sindy 2019/12/10
   
EXITSUB:
   Set rsTmp = Nothing
End Sub

' 案件進度檔
Private Function IsCaseProgressExist(ByVal strTM01 As String, ByVal strTM02 As String, ByVal strTM03 As String, ByVal strTM04 As String) As Boolean
Dim strSql As String
Dim rsTmp As New ADODB.Recordset
   
   IsCaseProgressExist = False
   strSql = "SELECT * from CaseProgress " & _
            "WHERE CP01 = '" & strTM01 & "' AND " & _
                  "CP02 = '" & strTM02 & "' AND " & _
                  "CP03 = '" & strTM03 & "' AND " & _
                  "CP04 = '" & strTM04 & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly   'edit by nickc 2005/08/04 原先是  動態開啟
   If rsTmp.RecordCount > 0 Then
      IsCaseProgressExist = True
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Function

' 更新 Create 及 Update 的人
Private Sub UpdateCUID(ByRef rsSrcTmp As ADODB.Recordset)
Dim strTemp As String
Dim strCName As String
Dim strCDate As String
Dim strCTime As String
Dim strUName As String
Dim strUDate As String
Dim strUTime As String
   
   If IsNull(rsSrcTmp.Fields("TM59")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("TM59")) = False Then
         strCName = GetStaffName(rsSrcTmp.Fields("TM59"), True)
      End If
   End If
   If IsNull(rsSrcTmp.Fields("TM60")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("TM60")) = False Then
         strTemp = TAIWANDATE(rsSrcTmp.Fields("TM60"))
         strCDate = Format(strTemp, "###/##/##")
      End If
   End If
   If IsNull(rsSrcTmp.Fields("TM61")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("TM61")) = False Then
         strTemp = rsSrcTmp.Fields("TM61")
         strCTime = Format(strTemp, "##:##")
      End If
   End If
   If IsNull(rsSrcTmp.Fields("TM62")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("TM62")) = False Then
         strUName = GetStaffName(rsSrcTmp.Fields("TM62"), True)
      End If
   End If
   If IsNull(rsSrcTmp.Fields("TM63")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("TM63")) = False Then
         strTemp = TAIWANDATE(rsSrcTmp.Fields("TM63"))
         strUDate = Format(strTemp, "###/##/##")
      End If
   End If
   If IsNull(rsSrcTmp.Fields("TM64")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("TM64")) = False Then
         strTemp = rsSrcTmp.Fields("TM64")
         strUTime = Format(strTemp, "##:##")
      End If
   End If
   
   ' 設定CUID中的文字
   textCUID_1 = "CREATE : " & strCName & " " & _
              " " & strCDate & " " & _
              " " & strCTime
   textCUID_2 = "UPDATE : " & strUName & " " & _
              " " & strUDate & " " & _
              " " & strUTime
End Sub

' 顯示資料
'Modify By Sindy 2016/11/28
'Private Sub ShowCurrRecord(ByVal strTM01 As String, ByVal strTM02 As String, ByVal strTM03 As String, ByVal strTM04 As String)
Public Sub ShowCurrRecord(ByVal strTM01 As String, ByVal strTM02 As String, ByVal strTM03 As String, ByVal strTM04 As String)
'2016/11/28 END
Dim strSql As String
Dim rsTmp As New ADODB.Recordset
   
   If IsRecordExist(strTM01, strTM02, strTM03, strTM04) = True Then
      m_CurrTM(0) = strTM01
      m_CurrTM(1) = strTM02
      m_CurrTM(2) = strTM03
      m_CurrTM(3) = strTM04
   Else
      strSql = "SELECT TM01,TM02,TM03,TM04 FROM TRADEMARK " & _
               "WHERE TM01 = '" & m_CurrTM(0) & "' AND " & _
                     "TM02 = '" & m_CurrTM(1) & "' AND " & _
                     "TM03 = '" & m_CurrTM(2) & "' AND " & _
                     "TM04 = (SELECT MIN(TM04) FROM TRADEMARK " & _
                             "WHERE TM01 = '" & m_CurrTM(0) & "' AND " & _
                                   "TM02 = '" & m_CurrTM(1) & "' AND " & _
                                   "TM03 = '" & m_CurrTM(2) & "' AND " & _
                                   "TM04 > '" & m_CurrTM(3) & "' )"
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         If IsNull(rsTmp.Fields("TM01")) = False Then: m_CurrTM(0) = rsTmp.Fields("TM01")
         If IsNull(rsTmp.Fields("TM02")) = False Then: m_CurrTM(1) = rsTmp.Fields("TM02")
         If IsNull(rsTmp.Fields("TM03")) = False Then: m_CurrTM(2) = rsTmp.Fields("TM03")
         If IsNull(rsTmp.Fields("TM04")) = False Then: m_CurrTM(3) = rsTmp.Fields("TM04")
         rsTmp.Close
         UpdateCtrlData
         GoTo EXITSUB
      End If
      rsTmp.Close
   
      strSql = "SELECT TM01,TM02,TM03,TM04 FROM TRADEMARK " & _
               "WHERE TM01 = '" & m_CurrTM(0) & "' AND " & _
                     "TM02 = '" & m_CurrTM(1) & "' AND " & _
                     "TM03 = (SELECT MIN(TM03) FROM TRADEMARK " & _
                             "WHERE TM01 = '" & m_CurrTM(0) & "' AND " & _
                                   "TM02 = '" & m_CurrTM(1) & "' AND " & _
                                   "TM03 > '" & m_CurrTM(2) & "') AND " & _
                     "TM04 = (SELECT MIN(TM04) FROM TRADEMARK " & _
                             "WHERE TM01 = '" & m_CurrTM(0) & "' AND " & _
                                   "TM02 = '" & m_CurrTM(1) & "' AND " & _
                                   "TM03 = (SELECT MIN(TM03) FROM TRADEMARK " & _
                                           "WHERE TM01 = '" & m_CurrTM(0) & "' AND " & _
                                                 "TM02 = '" & m_CurrTM(1) & "' AND " & _
                                                 "TM03 > '" & m_CurrTM(2) & "'))"
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         If IsNull(rsTmp.Fields("TM01")) = False Then: m_CurrTM(0) = rsTmp.Fields("TM01")
         If IsNull(rsTmp.Fields("TM02")) = False Then: m_CurrTM(1) = rsTmp.Fields("TM02")
         If IsNull(rsTmp.Fields("TM03")) = False Then: m_CurrTM(2) = rsTmp.Fields("TM03")
         If IsNull(rsTmp.Fields("TM04")) = False Then: m_CurrTM(3) = rsTmp.Fields("TM04")
         rsTmp.Close
         UpdateCtrlData
         GoTo EXITSUB
      End If
      rsTmp.Close
                                
      strSql = "SELECT TM01,TM02,TM03,TM04 FROM TRADEMARK " & _
               "WHERE TM01 = '" & m_CurrTM(0) & "' AND " & _
                     "TM02 = (SELECT MIN(TM02) FROM TRADEMARK " & _
                             "WHERE TM01 = '" & m_CurrTM(0) & "' AND " & _
                                   "TM02 > '" & m_CurrTM(1) & "') AND " & _
                     "TM03 = (SELECT MIN(TM03) FROM TRADEMARK " & _
                             "WHERE TM01 = '" & m_CurrTM(0) & "' AND " & _
                                   "TM02 = (SELECT MIN(TM02) FROM TRADEMARK " & _
                                           "WHERE TM01 = '" & m_CurrTM(0) & "' AND " & _
                                                 "TM02 > '" & m_CurrTM(1) & "')) AND " & _
                     "TM04 = (SELECT MIN(TM04) FROM TRADEMARK " & _
                             "WHERE TM01 = '" & m_CurrTM(0) & "' AND " & _
                                   "TM02 = (SELECT MIN(TM02) FROM TRADEMARK " & _
                                           "WHERE TM01 = '" & m_CurrTM(0) & "' AND " & _
                                                 "TM02 > '" & m_CurrTM(1) & "') AND " & _
                                                 "TM03 = (SELECT MIN(TM03) FROM TRADEMARK " & _
                                                         "WHERE TM01 = '" & m_CurrTM(0) & "' AND " & _
                                                               "TM02 = (SELECT MIN(TM02) FROM TRADEMARK " & _
                                                                       "WHERE TM01 = '" & m_CurrTM(0) & "' AND " & _
                                                                             "TM02 > '" & m_CurrTM(1) & "'))) "
   
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         If IsNull(rsTmp.Fields("TM01")) = False Then: m_CurrTM(0) = rsTmp.Fields("TM01")
         If IsNull(rsTmp.Fields("TM02")) = False Then: m_CurrTM(1) = rsTmp.Fields("TM02")
         If IsNull(rsTmp.Fields("TM03")) = False Then: m_CurrTM(2) = rsTmp.Fields("TM03")
         If IsNull(rsTmp.Fields("TM04")) = False Then: m_CurrTM(3) = rsTmp.Fields("TM04")
      Else
         ShowLastRecord
         GoTo EXITSUB
      End If
      rsTmp.Close
   
      'Select Case m_SysKind
      '   Case 0:
      '      strSQL = "SELECT TM01,TM02,TM03,TM04 FROM TradeMark " & _
      '               "WHERE (TM01||TM02||TM03||TM04) IN (SELECT MIN(TM01||TM02||TM03||TM04) FROM TradeMark " & _
      '                                                  "WHERE (TM01||TM02||TM03||TM04) > '" & strTM01 & strTM02 & strTM03 & strTM04 & "' AND " & _
      '                                                  "(TM01 = 'T' OR TM01 = 'TF'))"
      '   Case 1:
      '      strSQL = "SELECT TM01,TM02,TM03,TM04 FROM TradeMark " & _
      '               "WHERE (TM01||TM02||TM03||TM04) IN (SELECT MIN(TM01||TM02||TM03||TM04) FROM TradeMark " & _
      '                                                  "WHERE (TM01||TM02||TM03||TM04) > '" & strTM01 & strTM02 & strTM03 & strTM04 & "' AND " & _
      '                                                  "(TM01 = 'CFT' OR TM01 = 'FCT'))"
      'End Select
      'rsTmp.CursorLocation = adUseClient
      'rsTmp.Open strSQL, cnnConnection, adOpenStatic, adLockReadOnly
      'If rsTmp.RecordCount > 0 Then
      '   If IsNull(rsTmp.Fields("TM01")) = False Then: m_CurrTM(0) = rsTmp.Fields("TM01")
      '   If IsNull(rsTmp.Fields("TM02")) = False Then: m_CurrTM(1) = rsTmp.Fields("TM02")
      '   If IsNull(rsTmp.Fields("TM03")) = False Then: m_CurrTM(2) = rsTmp.Fields("TM03")
      '   If IsNull(rsTmp.Fields("TM04")) = False Then: m_CurrTM(3) = rsTmp.Fields("TM04")
      'Else
      '   'RefreshRange
      '   m_CurrTM(0) = m_LastTM(0)
      '   m_CurrTM(1) = m_LastTM(1)
      '   m_CurrTM(2) = m_LastTM(2)
      '   m_CurrTM(3) = m_LastTM(3)
      'End If
      'rsTmp.Close
   End If
   UpdateCtrlData
   Command2.Enabled = True   'add by sonia 2025/5/14  T案證明標章、團體標章不可輸入商品及服務，故此處要打開
EXITSUB:
End Sub

' 顯示第一筆資料
Private Sub ShowFirstRecord()
   m_CurrTM(0) = m_FirstTM(0)
   m_CurrTM(1) = m_FirstTM(1)
   m_CurrTM(2) = m_FirstTM(2)
   m_CurrTM(3) = m_FirstTM(3)
   
   UpdateCtrlData
End Sub

' 顯示上一筆資料
Private Sub ShowPrevRecord()
Dim strSql As String
Dim rsTmp As New ADODB.Recordset
   
   If m_CurrTM(0) = m_FirstTM(0) And m_CurrTM(1) = m_FirstTM(1) And m_CurrTM(2) = m_FirstTM(2) And m_CurrTM(3) = m_FirstTM(3) Then
      ShowMsg MsgText(9008)
      GoTo EXITSUB
   End If
   
   'Select Case m_SysKind
   '   Case 0:
   '      strSQL = "SELECT TM01,TM02,TM03,TM04 FROM TradeMark " & _
   '               "WHERE (TM01||TM02||TM03||TM04) IN (SELECT MAX(TM01||TM02||TM03||TM04) FROM TradeMark " & _
   '                                                  "WHERE (TM01||TM02||TM03||TM04) < '" & m_CurrTM(0) & m_CurrTM(1) & m_CurrTM(2) & m_CurrTM(3) & "' AND " & _
   '                                                         "(TM01 = 'T' OR TM01 = 'TF'))"
   '   Case 1:
   '      strSQL = "SELECT TM01,TM02,TM03,TM04 FROM TradeMark " & _
   '               "WHERE (TM01||TM02||TM03||TM04) IN (SELECT MAX(TM01||TM02||TM03||TM04) FROM TradeMark " & _
   '                                                  "WHERE (TM01||TM02||TM03||TM04) < '" & m_CurrTM(0) & m_CurrTM(1) & m_CurrTM(2) & m_CurrTM(3) & "' AND " & _
   '                                                         "(TM01 = 'CFT' OR TM01 = 'FCT'))"
   'End Select
   
   strSql = "SELECT TM01,TM02,TM03,TM04 FROM TRADEMARK " & _
            "WHERE TM01 = '" & m_CurrTM(0) & "' AND " & _
                  "TM02 = '" & m_CurrTM(1) & "' AND " & _
                  "TM03 = '" & m_CurrTM(2) & "' AND " & _
                  "TM04 = (SELECT MAX(TM04) FROM TRADEMARK " & _
                          "WHERE TM01 = '" & m_CurrTM(0) & "' AND " & _
                                "TM02 = '" & m_CurrTM(1) & "' AND " & _
                                "TM03 = '" & m_CurrTM(2) & "' AND " & _
                                "TM04 < '" & m_CurrTM(3) & "' )"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("TM01")) = False Then: m_CurrTM(0) = rsTmp.Fields("TM01")
      If IsNull(rsTmp.Fields("TM02")) = False Then: m_CurrTM(1) = rsTmp.Fields("TM02")
      If IsNull(rsTmp.Fields("TM03")) = False Then: m_CurrTM(2) = rsTmp.Fields("TM03")
      If IsNull(rsTmp.Fields("TM04")) = False Then: m_CurrTM(3) = rsTmp.Fields("TM04")
      rsTmp.Close
      UpdateCtrlData
      GoTo EXITSUB
   End If
   rsTmp.Close
   
   strSql = "SELECT TM01,TM02,TM03,TM04 FROM TRADEMARK " & _
            "WHERE TM01 = '" & m_CurrTM(0) & "' AND " & _
                  "TM02 = '" & m_CurrTM(1) & "' AND " & _
                  "TM03 = (SELECT MAX(TM03) FROM TRADEMARK " & _
                          "WHERE TM01 = '" & m_CurrTM(0) & "' AND " & _
                                "TM02 = '" & m_CurrTM(1) & "' AND " & _
                                "TM03 < '" & m_CurrTM(2) & "') AND " & _
                  "TM04 = (SELECT MAX(TM04) FROM TRADEMARK " & _
                          "WHERE TM01 = '" & m_CurrTM(0) & "' AND " & _
                                "TM02 = '" & m_CurrTM(1) & "' AND " & _
                                "TM03 = (SELECT MAX(TM03) FROM TRADEMARK " & _
                                        "WHERE TM01 = '" & m_CurrTM(0) & "' AND " & _
                                              "TM02 = '" & m_CurrTM(1) & "' AND " & _
                                              "TM03 < '" & m_CurrTM(2) & "'))"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("TM01")) = False Then: m_CurrTM(0) = rsTmp.Fields("TM01")
      If IsNull(rsTmp.Fields("TM02")) = False Then: m_CurrTM(1) = rsTmp.Fields("TM02")
      If IsNull(rsTmp.Fields("TM03")) = False Then: m_CurrTM(2) = rsTmp.Fields("TM03")
      If IsNull(rsTmp.Fields("TM04")) = False Then: m_CurrTM(3) = rsTmp.Fields("TM04")
      rsTmp.Close
      UpdateCtrlData
      GoTo EXITSUB
   End If
   rsTmp.Close
   
   strSql = "SELECT TM01,TM02,TM03,TM04 FROM TRADEMARK " & _
            "WHERE TM01 = '" & m_CurrTM(0) & "' AND " & _
                  "TM02 = (SELECT MAX(TM02) FROM TRADEMARK " & _
                          "WHERE TM01 = '" & m_CurrTM(0) & "' AND " & _
                                "TM02 < '" & m_CurrTM(1) & "') AND " & _
                  "TM03 = (SELECT MAX(TM03) FROM TRADEMARK " & _
                          "WHERE TM01 = '" & m_CurrTM(0) & "' AND " & _
                                "TM02 = (SELECT MAX(TM02) FROM TRADEMARK " & _
                                        "WHERE TM01 = '" & m_CurrTM(0) & "' AND " & _
                                              "TM02 < '" & m_CurrTM(1) & "')) AND " & _
                  "TM04 = (SELECT MAX(TM04) FROM TRADEMARK " & _
                          "WHERE TM01 = '" & m_CurrTM(0) & "' AND " & _
                                "TM02 = (SELECT MAX(TM02) FROM TRADEMARK " & _
                                        "WHERE TM01 = '" & m_CurrTM(0) & "' AND " & _
                                              "TM02 < '" & m_CurrTM(1) & "') AND " & _
                                              "TM03 = (SELECT MAX(TM03) FROM TRADEMARK " & _
                                                      "WHERE TM01 = '" & m_CurrTM(0) & "' AND " & _
                                                            "TM02 = (SELECT MAX(TM02) FROM TRADEMARK " & _
                                                                    "WHERE TM01 = '" & m_CurrTM(0) & "' AND " & _
                                                                          "TM02 < '" & m_CurrTM(1) & "'))) "
                  
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("TM01")) = False Then: m_CurrTM(0) = rsTmp.Fields("TM01")
      If IsNull(rsTmp.Fields("TM02")) = False Then: m_CurrTM(1) = rsTmp.Fields("TM02")
      If IsNull(rsTmp.Fields("TM03")) = False Then: m_CurrTM(2) = rsTmp.Fields("TM03")
      If IsNull(rsTmp.Fields("TM04")) = False Then: m_CurrTM(3) = rsTmp.Fields("TM04")
   End If
   rsTmp.Close
   UpdateCtrlData
   
EXITSUB:
   Set rsTmp = Nothing
End Sub

' 顯示下一筆資料
Private Sub ShowNextRecord()
Dim strSql As String
Dim rsTmp As New ADODB.Recordset
   
   If m_CurrTM(0) = m_LastTM(0) And m_CurrTM(1) = m_LastTM(1) And m_CurrTM(2) = m_LastTM(2) And m_CurrTM(3) = m_LastTM(3) Then
      ShowMsg MsgText(9009)
      GoTo EXITSUB
   End If
   
   'Select Case m_SysKind
   '   Case 0:
   '      strSQL = "SELECT TM01,TM02,TM03,TM04 FROM TradeMark " & _
   '               "WHERE (TM01||TM02||TM03||TM04) IN (SELECT MIN(TM01||TM02||TM03||TM04) FROM TradeMark " & _
   '                                                  "WHERE (TM01||TM02||TM03||TM04) > '" & m_CurrTM(0) & m_CurrTM(1) & m_CurrTM(2) & m_CurrTM(3) & "' AND " & _
   '                                                         "(TM01 = 'T' OR TM01 = 'TF'))"
   '   Case 1:
   '      strSQL = "SELECT TM01,TM02,TM03,TM04 FROM TradeMark " & _
   '               "WHERE (TM01||TM02||TM03||TM04) IN (SELECT MIN(TM01||TM02||TM03||TM04) FROM TradeMark " & _
   '                                                  "WHERE (TM01||TM02||TM03||TM04) > '" & m_CurrTM(0) & m_CurrTM(1) & m_CurrTM(2) & m_CurrTM(3) & "' AND " & _
   '                                                         "(TM01 = 'CFT' OR TM01 = 'FCT'))"
   'End Select
   
   strSql = "SELECT TM01,TM02,TM03,TM04 FROM TRADEMARK " & _
            "WHERE TM01 = '" & m_CurrTM(0) & "' AND " & _
                  "TM02 = '" & m_CurrTM(1) & "' AND " & _
                  "TM03 = '" & m_CurrTM(2) & "' AND " & _
                  "TM04 = (SELECT MIN(TM04) FROM TRADEMARK " & _
                          "WHERE TM01 = '" & m_CurrTM(0) & "' AND " & _
                                "TM02 = '" & m_CurrTM(1) & "' AND " & _
                                "TM03 = '" & m_CurrTM(2) & "' AND " & _
                                "TM04 > '" & m_CurrTM(3) & "' )"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("TM01")) = False Then: m_CurrTM(0) = rsTmp.Fields("TM01")
      If IsNull(rsTmp.Fields("TM02")) = False Then: m_CurrTM(1) = rsTmp.Fields("TM02")
      If IsNull(rsTmp.Fields("TM03")) = False Then: m_CurrTM(2) = rsTmp.Fields("TM03")
      If IsNull(rsTmp.Fields("TM04")) = False Then: m_CurrTM(3) = rsTmp.Fields("TM04")
      rsTmp.Close
      UpdateCtrlData
      GoTo EXITSUB
   End If
   rsTmp.Close
   
   strSql = "SELECT TM01,TM02,TM03,TM04 FROM TRADEMARK " & _
            "WHERE TM01 = '" & m_CurrTM(0) & "' AND " & _
                  "TM02 = '" & m_CurrTM(1) & "' AND " & _
                  "TM03 = (SELECT MIN(TM03) FROM TRADEMARK " & _
                          "WHERE TM01 = '" & m_CurrTM(0) & "' AND " & _
                                "TM02 = '" & m_CurrTM(1) & "' AND " & _
                                "TM03 > '" & m_CurrTM(2) & "') AND " & _
                  "TM04 = (SELECT MIN(TM04) FROM TRADEMARK " & _
                          "WHERE TM01 = '" & m_CurrTM(0) & "' AND " & _
                                "TM02 = '" & m_CurrTM(1) & "' AND " & _
                                "TM03 = (SELECT MIN(TM03) FROM TRADEMARK " & _
                                        "WHERE TM01 = '" & m_CurrTM(0) & "' AND " & _
                                              "TM02 = '" & m_CurrTM(1) & "' AND " & _
                                              "TM03 > '" & m_CurrTM(2) & "'))"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("TM01")) = False Then: m_CurrTM(0) = rsTmp.Fields("TM01")
      If IsNull(rsTmp.Fields("TM02")) = False Then: m_CurrTM(1) = rsTmp.Fields("TM02")
      If IsNull(rsTmp.Fields("TM03")) = False Then: m_CurrTM(2) = rsTmp.Fields("TM03")
      If IsNull(rsTmp.Fields("TM04")) = False Then: m_CurrTM(3) = rsTmp.Fields("TM04")
      rsTmp.Close
      UpdateCtrlData
      GoTo EXITSUB
   End If
   rsTmp.Close
                                
   strSql = "SELECT TM01,TM02,TM03,TM04 FROM TRADEMARK " & _
            "WHERE TM01 = '" & m_CurrTM(0) & "' AND " & _
                  "TM02 = (SELECT MIN(TM02) FROM TRADEMARK " & _
                          "WHERE TM01 = '" & m_CurrTM(0) & "' AND " & _
                                "TM02 > '" & m_CurrTM(1) & "') AND " & _
                  "TM03 = (SELECT MIN(TM03) FROM TRADEMARK " & _
                          "WHERE TM01 = '" & m_CurrTM(0) & "' AND " & _
                                "TM02 = (SELECT MIN(TM02) FROM TRADEMARK " & _
                                        "WHERE TM01 = '" & m_CurrTM(0) & "' AND " & _
                                              "TM02 > '" & m_CurrTM(1) & "')) AND " & _
                  "TM04 = (SELECT MIN(TM04) FROM TRADEMARK " & _
                          "WHERE TM01 = '" & m_CurrTM(0) & "' AND " & _
                                "TM02 = (SELECT MIN(TM02) FROM TRADEMARK " & _
                                        "WHERE TM01 = '" & m_CurrTM(0) & "' AND " & _
                                              "TM02 > '" & m_CurrTM(1) & "') AND " & _
                                              "TM03 = (SELECT MIN(TM03) FROM TRADEMARK " & _
                                                      "WHERE TM01 = '" & m_CurrTM(0) & "' AND " & _
                                                            "TM02 = (SELECT MIN(TM02) FROM TRADEMARK " & _
                                                                    "WHERE TM01 = '" & m_CurrTM(0) & "' AND " & _
                                                                          "TM02 > '" & m_CurrTM(1) & "'))) "
   
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("TM01")) = False Then: m_CurrTM(0) = rsTmp.Fields("TM01")
      If IsNull(rsTmp.Fields("TM02")) = False Then: m_CurrTM(1) = rsTmp.Fields("TM02")
      If IsNull(rsTmp.Fields("TM03")) = False Then: m_CurrTM(2) = rsTmp.Fields("TM03")
      If IsNull(rsTmp.Fields("TM04")) = False Then: m_CurrTM(3) = rsTmp.Fields("TM04")
   End If
   rsTmp.Close
   
   UpdateCtrlData
   
EXITSUB:
   Set rsTmp = Nothing
End Sub

' 顯示最後一筆資料
Private Sub ShowLastRecord()
   m_CurrTM(0) = m_LastTM(0)
   m_CurrTM(1) = m_LastTM(1)
   m_CurrTM(2) = m_LastTM(2)
   m_CurrTM(3) = m_LastTM(3)
   
   UpdateCtrlData
End Sub

' 更新toolbar上按紐的狀態
Private Sub UpdateToolbarState()
   Select Case m_EditMode
      ' 無任何動作
      Case 0:
         ' 90.07.13 modify by louis (依照權限設定其工具列的按紐狀態)
         'tlbar.Buttons(1).Enabled = True
         'tlbar.Buttons(2).Enabled = True
         'tlbar.Buttons(3).Enabled = True
         'tlbar.Buttons(4).Enabled = True
         'tlbar.Buttons(6).Enabled = True
         'tlbar.Buttons(7).Enabled = True
         'tlbar.Buttons(8).Enabled = True
         'tlbar.Buttons(9).Enabled = True
         'tlbar.Buttons(11).Enabled = False
         'tlbar.Buttons(12).Enabled = False
         'tlbar.Buttons(14).Enabled = True
         
         If m_bInsert Then
            tlbar.Buttons(1).Enabled = True
         Else
            tlbar.Buttons(1).Enabled = False
         End If
         If m_bUpdate Then
            tlbar.Buttons(2).Enabled = True
         Else
            tlbar.Buttons(2).Enabled = False
         End If
         If m_bDelete Then
            tlbar.Buttons(3).Enabled = True
         Else
            tlbar.Buttons(3).Enabled = False
         End If
         If m_bQuery Then
            tlbar.Buttons(4).Enabled = True
         Else
            tlbar.Buttons(4).Enabled = False
         End If
         If m_bQuery Then
            If Not IsEmptyText(m_FirstTM(0)) And Not IsEmptyText(m_FirstTM(1)) And Not IsEmptyText(m_FirstTM(2)) And Not IsEmptyText(m_FirstTM(3)) Then
               tlbar.Buttons(6).Enabled = True
               tlbar.Buttons(7).Enabled = True
            Else
               tlbar.Buttons(6).Enabled = False
               tlbar.Buttons(7).Enabled = False
            End If
            If Not IsEmptyText(m_LastTM(0)) And Not IsEmptyText(m_LastTM(1)) And Not IsEmptyText(m_LastTM(2)) And Not IsEmptyText(m_LastTM(3)) Then
               tlbar.Buttons(8).Enabled = True
               tlbar.Buttons(9).Enabled = True
            Else
               tlbar.Buttons(8).Enabled = False
               tlbar.Buttons(9).Enabled = False
            End If
         Else
            tlbar.Buttons(6).Enabled = False
            tlbar.Buttons(7).Enabled = False
            tlbar.Buttons(8).Enabled = False
            tlbar.Buttons(9).Enabled = False
         End If
         tlbar.Buttons(11).Enabled = False
         tlbar.Buttons(12).Enabled = False
         tlbar.Buttons(14).Enabled = True
         ' 新增
      Case 1, 2, 3, 4:
         tlbar.Buttons(1).Enabled = False
         tlbar.Buttons(2).Enabled = False
         tlbar.Buttons(3).Enabled = False
         tlbar.Buttons(4).Enabled = False
         tlbar.Buttons(6).Enabled = False
         tlbar.Buttons(7).Enabled = False
         tlbar.Buttons(8).Enabled = False
         tlbar.Buttons(9).Enabled = False
         tlbar.Buttons(11).Enabled = True
         tlbar.Buttons(12).Enabled = True
         tlbar.Buttons(14).Enabled = False
   End Select
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   m_FirstTM(0) = Empty
   m_FirstTM(1) = Empty
   m_FirstTM(2) = Empty
   m_FirstTM(3) = Empty
   m_CurrTM(0) = Empty
   m_CurrTM(1) = Empty
   m_CurrTM(2) = Empty
   m_CurrTM(3) = Empty
   m_LastTM(0) = Empty
   m_LastTM(1) = Empty
   m_LastTM(2) = Empty
   m_LastTM(3) = Empty
   m_EditMode = 0
   
   'Add By Sindy 2012/6/1
   If m_form Is Nothing = False Then
      m_form.Enabled = True
      If m_form.Name = "frm020101_02" Or m_form.Name = "frm030201_02" Then
         m_form.QueryMainFile
      End If
      Set m_form = Nothing
   End If
   '2012/6/1 End
   
   'Add By Cheng 2002/07/18
   Set frm020501 = Nothing
End Sub

'Add By Sindy 2024/6/14
Private Sub tabCtrl_Click(PreviousTab As Integer)
   If PreviousTab = 0 And tabCtrl.Tab = 7 Then
      Call SetTM72forCol(Left(Me.cboTM72, 1))
   End If
End Sub

'Modified by Lydia 2021/11/29 改成Form 2.0; KeyAscii As Integer=>MSForms.ReturnInteger
Private Sub textTM01_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = UpperCase(KeyAscii)
End Sub

' 系統類別
Private Sub textTM01_Validate(Cancel As Boolean)
Dim strTit As String
Dim strMsg As String
Dim nResponse
   
   Cancel = False
   textTM02_1.MaxLength = 6
   If IsEmptyText(textTM01) = False Then
      If Not IsCorrectSysKind(textTM01) Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "系統類別不正確"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textTM01_GotFocus
        'Add By Cheng 2003/05/22
        Me.textTM01.Text = ""
         GoTo EXITSUB
      End If
      
      ' 檢查使用者是否有使用該系統類別的權限
      If IsUserHasRightOfSystem(strUserNum, textTM01) = False Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "您沒有使有此系統別的權限"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textTM01_GotFocus
         GoTo EXITSUB
      End If
      
      ' 設定欄位的長度及顯示
      Select Case textTM01
         Case "T"
            textTM02_2.Locked = True
            textTM02_2.TabStop = False
            textTM02_2.Visible = False
            textTM02_1.MaxLength = 6
         Case "TF"
            textTM02_2.Locked = False
            textTM02_2.TabStop = True
            textTM02_2.Visible = True
            textTM02_1.MaxLength = 5
         Case Else
            textTM02_2.Locked = True
            textTM02_2.TabStop = False
            textTM02_2.Visible = False
            textTM02_1.MaxLength = 6
      End Select
   Else
      textTM02_2.Locked = True
      textTM02_2.TabStop = False
      textTM02_2.Visible = False
   End If
EXITSUB:
End Sub

'Modified by Lydia 2021/11/29 改成Form 2.0; KeyAscii As Integer=>MSForms.ReturnInteger
Private Sub textTM03_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = UpperCase(KeyAscii)
End Sub

' 本所案號輸入完後
Private Sub textTM04_LostFocus()
Dim strTit As String
Dim strMsg As String
Dim nResponse
Dim strTM01 As String
Dim strTM02 As String
Dim strTM03 As String
Dim strTM04 As String
   
   strTM01 = textTM01
   strTM02 = textTM02_1
   If strTM01 = "TF" Then: strTM02 = strTM02 & textTM02_2
   strTM03 = textTM03
   If IsEmptyText(strTM03) = True Then: strTM03 = "0"
   strTM04 = textTM04
   If IsEmptyText(strTM04) = True Then: strTM04 = "00"
   
   ' 新增模式下檢查資料是否已存在資料庫中
   If m_EditMode = 1 Then
      If IsRecordExist(strTM01, strTM02, strTM03, strTM04) = True Then
         strTit = "檢核資料"
         strMsg = "此筆資料已存在資料庫中"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textTM02_1.SetFocus
         GoTo EXITSUB
      End If
      ' 檢查是否超過自動編號
      If IsOverAutoNumber(strTM01, DBYEAR(SystemDate()), strTM02) = True Then
         strTit = "檢核資料"
         strMsg = "本所案號中的流水號超過自動編號"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textTM02_1.SetFocus
         GoTo EXITSUB
      End If
   End If
EXITSUB:
End Sub

' 案件名稱(中)
Private Sub textTM05_Validate(Cancel As Boolean)
Dim strTit As String
Dim strMsg As String
Dim nResponse
   
   Cancel = False
   If CheckLengthIsOK(textTM05, textTM05.MaxLength) = False Then
      Cancel = True
      textTM05_GotFocus
   End If
   'edit by nickc 2007/06/06 切換輸入法改用API
   If Cancel = False Then CloseIme
End Sub

' 商標種類
Private Sub textTM08_1_Validate(Cancel As Boolean)
Dim strTit As String
Dim strMsg As String
Dim nResponse
   
   If m_EditMode <> 1 And m_EditMode <> 2 Then Exit Sub
   
   Cancel = False
   textTM08_2 = Empty
   If IsEmptyText(textTM08_1) = False Then
      'Modify By Sindy 2015/8/13
      'textTM08_2 = GetTradeMarkName(textTM08_1, 0)
      textTM08_2 = GetTradeMarkName(textTM08_1, IIf(textTM10_1 = "020", 1, 0))
      '2015/8/13 END
      If IsEmptyText(textTM08_2) = True Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "商標種類代號不正確"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textTM08_1_GotFocus
      'Add By Sindy 2015/6/30 T,FCT商標種類為證明標章時,商品類別為證
      'modify by sonia 2020/12/28 +textTM09 = ""條件
      ElseIf textTM08_1 = "7" And (textTM01 = "T" Or textTM01 = "FCT") And textTM09 = "" Then '7證明標章
         textTM09 = "證"
      '2015/6/30 END
      'add by sonia 2020/12/28
      ElseIf textTM08_1 = "8" And (textTM01 = "T" Or textTM01 = "FCT") And textTM09 = "" Then '8團體標章
         textTM09 = "團"
      'end 2020/12/28
      '2015/6/30 END
      End If
   End If
End Sub


' 商品類別
Private Sub textTM09_Validate(Cancel As Boolean)
Dim nCount As Integer
Dim nIndex As Integer
Dim strTit As String
Dim strMsg As String
Dim strTemp As String
Dim nResponse
Dim strTempTM09 As String 'Add By Sindy 2017/10/13
   
   If m_EditMode = 4 Then Exit Sub
   textTM09 = Replace(textTM09, " ", "")
   If Trim(textTM09) = "" Then Exit Sub 'Add By Sindy 2017/10/13
   nCount = GetSubStringCount(textTM09)
   For nIndex = 1 To nCount
      'strTemp = GetSubString(textTM09, nIndex)
      strTemp = Format(GetSubString(textTM09, nIndex), "00") '補足類別格式為2碼
      strTempTM09 = strTempTM09 & strTemp & "," 'Add By Sindy 2017/10/13
      For nCount = 1 To nCount
         If nIndex <> nCount Then
            'If strTemp = GetSubString(textTM09, nCount) Then
            If strTemp = Format(GetSubString(textTM09, nCount), "00") Then
               Cancel = True
               strTit = "檢核資料"
               strMsg = "商品類別<" & strTemp & ">不可重覆"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textTM09_GotFocus
               GoTo EXITSUB
            End If
         End If
      Next nCount
   Next nIndex
   'add by nickc 2005/06/03
   strTempTM09 = Mid(strTempTM09, 1, Len(strTempTM09) - 1) 'Add By Sindy 2017/10/13
   textTM09 = strTempTM09
EXITSUB:
End Sub

' 申請國家
Private Sub textTM10_1_Validate(Cancel As Boolean)
Dim strTit As String
Dim strMsg As String
Dim nResponse
   
   Cancel = False
   textTM10_2 = Empty
   If IsEmptyText(textTM10_1) = False Then
      ' 申請國家不可輸入 001 - 008
      Select Case textTM10_1
         Case "001", "002", "003", "004", "005", "006", "007", "008":
            Cancel = True
            strTit = "檢核資料"
            strMsg = "申請國家不可輸入001-008"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textTM10_1_GotFocus
            GoTo EXITSUB
         Case Else
      End Select
      ' 取得國家代碼
      textTM10_2 = GetNationName(textTM10_1, 0)
      If IsEmptyText(textTM10_2) = True Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "國別代號不正確"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textTM10_1_GotFocus
      End If
   End If
EXITSUB:
End Sub

' 申請日
Private Sub textTM11_Validate(Cancel As Boolean)
Dim strTit As String
Dim strMsg As String
Dim nResponse
   
   Cancel = False
   If IsEmptyText(textTM11) = False Then
      If CheckIsTaiwanDate(textTM11, False) = False Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "申請日日期格式不正確"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textTM11_GotFocus
      End If
   End If
End Sub

'Add By Sindy 2015/3/25
'Modified by Lydia 2021/11/29 改成Form 2.0; KeyAscii As Integer=>MSForms.ReturnInteger
Private Sub textTM12_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = UpperCase(KeyAscii)
End Sub

'Add By Sindy 2010/9/1
Private Sub textTM12_Validate(Cancel As Boolean)
Dim strRetrunText As String 'Add By Sindy 2017/5/17
   
   If m_EditMode <> 1 And m_EditMode <> 2 Then Exit Sub
   If IsEmptyText(textTM12) = False Then
      '檢查申請案號所輸入的長度是否正確
      'Add By Sindy 2017/5/17 + strRetrunText
      If PUB_ChkTm12Tm15Length("1", textTM12, textTM01, textTM02_1 & textTM02_2, textTM03, textTM04, textTM10_1, , True, strRetrunText) = False Then
         Cancel = True
         textTM12_GotFocus
         Exit Sub
      'Add By Sindy 2017/5/17
      Else
         textTM12 = strRetrunText
      '2017/5/17 END
      End If
   End If
End Sub

Private Sub textTM121_GotFocus()
   CloseIme
   TextInverse textTM121
End Sub

'Modified by Lydia 2021/11/29 改成Form 2.0; KeyAscii As Integer=>MSForms.ReturnInteger
Private Sub textTM121_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = UpperCase(KeyAscii)
   'Modified by Morgan 2014/6/4 +可輸D
   If KeyAscii <> 8 And KeyAscii <> Asc("Y") And KeyAscii <> Asc("D") Then
      KeyAscii = 0
      Beep
   End If
End Sub
'Add by Morgan 2009/10/20
Private Sub textTM121_Validate(Cancel As Boolean)
   If (textTM121 = "" And textTM126 = "Y") Then
      MsgBox "【EMail 同時寄紙本】為 Y 時，【以EMail 通知】欄位也必須為 Y！"
      Cancel = True
   End If
End Sub

Private Sub textTM122_GotFocus()
   InverseTextBox textTM122
End Sub

'add by Toni 2008/10/21
'Modified by Lydia 2021/11/29 改成Form 2.0; KeyAscii As Integer=>MSForms.ReturnInteger
Private Sub textTM122_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 8 And KeyAscii <> Asc("Y") Then
      KeyAscii = 0
      Beep
   End If
End Sub
'end 2008/10/21

Private Sub textTM122_Validate(Cancel As Boolean)
Dim strTit As String
Dim strMsg As String
Dim nResponse
   
   Cancel = False
   If IsYesOrSpace(textTM122) = False Or textTM122 = " " Then
      Cancel = True
      strTit = "檢核資料"
      strMsg = "請輸入Y,不可輸入空白"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textTM122_GotFocus
   End If
End Sub

'Add By Sindy 2009/09/09
'Modified by Lydia 2021/11/29 改成Form 2.0; KeyAscii As Integer=>MSForms.ReturnInteger
Private Sub textTM124_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 8 And Not IsNumeric(Chr(KeyAscii)) Then
      KeyAscii = 0
      Beep
   End If
End Sub
'Modified by Lydia 2021/11/29 改成Form 2.0; KeyAscii As Integer=>MSForms.ReturnInteger
Private Sub textTM125_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 8 And Not IsNumeric(Chr(KeyAscii)) Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub textTM127_GotFocus()
   TextInverse textTM127
   CloseIme
End Sub

'Add By Sindy 2011/3/4
Private Sub textTM129_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   Select Case textTM129
      Case "Y", "":
      Case Else:
         Select Case m_EditMode
            Case 1, 2:
               Cancel = True
               strTit = "檢核資料"
               strMsg = "不催延展只可輸入Y或空白"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textTM129_GotFocus
         End Select
   End Select
End Sub
'2011/3/4 End

' 審定來函日
Private Sub textTM13_Validate(Cancel As Boolean)
Dim strTit As String
Dim strMsg As String
Dim nResponse
   
   Cancel = False
   If IsEmptyText(textTM13) = False Then
      If CheckIsTaiwanDate(textTM13, False) = False Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "審定來函日日期格式不正確"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textTM13_GotFocus
      End If
   End If
End Sub

'Add By Sindy 2015/6/30
Private Sub textTM131_GotFocus()
   InverseTextBox textTM131
   '切換輸入法改用API
   OpenIme
End Sub
' 定稿商標名稱
Private Sub textTM131_Validate(Cancel As Boolean)
Dim strTit As String
Dim strMsg As String
Dim nResponse
   
   Cancel = False
   If CheckLengthIsOK(textTM131, textTM131.MaxLength) = False Then
      Cancel = True
      textTM131_GotFocus
   End If
   '切換輸入法改用API
   If Cancel = False Then CloseIme
End Sub
'2015/6/30 END

'Add By Sindy 2024/6/14
Private Sub textTM137_GotFocus()
   InverseTextBox textTM137
   '切換輸入法改用API
   OpenIme
End Sub
' 商標描述中文
Private Sub textTM137_Validate(Cancel As Boolean)
Dim strTit As String
Dim strMsg As String
Dim nResponse
   
   Cancel = False
   If CheckLengthIsOK(textTM137, textTM137.MaxLength) = False Then
      Cancel = True
      textTM137_GotFocus
   End If
   '切換輸入法改用API
   If Cancel = False Then CloseIme
End Sub
Private Sub textTM138_GotFocus()
   InverseTextBox textTM138
   '切換輸入法改用API
   OpenIme
End Sub
' 商標描述英文
Private Sub textTM138_Validate(Cancel As Boolean)
Dim strTit As String
Dim strMsg As String
Dim nResponse
   
   Cancel = False
   If CheckLengthIsOK(textTM138, textTM138.MaxLength) = False Then
      Cancel = True
      textTM138_GotFocus
   End If
   '切換輸入法改用API
   If Cancel = False Then CloseIme
End Sub
Private Sub textTM139_GotFocus()
   InverseTextBox textTM139
   '切換輸入法改用API
   OpenIme
End Sub
' 商標描述日文
Private Sub textTM139_Validate(Cancel As Boolean)
Dim strTit As String
Dim strMsg As String
Dim nResponse
   
   Cancel = False
   If CheckLengthIsOK(textTM139, textTM139.MaxLength) = False Then
      Cancel = True
      textTM139_GotFocus
   End If
   '切換輸入法改用API
   If Cancel = False Then CloseIme
End Sub
'2024/6/14 END

'Added by Morgan 2022/12/1
Private Sub textTM136_GotFocus()
   InverseTextBox textTM136
End Sub

Private Sub textTM136_KeyPress(KeyAscii As MSForms.ReturnInteger)
   If KeyAscii <> 49 And KeyAscii <> 50 And KeyAscii <> 8 Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub textTM136_Validate(Cancel As Boolean)
Dim strTit As String
Dim strMsg As String
Dim nResponse
   
   Cancel = False
   If IsEmptyText(textTM136) = False Then
      Select Case textTM136
         Case "1", "2"
         Case Else:
            Cancel = True
            strTit = "檢核資料"
            strMsg = "請輸入 1 或 2"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textTM136_GotFocus
      End Select
   End If
End Sub
'end 2022/12/1

' 公告日
Private Sub textTM14_Validate(Cancel As Boolean)
Dim strTit As String
Dim strMsg As String
Dim nResponse
   
   Cancel = False
   If IsEmptyText(textTM14) = False Then
      If CheckIsTaiwanDate(textTM14, False) = False Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "公告日日期格式不正確"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textTM14_GotFocus
      End If
   End If
End Sub

'Add By Sindy 2015/3/25
'Modified by Lydia 2021/11/29 改成Form 2.0; KeyAscii As Integer=>MSForms.ReturnInteger
Private Sub textTM15_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = UpperCase(KeyAscii)
End Sub

' 審定號
Private Sub textTM15_Validate(Cancel As Boolean)
Dim strSql As String
Dim rsTmp As ADODB.Recordset
Dim strTit As String
Dim strMsg As String
Dim nResponse
Dim strRetrunText As String 'Add By Sindy 2017/5/17
   
   If m_EditMode <> 1 And m_EditMode <> 2 Then Exit Sub
   Cancel = False
   textSP32 = Empty
   If IsEmptyText(textTM15) = False Then
      strSql = "SELECT * FROM SERVICEPRACTICE " & _
               "WHERE SP32 = '" & textTM15 & "' AND " & _
                     "SP01 = '" & "TM" & "' "
      Set rsTmp = New ADODB.Recordset
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         rsTmp.MoveFirst
         textSP32 = rsTmp.Fields("SP01") & "-" & rsTmp.Fields("SP02") & "-" & rsTmp.Fields("SP03") & "-" & rsTmp.Fields("SP04")
      End If
      rsTmp.Close
      Set rsTmp = Nothing
      'Modify By Sindy 2010/9/1
      '檢查審定號所輸入的長度是否正確
      If bolNewAppNoFormat Then
         '2011/1/14 MODIFY BY SONIA 台灣核駁審定號0+6碼數字
         'If PUB_ChkTm12Tm15Length("2", textTM15, textTM01, textTM02_1 & textTM02_2, textTM03, textTM04, textTM10_1) = False Then
         'Add By Sindy 2017/5/17 + strRetrunText
         If PUB_ChkTm12Tm15Length("2", textTM15, textTM01, textTM02_1 & textTM02_2, textTM03, textTM04, textTM10_1, "" & textTM16, True, strRetrunText) = False Then
            Cancel = True
            textTM15_GotFocus
            Exit Sub
         'Add By Sindy 2017/5/17
         Else
            textTM15 = strRetrunText
         '2017/5/17 END
         End If
      Else
         '2008/12/4 add by sonia
         If textTM10_1 = "000" Then
            If GetTextLength(Trim(textTM15)) <> 8 Then
               Cancel = True
               strTit = "檢核資料"
               strMsg = "申請國家為台灣,審定號只可為8碼數字,不足8碼請在前面補0 !"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textTM15_GotFocus
   '2010/6/28 CANCEL BY SONIA T-139794
   '         ElseIf Not IsNumeric(textTM15) Then
   '            Cancel = True
   '            strTit = "檢核資料"
   '            strMsg = "申請國家為台灣,審定號只可為8碼數字,不足8碼請在前面補0 !"
   '            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
   '            textTM15_GotFocus
   '2010/6/28 END
           End If
         End If
         '2008/12/4 end
      End If
   End If
End Sub

' 目前准駁
Private Sub textTM16_Validate(Cancel As Boolean)
Dim strTit As String
Dim strMsg As String
Dim nResponse
   
   Cancel = False
   If IsEmptyText(textTM16) = False Then
      Select Case textTM16
         Case "", "1", "2":
         Case Else:
            Cancel = True
            strTit = "檢核資料"
            strMsg = "請輸入 1 或 2 "
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textTM16_GotFocus
      End Select
   End If
   'add by sonia 2022/10/12
   If textTM16 = "2" And textTM10_1 = "020" And textTM15 <> "" Then
      Cancel = True
      strTit = "檢核資料"
      strMsg = "大陸案核駁沒有審定號, 目前准駁不可輸入 2 !"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textTM16_GotFocus
   End If
   'end 2022/10/12
   '2011/4/28 ADD BY SONIA
   If IsEmptyText(textTM22) = False Then
      If textTM16 <> "1" Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "有專用期, 目前准駁請輸入 1 !"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textTM16_GotFocus
      End If
   End If
   '2011/4/28 END
End Sub

'Modified by Lydia 2021/11/29 改成Form 2.0; KeyAscii As Integer=>MSForms.ReturnInteger
Private Sub textTM17_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = UpperCase(KeyAscii)
End Sub

' 專用權是否存在
Private Sub textTM17_Validate(Cancel As Boolean)
Dim strTit As String
Dim strMsg As String
Dim nResponse
   
   Cancel = False
   If IsEmptyText(textTM17) = False Then
      Select Case textTM17
         Case "Y", "N", "":
         Case Else:
            Cancel = True
            strTit = "檢核資料"
            strMsg = "專用權是否存在只可輸入Y或N"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textTM17_GotFocus
      End Select
   '2008/3/28 add by sonia
   Else
      If IsEmptyText(textTM22) = False Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "有專用期間，專用權是否存在欄不可空白!"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textTM17_GotFocus
      End If
   '2008/3/28 end
   End If
End Sub

'Modified by Lydia 2021/11/29 改成Form 2.0; KeyAscii As Integer=>MSForms.ReturnInteger
Private Sub textTM18_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = UpperCase(KeyAscii)
End Sub

' 是否有救濟程序
Private Sub textTM18_Validate(Cancel As Boolean)
Dim strTit As String
Dim strMsg As String
Dim nResponse
   
   Cancel = False
   If IsYesOrSpace(textTM18) = False Then
      Cancel = True
      strTit = "檢核資料"
      strMsg = "請輸入Y或空白"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textTM18_GotFocus
   End If
End Sub

'Modified by Lydia 2021/11/29 改成Form 2.0; KeyAscii As Integer=>MSForms.ReturnInteger
Private Sub textTM19_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = UpperCase(KeyAscii)
End Sub

' 是否有爭議程序
Private Sub textTM19_Validate(Cancel As Boolean)
Dim strTit As String
Dim strMsg As String
Dim nResponse
   
   Cancel = False
   If IsYesOrSpace(textTM19) = False Then
      Cancel = True
      strTit = "檢核資料"
      strMsg = "請輸入Y或空白"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textTM19_GotFocus
   End If
End Sub
' 發證日
Private Sub textTM20_Validate(Cancel As Boolean)
Dim strTit As String
Dim strMsg As String
Dim nResponse
   
   Cancel = False
   If IsEmptyText(textTM20) = False Then
      If CheckIsTaiwanDate(textTM20, False) = False Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "發證日日期格式不正確"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textTM20_GotFocus
      End If
   End If
End Sub
' 專用期限 (起)
Private Sub textTM21_Validate(Cancel As Boolean)
Dim strTit As String
Dim strMsg As String
Dim nResponse
   
   Cancel = False
   If IsEmptyText(textTM21) = False Then
      If CheckIsDate(textTM21, False) = False Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "專用期限起日格式不正確"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textTM21_GotFocus
         Exit Sub          '2011/9/28 add by sonia
      End If
      '2011/9/28 add by sonia
      If textTM28 <> "申請" Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "卷宗性質非申請, 不可輸入專用期限"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textTM21_GotFocus
         Exit Sub
      End If
      '2011/9/28 end
   End If
End Sub
' 專用期限 (迄)
Private Sub textTM22_Validate(Cancel As Boolean)
Dim strTit As String
Dim strMsg As String
Dim nResponse
   
   Cancel = False
   If IsEmptyText(textTM22) = False Then
      If CheckIsDate(textTM22, False) = False Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "專用期限止日格式不正確"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textTM22_GotFocus
      End If
      '2011/9/28 add by sonia
      If textTM28 <> "申請" Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "卷宗性質非申請, 不可輸入專用期限"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textTM22_GotFocus
         Exit Sub
      End If
      '2011/9/28 end
   End If
   '2008/9/11 ADD BY SONIA
   If textTM21 <> "" And textTM22 <> "" Then
      If Not ChkRange(textTM21, textTM22, "專用期限") Then
         Cancel = True
         textTM22.SetFocus
      End If
   End If
   '2008/9/11 END

   'add by sonia 2022/10/12 補輸專用期時上核准及專用權存在
   If textTM22.Tag = "" And textTM22 <> "" Then
      If textTM16 = "" Then
         textTM16 = "1"
      End If
      If textTM17 = "" Then
         textTM17 = "Y"
      End If
   End If
   'end 2022/10/12
   
End Sub

'Modified by Lydia 2021/11/29 改成Form 2.0; KeyAscii As Integer=>MSForms.ReturnInteger
Private Sub textTM23_1_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = UpperCase(KeyAscii)
End Sub

' 申請人欄位
Private Sub textTM23_1_Validate(Cancel As Boolean)
Dim strTit As String
Dim strMsg As String
Dim nResponse
Dim rsTmp As ADODB.Recordset
Dim strKey As String
Dim strSql As String
Dim strData As String
       
    'Modify By Cheng 2003/08/07
    '若有更改申請人
    If Left(Me.textTM23_1.Text & "000000000", 9) <> m_TM23 Then
       Cancel = False
       textTM23_2 = Empty
       textCU72 = Empty
       textCU79 = Empty
        'Add By Cheng 2003/08/07
       textTM24 = Empty
       textTM25 = Empty
       textTM26 = Empty
       ' 不滿九碼補0
       If IsEmptyText(textTM23_1) = False Then
          strData = textTM23_1 & String(9 - Len(textTM23_1), "0")
          Select Case Mid(strData, 1, 1)
          Case "X", "x":
             Set rsTmp = New ADODB.Recordset
             If Len(strData) > 8 Then
                strSql = "SELECT * FROM Customer " & _
                         "WHERE CU01 = '" & Mid(strData, 1, 8) & "' AND " & _
                               "CU02 = '" & Mid(strData, 9, 1) & "'"
             Else
                strSql = "SELECT * FROM Customer " & _
                         "WHERE CU01 = '" & Mid(strData, 1, 8) & "'"
             End If
             rsTmp.CursorLocation = adUseClient
             rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly   'edit by nickc 2005/08/04 原先是  動態開啟
             If rsTmp.RecordCount > 0 Then
                rsTmp.MoveFirst
                If IsNull(rsTmp.Fields("CU04")) = False Then
                   textTM23_2 = rsTmp.Fields("CU04")
                ElseIf IsNull(rsTmp.Fields("CU05")) = False Then
                   textTM23_2 = rsTmp.Fields("CU05")
                ElseIf IsNull(rsTmp.Fields("CU06")) = False Then
                   textTM23_2 = rsTmp.Fields("CU06")
                End If
                If IsNull(rsTmp.Fields("CU79")) = False Then
                   textCU79 = rsTmp.Fields("CU79")
                End If
                If IsNull(rsTmp.Fields("CU72")) = False Then
                   textCU72 = rsTmp.Fields("CU72")
                End If
                ' 帶出中英日地址
                If IsNull(rsTmp.Fields("CU23")) = False Then
                   textTM24 = rsTmp.Fields("CU23")
                End If
                If IsNull(rsTmp.Fields("CU24")) = False Then
    '               textTM25 = rsTmp.Fields("CU24")
                   textTM25 = rsTmp.Fields("CU24") & _
                               IIf(IsNull(rsTmp.Fields("CU25")), "", " " & rsTmp.Fields("CU25")) & _
                               IIf(IsNull(rsTmp.Fields("CU26")), "", " " & rsTmp.Fields("CU26")) & _
                               IIf(IsNull(rsTmp.Fields("CU27")), "", " " & rsTmp.Fields("CU27")) & _
                               IIf(IsNull(rsTmp.Fields("CU28")), "", " " & rsTmp.Fields("CU28")) & _
                               IIf(IsNull(rsTmp.Fields("CU102")), "", " " & rsTmp.Fields("CU102"))
                End If
                If IsNull(rsTmp.Fields("CU29")) = False Then
                   textTM26 = rsTmp.Fields("CU29")
                End If
             Else
                Cancel = True
                strTit = "檢核資料"
                strMsg = "申請人代號不存在"
                nResponse = MsgBox(strMsg, vbOKOnly, strTit)
                textTM23_1_GotFocus
             End If
             rsTmp.Close
          Case "Y", "y":
             Set rsTmp = New ADODB.Recordset
             If Len(strData) > 8 Then
                strSql = "SELECT * FROM FAGENT " & _
                         "WHERE FA01 = '" & Mid(strData, 1, 8) & "' AND " & _
                               "FA02 = '" & Mid(strData, 9, 1) & "'"
             Else
                strSql = "SELECT * FROM FAGENT " & _
                         "WHERE FA01 = '" & Mid(strData, 1, 8) & "' AND " & _
                               "FA02 = '0' "
             End If
             rsTmp.CursorLocation = adUseClient
             rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly   'edit by nickc 2005/08/04 原先是  動態開啟
             ' 檢查讀取的資料筆數
             If rsTmp.RecordCount > 0 Then
                rsTmp.MoveFirst
                If IsNull(rsTmp.Fields("FA03")) = False Then
                   strKey = rsTmp.Fields("FA03")
                   textTM23_1 = strKey
                   rsTmp.Close
                   If Len(strKey) > 8 Then
                      strSql = "SELECT * FROM Customer " & _
                            "WHERE CU01 = '" & Mid(strKey, 1, 8) & "' AND " & _
                                  "CU02 = '" & Mid(strKey, 9, 1) & "'"
                   Else
                      strSql = "SELECT * FROM Customer " & _
                            "WHERE CU01 = '" & Mid(strKey, 1, 8) & "' AND " & _
                                  "CU02 = '0' "
                   End If
                   rsTmp.CursorLocation = adUseClient
                   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly   'ok edit by nickc 2005/08/04 原先是  動態開啟
                   If rsTmp.RecordCount > 0 Then
                      rsTmp.MoveFirst
                      If IsNull(rsTmp.Fields("CU04")) = False Then
                         textTM23_2 = rsTmp.Fields("CU04")
                      ElseIf IsNull(rsTmp.Fields("CU05")) = False Then
                         textTM23_2 = rsTmp.Fields("CU05")
                      ElseIf IsNull(rsTmp.Fields("CU06")) = False Then
                         textTM23_2 = rsTmp.Fields("CU06")
                      End If
                      If IsNull(rsTmp.Fields("CU79")) = False Then
                         textCU79 = rsTmp.Fields("CU79")
                      End If
                      If IsNull(rsTmp.Fields("CU72")) = False Then
                         textCU72 = rsTmp.Fields("CU72")
                      End If
                      ' 帶出中英日地址
                      If IsNull(rsTmp.Fields("CU23")) = False Then
                         textTM24 = rsTmp.Fields("CU23")
                      End If
                      If IsNull(rsTmp.Fields("CU24")) = False Then
                         textTM25 = rsTmp.Fields("CU24")
                      End If
                      If IsNull(rsTmp.Fields("CU29")) = False Then
                         textTM26 = rsTmp.Fields("CU29")
                      End If
                   Else
                      Cancel = True
                      strTit = "檢核資料"
                      strMsg = "申請人代號不存在"
                      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
                      textTM23_1_GotFocus
                   End If
                Else
                   Cancel = True
                   strTit = "檢核資料"
                   strMsg = "申請人代號不存在"
                   nResponse = MsgBox(strMsg, vbOKOnly, strTit)
                   textTM23_1_GotFocus
                End If
             Else
                Cancel = True
                strTit = "檢核資料"
                strMsg = "申請人代號不存在"
                nResponse = MsgBox(strMsg, vbOKOnly, strTit)
                textTM23_1_GotFocus
             End If
             rsTmp.Close
          Case Else:
             Cancel = True
             strTit = "檢核資料"
             strMsg = "申請人代號不正確"
             nResponse = MsgBox(strMsg, vbOKOnly, strTit)
             textTM23_1_GotFocus
          End Select
       End If
       Set rsTmp = Nothing
    End If
End Sub

' 申請地址(中)
Private Sub textTM24_Validate(Cancel As Boolean)
Dim strTit As String
Dim strMsg As String
Dim nResponse
   
   Cancel = False
   'edit by nickc 2007/05/03 長度不符
   'If CheckLengthIsOK(textTM24, 70) = False Then
   If CheckLengthIsOK(textTM24, textTM24.MaxLength) = False Then
      Cancel = True
      textTM24_GotFocus
   End If
   'edit by nickc 2007/06/06 切換輸入法改用API
   'If Cancel = False Then: textTM24.IMEMode = 2
   If Cancel = False Then CloseIme
End Sub

' 申請地址(日)
Private Sub textTM26_Validate(Cancel As Boolean)
Dim strTit As String
Dim strMsg As String
Dim nResponse
   
   Cancel = False
   If CheckLengthIsOK(textTM26, textTM26.MaxLength) = False Then
      Cancel = True
      textTM26_GotFocus
   End If
   'edit by nickc 2007/06/06 切換輸入法改用API
   'If Cancel = False Then: textTM26.IMEMode = 2
   If Cancel = False Then CloseIme
End Sub

'Modified by Lydia 2021/11/29 改成Form 2.0; KeyAscii As Integer=>MSForms.ReturnInteger
Private Sub textTM29_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = UpperCase(KeyAscii)
End Sub

' 是否閉卷
Private Sub textTM29_Validate(Cancel As Boolean)
Dim strTit As String
Dim strMsg As String
Dim nResponse
   
   Cancel = False
   If IsYesOrSpace(textTM29) = False Then
      Cancel = True
      strTit = "檢核資料"
      strMsg = "請輸入Y或空白"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textTM29_GotFocus
   End If
End Sub
' 閉卷日期
Private Sub textTM30_Validate(Cancel As Boolean)
Dim strTit As String
Dim strMsg As String
Dim nResponse
   
   Cancel = False
   If IsEmptyText(textTM30) = False Then
      If CheckIsTaiwanDate(textTM30, False) = False Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "閉卷日期日期格式不正確"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textTM30_GotFocus
      End If
   End If
End Sub

' 商品組群
Private Sub textTM32_Validate(Cancel As Boolean)
Dim strTit As String
Dim strMsg As String
Dim nResponse
Dim nCount As Integer
Dim nIndex As Integer
Dim strSql As String
Dim strTemp As String
   
   Cancel = False
   ' 無資料時不做任何檢查
   If IsEmptyText(textTM32) = True And m_EditMode = 4 Then
      GoTo EXITSUB
   End If
   
   'Modify By Sindy 2024/4/18 商品組群欄人員貼上資料後將全形或半形的「；」分號，轉為半形的逗號存入TM32。
   textTM32 = Replace(Replace(textTM32, ";", ","), "；", ",")
   '2024/4/18 END
   nCount = GetSubStringCount(textTM32)
   For nIndex = 1 To nCount
      strTemp = GetSubString(textTM32, nIndex)
      If Len(strTemp) > 6 Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "商品組群<" & strTemp & ">不正確"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textTM32_GotFocus
         GoTo EXITSUB
      End If
   Next nIndex
   
   For nIndex = 1 To nCount
      strTemp = GetSubString(textTM32, nIndex)
      For nCount = 1 To nCount
         If nIndex <> nCount Then
            If strTemp = GetSubString(textTM32, nCount) Then
               Cancel = True
               strTit = "檢核資料"
               strMsg = "商品組群<" & strTemp & ">不可重覆"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textTM32_GotFocus
               GoTo EXITSUB
            End If
         End If
      Next nCount
   Next nIndex
   
EXITSUB:
End Sub

'Modified by Lydia 2021/11/29 改成Form 2.0; KeyAscii As Integer=>MSForms.ReturnInteger
Private Sub textTM33_1_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = UpperCase(KeyAscii)
End Sub

'Modified by Lydia 2021/11/29 改成Form 2.0; KeyAscii As Integer=>MSForms.ReturnInteger
Private Sub textTM34_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textTM34_Validate(Cancel As Boolean)
Dim strTit As String
Dim strMsg As String
Dim nResponse
   
   Cancel = False
   If CheckLengthIsOK(textTM34, textTM34.MaxLength) = False Then
      Cancel = True
      strTit = "檢核資料"
      strMsg = "分所號內容太長"
      '911111 nick
      textTM34.SetFocus
      
      textTM34_GotFocus
   End If
End Sub

' 聯絡人1(中)
Private Sub textTM38_Validate(Cancel As Boolean)
Dim strTit As String
Dim strMsg As String
Dim nResponse
   
   Cancel = False
    'Modified by Lydia 2017/06/14 聯絡人(中)改為30字
    'If CheckLengthIsOK(textTM38, textTM38.MaxLength) = False Then
    If CheckLengthIsOK(textTM38, 30) = False Then
      Cancel = True
      textTM38_GotFocus
   End If
   'edit by nickc 2007/06/06 切換輸入法改用API
   'If Cancel = False Then: textTM38.IMEMode = 2
   If Cancel = False Then CloseIme
End Sub

' 聯絡人1(日)
Private Sub textTM40_Validate(Cancel As Boolean)
Dim strTit As String
Dim strMsg As String
Dim nResponse
   
   Cancel = False
   'Modified by Lydia 2017/06/14
   'If CheckLengthIsOK(textTM40, textTM40.MaxLength) = False Then
   If CheckLengthIsOK(textTM40, 60) = False Then
      Cancel = True
      textTM40_GotFocus
   End If
   'edit by nickc 2007/06/06 切換輸入法改用API
   'If Cancel = False Then: textTM40.IMEMode = 2
   If Cancel = False Then CloseIme
End Sub

' 聯絡人2(中)
Private Sub textTM41_Validate(Cancel As Boolean)
Dim strTit As String
Dim strMsg As String
Dim nResponse
   
   Cancel = False
   'Modified by Lydia 2017/06/14 聯絡人(中)改為30字
   'If CheckLengthIsOK(textTM41, textTM41.MaxLength) = False Then
   If CheckLengthIsOK(textTM41, 30) = False Then
      Cancel = True
      textTM41_GotFocus
   End If
   'edit by nickc 2007/06/06 切換輸入法改用API
   'If Cancel = False Then: textTM41.IMEMode = 2
   If Cancel = False Then CloseIme
End Sub

' 聯絡人2(日)
Private Sub textTM43_Validate(Cancel As Boolean)
Dim strTit As String
Dim strMsg As String
Dim nResponse
   
   Cancel = False
   'Modified by Lydia 2017/06/14
   'If CheckLengthIsOK(textTM43, textTM43.MaxLength) = False Then
   If CheckLengthIsOK(textTM43, 60) = False Then
      Cancel = True
      textTM43_GotFocus
   End If
   'edit by nickc 2007/06/06 切換輸入法改用API
   'If Cancel = False Then: textTM43.IMEMode = 2
   If Cancel = False Then CloseIme
End Sub

'Modified by Lydia 2021/11/29 改成Form 2.0; KeyAscii As Integer=>MSForms.ReturnInteger
Private Sub textTM44_1_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = UpperCase(KeyAscii)
End Sub

'Modified by Lydia 2021/11/29 改成Form 2.0; KeyAscii As Integer=>MSForms.ReturnInteger
Private Sub textTM46_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = UpperCase(KeyAscii)
End Sub

' D/N是否列印申請人
Private Sub textTM46_Validate(Cancel As Boolean)
Dim strTit As String
Dim strMsg As String
Dim nResponse
   
   Cancel = False
   If IsYesOrSpace(textTM46) = False Then
      Cancel = True
      strTit = "檢核資料"
      strMsg = "請輸入Y或空白"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textTM46_GotFocus
   End If
End Sub

' 代表人1(中)
Private Sub textTM47_Validate(Cancel As Boolean)
Dim strTit As String
Dim strMsg As String
Dim nResponse
   
   Cancel = False
   If CheckLengthIsOK(textTM47, textTM47.MaxLength) = False Then
      Cancel = True
      textTM47_GotFocus
   End If
   'edit by nickc 2007/06/06 切換輸入法改用API
   'If Cancel = False Then: textTM47.IMEMode = 2
   If Cancel = False Then CloseIme
End Sub

' 代表人1(日)
Private Sub textTM49_Validate(Cancel As Boolean)
Dim strTit As String
Dim strMsg As String
Dim nResponse
   
   Cancel = False
   If CheckLengthIsOK(textTM49, textTM49.MaxLength) = False Then
      Cancel = True
      textTM49_GotFocus
   End If
   'edit by nickc 2007/06/06 切換輸入法改用API
   'If Cancel = False Then: textTM49.IMEMode = 2
   If Cancel = False Then CloseIme
End Sub

' 代表人2(中)
Private Sub textTM50_Validate(Cancel As Boolean)
Dim strTit As String
Dim strMsg As String
Dim nResponse
   
   Cancel = False
   If CheckLengthIsOK(textTM50, textTM50.MaxLength) = False Then
      Cancel = True
      textTM50_GotFocus
   End If
   'edit by nickc 2007/06/06 切換輸入法改用API
   'If Cancel = False Then: textTM50.IMEMode = 2
   If Cancel = False Then CloseIme
End Sub

' 代表人2(日)
Private Sub textTM52_Validate(Cancel As Boolean)
Dim strTit As String
Dim strMsg As String
Dim nResponse
   
   Cancel = False
   If CheckLengthIsOK(textTM52, textTM52.MaxLength) = False Then
      Cancel = True
      textTM52_GotFocus
   End If
   'edit by nickc 2007/06/06 切換輸入法改用API
   'If Cancel = False Then: textTM52.IMEMode = 2
   If Cancel = False Then CloseIme
End Sub

'Modified by Lydia 2021/11/29 改成Form 2.0; KeyAscii As Integer=>MSForms.ReturnInteger
Private Sub textTM54_1_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = UpperCase(KeyAscii)
End Sub

'Modified by Lydia 2021/11/29 改成Form 2.0; KeyAscii As Integer=>MSForms.ReturnInteger
Private Sub textTM56_1_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = UpperCase(KeyAscii)
End Sub

'Modified by Lydia 2021/11/29 改成Form 2.0; KeyAscii As Integer=>MSForms.ReturnInteger
Private Sub textTM57_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = UpperCase(KeyAscii)
End Sub

' 是否銷卷
Private Sub textTM57_Validate(Cancel As Boolean)
Dim strTit As String
Dim strMsg As String
Dim nResponse
   
   Cancel = False
   'Remark by Lydia 2021/12/01 銷卷日期為數字
   'If IsYesOrSpace(textTM57) = False Then
   '   Cancel = True
   '   strTit = "檢核資料"
   '   strMsg = "請輸入Y或空白"
   '   nResponse = MsgBox(strMsg, vbOKOnly, strTit)
   '   textTM57_GotFocus
   'End If
   'end 2021/12/01
End Sub
' 閉卷原因
Private Sub textTM31_1_Validate(Cancel As Boolean)
Dim strTit As String
Dim strMsg As String
Dim nResponse
Dim rsTmp As ADODB.Recordset
Dim strSql As String
   
   Cancel = False
   textTM31_2 = Empty
   If IsEmptyText(textTM31_1) = False Then
      Set rsTmp = New ADODB.Recordset
      strSql = "SELECT * FROM ReasonOfRelief " & _
               "WHERE ROR01 = '" & textTM31_1 & "'"
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly   'ok edit by nickc 2005/08/04 原先是  動態開啟
      ' 檢查讀取的資料筆數
      If rsTmp.RecordCount > 0 Then
         rsTmp.MoveFirst
         If IsNull(rsTmp.Fields("ROR02")) = False Then
            textTM31_2 = rsTmp.Fields("ROR02")
         End If
      Else
         Cancel = True
         strTit = "檢核資料"
         strMsg = "閉卷原因不存在"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textTM31_1_GotFocus
      End If
      rsTmp.Close
   End If
   Set rsTmp = Nothing
End Sub
' FC代理人
Private Sub textTM44_1_Validate(Cancel As Boolean)
Dim strTit As String
Dim strMsg As String
Dim nResponse
Dim strAgent As String
Dim rsTmp As ADODB.Recordset
Dim strKey As String
Dim strSql As String
Dim strData As String
   
   Cancel = False
   textTM44_2 = Empty
   textFA29 = Empty
   textFA39 = Empty
   If IsEmptyText(textTM44_1) = False Then
      strData = textTM44_1 & String(9 - Len(textTM44_1), "0")
      Select Case Mid(strData, 1, 1)
      Case "X", "x":
         Set rsTmp = New ADODB.Recordset
         If Len(strData) > 8 Then
            strSql = "SELECT * FROM Customer " & _
                     "WHERE CU01 = '" & Mid(strData, 1, 8) & "' AND " & _
                           "CU02 = '" & Mid(strData, 9, 1) & "'"
         Else
            strSql = "SELECT * FROM Customer " & _
                     "WHERE CU01 = '" & Mid(strData, 1, 8) & "' AND " & _
                           "CU02 = '" & "0" & "' "
         End If
         rsTmp.CursorLocation = adUseClient
         rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly   'edit by nickc 2005/08/04 原先是  動態開啟
         If rsTmp.RecordCount > 0 Then
            rsTmp.MoveFirst
            If IsNull(rsTmp.Fields("CU03")) = False Then
               strKey = rsTmp.Fields("CU03")
               textTM44_1 = strKey
               rsTmp.Close
               If Len(strKey) > 8 Then
                  strSql = "SELECT * FROM FAGENT " & _
                           "WHERE FA01 = '" & Mid(strKey, 1, 8) & "' AND" & _
                                 "FA02 = '" & Mid(strKey, 9, 1) & "'"
               Else
                  strSql = "SELECT * FROM FAGENT " & _
                           "WHERE FA01 = '" & Mid(strKey, 1, 8) & "'"
               End If
               rsTmp.CursorLocation = adUseClient
               rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly   'edit by nickc 2005/08/04 原先是  動態開啟
               If rsTmp.RecordCount > 0 Then
                  rsTmp.MoveFirst
                  If IsNull(rsTmp.Fields("FA05")) = False Then
                     textTM44_2 = rsTmp.Fields("FA05")
                  ElseIf IsNull(rsTmp.Fields("FA04")) = False Then
                     textTM44_2 = rsTmp.Fields("FA04")
                  ElseIf IsNull(rsTmp.Fields("FA06")) = False Then
                     textTM44_2 = rsTmp.Fields("FA06")
                  End If
                  If IsNull(rsTmp.Fields("FA29")) = False Then
                     textFA29 = rsTmp.Fields("FA29")
                  End If
                  If IsNull(rsTmp.Fields("FA39")) = False Then
                     textFA39 = rsTmp.Fields("FA39")
                  End If
               Else
                  Cancel = True
                  strTit = "檢核資料"
                  strMsg = "FC代理人代號不存在"
                  nResponse = MsgBox(strMsg, vbOKOnly, strTit)
                  textTM44_1_GotFocus
               End If
            Else
               Cancel = True
               strTit = "檢核資料"
               strMsg = "FC代理人代號不存在"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textTM44_1_GotFocus
            End If
         Else
            Cancel = True
            strTit = "檢核資料"
            strMsg = "FC代理人代號不存在"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textTM44_1_GotFocus
         End If
         rsTmp.Close
      Case "Y", "y":
         Set rsTmp = New ADODB.Recordset
         If Len(strData) > 8 Then
            strSql = "SELECT * FROM FAGENT " & _
                     "WHERE FA01 = '" & Mid(strData, 1, 8) & "' AND " & _
                           "FA02 = '" & Mid(strData, 9, 1) & "'"
         Else
            strSql = "SELECT * FROM FAGENT " & _
                     "WHERE FA01 = '" & Mid(strData, 1, 8) & "'"
         End If
         rsTmp.CursorLocation = adUseClient
         rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly   'edit by nickc 2005/08/04 原先是  動態開啟
         If rsTmp.RecordCount > 0 Then
            rsTmp.MoveFirst
'edit by nickc 2006/07/06 秀玲說代理人名稱顯示方式跟專利相同
'            If IsNull(rsTmp.Fields("FA05")) = False Then
'               textTM44_2 = rsTmp.Fields("FA05")
'            ElseIf IsNull(rsTmp.Fields("FA04")) = False Then
'               textTM44_2 = rsTmp.Fields("FA04")
'            ElseIf IsNull(rsTmp.Fields("FA06")) = False Then
'               textTM44_2 = rsTmp.Fields("FA06")
'            End If
            Dim CheckIn As Boolean
            Dim strTempName As String
            CheckIn = PUB_GetAgentName(textTM01.Text, strData, strTempName)
            textTM44_2 = strTempName
            If IsNull(rsTmp.Fields("FA29")) = False Then
               textFA29 = rsTmp.Fields("FA29")
            End If
            If IsNull(rsTmp.Fields("FA39")) = False Then
               textFA39 = rsTmp.Fields("FA39")
            End If
         Else
            Cancel = True
            strTit = "檢核資料"
            strMsg = "FC代理人代號不存在"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textTM44_1_GotFocus
         End If
         rsTmp.Close
      Case Else:
         Cancel = True
         strTit = "檢核資料"
         strMsg = "FC代理人代號不正確"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textTM44_1_GotFocus
      End Select
   End If
   Set rsTmp = Nothing
End Sub
' 定稿語文
Private Sub textTM53_Validate(Cancel As Boolean)
Dim strTit As String
Dim strMsg As String
Dim nResponse
   
   Cancel = False
   If IsEmptyText(textTM53) = False Then
      Select Case textTM53
         Case "1", "2", "3":
         Case Else:
            Cancel = True
            strTit = "檢核資料"
            strMsg = "請輸入 1 或 2 或 3"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textTM53_GotFocus
      End Select
   End If
End Sub

'Added by Lydia 2023/11/14
Private Sub textTM72_1_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = UpperCase(KeyAscii)
End Sub

'Add By Sindy 2009/09/09
'Modified by Lydia 2021/11/29 改成Form 2.0; KeyAscii As Integer=>MSForms.ReturnInteger
Private Sub textTM77_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = UpperCase(KeyAscii)
End Sub

'Add By Sindy 2009/09/09
' 畫面上定稿語言
Private Sub textTM77_Validate(Cancel As Boolean)
Dim strTit As String
Dim strMsg As String
Dim nResponse
   
   Cancel = False
   If IsEmptyText(textTM77) = False Then
      Select Case textTM77
         Case "N", "1", "2", "3":
         Case Else:
            Cancel = True
            strTit = "檢核資料"
            strMsg = "請輸入 N 或 1 或 2 或 3"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textTM77_GotFocus
      End Select
   End If
End Sub

' 延展通知人
Private Sub textTM33_1_Validate(Cancel As Boolean)
Dim strTemp As String
Dim strTit As String
Dim strMsg As String
Dim nResponse
'Add By Cheng 2002/07/09
Dim strTempName As String
   
   Cancel = False
   If IsEmptyText(textTM33_1) = False Then
      strTemp = textTM33_1
      ' 不滿八碼補0
      If Len(strTemp) < 8 Then: strTemp = strTemp & String(8 - Len(strTemp), "0")
      'Modify By Cheng 2002/07/09
'      textTM33_2 = GetAgentOrCustName(strTemp)
      If Left(Me.textTM33_1.Text, 1) = "X" Then
         textTM33_2 = GetAgentOrCustName(strTemp)
      Else
         If PUB_GetAgentName(Me.textTM01.Text, strTemp, strTempName) Then
            textTM33_2 = strTempName
         Else
            textTM33_2 = ""
         End If
      End If
      If IsEmptyText(textTM33_2) = True Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "延展代理人代號不存在"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textTM33_1_GotFocus
      End If
   End If
End Sub

' 副本收受人
Private Sub textTM54_1_Validate(Cancel As Boolean)
Dim strTemp As String
Dim strTit As String
Dim strMsg As String
Dim nResponse
'Add By Cheng 2002/07/09
Dim strTempName As String
   
   Cancel = False
   If IsEmptyText(textTM54_1) = False Then
      strTemp = textTM54_1
      If Len(strTemp) < 8 Then: strTemp = strTemp & String(8 - Len(strTemp), "0")
      'Modify By Cheng 2002/07/09
'      textTM54_2 = GetAgentOrCustName(strTemp)
      If Left(Me.textTM54_1.Text, 1) = "X" Then
         textTM54_2 = GetAgentOrCustName(strTemp)
      Else
         If PUB_GetAgentName(Me.textTM01.Text, strTemp, strTempName) Then
            textTM54_2 = strTempName
         Else
            textTM54_2 = ""
         End If
      End If
      
      If IsEmptyText(textTM54_2) = True Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "副本收受人代號不存在"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textTM54_1_GotFocus
      End If
   End If
End Sub

' 固定請款對象
Private Sub textTM56_1_Validate(Cancel As Boolean)
Dim strTemp As String
Dim strTit As String
Dim strMsg As String
Dim nResponse
'Add By Cheng 2002/07/09
Dim strTempName As String
   
   Cancel = False
   If IsEmptyText(textTM56_1) = False Then
      strTemp = textTM56_1
      If Len(strTemp) < 8 Then: strTemp = strTemp & String(8 - Len(strTemp), "0")
      'Modify By Cheng 2002/07/09
'      textTM56_2 = GetAgentOrCustName(strTemp)
      If Left(Me.textTM56_1.Text, 1) = "X" Then
         textTM56_2 = GetAgentOrCustName(strTemp)
      Else
         If PUB_GetAgentName(Me.textTM01.Text, strTemp, strTempName) Then
            textTM56_2 = strTempName
         Else
            textTM56_2 = ""
         End If
      End If
      
      If IsEmptyText(textTM56_2) = True Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "固定請款對象代號不存在"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textTM56_1_GotFocus
      End If
   End If
End Sub

' 案件備註
Private Sub textTM58_Validate(Cancel As Boolean)
Dim strTit As String
Dim strMsg As String
Dim nResponse
   
   Cancel = False
   If CheckLengthIsOK(textTM58, textTM58.MaxLength) = False Then
      Cancel = True
      textTM58_GotFocus
   End If
   'edit by nickc 2007/06/06 切換輸入法改用API
   'If Cancel = False Then: textTM58.IMEMode = 2
   If Cancel = False Then CloseIme
End Sub

'Modified by Lydia 2021/11/29 改成Form 2.0; KeyAscii As Integer=>MSForms.ReturnInteger
Private Sub textTM66_1_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = UpperCase(KeyAscii)
End Sub

' 延展請款對象
Private Sub textTM66_1_Validate(Cancel As Boolean)
Dim strTemp As String
Dim strTit As String
Dim strMsg As String
Dim nResponse
'Add By Cheng 2002/07/09
Dim strTempName As String
   
   Cancel = False
   If IsEmptyText(textTM66_1) = False Then
      strTemp = textTM66_1
      ' 不滿八碼補0
      If Len(strTemp) < 8 Then: strTemp = strTemp & String(8 - Len(strTemp), "0")
      'Modify By Cheng 2002/07/09
'      textTM66_2 = GetAgentOrCustName(strTemp)
      If Left(Me.textTM66_1.Text, 1) = "X" Then
         textTM66_2 = GetAgentOrCustName(strTemp)
      Else
         If PUB_GetAgentName(Me.textTM01.Text, strTemp, strTempName) Then
            textTM66_2 = strTempName
         Else
            textTM66_2 = ""
         End If
      End If
      
      If IsEmptyText(textTM66_2) = True Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "延展請款對象代號不存在"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textTM66_1_GotFocus
      End If
   End If
End Sub

' 放棄專用權
Private Sub textTM67_Validate(Cancel As Boolean)
Dim strTit As String
Dim strMsg As String
Dim nResponse
   
   Cancel = False
    'Modify By Cheng 2003/05/05
'   If CheckLengthIsOK(textTM67, 40) = False Then
   If CheckLengthIsOK(textTM67, Me.textTM67.MaxLength) = False Then
      Cancel = True
      textTM67_GotFocus
   End If
   'edit by nickc 2007/06/06 切換輸入法改用API
   'If Cancel = False Then: textTM67.IMEMode = 2
   If Cancel = False Then CloseIme
End Sub

'Modified by Lydia 2021/11/29 改成Form 2.0; KeyAscii As Integer=>MSForms.ReturnInteger
Private Sub textTM68_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = UpperCase(KeyAscii)
End Sub

'Add By Sindy 2009/09/09
'Modified by Lydia 2021/11/29 改成Form 2.0; KeyAscii As Integer=>MSForms.ReturnInteger
Private Sub textTM126_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = UpperCase(KeyAscii)
End Sub

'Add By Sindy 2013/8/15
'Modified by Lydia 2021/11/29 改成Form 2.0; KeyAscii As Integer=>MSForms.ReturnInteger
Private Sub textTM129_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = UpperCase(KeyAscii)
End Sub

' 延展單筆不跑
Private Sub textTM68_Validate(Cancel As Boolean)
Dim strTit As String
Dim strMsg As String
Dim nResponse
   
   Cancel = False
   If IsEmptyText(textTM68) = False Then
      If IsYesOrSpace(textTM68) = False Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "延展單筆不跑請輸入Y或空白"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textTM68_GotFocus
      End If
   End If
End Sub

'Add By Sindy 2009/09/09
' EMail同時寄紙本
Private Sub textTM126_Validate(Cancel As Boolean)
Dim strTit As String
Dim strMsg As String
Dim nResponse
   
   Cancel = False
   If IsEmptyText(textTM126) = False Then
      If IsYesOrSpace(textTM126) = False Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "EMail同時寄紙本請輸入Y或空白"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textTM126_GotFocus
      End If
   End If
   'Add by Morgan 2009/10/20
   If Cancel = False Then
      If (textTM121 = "" And textTM126 = "Y") Then
         MsgBox "【EMail 同時寄紙本】為 Y 時，【以EMail 通知】欄位也必須為 Y！"
         Cancel = True
      End If
   End If
   'end 2009/10/20
End Sub

' 取得代理人名稱
Private Function GetAgentName(ByVal strData As String) As String
Dim rsTmp As ADODB.Recordset
Dim strKey As String
Dim strSql As String
   
   GetAgentName = Empty
   If IsEmptyText(strData) = False Then
      ' 不滿8碼自動補0
      If Len(strData) < 8 Then: strData = strData & String(8 - Len(strData), "0")
      Select Case Mid(strData, 1, 1)
      Case "X", "x":
         Set rsTmp = New ADODB.Recordset
         If Len(strData) > 8 Then
            strSql = "SELECT * FROM Customer " & _
                     "WHERE CU01 = '" & Mid(strData, 1, 8) & "' AND " & _
                           "CU02 = '" & Mid(strData, 9, 1) & "'"
         Else
            strSql = "SELECT * FROM Customer " & _
                     "WHERE CU01 = '" & Mid(strData, 1, 8) & "' AND " & _
                           "CU02 = '0' "
         End If
         rsTmp.CursorLocation = adUseClient
         rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly   'edit by nickc 2005/08/04 原先是  動態開啟
         If rsTmp.RecordCount > 0 Then
            rsTmp.MoveFirst
            If IsNull(rsTmp.Fields("CU03")) = False Then
               strKey = rsTmp.Fields("CU03")
               rsTmp.Close
               If Len(strKey) > 8 Then
                  strSql = "SELECT * FROM FAGENT " & _
                           "WHERE FA01 = '" & Mid(strKey, 1, 8) & "' AND" & _
                                 "FA02 = '" & Mid(strKey, 9, 1) & "'"
               Else
                  strSql = "SELECT * FROM FAGENT " & _
                           "WHERE FA01 = '" & Mid(strKey, 1, 8) & "' AND " & _
                                 "FA02 = '0' "
               End If
               rsTmp.CursorLocation = adUseClient
               rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly   'edit by nickc 2005/08/04 原先是  動態開啟
               If rsTmp.RecordCount > 0 Then
                  rsTmp.MoveFirst
                  If IsNull(rsTmp.Fields("FA05")) = False Then
                     GetAgentName = rsTmp.Fields("FA05")
                  ElseIf IsNull(rsTmp.Fields("FA04")) = False Then
                     GetAgentName = rsTmp.Fields("FA04")
                  ElseIf IsNull(rsTmp.Fields("FA06")) = False Then
                     GetAgentName = rsTmp.Fields("FA06")
                  End If
               End If
            End If
         End If
         rsTmp.Close
      Case "Y", "y":
         Set rsTmp = New ADODB.Recordset
         If Len(strData) > 8 Then
            strSql = "SELECT * FROM FAGENT " & _
                     "WHERE FA01 = '" & Mid(strData, 1, 8) & "' AND " & _
                           "FA02 = '" & Mid(strData, 9, 1) & "'"
         Else
            strSql = "SELECT * FROM FAGENT " & _
                     "WHERE FA01 = '" & Mid(strData, 1, 8) & "' AND " & _
                           "FA02 = '0' "
         End If
         rsTmp.CursorLocation = adUseClient
         rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly   'edit by nickc 2005/08/04 原先是  動態開啟
         If rsTmp.RecordCount > 0 Then
            rsTmp.MoveFirst
            If IsNull(rsTmp.Fields("FA05")) = False Then
               GetAgentName = rsTmp.Fields("FA05")
            ElseIf IsNull(rsTmp.Fields("FA04")) = False Then
               GetAgentName = rsTmp.Fields("FA04")
            ElseIf IsNull(rsTmp.Fields("FA06")) = False Then
               GetAgentName = rsTmp.Fields("FA06")
            End If
         End If
         rsTmp.Close
      End Select
   End If
   Set rsTmp = Nothing
End Function

' 取得客戶名稱
Private Function GetCustName(ByVal strData As String) As String
Dim rsTmp As ADODB.Recordset
Dim strKey As String
Dim strSql As String
   
   GetCustName = Empty
   If IsEmptyText(strData) = False Then
      ' 不滿8碼自動補0
      If Len(strData) < 8 Then: strData = strData & String(8 - Len(strData), "0")
      Select Case Mid(strData, 1, 1)
      Case "X", "x":
         Set rsTmp = New ADODB.Recordset
         If Len(strData) > 8 Then
            strSql = "SELECT * FROM Customer " & _
                     "WHERE CU01 = '" & Mid(strData, 1, 8) & "' AND " & _
                           "CU02 = '" & Mid(strData, 9, 1) & "'"
         Else
            strSql = "SELECT * FROM Customer " & _
                     "WHERE CU01 = '" & Mid(strData, 1, 8) & "' AND " & _
                           "CU02 = '0' "
         End If
         rsTmp.CursorLocation = adUseClient
         rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly   'edit by nickc 2005/08/04 原先是  動態開啟
         If rsTmp.RecordCount > 0 Then
            rsTmp.MoveFirst
            If IsNull(rsTmp.Fields("CU04")) = False Then
               GetCustName = rsTmp.Fields("CU04")
            ElseIf IsNull(rsTmp.Fields("CU05")) = False Then
               GetCustName = rsTmp.Fields("CU05")
            ElseIf IsNull(rsTmp.Fields("CU06")) = False Then
               GetCustName = rsTmp.Fields("CU06")
            End If
         End If
         rsTmp.Close
      Case "Y", "y":
         Set rsTmp = New ADODB.Recordset
         If Len(strData) > 8 Then
            strSql = "SELECT * FROM FAGENT " & _
                     "WHERE FA01 = '" & Mid(strData, 1, 8) & "' AND " & _
                           "FA02 = '" & Mid(strData, 9, 1) & "'"
         Else
            strSql = "SELECT * FROM FAGENT " & _
                     "WHERE FA01 = '" & Mid(strData, 1, 8) & "' AND " & _
                           "FA02 = '0' "
         End If
         rsTmp.CursorLocation = adUseClient
         rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly   'edit by nickc 2005/08/04 原先是  動態開啟
         ' 檢查讀取的資料筆數
         If rsTmp.RecordCount > 0 Then
            rsTmp.MoveFirst
            If IsNull(rsTmp.Fields("FA03")) = False Then
               strKey = rsTmp.Fields("FA03")
               rsTmp.Close
               If Len(strKey) > 8 Then
                  strSql = "SELECT * FROM Customer " & _
                        "WHERE CU01 = '" & Mid(strKey, 1, 8) & "' AND " & _
                              "CU02 = '" & Mid(strKey, 9, 1) & "'"
               Else
                  strSql = "SELECT * FROM Customer " & _
                        "WHERE CU01 = '" & Mid(strKey, 1, 8) & "' AND " & _
                              "CU02 = '0' "
               End If
               rsTmp.CursorLocation = adUseClient
               rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly   'edit by nickc 2005/08/04 原先是  動態開啟
               If rsTmp.RecordCount > 0 Then
                  rsTmp.MoveFirst
                  If IsNull(rsTmp.Fields("CU04")) = False Then
                     GetCustName = rsTmp.Fields("CU04")
                  ElseIf IsNull(rsTmp.Fields("CU05")) = False Then
                     GetCustName = rsTmp.Fields("CU05")
                  ElseIf IsNull(rsTmp.Fields("CU06")) = False Then
                     GetCustName = rsTmp.Fields("CU06")
                  End If
               End If
            End If
         End If
         rsTmp.Close
      End Select
   End If
   Set rsTmp = Nothing
End Function
' 取得客戶或是代理人名稱
Private Function GetAgentOrCustName(ByVal strData As String) As String
Dim rsTmp As ADODB.Recordset
Dim strSql As String
   
   GetAgentOrCustName = Empty
   If IsEmptyText(strData) = False Then
      ' 不滿8碼自動補0
      If Len(strData) < 8 Then: strData = strData & String(8 - Len(strData), "0")
      Select Case Mid(strData, 1, 1)
      Case "X", "x":
         Set rsTmp = New ADODB.Recordset
         If Len(strData) > 8 Then
            strSql = "SELECT * FROM Customer " & _
                     "WHERE CU01 = '" & Mid(strData, 1, 8) & "' AND " & _
                           "CU02 = '" & Mid(strData, 9, 1) & "'"
         Else
            strSql = "SELECT * FROM Customer " & _
                     "WHERE CU01 = '" & Mid(strData, 1, 8) & "' AND " & _
                           "CU02 = '0' "
         End If
         rsTmp.CursorLocation = adUseClient
         rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly   'edit by nickc 2005/08/04 原先是  動態開啟
         If rsTmp.RecordCount > 0 Then
            rsTmp.MoveFirst
            If IsNull(rsTmp.Fields("CU05")) = False Then
               GetAgentOrCustName = rsTmp.Fields("CU05")
            ElseIf IsNull(rsTmp.Fields("CU04")) = False Then
               GetAgentOrCustName = rsTmp.Fields("CU04")
            ElseIf IsNull(rsTmp.Fields("CU06")) = False Then
               GetAgentOrCustName = rsTmp.Fields("CU06")
            End If
         End If
         rsTmp.Close
      Case "Y", "y":
         Set rsTmp = New ADODB.Recordset
         If Len(strData) > 8 Then
            strSql = "SELECT * FROM FAGENT " & _
                     "WHERE FA01 = '" & Mid(strData, 1, 8) & "' AND " & _
                           "FA02 = '" & Mid(strData, 9, 1) & "'"
         Else
            strSql = "SELECT * FROM FAGENT " & _
                     "WHERE FA01 = '" & Mid(strData, 1, 8) & "' AND " & _
                           "FA02 = '0' "
         End If
         rsTmp.CursorLocation = adUseClient
         rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly   'edit by nickc 2005/08/04 原先是  動態開啟
         If rsTmp.RecordCount > 0 Then
            rsTmp.MoveFirst
            If IsNull(rsTmp.Fields("FA05")) = False Then
               GetAgentOrCustName = rsTmp.Fields("FA05")
            ElseIf IsNull(rsTmp.Fields("FA04")) = False Then
               GetAgentOrCustName = rsTmp.Fields("FA04")
            ElseIf IsNull(rsTmp.Fields("FA06")) = False Then
               GetAgentOrCustName = rsTmp.Fields("FA06")
            End If
         End If
         rsTmp.Close
      End Select
   End If
   Set rsTmp = Nothing
End Function

' 檢查是否為Y或空白
Private Function IsYesOrSpace(ByVal strData As String) As Boolean
Dim strTit As String
Dim strMsg As String
Dim nResponse
   
   IsYesOrSpace = False
   Select Case strData
      Case "", "Y", " ":
         IsYesOrSpace = True
      Case Else:
         IsYesOrSpace = False
   End Select
End Function

' 按下按鍵
'edit by nickc 2006/12/07
'Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Public Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'Memo by Lydia 2021/11/29 原程式搬到Form_KeyUp
End Sub

' 執行指令
Private Sub OnAction(ByVal KeyCode As Integer)
Dim strTit As String
Dim strMsg As String
Dim nResponse
   
   m_SubMode = 0
   Select Case KeyCode
      ' 新增
      Case vbKeyF2:
         m_EditMode = 1
         ClearField
         SetCtrlReadOnly False
         UpdateToolbarState
         SetInputEntry
      ' 修改
      Case vbKeyF3:
         m_EditMode = 2
         SetCtrlReadOnly False
         SetKeyReadOnly True
         'Added by Lydia 2022/09/13 外商承辦F11判斷人員職稱等級決定是否鎖住「閉卷」
         If bolUpdClose = False Then
             textTM29.Locked = True
             textTM30.Locked = True
             textTM31_1.Locked = True
         End If
         'end 2022/09/13
         UpdateToolbarState
         SetInputEntry
         'add by sonia 2025/5/14 T案證明標章、團體標章不可輸入商品及服務
         If textTM01 = "T" And (cboTM08 = "7  證明標章" Or cboTM08 = "8  團體標章") Then
            Command2.Enabled = False
         Else
            Command2.Enabled = True
         End If
         'end 2025/5/14
      ' 刪除
      Case vbKeyF5:
         If IsCaseProgressExist(textTM01, textTM02_1 & textTM02_2, textTM03, textTM04) = True Then
            strTit = "檢核資料"
            strMsg = "此本所案號在案件進度檔中仍有資料, 不可刪除!"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         Else
            'Add By Sindy 2010/7/1
            If ChkCaseCode("NP", textTM01, textTM02_1 & textTM02_2, textTM03, textTM04) = False Then Exit Sub
            '2010/7/1 End
            strTit = "詢問"
            strMsg = "是否要刪除此筆資料?"
            nResponse = MsgBox(strMsg, vbYesNo + vbCritical + vbDefaultButton2, strTit)
            If nResponse = vbYes Then
               m_EditMode = 3
               OnWork
               UpdateToolbarState
            End If
         End If
      ' 查詢
      Case vbKeyF4:
         m_EditMode = 4
         SetCtrlReadOnly True
         SetKeyReadOnly False
         ClearField
         UpdateToolbarState
         SetInputEntry
      ' 第一筆
      Case vbKeyHome:
         ShowFirstRecord
      ' 前一筆
      Case vbKeyPageUp:
         ShowPrevRecord
      ' 後一筆
      Case vbKeyPageDown:
         ShowNextRecord
      ' 最後一筆
      Case vbKeyEnd:
         ShowLastRecord
      ' 確定
      Case vbKeyF9:
         ' 將所有欄位的內容更新到欄位串列中的欄位內容項目
'edit by nickc 2008/03/28 還沒檢查完資料就先更新，有些資料在檢查時才上，會更新不到
'         UpdateFieldNewData
         OnWork
         UpdateToolbarState
      ' 取消
      Case vbKeyF10:
         Select Case m_EditMode
            Case 1, 2:
               strTit = "詢問"
               strMsg = "你並未存檔, 確定離開嗎?"
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
      ' 離開
      Case vbKeyEscape:
         Unload Me
   End Select
End Sub

Private Sub textTM69_1_GotFocus()
    TextInverse Me.textTM69_1
End Sub

'Modified by Lydia 2021/11/29 改成Form 2.0; KeyAscii As Integer=>MSForms.ReturnInteger
Private Sub textTM69_1_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textTM69_1_Validate(Cancel As Boolean)
Dim strTemp As String
Dim strTit As String
Dim strMsg As String
Dim nResponse
'Add By Cheng 2002/07/09
Dim strTempName As String
   
   Cancel = False
   If IsEmptyText(textTM69_1) = False Then
      strTemp = textTM69_1
      ' 不滿八碼補0
      If Len(strTemp) < 8 Then: strTemp = strTemp & String(8 - Len(strTemp), "0")
      'Modify By Cheng 2002/07/09
'      textTM69_2 = GetAgentOrCustName(strTemp)
      If Left(Me.textTM69_1.Text, 1) = "X" Then
         textTM69_2 = GetAgentOrCustName(strTemp)
      Else
         If PUB_GetAgentName(Me.textTM01.Text, strTemp, strTempName) Then
            textTM69_2 = strTempName
         Else
            textTM69_2 = ""
         End If
      End If
      If IsEmptyText(textTM69_2) = True Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "D/N固定列印對象代號不存在"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textTM69_1_GotFocus
      End If
   End If
End Sub

Private Sub textTM70_1_GotFocus()
    TextInverse Me.textTM70_1
End Sub

'Modified by Lydia 2021/11/29 改成Form 2.0; KeyAscii As Integer=>MSForms.ReturnInteger
Private Sub textTM70_1_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textTM70_1_Validate(Cancel As Boolean)
Dim strTemp As String
Dim strTit As String
Dim strMsg As String
Dim nResponse
'Add By Cheng 2002/07/09
Dim strTempName As String
   
   Cancel = False
   If IsEmptyText(textTM70_1) = False Then
      strTemp = textTM70_1
      ' 不滿八碼補0
      If Len(strTemp) < 8 Then: strTemp = strTemp & String(8 - Len(strTemp), "0")
      'Modify By Cheng 2002/07/09
'      textTM70_2 = GetAgentOrCustName(strTemp)
      If Left(Me.textTM70_1.Text, 1) = "X" Then
         textTM70_2 = GetAgentOrCustName(strTemp)
      Else
         If PUB_GetAgentName(Me.textTM01.Text, strTemp, strTempName) Then
            textTM70_2 = strTempName
         Else
            textTM70_2 = ""
         End If
      End If
      If IsEmptyText(textTM70_2) = True Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "延展D/N列印對象代號不存在"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textTM70_1_GotFocus
      End If
   End If
End Sub

Private Sub textTM71_1_GotFocus()
    TextInverse Me.textTM71_1
End Sub

Private Sub textTM72_1_GotFocus()
    TextInverse Me.textTM72_1
End Sub

Private Sub textTM72_1_Validate(Cancel As Boolean)
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
    
    'Modified by Lydia 2023/11/16 +Trim
    If Trim(Me.textTM72_1.Text) <> "" Then
        Me.textTM72_2.Text = PUB_GetSpecialPTName("2", Me.textTM72_1.Text)
        If Me.textTM72_2.Text = "" Then
            MsgBox "特殊商標代碼輸入錯誤!!!", vbExclamation + vbOKOnly
            Cancel = True
        End If
    Else
        Me.textTM72_1.Text = "" 'Added by Lydia 2023/11/16
        Me.textTM72_2.Text = ""
    End If
    If Cancel = True Then textTM72_1_GotFocus
End Sub

Private Sub textTM76_GotFocus()
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textTM76.IMEMode = 1
   OpenIme
   InverseTextBox textTM76
End Sub

Private Sub textTM76_Validate(Cancel As Boolean)
Dim strTit As String
Dim strMsg As String
Dim nResponse
   
   Cancel = False
   'Modified by Lydia 2017/06/14
   'If CheckLengthIsOK(textTM76, textTM76.MaxLength) = False Then
   If CheckLengthIsOK(textTM76, 60) = False Then
      Cancel = True
      textTM76_GotFocus
   End If
   'edit by nickc 2007/06/06 切換輸入法改用API
   'If Cancel = False Then: textTM76.IMEMode = 2
   If Cancel = False Then CloseIme
End Sub

'add by nickc 2006/12/11
Private Sub textTM78_1_Change()
   '若申請人為空白, 自動清除相關地址及代表人
   If Me.textTM78_1.Text <> "" Then Exit Sub
   Me.textTM82.Text = Empty:   Me.textTM86.Text = Empty
   Me.textTM90.Text = Empty
   Me.textTM94.Text = Empty:   Me.textTM95.Text = Empty
   Me.textTM96.Text = Empty:   Me.textTM97.Text = Empty
   Me.textTM98.Text = Empty:   Me.textTM99.Text = Empty
End Sub
Private Sub textTM78_1_GotFocus()
   InverseTextBox textTM78_1
End Sub

'Modified by Lydia 2021/11/29 改成Form 2.0; KeyAscii As Integer=>MSForms.ReturnInteger
Private Sub textTM78_1_KeyPress(KeyAscii As MSForms.ReturnInteger)
KeyAscii = UpperCase(KeyAscii)
End Sub
Private Sub textTM78_1_Validate(Cancel As Boolean)
Dim strTit As String
Dim strMsg As String
Dim nResponse
Dim rsTmp As ADODB.Recordset
Dim strKey As String
Dim strSql As String
Dim strData As String
       
    'Modify By Cheng 2003/08/07
    '若有更改申請人
    If Left(Me.textTM78_1.Text & "000000000", 9) <> m_TM78 Then
       Cancel = False
       textTM78_2 = Empty
       textTM82 = Empty
       textTM86 = Empty
       textTM90 = Empty
       ' 不滿九碼補0
       If IsEmptyText(textTM78_1) = False Then
          strData = textTM78_1 & String(9 - Len(textTM78_1), "0")
          Select Case Mid(strData, 1, 1)
          Case "X", "x":
             Set rsTmp = New ADODB.Recordset
             If Len(strData) > 8 Then
                strSql = "SELECT * FROM Customer " & _
                         "WHERE CU01 = '" & Mid(strData, 1, 8) & "' AND " & _
                               "CU02 = '" & Mid(strData, 9, 1) & "'"
             Else
                strSql = "SELECT * FROM Customer " & _
                         "WHERE CU01 = '" & Mid(strData, 1, 8) & "'"
             End If
             rsTmp.CursorLocation = adUseClient
             rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
             If rsTmp.RecordCount > 0 Then
                rsTmp.MoveFirst
                If IsNull(rsTmp.Fields("CU04")) = False Then
                   textTM78_2 = rsTmp.Fields("CU04")
                ElseIf IsNull(rsTmp.Fields("CU05")) = False Then
                   textTM78_2 = rsTmp.Fields("CU05")
                ElseIf IsNull(rsTmp.Fields("CU06")) = False Then
                   textTM78_2 = rsTmp.Fields("CU06")
                End If
                ' 帶出中英日地址
                If IsNull(rsTmp.Fields("CU23")) = False Then
                   textTM82 = rsTmp.Fields("CU23")
                End If
                If IsNull(rsTmp.Fields("CU24")) = False Then
                   textTM86 = rsTmp.Fields("CU24") & _
                               IIf(IsNull(rsTmp.Fields("CU25")), "", " " & rsTmp.Fields("CU25")) & _
                               IIf(IsNull(rsTmp.Fields("CU26")), "", " " & rsTmp.Fields("CU26")) & _
                               IIf(IsNull(rsTmp.Fields("CU27")), "", " " & rsTmp.Fields("CU27")) & _
                               IIf(IsNull(rsTmp.Fields("CU28")), "", " " & rsTmp.Fields("CU28")) & _
                               IIf(IsNull(rsTmp.Fields("CU102")), "", " " & rsTmp.Fields("CU102"))
                End If
                If IsNull(rsTmp.Fields("CU29")) = False Then
                   textTM90 = rsTmp.Fields("CU29")
                End If
             Else
                Cancel = True
                strTit = "檢核資料"
                strMsg = "申請人代號不存在"
                nResponse = MsgBox(strMsg, vbOKOnly, strTit)
                textTM78_1_GotFocus
             End If
             rsTmp.Close
          Case "Y", "y":
             Set rsTmp = New ADODB.Recordset
             If Len(strData) > 8 Then
                strSql = "SELECT * FROM FAGENT " & _
                         "WHERE FA01 = '" & Mid(strData, 1, 8) & "' AND " & _
                               "FA02 = '" & Mid(strData, 9, 1) & "'"
             Else
                strSql = "SELECT * FROM FAGENT " & _
                         "WHERE FA01 = '" & Mid(strData, 1, 8) & "' AND " & _
                               "FA02 = '0' "
             End If
             rsTmp.CursorLocation = adUseClient
             rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
             ' 檢查讀取的資料筆數
             If rsTmp.RecordCount > 0 Then
                rsTmp.MoveFirst
                If IsNull(rsTmp.Fields("FA03")) = False Then
                   strKey = rsTmp.Fields("FA03")
                   textTM78_1 = strKey
                   rsTmp.Close
                   If Len(strKey) > 8 Then
                      strSql = "SELECT * FROM Customer " & _
                            "WHERE CU01 = '" & Mid(strKey, 1, 8) & "' AND " & _
                                  "CU02 = '" & Mid(strKey, 9, 1) & "'"
                   Else
                      strSql = "SELECT * FROM Customer " & _
                            "WHERE CU01 = '" & Mid(strKey, 1, 8) & "' AND " & _
                                  "CU02 = '0' "
                   End If
                   rsTmp.CursorLocation = adUseClient
                   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
                   If rsTmp.RecordCount > 0 Then
                      rsTmp.MoveFirst
                      If IsNull(rsTmp.Fields("CU04")) = False Then
                         textTM78_2 = rsTmp.Fields("CU04")
                      ElseIf IsNull(rsTmp.Fields("CU05")) = False Then
                         textTM78_2 = rsTmp.Fields("CU05")
                      ElseIf IsNull(rsTmp.Fields("CU06")) = False Then
                         textTM78_2 = rsTmp.Fields("CU06")
                      End If
                      ' 帶出中英日地址
                      If IsNull(rsTmp.Fields("CU23")) = False Then
                         textTM82 = rsTmp.Fields("CU23")
                      End If
                      If IsNull(rsTmp.Fields("CU24")) = False Then
                         textTM86 = rsTmp.Fields("CU24")
                      End If
                      If IsNull(rsTmp.Fields("CU29")) = False Then
                         textTM90 = rsTmp.Fields("CU29")
                      End If
                   Else
                      Cancel = True
                      strTit = "檢核資料"
                      strMsg = "申請人代號不存在"
                      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
                      textTM78_1_GotFocus
                   End If
                Else
                   Cancel = True
                   strTit = "檢核資料"
                   strMsg = "申請人代號不存在"
                   nResponse = MsgBox(strMsg, vbOKOnly, strTit)
                   textTM78_1_GotFocus
                End If
             Else
                Cancel = True
                strTit = "檢核資料"
                strMsg = "申請人代號不存在"
                nResponse = MsgBox(strMsg, vbOKOnly, strTit)
                textTM78_1_GotFocus
             End If
             rsTmp.Close
          Case Else:
             Cancel = True
             strTit = "檢核資料"
             strMsg = "申請人代號不正確"
             nResponse = MsgBox(strMsg, vbOKOnly, strTit)
             textTM78_1_GotFocus
          End Select
       End If
       Set rsTmp = Nothing
    End If
End Sub
Private Sub textTM79_1_Change()
   '若申請人為空白, 自動清除相關地址及代表人
   If Me.textTM79_1.Text <> "" Then Exit Sub
   Me.textTM83.Text = Empty:   Me.textTM87.Text = Empty
   Me.textTM91.Text = Empty
   Me.textTM100.Text = Empty:   Me.textTM101.Text = Empty
   Me.textTM102.Text = Empty:   Me.textTM103.Text = Empty
   Me.textTM104.Text = Empty:   Me.textTM105.Text = Empty
End Sub
Private Sub textTM79_1_GotFocus()
   InverseTextBox textTM79_1
End Sub

'Modified by Lydia 2021/11/29 改成Form 2.0; KeyAscii As Integer=>MSForms.ReturnInteger
Private Sub textTM79_1_KeyPress(KeyAscii As MSForms.ReturnInteger)
KeyAscii = UpperCase(KeyAscii)
End Sub
Private Sub textTM79_1_Validate(Cancel As Boolean)
Dim strTit As String
Dim strMsg As String
Dim nResponse
Dim rsTmp As ADODB.Recordset
Dim strKey As String
Dim strSql As String
Dim strData As String
       
    'Modify By Cheng 2003/08/07
    '若有更改申請人
    If Left(Me.textTM79_1.Text & "000000000", 9) <> m_TM79 Then
       Cancel = False
       textTM79_2 = Empty
       textTM83 = Empty
       textTM87 = Empty
       textTM91 = Empty
       ' 不滿九碼補0
       If IsEmptyText(textTM79_1) = False Then
          strData = textTM79_1 & String(9 - Len(textTM79_1), "0")
          Select Case Mid(strData, 1, 1)
          Case "X", "x":
             Set rsTmp = New ADODB.Recordset
             If Len(strData) > 8 Then
                strSql = "SELECT * FROM Customer " & _
                         "WHERE CU01 = '" & Mid(strData, 1, 8) & "' AND " & _
                               "CU02 = '" & Mid(strData, 9, 1) & "'"
             Else
                strSql = "SELECT * FROM Customer " & _
                         "WHERE CU01 = '" & Mid(strData, 1, 8) & "'"
             End If
             rsTmp.CursorLocation = adUseClient
             rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
             If rsTmp.RecordCount > 0 Then
                rsTmp.MoveFirst
                If IsNull(rsTmp.Fields("CU04")) = False Then
                   textTM79_2 = rsTmp.Fields("CU04")
                ElseIf IsNull(rsTmp.Fields("CU05")) = False Then
                   textTM79_2 = rsTmp.Fields("CU05")
                ElseIf IsNull(rsTmp.Fields("CU06")) = False Then
                   textTM79_2 = rsTmp.Fields("CU06")
                End If
                ' 帶出中英日地址
                If IsNull(rsTmp.Fields("CU23")) = False Then
                   textTM83 = rsTmp.Fields("CU23")
                End If
                If IsNull(rsTmp.Fields("CU24")) = False Then
                   textTM87 = rsTmp.Fields("CU24") & _
                               IIf(IsNull(rsTmp.Fields("CU25")), "", " " & rsTmp.Fields("CU25")) & _
                               IIf(IsNull(rsTmp.Fields("CU26")), "", " " & rsTmp.Fields("CU26")) & _
                               IIf(IsNull(rsTmp.Fields("CU27")), "", " " & rsTmp.Fields("CU27")) & _
                               IIf(IsNull(rsTmp.Fields("CU28")), "", " " & rsTmp.Fields("CU28")) & _
                               IIf(IsNull(rsTmp.Fields("CU102")), "", " " & rsTmp.Fields("CU102"))
                End If
                If IsNull(rsTmp.Fields("CU29")) = False Then
                   textTM91 = rsTmp.Fields("CU29")
                End If
             Else
                Cancel = True
                strTit = "檢核資料"
                strMsg = "申請人代號不存在"
                nResponse = MsgBox(strMsg, vbOKOnly, strTit)
                textTM79_1_GotFocus
             End If
             rsTmp.Close
          Case "Y", "y":
             Set rsTmp = New ADODB.Recordset
             If Len(strData) > 8 Then
                strSql = "SELECT * FROM FAGENT " & _
                         "WHERE FA01 = '" & Mid(strData, 1, 8) & "' AND " & _
                               "FA02 = '" & Mid(strData, 9, 1) & "'"
             Else
                strSql = "SELECT * FROM FAGENT " & _
                         "WHERE FA01 = '" & Mid(strData, 1, 8) & "' AND " & _
                               "FA02 = '0' "
             End If
             rsTmp.CursorLocation = adUseClient
             rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
             ' 檢查讀取的資料筆數
             If rsTmp.RecordCount > 0 Then
                rsTmp.MoveFirst
                If IsNull(rsTmp.Fields("FA03")) = False Then
                   strKey = rsTmp.Fields("FA03")
                   textTM79_1 = strKey
                   rsTmp.Close
                   If Len(strKey) > 8 Then
                      strSql = "SELECT * FROM Customer " & _
                            "WHERE CU01 = '" & Mid(strKey, 1, 8) & "' AND " & _
                                  "CU02 = '" & Mid(strKey, 9, 1) & "'"
                   Else
                      strSql = "SELECT * FROM Customer " & _
                            "WHERE CU01 = '" & Mid(strKey, 1, 8) & "' AND " & _
                                  "CU02 = '0' "
                   End If
                   rsTmp.CursorLocation = adUseClient
                   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
                   If rsTmp.RecordCount > 0 Then
                      rsTmp.MoveFirst
                      If IsNull(rsTmp.Fields("CU04")) = False Then
                         textTM79_2 = rsTmp.Fields("CU04")
                      ElseIf IsNull(rsTmp.Fields("CU05")) = False Then
                         textTM79_2 = rsTmp.Fields("CU05")
                      ElseIf IsNull(rsTmp.Fields("CU06")) = False Then
                         textTM79_2 = rsTmp.Fields("CU06")
                      End If
                      ' 帶出中英日地址
                      If IsNull(rsTmp.Fields("CU23")) = False Then
                         textTM83 = rsTmp.Fields("CU23")
                      End If
                      If IsNull(rsTmp.Fields("CU24")) = False Then
                         textTM87 = rsTmp.Fields("CU24")
                      End If
                      If IsNull(rsTmp.Fields("CU29")) = False Then
                         textTM91 = rsTmp.Fields("CU29")
                      End If
                   Else
                      Cancel = True
                      strTit = "檢核資料"
                      strMsg = "申請人代號不存在"
                      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
                      textTM79_1_GotFocus
                   End If
                Else
                   Cancel = True
                   strTit = "檢核資料"
                   strMsg = "申請人代號不存在"
                   nResponse = MsgBox(strMsg, vbOKOnly, strTit)
                   textTM79_1_GotFocus
                End If
             Else
                Cancel = True
                strTit = "檢核資料"
                strMsg = "申請人代號不存在"
                nResponse = MsgBox(strMsg, vbOKOnly, strTit)
                textTM79_1_GotFocus
             End If
             rsTmp.Close
          Case Else:
             Cancel = True
             strTit = "檢核資料"
             strMsg = "申請人代號不正確"
             nResponse = MsgBox(strMsg, vbOKOnly, strTit)
             textTM79_1_GotFocus
          End Select
       End If
       Set rsTmp = Nothing
    End If
End Sub
Private Sub textTM80_1_Change()
   '若申請人為空白, 自動清除相關地址及代表人
   If Me.textTM80_1.Text <> "" Then Exit Sub
   Me.textTM84.Text = Empty:   Me.textTM88.Text = Empty
   Me.textTM92.Text = Empty
   Me.textTM106.Text = Empty:   Me.textTM107.Text = Empty
   Me.textTM108.Text = Empty:   Me.textTM109.Text = Empty
   Me.textTM110.Text = Empty:   Me.textTM111.Text = Empty
End Sub
Private Sub textTM80_1_GotFocus()
   InverseTextBox textTM80_1
End Sub

'Modified by Lydia 2021/11/29 改成Form 2.0; KeyAscii As Integer=>MSForms.ReturnInteger
Private Sub textTM80_1_KeyPress(KeyAscii As MSForms.ReturnInteger)
KeyAscii = UpperCase(KeyAscii)
End Sub
Private Sub textTM80_1_Validate(Cancel As Boolean)
Dim strTit As String
Dim strMsg As String
Dim nResponse
Dim rsTmp As ADODB.Recordset
Dim strKey As String
Dim strSql As String
Dim strData As String
       
    'Modify By Cheng 2003/08/07
    '若有更改申請人
    If Left(Me.textTM80_1.Text & "000000000", 9) <> m_TM80 Then
       Cancel = False
       textTM80_2 = Empty
       textTM84 = Empty
       textTM88 = Empty
       textTM92 = Empty
       ' 不滿九碼補0
       If IsEmptyText(textTM80_1) = False Then
          strData = textTM80_1 & String(9 - Len(textTM80_1), "0")
          Select Case Mid(strData, 1, 1)
          Case "X", "x":
             Set rsTmp = New ADODB.Recordset
             If Len(strData) > 8 Then
                strSql = "SELECT * FROM Customer " & _
                         "WHERE CU01 = '" & Mid(strData, 1, 8) & "' AND " & _
                               "CU02 = '" & Mid(strData, 9, 1) & "'"
             Else
                strSql = "SELECT * FROM Customer " & _
                         "WHERE CU01 = '" & Mid(strData, 1, 8) & "'"
             End If
             rsTmp.CursorLocation = adUseClient
             rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
             If rsTmp.RecordCount > 0 Then
                rsTmp.MoveFirst
                If IsNull(rsTmp.Fields("CU04")) = False Then
                   textTM80_2 = rsTmp.Fields("CU04")
                ElseIf IsNull(rsTmp.Fields("CU05")) = False Then
                   textTM80_2 = rsTmp.Fields("CU05")
                ElseIf IsNull(rsTmp.Fields("CU06")) = False Then
                   textTM80_2 = rsTmp.Fields("CU06")
                End If
                ' 帶出中英日地址
                If IsNull(rsTmp.Fields("CU23")) = False Then
                   textTM84 = rsTmp.Fields("CU23")
                End If
                If IsNull(rsTmp.Fields("CU24")) = False Then
                   textTM88 = rsTmp.Fields("CU24") & _
                               IIf(IsNull(rsTmp.Fields("CU25")), "", " " & rsTmp.Fields("CU25")) & _
                               IIf(IsNull(rsTmp.Fields("CU26")), "", " " & rsTmp.Fields("CU26")) & _
                               IIf(IsNull(rsTmp.Fields("CU27")), "", " " & rsTmp.Fields("CU27")) & _
                               IIf(IsNull(rsTmp.Fields("CU28")), "", " " & rsTmp.Fields("CU28")) & _
                               IIf(IsNull(rsTmp.Fields("CU102")), "", " " & rsTmp.Fields("CU102"))
                End If
                If IsNull(rsTmp.Fields("CU29")) = False Then
                   textTM92 = rsTmp.Fields("CU29")
                End If
             Else
                Cancel = True
                strTit = "檢核資料"
                strMsg = "申請人代號不存在"
                nResponse = MsgBox(strMsg, vbOKOnly, strTit)
                textTM80_1_GotFocus
             End If
             rsTmp.Close
          Case "Y", "y":
             Set rsTmp = New ADODB.Recordset
             If Len(strData) > 8 Then
                strSql = "SELECT * FROM FAGENT " & _
                         "WHERE FA01 = '" & Mid(strData, 1, 8) & "' AND " & _
                               "FA02 = '" & Mid(strData, 9, 1) & "'"
             Else
                strSql = "SELECT * FROM FAGENT " & _
                         "WHERE FA01 = '" & Mid(strData, 1, 8) & "' AND " & _
                               "FA02 = '0' "
             End If
             rsTmp.CursorLocation = adUseClient
             rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
             ' 檢查讀取的資料筆數
             If rsTmp.RecordCount > 0 Then
                rsTmp.MoveFirst
                If IsNull(rsTmp.Fields("FA03")) = False Then
                   strKey = rsTmp.Fields("FA03")
                   textTM80_1 = strKey
                   rsTmp.Close
                   If Len(strKey) > 8 Then
                      strSql = "SELECT * FROM Customer " & _
                            "WHERE CU01 = '" & Mid(strKey, 1, 8) & "' AND " & _
                                  "CU02 = '" & Mid(strKey, 9, 1) & "'"
                   Else
                      strSql = "SELECT * FROM Customer " & _
                            "WHERE CU01 = '" & Mid(strKey, 1, 8) & "' AND " & _
                                  "CU02 = '0' "
                   End If
                   rsTmp.CursorLocation = adUseClient
                   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
                   If rsTmp.RecordCount > 0 Then
                      rsTmp.MoveFirst
                      If IsNull(rsTmp.Fields("CU04")) = False Then
                         textTM80_2 = rsTmp.Fields("CU04")
                      ElseIf IsNull(rsTmp.Fields("CU05")) = False Then
                         textTM80_2 = rsTmp.Fields("CU05")
                      ElseIf IsNull(rsTmp.Fields("CU06")) = False Then
                         textTM80_2 = rsTmp.Fields("CU06")
                      End If
                      ' 帶出中英日地址
                      If IsNull(rsTmp.Fields("CU23")) = False Then
                         textTM84 = rsTmp.Fields("CU23")
                      End If
                      If IsNull(rsTmp.Fields("CU24")) = False Then
                         textTM88 = rsTmp.Fields("CU24")
                      End If
                      If IsNull(rsTmp.Fields("CU29")) = False Then
                         textTM92 = rsTmp.Fields("CU29")
                      End If
                   Else
                      Cancel = True
                      strTit = "檢核資料"
                      strMsg = "申請人代號不存在"
                      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
                      textTM80_1_GotFocus
                   End If
                Else
                   Cancel = True
                   strTit = "檢核資料"
                   strMsg = "申請人代號不存在"
                   nResponse = MsgBox(strMsg, vbOKOnly, strTit)
                   textTM80_1_GotFocus
                End If
             Else
                Cancel = True
                strTit = "檢核資料"
                strMsg = "申請人代號不存在"
                nResponse = MsgBox(strMsg, vbOKOnly, strTit)
                textTM80_1_GotFocus
             End If
             rsTmp.Close
          Case Else:
             Cancel = True
             strTit = "檢核資料"
             strMsg = "申請人代號不正確"
             nResponse = MsgBox(strMsg, vbOKOnly, strTit)
             textTM80_1_GotFocus
          End Select
       End If
       Set rsTmp = Nothing
    End If
End Sub
Private Sub textTM81_1_Change()
   '若申請人為空白, 自動清除相關地址及代表人
   If Me.textTM81_1.Text <> "" Then Exit Sub
   Me.textTM85.Text = Empty:   Me.textTM89.Text = Empty
   Me.textTM93.Text = Empty
   Me.textTM112.Text = Empty:   Me.textTM113.Text = Empty
   Me.textTM114.Text = Empty:   Me.textTM115.Text = Empty
   Me.textTM116.Text = Empty:   Me.textTM117.Text = Empty
End Sub
Private Sub textTM81_1_GotFocus()
   InverseTextBox textTM81_1
End Sub

'Modified by Lydia 2021/11/29 改成Form 2.0; KeyAscii As Integer=>MSForms.ReturnInteger
Private Sub textTM81_1_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = UpperCase(KeyAscii)
End Sub
Private Sub textTM81_1_Validate(Cancel As Boolean)
Dim strTit As String
Dim strMsg As String
Dim nResponse
Dim rsTmp As ADODB.Recordset
Dim strKey As String
Dim strSql As String
Dim strData As String
       
    'Modify By Cheng 2003/08/07
    '若有更改申請人
    If Left(Me.textTM81_1.Text & "000000000", 9) <> m_TM81 Then
       Cancel = False
       textTM81_2 = Empty
       textTM85 = Empty
       textTM89 = Empty
       textTM93 = Empty
       ' 不滿九碼補0
       If IsEmptyText(textTM81_1) = False Then
          strData = textTM81_1 & String(9 - Len(textTM81_1), "0")
          Select Case Mid(strData, 1, 1)
          Case "X", "x":
             Set rsTmp = New ADODB.Recordset
             If Len(strData) > 8 Then
                strSql = "SELECT * FROM Customer " & _
                         "WHERE CU01 = '" & Mid(strData, 1, 8) & "' AND " & _
                               "CU02 = '" & Mid(strData, 9, 1) & "'"
             Else
                strSql = "SELECT * FROM Customer " & _
                         "WHERE CU01 = '" & Mid(strData, 1, 8) & "'"
             End If
             rsTmp.CursorLocation = adUseClient
             rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
             If rsTmp.RecordCount > 0 Then
                rsTmp.MoveFirst
                If IsNull(rsTmp.Fields("CU04")) = False Then
                   textTM81_2 = rsTmp.Fields("CU04")
                ElseIf IsNull(rsTmp.Fields("CU05")) = False Then
                   textTM81_2 = rsTmp.Fields("CU05")
                ElseIf IsNull(rsTmp.Fields("CU06")) = False Then
                   textTM81_2 = rsTmp.Fields("CU06")
                End If
                ' 帶出中英日地址
                If IsNull(rsTmp.Fields("CU23")) = False Then
                   textTM85 = rsTmp.Fields("CU23")
                End If
                If IsNull(rsTmp.Fields("CU24")) = False Then
                   textTM89 = rsTmp.Fields("CU24") & _
                               IIf(IsNull(rsTmp.Fields("CU25")), "", " " & rsTmp.Fields("CU25")) & _
                               IIf(IsNull(rsTmp.Fields("CU26")), "", " " & rsTmp.Fields("CU26")) & _
                               IIf(IsNull(rsTmp.Fields("CU27")), "", " " & rsTmp.Fields("CU27")) & _
                               IIf(IsNull(rsTmp.Fields("CU28")), "", " " & rsTmp.Fields("CU28")) & _
                               IIf(IsNull(rsTmp.Fields("CU102")), "", " " & rsTmp.Fields("CU102"))
                End If
                If IsNull(rsTmp.Fields("CU29")) = False Then
                   textTM93 = rsTmp.Fields("CU29")
                End If
             Else
                Cancel = True
                strTit = "檢核資料"
                strMsg = "申請人代號不存在"
                nResponse = MsgBox(strMsg, vbOKOnly, strTit)
                textTM81_1_GotFocus
             End If
             rsTmp.Close
          Case "Y", "y":
             Set rsTmp = New ADODB.Recordset
             If Len(strData) > 8 Then
                strSql = "SELECT * FROM FAGENT " & _
                         "WHERE FA01 = '" & Mid(strData, 1, 8) & "' AND " & _
                               "FA02 = '" & Mid(strData, 9, 1) & "'"
             Else
                strSql = "SELECT * FROM FAGENT " & _
                         "WHERE FA01 = '" & Mid(strData, 1, 8) & "' AND " & _
                               "FA02 = '0' "
             End If
             rsTmp.CursorLocation = adUseClient
             rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
             ' 檢查讀取的資料筆數
             If rsTmp.RecordCount > 0 Then
                rsTmp.MoveFirst
                If IsNull(rsTmp.Fields("FA03")) = False Then
                   strKey = rsTmp.Fields("FA03")
                   textTM81_1 = strKey
                   rsTmp.Close
                   If Len(strKey) > 8 Then
                      strSql = "SELECT * FROM Customer " & _
                            "WHERE CU01 = '" & Mid(strKey, 1, 8) & "' AND " & _
                                  "CU02 = '" & Mid(strKey, 9, 1) & "'"
                   Else
                      strSql = "SELECT * FROM Customer " & _
                            "WHERE CU01 = '" & Mid(strKey, 1, 8) & "' AND " & _
                                  "CU02 = '0' "
                   End If
                   rsTmp.CursorLocation = adUseClient
                   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
                   If rsTmp.RecordCount > 0 Then
                      rsTmp.MoveFirst
                      If IsNull(rsTmp.Fields("CU04")) = False Then
                         textTM81_2 = rsTmp.Fields("CU04")
                      ElseIf IsNull(rsTmp.Fields("CU05")) = False Then
                         textTM81_2 = rsTmp.Fields("CU05")
                      ElseIf IsNull(rsTmp.Fields("CU06")) = False Then
                         textTM81_2 = rsTmp.Fields("CU06")
                      End If
                      ' 帶出中英日地址
                      If IsNull(rsTmp.Fields("CU23")) = False Then
                         textTM85 = rsTmp.Fields("CU23")
                      End If
                      If IsNull(rsTmp.Fields("CU24")) = False Then
                         textTM89 = rsTmp.Fields("CU24")
                      End If
                      If IsNull(rsTmp.Fields("CU29")) = False Then
                         textTM93 = rsTmp.Fields("CU29")
                      End If
                   Else
                      Cancel = True
                      strTit = "檢核資料"
                      strMsg = "申請人代號不存在"
                      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
                      textTM81_1_GotFocus
                   End If
                Else
                   Cancel = True
                   strTit = "檢核資料"
                   strMsg = "申請人代號不存在"
                   nResponse = MsgBox(strMsg, vbOKOnly, strTit)
                   textTM81_1_GotFocus
                End If
             Else
                Cancel = True
                strTit = "檢核資料"
                strMsg = "申請人代號不存在"
                nResponse = MsgBox(strMsg, vbOKOnly, strTit)
                textTM81_1_GotFocus
             End If
             rsTmp.Close
          Case Else:
             Cancel = True
             strTit = "檢核資料"
             strMsg = "申請人代號不正確"
             nResponse = MsgBox(strMsg, vbOKOnly, strTit)
             textTM81_1_GotFocus
          End Select
       End If
       Set rsTmp = Nothing
    End If
End Sub

Private Sub textTM82_GotFocus()
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textTM82.IMEMode = 1
   OpenIme
   InverseTextBox textTM82
End Sub
Private Sub textTM83_GotFocus()
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textTM83.IMEMode = 1
   OpenIme
   InverseTextBox textTM83
End Sub
Private Sub textTM84_GotFocus()
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textTM84.IMEMode = 1
   OpenIme
   InverseTextBox textTM84
End Sub
Private Sub textTM85_GotFocus()
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textTM85.IMEMode = 1
   OpenIme
   InverseTextBox textTM85
End Sub
Private Sub textTM86_GotFocus()
   InverseTextBox textTM86
End Sub
Private Sub textTM87_GotFocus()
   InverseTextBox textTM87
End Sub
Private Sub textTM88_GotFocus()
   InverseTextBox textTM88
End Sub
Private Sub textTM89_GotFocus()
   InverseTextBox textTM89
End Sub
Private Sub textTM90_GotFocus()
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textTM90.IMEMode = 1
   OpenIme
   InverseTextBox textTM90
End Sub
Private Sub textTM91_GotFocus()
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textTM91.IMEMode = 1
   OpenIme
   InverseTextBox textTM91
End Sub
Private Sub textTM92_GotFocus()
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textTM92.IMEMode = 1
   OpenIme
   InverseTextBox textTM92
End Sub
Private Sub textTM93_GotFocus()
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textTM93.IMEMode = 1
   OpenIme
   InverseTextBox textTM93
End Sub
Private Sub textTM82_Validate(Cancel As Boolean)
Dim strTit As String
Dim strMsg As String
Dim nResponse
   
   Cancel = False
   If CheckLengthIsOK(textTM82, textTM82.MaxLength) = False Then
      Cancel = True
      textTM82_GotFocus
   End If
   'edit by nickc 2007/06/06 切換輸入法改用API
   'If Cancel = False Then: textTM82.IMEMode = 2
   If Cancel = False Then CloseIme
End Sub
Private Sub textTM83_Validate(Cancel As Boolean)
Dim strTit As String
Dim strMsg As String
Dim nResponse
   
   Cancel = False
   If CheckLengthIsOK(textTM83, textTM83.MaxLength) = False Then
      Cancel = True
      textTM83_GotFocus
   End If
   'edit by nickc 2007/06/06 切換輸入法改用API
   'If Cancel = False Then: textTM83.IMEMode = 2
   If Cancel = False Then CloseIme
End Sub
Private Sub textTM84_Validate(Cancel As Boolean)
Dim strTit As String
Dim strMsg As String
Dim nResponse
   
   Cancel = False
   If CheckLengthIsOK(textTM84, textTM84.MaxLength) = False Then
      Cancel = True
      textTM84_GotFocus
   End If
   'edit by nickc 2007/06/06 切換輸入法改用API
   'If Cancel = False Then: textTM84.IMEMode = 2
   If Cancel = False Then CloseIme
End Sub
Private Sub textTM85_Validate(Cancel As Boolean)
Dim strTit As String
Dim strMsg As String
Dim nResponse
   
   Cancel = False
   If CheckLengthIsOK(textTM85, textTM85.MaxLength) = False Then
      Cancel = True
      textTM85_GotFocus
   End If
   'edit by nickc 2007/06/06 切換輸入法改用API
   'If Cancel = False Then: textTM85.IMEMode = 2
   If Cancel = False Then CloseIme
End Sub
Private Sub textTM90_Validate(Cancel As Boolean)
Dim strTit As String
Dim strMsg As String
Dim nResponse
   
   Cancel = False
   If CheckLengthIsOK(textTM90, textTM90.MaxLength) = False Then
      Cancel = True
      textTM90_GotFocus
   End If
   'edit by nickc 2007/06/06 切換輸入法改用API
   'If Cancel = False Then: textTM90.IMEMode = 2
   If Cancel = False Then CloseIme
End Sub
Private Sub textTM91_Validate(Cancel As Boolean)
Dim strTit As String
Dim strMsg As String
Dim nResponse
   
   Cancel = False
   If CheckLengthIsOK(textTM91, textTM91.MaxLength) = False Then
      Cancel = True
      textTM91_GotFocus
   End If
   'edit by nickc 2007/06/06 切換輸入法改用API
   'If Cancel = False Then: textTM91.IMEMode = 2
   If Cancel = False Then CloseIme
End Sub
Private Sub textTM92_Validate(Cancel As Boolean)
Dim strTit As String
Dim strMsg As String
Dim nResponse
   
   Cancel = False
   If CheckLengthIsOK(textTM92, textTM92.MaxLength) = False Then
      Cancel = True
      textTM92_GotFocus
   End If
   'edit by nickc 2007/06/06 切換輸入法改用API
   'If Cancel = False Then: textTM92.IMEMode = 2
   If Cancel = False Then CloseIme
End Sub
Private Sub textTM93_Validate(Cancel As Boolean)
Dim strTit As String
Dim strMsg As String
Dim nResponse
   
   Cancel = False
   If CheckLengthIsOK(textTM93, textTM93.MaxLength) = False Then
      Cancel = True
      textTM93_GotFocus
   End If
   'edit by nickc 2007/06/06 切換輸入法改用API
   'If Cancel = False Then: textTM93.IMEMode = 2
   If Cancel = False Then CloseIme
End Sub

' 按下 ToolBar 的 Button
Private Sub tlbar_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim strTit As String
Dim strMsg As String
Dim nResponse
   
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

' 檢查記錄是否已經存在
Private Function IsRecordExist(ByVal strTM01 As String, ByVal strTM02 As String, ByVal strTM03 As String, ByVal strTM04 As String) As Boolean
Dim rsTmp As New ADODB.Recordset
Dim strSql As String
   
   IsRecordExist = False
   strSql = "SELECT * FROM TradeMark " & _
            "WHERE TM01 = '" & strTM01 & "' AND " & _
                  "TM02 = '" & strTM02 & "' AND " & _
                  "TM03 = '" & strTM03 & "' AND " & _
                  "TM04 = '" & strTM04 & "'"
                  
   ' 讀取資料庫
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly   'edit by nickc 2005/08/04 原先是  動態開啟
   ' 檢查讀取的資料筆數
   If rsTmp.RecordCount > 0 Then
      IsRecordExist = True
   Else
      IsRecordExist = False
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Function

' 新增記錄
'edit by nickc 2006/06/08
'Private Sub AddRecord()
Private Function AddRecord() As Boolean
Dim strSql As String
Dim strTmp As String
Dim strTit As String
Dim strMsg As String
Dim nResponse
Dim nIndex As Integer
Dim bDifference As Boolean
Dim bFirst As Boolean
Dim strTM01, strTM02, strTM03, strTM04 As String
Dim strUpdTM58_716 As String, strUpdTM58_102 As String, strUpdTM58_105 As String
   
   'add by nickc 2006/06/08
   AddRecord = False
   
   strTM01 = textTM01
   strTM02 = textTM02_1 & textTM02_2
   strTM03 = textTM03
   strTM04 = textTM04
   
   ' 檢查記錄是否已存在
   If IsRecordExist(strTM01, strTM02, strTM03, strTM04) = True Then
      strTit = "新增資料"
      strMsg = "該筆記錄已存在"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      UpdateCtrlData
      GoTo EXITSUB
   End If
   
   bFirst = True
   bDifference = False
   strSql = "INSERT INTO TradeMark ("
   For nIndex = 0 To MAX_FIELD - 1
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
   For nIndex = 0 To MAX_FIELD - 1
      strTmp = Empty
      If m_FieldList(nIndex).fiOldData <> m_FieldList(nIndex).fiNewData Then
         If m_FieldList(nIndex).fiType = 0 Then
            ' 91.03.25 modify by louis (單引號)
            'strTmp = "'" & m_FieldList(nIndex).fiNewData & "'"
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

    'add by nickc 2006/03/16 紀錄分析語法
    On Error GoTo oErr
    cnnConnection.BeginTrans
    Pub_SeekTbLog strSql
    cnnConnection.Execute strSql
      
      'Modify By Sindy 2013/5/14 移為函數
      strTM01 = textTM01
      strTM02 = textTM02_1
      If textTM01 = "TF" Then: strTM02 = strTM02 & textTM02_2
      strTM03 = textTM03
      If IsEmptyText(strTM03) = True Then: strTM03 = "0"
      strTM04 = textTM04
      If IsEmptyText(strTM04) = True Then: strTM04 = "00"
      If strTM01 = "TF" And strTM03 & strTM04 = "000" Then
         '母案在存檔時執行檢核, 子案要抓取母案資料才能進行
         strExc(0) = "select TM01,TM02,TM03,TM04,TM10,TM29,TM11,TM14,TM20," & _
                     "TM21,TM22,decode(TM28,'1','申請','2','異議','3','評定','4','廢止') as TM28,TM58" & _
                     " from trademark where tm01='" & strTM01 & "' and substr(tm02,1,5)='" & Left(strTM02, 5) & "'" & _
                     " order by TM01,TM02,TM03,TM04 asc"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         strUpdTM58_716 = "": strUpdTM58_102 = "": strUpdTM58_105 = ""
         If intI = 1 Then
            RsTemp.MoveFirst
            Do While Not RsTemp.EOF
               Call CheckDeadLine(RsTemp.Fields("TM01"), RsTemp.Fields("TM02"), RsTemp.Fields("TM03"), RsTemp.Fields("TM04"), _
                                  "" & RsTemp.Fields("TM10"), "" & RsTemp.Fields("TM29"), "" & RsTemp.Fields("TM58"), _
                                  textTM11, textTM14, textTM20, textTM21, textTM22, textTM28, _
                                  strUpdTM58_716, strUpdTM58_102, strUpdTM58_105)
               Call UpdateData(RsTemp.Fields("TM01"), Mid(Trim(RsTemp.Fields("TM02")), 1, 5), Mid(Trim(RsTemp.Fields("TM02")), 6, 1), RsTemp.Fields("TM03"), RsTemp.Fields("TM04"))
               
               RsTemp.MoveNext
            Loop
            If strUpdTM58_716 <> "" Or strUpdTM58_102 <> "" Or strUpdTM58_105 <> "" Then
               strSql = "update trademark set" & _
                        " tm58='" & Trim(strUpdTM58_716) & Trim(strUpdTM58_102) & Trim(strUpdTM58_105) & "'||tm58" & _
                        " where TM01='" & textTM01 & "' and TM02='" & textTM02_1 & "0' and TM03='0' and TM04='00'"
               cnnConnection.Execute strSql
            End If
         End If
      Else
         Call UpdateData(textTM01, textTM02_1, textTM02_2, textTM03, textTM04) 'Modify By Sindy 2013/5/3 統一寫在此函數
      End If
      '2013/5/14 End
      
'      'add by nickc 2008/02/13 加掛期限
'      '716
'      strCP09 = GetLastA(textTM01, textTM02_1, textTM02_2, textTM03, textTM04)
'      If IsCreate716 Then
'           If m_716CP06 <> "" And m_716CP07 <> "" And m_716Key <> "" Then
'                cnnConnection.Execute "update caseprogress set cp06='" & m_716CP06 & "',cp07='" & m_716CP07 & "' where cp09='" & m_716Key & "' "
'           ElseIf m_716NP08 <> "" And m_716NP09 <> "" Then
'                If m_716Key <> "" Then
'                    'Modify By Sindy 2011/5/27
'                    'cnnConnection.Execute "update nextprogress set np08='" & m_716NP08 & "',np09='" & m_716NP09 & "' where np01='" & m_716Key & "' "
'                    cnnConnection.Execute "update nextprogress set np08='" & m_716NP08 & "',np09='" & m_716NP09 & "' where np01='" & m_716Key & "' and np07=716 and np06 is null "
'                Else
'                    If strCP09 <> "" Then
'                        strNP22 = GetNextProgressNo()
'                        strSql = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
'                                 "VALUES ('" & strCP09 & "','" & textTM01 & "','" & textTM02_1 & textTM02_2 & "','" & textTM03 & "','" & textTM04 & "',716," & _
'                                            DBDATE(m_716NP08) & "," & DBDATE(m_716NP09) & ",'" & PUB_GetAKindSalesNo(textTM01, textTM02_1 & textTM02_2, textTM03, textTM04) & "'," & strNP22 & ")"
'                        cnnConnection.Execute strSql
'                    End If
'                End If
'           End If
'      End If
'      '102
'      If IsCreate102 Then
'           If m_102CP06 <> "" And m_102CP07 <> "" And m_102Key <> "" Then
'                cnnConnection.Execute "update caseprogress set cp06='" & m_102CP06 & "',cp07='" & m_102CP07 & "' where cp09='" & m_102Key & "' "
'           ElseIf m_102NP08 <> "" And m_102NP09 <> "" Then
'                If m_102Key <> "" Then
'                    'Modify By Sindy 2011/5/27
'                    'cnnConnection.Execute "update nextprogress set np08='" & m_102NP08 & "',np09='" & m_102NP09 & "' where np01='" & m_102Key & "' "
'                    cnnConnection.Execute "update nextprogress set np08='" & m_102NP08 & "',np09='" & m_102NP09 & "' where np01='" & m_102Key & "' and np07=102 and np06 is null "
'                Else
'                    If strCP09 <> "" Then
'                        strNP22 = GetNextProgressNo()
'                        strSql = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
'                                 "VALUES ('" & strCP09 & "','" & textTM01 & "','" & textTM02_1 & textTM02_2 & "','" & textTM03 & "','" & textTM04 & "',102," & _
'                                            DBDATE(m_102NP08) & "," & DBDATE(m_102NP09) & ",'" & PUB_GetAKindSalesNo(textTM01, textTM02_1 & textTM02_2, textTM03, textTM04) & "'," & strNP22 & ")"
'                        cnnConnection.Execute strSql
'                    End If
'                End If
'           End If
'      End If
'      '105
'      If IsCreate105 Then
'           If m_105CP06 <> "" And m_105CP07 <> "" And m_105Key <> "" Then
'                cnnConnection.Execute "update caseprogress set cp06='" & m_105CP06 & "',cp07='" & m_105CP07 & "' where cp09='" & m_105Key & "' "
'           ElseIf m_105NP08 <> "" And m_105NP09 <> "" Then
'                If m_105Key <> "" Then
'                    'Modify By Sindy 2011/5/27
'                    'cnnConnection.Execute "update nextprogress set np08='" & m_105NP08 & "',np09='" & m_105NP09 & "' where np01='" & m_105Key & "' "
'                    cnnConnection.Execute "update nextprogress set np08='" & m_105NP08 & "',np09='" & m_105NP09 & "' where np01='" & m_105Key & "' and np07=105 and np06 is null "
'                Else
'                    If strCP09 <> "" Then
'                        strNP22 = GetNextProgressNo()
'                        strSql = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
'                                 "VALUES ('" & strCP09 & "','" & textTM01 & "','" & textTM02_1 & textTM02_2 & "','" & textTM03 & "','" & textTM04 & "',105," & _
'                                            DBDATE(m_105NP08) & "," & DBDATE(m_105NP09) & ",'" & PUB_GetAKindSalesNo(textTM01, textTM02_1 & textTM02_2, textTM03, textTM04) & "'," & strNP22 & ")"
'                        cnnConnection.Execute strSql
'                    End If
'                End If
'           End If
'      End If
   
   'Add By Sindy 2013/9/16 申請人為X13175010工研院者且有專用期間者設定為不催延展
   If (textTM01 = "T" Or textTM01 = "TF") And _
      (textTM23_1 = "X13175010" Or textTM78_1 = "X13175010" Or textTM79_1 = "X13175010" Or textTM80_1 = "X13175010" Or textTM81_1 = "X13175010") And _
      Val(textTM21) > 0 And _
      Val(textTM22) > 0 Then
      strSql = "update trademark set" & _
               " tm129='Y'" & _
               " where TM01='" & textTM01 & "' and TM02='" & textTM02_1 & textTM02_2 & _
                "' and TM03='" & textTM03 & "' and TM04='" & textTM04 & "'"
      cnnConnection.Execute strSql
   End If
   '2013/9/16 END
     
   'add by nickc 2006/06/08
   cnnConnection.CommitTrans
   
   If ((strTM01 & strTM02 & strTM03 & strTM04) < (m_FirstTM(0) & m_FirstTM(1) & m_FirstTM(2) & m_FirstTM(3))) Or ((strTM01 & strTM02 & strTM03 & strTM04) > (m_LastTM(0) & m_LastTM(1) & m_LastTM(2) & m_LastTM(3))) Then
      RefreshRange
   End If
   
   ShowCurrRecord strTM01, strTM02, strTM03, strTM04
   AddRecord = True
EXITSUB:
'add by nickc 2006/06/08
Exit Function
oErr:
    cnnConnection.RollbackTrans
    MsgBox Err.Description
End Function

' 修改記錄
'edit by nickc 2006/06/08
'Private Sub ModRecord()
Private Function ModRecord() As Boolean
Dim strSql As String
Dim strTmp As String
Dim strTit As String
Dim strMsg As String
Dim nResponse
Dim nIndex As Integer
Dim bDifference As Boolean
Dim bFirst As Boolean
Dim strTM01, strTM02, strTM03, strTM04 As String
Dim strUpdTM58_716 As String, strUpdTM58_102  As String, strUpdTM58_105 As String
'Added by Lydia 2016/12/19
Dim strCP09 As String
Dim strNP22 As String
Dim bolData As Boolean, strMCTF(0) As String 'Add by Amy 2017/03/15

   'add by nickc 2006/06/08
   ModRecord = False
   
   strTM01 = textTM01
   strTM02 = textTM02_1 & textTM02_2
   strTM03 = textTM03
   strTM04 = textTM04
   '910910  nick tigger
   '***** start
   'strSQL = "UPDATE TradeMark SET "
   'edit by nickc 2006/06/07 紀錄 log
   'StrSql = "begin user_data.user_enabled:=1; UPDATE TradeMark SET "
   strSql = " UPDATE TradeMark SET "
   '***** end
   bFirst = True
   bDifference = False
   For nIndex = 0 To MAX_FIELD - 1
      strTmp = Empty
      '92.05.22 nick 跳過create & update 相關項目
      'edit by nickc 2006/07/12 加入跳過銷卷項目
      'If nIndex < 58 Or nIndex > 63 Then
      If (nIndex < 58 Or nIndex > 63) And nIndex <> 56 And nIndex <> 73 And nIndex <> 72 And nIndex <> 74 Then
        If m_FieldList(nIndex).fiOldData <> m_FieldList(nIndex).fiNewData Then
           If m_FieldList(nIndex).fiType = 0 Then
              If m_FieldList(nIndex).fiNewData = Empty Then
                 strTmp = m_FieldList(nIndex).fiName & " = NULL "
              Else
                 ' 91.03.25 modify by louis (單引號)
                 'strTmp = m_FieldList(nIndex).fiName & " = '" & m_FieldList(nIndex).fiNewData & "'"
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
   '910910 nick tigger
   '***** start
   'strSQL = strSQL & " " & _
                  "WHERE TM01 = '" & textTM01 & "' AND " & _
                     "TM02 = '" & textTM02_1 & textTM02_2 & "' AND " & _
                     "TM03 = '" & textTM03 & "' AND " & _
                     "TM04 = '" & textTM04 & "'"
   'edit by nickc 2006/06/07 紀錄 log
   'StrSql = StrSql & " " & _
                  "WHERE TM01 = '" & textTM01 & "' AND " & _
                     "TM02 = '" & textTM02_1 & textTM02_2 & "' AND " & _
                     "TM03 = '" & textTM03 & "' AND " & _
                     "TM04 = '" & textTM04 & "' ; end;"
   strSql = strSql & " " & _
                  "WHERE TM01 = '" & textTM01 & "' AND " & _
                     "TM02 = '" & textTM02_1 & textTM02_2 & "' AND " & _
                     "TM03 = '" & textTM03 & "' AND " & _
                     "TM04 = '" & textTM04 & "' "
                     
    '***** end
'910910 nick tigger
'***** start
On Error GoTo ErrHand
'***** end

   'If bDifference = True Then 'Modify By Sindy 2012/10/1 Mark
      '910910 nick tigger
      '**** start
      cnnConnection.BeginTrans
      '***** end
      'add by nickc 2006/03/16 紀錄分析語法
      If bDifference = True Then 'Modify By Sindy 2012/10/1 下列期限不管有無異動畫面資料均可異動
         'Add by Amy 2017/11/28 TF母案修改出名公司,子案一併更新 ex:CFP-029915
         If strTM01 = "TF" And Right(strTM02, 1) = "0" And strTM03 = "0" And strTM04 = "00" And m_FieldList(129).fiOldData <> textTM130 Then
             strExc(1) = "Update Trademark Set tm130=" & CNULL(textTM130) & " " & _
                         "WHERE tm01='" & strTM01 & "' and substr(tm02,1,5)='" & Left(strTM02, 5) & "' "
             cnnConnection.Execute strExc(1)
         End If
         'add by sonia 2021/10/4 TF母案修改專用期間，存活的子案也要一併改，但菲律賓030除外(桂英)
         If strTM01 = "TF" And Right(strTM02, 1) = "0" And strTM03 = "0" And strTM04 = "00" Then
             strExc(1) = "Update Trademark Set tm21=" & CNULL(textTM21) & ",tm22=" & CNULL(textTM22) & " " & _
                         "WHERE tm01='" & strTM01 & "' and tm02='" & strTM02 & "' And tm03<>'0' and tm04<>'00' and tm10<>'030'"
             cnnConnection.Execute strExc(1)
         End If
         'end 2021/10/4
         
         Pub_SeekTbLog strSql
         'edit by nickc 2006/06/07 紀錄 log
         'cnnConnection.Execute StrSql
         cnnConnection.Execute "begin user_data.user_enabled:=1; " & strSql & "; end;"
         'add by nickc 2005/08/23 紀錄修改案號
         pub_ModifyCaseNum = strTM01 & "-" & strTM02 & "-" & strTM03 & "-" & strTM04
      End If
      '910910 nick tigger
      
      'Modify By Sindy 2013/5/14 移為函數
      strTM01 = textTM01
      strTM02 = textTM02_1
      If textTM01 = "TF" Then: strTM02 = strTM02 & textTM02_2
      strTM03 = textTM03
      If IsEmptyText(strTM03) = True Then: strTM03 = "0"
      strTM04 = textTM04
      If IsEmptyText(strTM04) = True Then: strTM04 = "00"
      If strTM01 = "TF" And strTM03 & strTM04 = "000" Then
         '母案在存檔時執行檢核, 子案要抓取母案資料才能進行
         strExc(0) = "select TM01,TM02,TM03,TM04,TM10,TM29,TM11,TM14,TM20," & _
                     "TM21,TM22,decode(TM28,'1','申請','2','異議','3','評定','4','廢止') as TM28,TM58" & _
                     " from trademark where tm01='" & strTM01 & "' and substr(tm02,1,5)='" & Left(strTM02, 5) & "'" & _
                     " order by TM01,TM02,TM03,TM04 asc"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         strUpdTM58_716 = "": strUpdTM58_102 = "": strUpdTM58_105 = ""
         If intI = 1 Then
            RsTemp.MoveFirst
            Do While Not RsTemp.EOF
               Call CheckDeadLine(RsTemp.Fields("TM01"), RsTemp.Fields("TM02"), RsTemp.Fields("TM03"), RsTemp.Fields("TM04"), _
                                  "" & RsTemp.Fields("TM10"), "" & RsTemp.Fields("TM29"), "" & RsTemp.Fields("TM58"), _
                                  textTM11, textTM14, textTM20, textTM21, textTM22, textTM28, _
                                  strUpdTM58_716, strUpdTM58_102, strUpdTM58_105)
               Call UpdateData(RsTemp.Fields("TM01"), Mid(Trim(RsTemp.Fields("TM02")), 1, 5), Mid(Trim(RsTemp.Fields("TM02")), 6, 1), RsTemp.Fields("TM03"), RsTemp.Fields("TM04"))
               
               RsTemp.MoveNext
            Loop
            If strUpdTM58_716 <> "" Or strUpdTM58_102 <> "" Or strUpdTM58_105 <> "" Then
               strSql = "update trademark set" & _
                        " tm58='" & Trim(strUpdTM58_716) & Trim(strUpdTM58_102) & Trim(strUpdTM58_105) & "'||tm58" & _
                        " where TM01='" & textTM01 & "' and TM02='" & textTM02_1 & "0' and TM03='0' and TM04='00'"
               cnnConnection.Execute strSql
            End If
         End If
      Else
         Call UpdateData(textTM01, textTM02_1, textTM02_2, textTM03, textTM04) 'Modify By Sindy 2013/5/3 統一寫在此函數
      End If
      '2013/5/14 End
      
'      'add by nickc 2008/02/13 加掛期限
'      '716
'      strCP09 = GetLastA(textTM01, textTM02_1, textTM02_2, textTM03, textTM04)
'      If IsCreate716 Then
'           If m_716CP06 <> "" And m_716CP07 <> "" And m_716Key <> "" Then
'                cnnConnection.Execute "update caseprogress set cp06='" & m_716CP06 & "',cp07='" & m_716CP07 & "' where cp09='" & m_716Key & "' "
'           ElseIf m_716NP08 <> "" And m_716NP09 <> "" Then
'                If m_716Key <> "" Then
'                    'Modify By Sindy 2011/5/27
'                    'cnnConnection.Execute "update nextprogress set np08='" & m_716NP08 & "',np09='" & m_716NP09 & "' where np01='" & m_716Key & "' "
'                    cnnConnection.Execute "update nextprogress set np08='" & m_716NP08 & "',np09='" & m_716NP09 & "' where np01='" & m_716Key & "' and np07=716 and np06 is null "
'                Else
'                    If strCP09 <> "" Then
'                        strNP22 = GetNextProgressNo()
'                        strSql = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
'                                 "VALUES ('" & strCP09 & "','" & textTM01 & "','" & textTM02_1 & textTM02_2 & "','" & textTM03 & "','" & textTM04 & "',716," & _
'                                            DBDATE(m_716NP08) & "," & DBDATE(m_716NP09) & ",'" & PUB_GetAKindSalesNo(textTM01, textTM02_1 & textTM02_2, textTM03, textTM04) & "'," & strNP22 & ")"
'                        cnnConnection.Execute strSql
'                    End If
'                End If
'           End If
'      End If
'      '102
'      If IsCreate102 Then
'           If m_102CP06 <> "" And m_102CP07 <> "" And m_102Key <> "" Then
'                cnnConnection.Execute "update caseprogress set cp06='" & m_102CP06 & "',cp07='" & m_102CP07 & "' where cp09='" & m_102Key & "' "
'           ElseIf m_102NP08 <> "" And m_102NP09 <> "" Then
'                If m_102Key <> "" Then
'                    'Modify By Sindy 2011/5/27
'                    'cnnConnection.Execute "update nextprogress set np08='" & m_102NP08 & "',np09='" & m_102NP09 & "' where np01='" & m_102Key & "' "
'                    cnnConnection.Execute "update nextprogress set np08='" & m_102NP08 & "',np09='" & m_102NP09 & "' where np01='" & m_102Key & "' and np07=102 and np06 is null "
'                Else
'                    If strCP09 <> "" Then
'                        strNP22 = GetNextProgressNo()
'                        strSql = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
'                                 "VALUES ('" & strCP09 & "','" & textTM01 & "','" & textTM02_1 & textTM02_2 & "','" & textTM03 & "','" & textTM04 & "',102," & _
'                                            DBDATE(m_102NP08) & "," & DBDATE(m_102NP09) & ",'" & PUB_GetAKindSalesNo(textTM01, textTM02_1 & textTM02_2, textTM03, textTM04) & "'," & strNP22 & ")"
'                        cnnConnection.Execute strSql
'                    End If
'                End If
'           End If
'      End If
'      '105
'      If IsCreate105 Then
'           If m_105CP06 <> "" And m_105CP07 <> "" And m_105Key <> "" Then
'                cnnConnection.Execute "update caseprogress set cp06='" & m_105CP06 & "',cp07='" & m_105CP07 & "' where cp09='" & m_105Key & "' "
'           ElseIf m_105NP08 <> "" And m_105NP09 <> "" Then
'                If m_105Key <> "" Then
'                    'Modify By Sindy 2011/5/27
'                    'cnnConnection.Execute "update nextprogress set np08='" & m_105NP08 & "',np09='" & m_105NP09 & "' where np01='" & m_105Key & "' "
'                    cnnConnection.Execute "update nextprogress set np08='" & m_105NP08 & "',np09='" & m_105NP09 & "' where np01='" & m_105Key & "' and np07=105 and np06 is null "
'                Else
'                    If strCP09 <> "" Then
'                        strNP22 = GetNextProgressNo()
'                        strSql = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
'                                 "VALUES ('" & strCP09 & "','" & textTM01 & "','" & textTM02_1 & textTM02_2 & "','" & textTM03 & "','" & textTM04 & "',105," & _
'                                            DBDATE(m_105NP08) & "," & DBDATE(m_105NP09) & ",'" & PUB_GetAKindSalesNo(textTM01, textTM02_1 & textTM02_2, textTM03, textTM04) & "'," & strNP22 & ")"
'                        cnnConnection.Execute strSql
'                    End If
'                End If
'           End If
'      End If
      
      'Add By Sindy 2013/9/16 申請人為X13175010工研院者且有專用期間者設定為不催延展
      If (textTM01 = "T" Or textTM01 = "TF") And _
         (textTM23_1 = "X13175010" Or textTM78_1 = "X13175010" Or textTM79_1 = "X13175010" Or textTM80_1 = "X13175010" Or textTM81_1 = "X13175010") And _
         Val(textTM21) > 0 And _
         Val(textTM22) > 0 Then
         strSql = "update trademark set" & _
                  " tm129='Y'" & _
                  " where TM01='" & textTM01 & "' and TM02='" & textTM02_1 & textTM02_2 & _
                   "' and TM03='" & textTM03 & "' and TM04='" & textTM04 & "'"
         cnnConnection.Execute strSql
      End If
      '2013/9/16 END
      
      'Added by Lydia 2016/12/19 存檔時，非TF案件，若目前准駁TM16欄為空白時
      'Modified by Lydia 2016/12/22 卷宗性質為申請
      'Modified by Lydia 2019/07/02 排除已閉卷或銷卷的案件 textTM30 & textTM57
      If textTM01 <> "TF" And textTM16 = "" And textTM28 = "申請" And Trim(textTM30 & textTM57) = "" Then
          '檢查進度檔若無申請101或分割308進度 , 則自動產生B類「申請」進度，收文日=發文日=19221111；承辦人=操作人員，智權人員依各系統規則預設；
          strExc(0) = "select cp05,cp09,cp10 from caseprogress where cp01='" & textTM01 & "' and cp02='" & textTM02_1 & textTM02_2 & "' and cp03='" & textTM03 & "' and cp04='" & textTM04 & "' and cp10 in ('101','308') and cp159=0 order by cp05 "
          intI = 1
          Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
          If intI = 0 Then
             strCP09 = AutoNo("B", 6)
             'Added by Lydia 2017/08/30 抓最新收文的智權人員
             strExc(8) = PUB_GetAKindSalesNo(textTM01, textTM02_1 & textTM02_2, textTM03, textTM04)
             
             'Modified by Lydia 2017/08/30 代入最新收文的智權人員; PUB_GetAKindSalesNo(textTM01, textTM02_1 & textTM02_2, textTM03, textTM04) => strExc(8)
             strSql = "INSERT INTO CaseProgress (CP01,CP02,CP03,CP04,CP05,CP09,CP10,CP12,CP13,CP14,CP26,CP27,CP32) " & _
                       "VALUES ('" & textTM01 & "','" & textTM02_1 & textTM02_2 & "','" & textTM03 & "','" & textTM04 & "',19221111," & _
                         "'" & strCP09 & "','101','" & GetSalesArea(strExc(8)) & "','" & strExc(8) & "'," & _
                         "'" & strUserNum & "','N',19221111,'N')"
             cnnConnection.Execute strSql
             
             '抓不到審查時間或CF05=0時，仍可存檔但要發E-MAIL給特殊人員(V)外商陳經理
             Call PUB_SetChkResultDateT(textTM01, textTM10_1, "101", DBDATE(textTM11), strExc(5), textTM02_1 & textTM02_2, textTM03, textTM04)
             If strExc(5) <> "" Then
                'Added by Lydia 2017/08/30 改變預設NP10
                If textTM01 = "FCT" Then
                   '最新收文的智權人員
                   strExc(10) = strExc(8)
                'Modified by Lydia 2022/09/21 增加CFC案
                ElseIf textTM01 = "CFT" Or textTM01 = "CFC" Then
                     '模組判斷
                   Call GetNA69("", textTM10_1, strExc(8), strExc(10), textTM01, textTM02_1 & textTM02_2, textTM03, textTM04)
                Else
                   'T案-操作人員
                   strExc(10) = strUserNum
                   'Added by Lydia 2018/09/17 改成新案之承辦人(ex.T-216202)
                   strExc(0) = "select cp14 from caseprogress where cp01='" & textTM01 & "' and cp02='" & textTM02_1 & textTM02_2 & "' and cp03='" & textTM03 & "' and cp04='" & textTM04 & "' and cp31='Y' and cp159=0 "
                   intI = 1
                   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                   If intI = 1 Then
                       If "" & RsTemp.Fields(0) <> "" Then strExc(10) = "" & RsTemp.Fields(0)
                   End If
                   'end 2018/09/17
                End If
                'end 2017/08/30
                '新增下一程序B類「申請」進度的催審305期限：以系統類別+申請國家+101抓案件國家設定檔CASEFEE之審查時間(天)CF05(必須>0)，以申請日＋CF05計算催審期限(本所期限=法定期限)；
                strNP22 = GetNextProgressNo()
                'Modified by Lydia 2017/08/30 PUB_GetAKindSalesNo(textTM01, textTM02_1 & textTM02_2, textTM03, textTM04) =>strexc(10)
                'Modified by Lydia 2020/07/07 本所期限檢查：若本所期限非工作天則直接調整至最近的工作天
                'strSql = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
                         "VALUES ('" & strCP09 & "','" & textTM01 & "','" & textTM02_1 & textTM02_2 & "','" & textTM03 & "','" & textTM04 & "',305," & _
                                    strExc(5) & "," & strExc(5) & ",'" & strExc(10) & "'," & strNP22 & ")"
                strSql = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
                         "VALUES ('" & strCP09 & "','" & textTM01 & "','" & textTM02_1 & textTM02_2 & "','" & textTM03 & "','" & textTM04 & "',305," & _
                                    PUB_GetWorkDay1(strExc(5), True) & "," & strExc(5) & ",'" & strExc(10) & "'," & strNP22 & ")"
                cnnConnection.Execute strSql
             End If
          End If
           
      End If
      'end 2016/12/19
      'Add by Amy 2017/03/15 FC代理人修改為MCTF時更新客戶檔及下一程序
      'Modify by Amy 2017/03/22 拿掉 And textTM10_1 = "000" 申請國家為台灣之判斷 (X39289040 為MCTF但要收申請國家為大陸的案件)
      If Left(textTM01, 1) = "T" And Trim(Len(textTM44_1)) > 0 And m_FieldList(43).fiOldData <> ChangeCustomerL(textTM44_1) Then
        bolData = GetCusORFagentData(ChangeCustomerL(textTM44_1), "FA120", strMCTF())
        If Left(strMCTF(0), 4) = "MCTF" Then
            If UpdMCTF_NP(ChangeCustomerL(textTM44_1), strMCTF(0), textTM01 & textTM02_1 & textTM02_2 & textTM03 & textTM04) = False Then GoTo ErrHand
        End If
      End If

      'Added by Lydia 2025/09/12 TF基礎案號設定：發通知Email
      'Mark by Lydia 2025/10/23
      'If textTM01 = "TF" And FraTFbase.Visible = True Then
      '    lblTFbase02.Caption = PUB_GetTFbaseInfo(textTM01, textTM02_1 & textTM02_2, textTM03, textTM04, Trim(textTM06), Trim(textTM07), IIf(Trim(textTM06 & textTM07) <> Trim(textTM06.Tag & textTM07.Tag), "1", ""))
      '    If Left(lblTFbase02, 1) = "N" Then
      '       lblTFbase02.ForeColor = vbRed
      '    Else
      '        lblTFbase02.ForeColor = &H80000012
      '    End If
      'End If
      'end 2025/10/23
      '基礎案變更
      If (textTM01 = "T" Or textTM01 = "CFT") And Trim(textTM12 & textTM15 & textTM10_1 & textTM29 & textTM16) <> Trim(textTM12.Tag & textTM15.Tag & textTM10_1.Tag & textTM29.Tag & textTM16.Tag) Then
          'Modified by Lydia 2025/10/23
          'lblTFbase02.Caption = PUB_GetTFbaseInfo(textTM01, textTM02_1, textTM03, textTM04, Trim(textTM15), Trim(textTM10_1), "1", Trim(textTM12))
          strSql = PUB_GetTFbaseInfo(textTM01, textTM02_1, textTM03, textTM04, Trim(textTM15), Trim(textTM10_1), IIf(textTM12 & textTM15 & textTM10_1 <> textTM12.Tag & textTM15.Tag & textTM10_1.Tag, "1", "2"), Trim(textTM12))
      End If
      'end 2025/09/12
      '***** start
      cnnConnection.CommitTrans
      '***** end
      
      ShowCurrRecord strTM01, strTM02, strTM03, strTM04
   'End If
   
   'Add By Sindy 2010/7/8 檢查商品資料與基本檔商品類別是否一致
   If CheckTMGoodsErr(strTM01, strTM02, strTM03, strTM04, True) = True Then
      Call Command2_Click
      Exit Function
   End If
   
      'add by nickc 2006/06/08
      ModRecord = True
'910910 nick tigger
'***** start
   Exit Function
ErrHand:
    MsgBox (Err.Description)
    cnnConnection.RollbackTrans
'******* end
End Function

'Modify By Sindy 2013/5/3 統一寫在此函數,不用在新增和修改時重覆寫此程式
Private Sub UpdateData(strTM01 As String, strTM02_1 As String, strTM02_2 As String, strTM03 As String, strTM04 As String)
   'add by nickc 2008/02/13 加掛期限
   '716
   strCP09 = GetLastA(strTM01, strTM02_1, strTM02_2, strTM03, strTM04)
   'Add By Sindy 2015/3/20 因全期註冊費101/7/1開始實施,因此到了三年(104/7/1)後則無第二期註冊費的問題了,故不用再Run下段程式
   If Val(strSrvDate(1)) < 20150701 Then
      If IsCreate716 Then
           If m_716CP06 <> "" And m_716CP07 <> "" And m_716Key <> "" Then
                cnnConnection.Execute "update caseprogress set cp06='" & m_716CP06 & "',cp07='" & m_716CP07 & "' where cp09='" & m_716Key & "' "
           ElseIf m_716NP08 <> "" And m_716NP09 <> "" Then
                If m_716Key <> "" Then
                    'Modify By Sindy 2011/5/27
                    'cnnConnection.Execute "update nextprogress set np08='" & m_716NP08 & "',np09='" & m_716NP09 & "' where np01='" & m_716Key & "' "
                    cnnConnection.Execute "update nextprogress set np08='" & m_716NP08 & "',np09='" & m_716NP09 & "' where np01='" & m_716Key & "' and np07=716 and np06 is null "
                Else
                    If strCP09 <> "" Then
                        strNP22 = GetNextProgressNo()
                        strSql = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
                                 "VALUES ('" & strCP09 & "','" & strTM01 & "','" & strTM02_1 & strTM02_2 & "','" & strTM03 & "','" & strTM04 & "',716," & _
                                            DBDATE(m_716NP08) & "," & DBDATE(m_716NP09) & ",'" & PUB_GetAKindSalesNo(strTM01, strTM02_1 & strTM02_2, strTM03, strTM04) & "'," & strNP22 & ")"
                        cnnConnection.Execute strSql
                    End If
                End If
           End If
      End If
   End If '2015/3/20 END
   '102 延展期限
   If IsCreate102 Then
        If m_102CP06 <> "" And m_102CP07 <> "" And m_102Key <> "" Then
             cnnConnection.Execute "update caseprogress set cp06='" & m_102CP06 & "',cp07='" & m_102CP07 & "' where cp09='" & m_102Key & "' "
        ElseIf m_102NP08 <> "" And m_102NP09 <> "" Then
             If m_102Key <> "" Then
                 'Modify By Sindy 2011/5/27
                 'cnnConnection.Execute "update nextprogress set np08='" & m_102NP08 & "',np09='" & m_102NP09 & "' where np01='" & m_102Key & "' "
                 cnnConnection.Execute "update nextprogress set np08='" & m_102NP08 & "',np09='" & m_102NP09 & "' where np01='" & m_102Key & "' and np07=102 and np06 is null "
             Else
                 If strCP09 <> "" Then
                     strNP22 = GetNextProgressNo()
                     strSql = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
                              "VALUES ('" & strCP09 & "','" & strTM01 & "','" & strTM02_1 & strTM02_2 & "','" & strTM03 & "','" & strTM04 & "',102," & _
                                         DBDATE(m_102NP08) & "," & DBDATE(m_102NP09) & ",'" & PUB_GetAKindSalesNo(strTM01, strTM02_1 & strTM02_2, strTM03, strTM04) & "'," & strNP22 & ")"
                     cnnConnection.Execute strSql
                 End If
             End If
        End If
   End If
   'add by sonia 2019/4/11  大陸案有專用期間案件將下一程序檔的被異議續展109期限上N,因為會掛續展期限,T-154644,若不管制續展也不必管制被異議續展
   If textTM22 <> "" And textTM10_1 = "020" Then
      cnnConnection.Execute "update nextprogress set np06='N' where np02='" & strTM01 & "' and np03='" & strTM02_1 & strTM02_2 & "' and np04='" & strTM03 & "' and np05='" & strTM04 & "' and np07=109 and np06 is null "
   End If
   'end 2019/4/11
   '105 使用宣誓(發證後)
   If IsCreate105 Then
        If m_105CP06 <> "" And m_105CP07 <> "" And m_105Key <> "" Then
             cnnConnection.Execute "update caseprogress set cp06='" & m_105CP06 & "',cp07='" & m_105CP07 & "' where cp09='" & m_105Key & "' "
        ElseIf m_105NP08 <> "" And m_105NP09 <> "" Then
             If m_105Key <> "" Then
                 'Modify By Sindy 2011/5/27
                 'cnnConnection.Execute "update nextprogress set np08='" & m_105NP08 & "',np09='" & m_105NP09 & "' where np01='" & m_105Key & "' "
                 cnnConnection.Execute "update nextprogress set np08='" & m_105NP08 & "',np09='" & m_105NP09 & "' where np01='" & m_105Key & "' and np07=105 and np06 is null "
             Else
                 If strCP09 <> "" Then
                     strNP22 = GetNextProgressNo()
                     strSql = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
                              "VALUES ('" & strCP09 & "','" & strTM01 & "','" & strTM02_1 & strTM02_2 & "','" & strTM03 & "','" & strTM04 & "',105," & _
                                         DBDATE(m_105NP08) & "," & DBDATE(m_105NP09) & ",'" & PUB_GetAKindSalesNo(strTM01, strTM02_1 & strTM02_2, strTM03, strTM04) & "'," & strNP22 & ")"
                     cnnConnection.Execute strSql
                 End If
             End If
        End If
   End If
   'add by sonia 2021/9/22
   '105 使用宣誓(發證前)
   If IsCreate105Before Then
      If m_105CP06Before <> "" And m_105CP07Before <> "" And m_105KeyBefore <> "" Then
         cnnConnection.Execute "update caseprogress set cp06='" & m_105CP06Before & "',cp07='" & m_105CP07Before & "' where cp09='" & m_105KeyBefore & "' "
      ElseIf m_105NP08Before <> "" And m_105NP09Before <> "" Then
         If m_105KeyBefore <> "" Then
            cnnConnection.Execute "update nextprogress set np08='" & m_105NP08Before & "',np09='" & m_105NP09Before & "' where np01='" & m_105KeyBefore & "' and np07=105 and np06 is null "
         Else
            If strCP09 <> "" Then
               strNP22 = GetNextProgressNo()
               strSql = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
                        "VALUES ('" & strCP09 & "','" & strTM01 & "','" & strTM02_1 & strTM02_2 & "','" & strTM03 & "','" & strTM04 & "',105," & _
                                   DBDATE(m_105NP08Before) & "," & DBDATE(m_105NP09Before) & ",'" & PUB_GetAKindSalesNo(strTM01, strTM02_1 & strTM02_2, strTM03, strTM04) & "'," & strNP22 & ")"
               cnnConnection.Execute strSql
            End If
         End If
      End If
   End If
   'end 2021/9/22
End Sub

' 刪除記錄
'edit by nickc 2006/06/08
'Private Sub DelRecord()
Private Function DelRecord() As Boolean
Dim strSql As String
Dim strTM01 As String
Dim strTM02 As String
Dim strTM03 As String
Dim strTM04 As String
    
   'add by nickc 2006/06/08
   DelRecord = False
   
   strTM01 = textTM01
   strTM02 = textTM02_1 & textTM02_2
   strTM03 = textTM03
   strTM04 = textTM04
   
   If OnDataDeleteRecord(0, strTM01 & strTM02 & strTM03 & strTM04) <> 0 Then
      GoTo EXITSUB
   End If

   strSql = "DELETE FROM TradeMark " & _
            "WHERE TM01 = '" & textTM01 & "' AND " & _
                  "TM02 = '" & textTM02_1 & textTM02_2 & "' AND " & _
                  "TM03 = '" & textTM03 & "' AND " & _
                  "TM04 = '" & textTM04 & "'"
   'add by nickc 2006/03/16 紀錄分析語法
   On Error GoTo oErr
   cnnConnection.BeginTrans
    Pub_SeekTbLog strSql
    cnnConnection.Execute strSql
    
    'Added by Lydia 2016/11/24 一併刪除各項指示
    strSql = "DELETE FROM INSTRUCTIONS WHERE ITS01=" & CNULL(Pub_GetITS01Type(textTM01)) & " AND ITS02=" & CNULL(textTM01 & textTM02_1 & textTM02_2 & textTM03 & textTM04)
    Pub_SeekTbLog strSql
    cnnConnection.Execute strSql
    'end 2016/11/24
    
   'add by nickc 2006/06/08
   cnnConnection.CommitTrans
   DelRecord = True
   
   ' 只有刪除的是最後一筆才須重新取的第一筆及最後一筆的本所案號
   If (strTM01 = m_LastTM(0) And strTM02 = m_LastTM(1) And strTM03 = m_LastTM(2) And strTM04 = m_LastTM(3)) Or (strTM01 = m_FirstTM(0) And strTM02 = m_FirstTM(1) And strTM03 = m_FirstTM(2) And strTM04 = m_FirstTM(3)) Then
      RefreshRange
   End If
   ShowCurrRecord strTM01, strTM02, strTM03, strTM04
   
EXITSUB:
'add by nickc 2006/06/08
Exit Function
oErr:
    cnnConnection.RollbackTrans
    MsgBox Err.Description
End Function

' 查詢記錄
Private Function QueryRecord() As Boolean
   QueryRecord = False
   
   If IsEmptyText(textTM03) = True Then: textTM03 = "0"
   If IsEmptyText(textTM04) = True Then: textTM04 = "00"
   
   If IsRecordExist(textTM01, textTM02_1 & textTM02_2, textTM03, textTM04) = True Then
      m_CurrTM(0) = textTM01
      m_CurrTM(1) = textTM02_1 & textTM02_2
      m_CurrTM(2) = textTM03
      m_CurrTM(3) = textTM04
      QueryRecord = True
      UpdateCtrlData
   Else
      QueryRecord = False
   End If

   ' 當系統別不為原先所輸入的系統別時則需重新取得範圍
   If textTM01 <> m_CurrTM(0) Then
      RefreshRange
   End If

   UpdateToolbarState
End Function

' 使用者按下確定的按紐
Private Sub OnWork()
Dim strMsg As String
Dim strTit As String
Dim nResponse
Dim StrSQLa As String            '2009/8/19 ADD BY SONIA
Dim rsA As New ADODB.Recordset   '2009/8/19 ADD BY SONIA
   
   Select Case m_EditMode
      Case 1:
         'Add By Cheng 2002/05/23
         '重新檢查欄位有效性
         If TxtValidate = False Then Exit Sub
         If CheckDataValid() = True Then
            'add by nickc 2008/03/28  更新欄位
            UpdateFieldNewData
            'edit by nickc 2006/06/08
            'AddRecord
            If AddRecord = False Then Exit Sub
            RefreshRange
         Else
            GoTo EXITSUB
         End If
      Case 2:
         'Add By Cheng 2002/05/23
         '重新檢查欄位有效性
         If TxtValidate = False Then Exit Sub
         If CheckDataValid() = True Then
            'Added by Lydia 2017/06/19 (存檔前)檢查收文智權人員和客戶智權人員不同業務區要發mail的內文和通知人員
            strChkCuAreaMail = PUB_ChkSameCustSales(textTM01, textTM02_1.Text & IIf(textTM01.Text = "TF", textTM02_2.Text, ""), textTM03, textTM04, "", Trim(textTM23_1), Trim(textTM78_1), Trim(textTM79_1), Trim(textTM80_1), Trim(textTM81_1), strChkCuAreaMailTo)
            
            'add by nickc 2008/03/28  更新欄位
            UpdateFieldNewData
            
            'edit by nickc 2006/06/08
            'ModRecord
            If ModRecord = False Then Exit Sub
            
            Call PUB_SendMailCache 'Added by Lydia 2025/09/12
            
            'Added by Lydia 2017/06/19 檢查收文智權人員和客戶智權人員不同業務區要發mail的內文和通知人員
            If strChkCuAreaMail <> "" Then
               PUB_SendMail strUserNum, strChkCuAreaMailTo, "", "案件收文通知--此案收文非原智權人員(區)！", strChkCuAreaMail
            End If
            'end 2017/06/19
         Else
            GoTo EXITSUB
         End If
      Case 3:
         'edit by nickc 2006/06/08
         'DelRecord
         If DelRecord = False Then Exit Sub
        'add by nickc 2008/03/28  更新欄位
        UpdateFieldNewData
         
         RefreshRange
      Case 4:
         If QueryRecord = False Then
            strMsg = "無此資料"
            strTit = "查詢資料"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            UpdateCtrlData
         '2009/9/8 加註 by sonia T非台灣案非外商收文之案件不必寫程式控制,因為在系統類別外商人員即不可使用T案件
         '2009/8/19 add by sonia FCT無爭議程序之案件內商人員不可查詢(該案有內商承辦人者為FCT爭議案)
         Else
            If textTM01 = "FCT" And Mid(GetStaffDepartment(strUserNum), 1, 2) = "P2" Then
               'modify by sonia 2021/9/23 FCT-047943異議案林靖傑承辦,桂英無法操作
               'StrSQLa = "Select * From CASEPROGRESS,STAFF Where CP01='" & textTM01 & "' AND CP02='" & textTM02_1 & textTM02_2 & "' AND CP03='" & textTM03 & "' AND CP04='" & textTM04 & "' AND CP14=ST01(+) AND SUBSTR(ST03,1,2)='P2' "
               'Modify By Sindy 2025/7/28 FCT案727分析不屬於爭議案 + FCT_NotTMdebate
               StrSQLa = "Select * From CASEPROGRESS,STAFF Where" & _
                         " CP01='" & textTM01 & "' AND CP02='" & textTM02_1 & textTM02_2 & "' AND CP03='" & textTM03 & "' AND CP04='" & textTM04 & "'" & _
                         " AND CP14=ST01(+)" & _
                         " AND (SUBSTR(ST03,1,2)='P2' or (cp10 in (" & TMdebate & ") AND not(CP01='FCT' AND cp10 in (" & FCT_NotTMdebate & ")))" & _
                              ")"
               rsA.CursorLocation = adUseClient
               rsA.Open StrSQLa, cnnConnection, adOpenStatic
               If rsA.RecordCount = 0 Then
                  ClearField
                  textTM01 = m_CurrTM(0): textTM02_1 = m_CurrTM(1): textTM03 = m_CurrTM(2): textTM04 = m_CurrTM(3)
                  strMsg = "非FCT爭議案，您沒有使用該案號資料的權限"
                  strTit = "查詢資料"
                  nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               End If
               If rsA.State <> adStateClosed Then rsA.Close
               Set rsA = Nothing
            End If
         '2009/8/19 END
         End If
   End Select
   m_EditMode = 0
   SetCtrlReadOnly True
EXITSUB:
End Sub

' 開始輸入資料
Private Sub SetInputEntry()
   Select Case m_EditMode
      Case 1: textTM01.SetFocus
      'Modified by Lydia 2023/11/16
      'Case 2: textTM08_1.SetFocus
      Case 2: cboTM08.SetFocus
      '2011/10/27 MODIFY BY SONIA
      'Case 4: textTM01.SetFocus
      Case 4:
         'Modified by Lydia 2025/09/12 TF案也要從流水號開始
         'If textTM01 = "TF" Then
         '   textTM02_2.SetFocus
         'Else
         '   textTM02_1.SetFocus
         'End If
         textTM02_1.SetFocus
         'end 2025/09/12
      '2011/10/27 END
   End Select
End Sub

Private Sub textTM01_GotFocus()
   InverseTextBox textTM01
   CloseIme
End Sub

Private Sub textTM02_1_GotFocus()
   InverseTextBox textTM02_1
End Sub

Private Sub textTM02_2_GotFocus()
   InverseTextBox textTM02_2
End Sub

Private Sub textTM03_GotFocus()
   InverseTextBox textTM03
End Sub

Private Sub textTM04_GotFocus()
   InverseTextBox textTM04
End Sub

Private Sub textTM05_GotFocus()
   InverseTextBox textTM05
   'edit by nickc 2007/06/06 切換輸入法改用API
   OpenIme
End Sub

Private Sub textTM08_1_GotFocus()
   InverseTextBox textTM08_1
End Sub

Private Sub textTM09_GotFocus()
   InverseTextBox textTM09
End Sub

Private Sub textTM10_1_GotFocus()
   InverseTextBox textTM10_1
End Sub

Private Sub textTM11_GotFocus()
   InverseTextBox textTM11
End Sub

Private Sub textTM12_GotFocus()
   InverseTextBox textTM12
End Sub

Private Sub textTM13_GotFocus()
   InverseTextBox textTM13
End Sub

Private Sub textTM14_GotFocus()
   InverseTextBox textTM14
End Sub

Private Sub textTM15_GotFocus()
   InverseTextBox textTM15
End Sub

Private Sub textTM16_GotFocus()
   InverseTextBox textTM16
End Sub

'Add By Sindy 2009/09/09
Private Sub textTM124_GotFocus()
   InverseTextBox textTM124
End Sub
Private Sub textTM125_GotFocus()
   InverseTextBox textTM125
End Sub

Private Sub textTM17_GotFocus()
   InverseTextBox textTM17
End Sub

Private Sub textTM18_GotFocus()
   InverseTextBox textTM18
End Sub

Private Sub textTM19_GotFocus()
   InverseTextBox textTM19
End Sub

Private Sub textTM20_GotFocus()
   InverseTextBox textTM20
End Sub

Private Sub textTM21_GotFocus()
   InverseTextBox textTM21
End Sub

Private Sub textTM22_GotFocus()
   InverseTextBox textTM22
End Sub

Private Sub textTM23_1_GotFocus()
   InverseTextBox textTM23_1
End Sub

Private Sub textTM24_GotFocus()
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textTM24.IMEMode = 1
   OpenIme
   InverseTextBox textTM24
End Sub

Private Sub textTM25_GotFocus()
   InverseTextBox textTM25
End Sub

Private Sub textTM26_GotFocus()
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textTM26.IMEMode = 1
   OpenIme
   InverseTextBox textTM26
End Sub

Private Sub textTM27_GotFocus()
   InverseTextBox textTM27
End Sub

Private Sub textTM29_GotFocus()
   InverseTextBox textTM29
End Sub

Private Sub textTM30_GotFocus()
   InverseTextBox textTM30
End Sub

Private Sub textTM31_1_GotFocus()
   InverseTextBox textTM31_1
End Sub

Private Sub textTM32_GotFocus()
   InverseTextBox textTM32
End Sub

Private Sub textTM33_1_GotFocus()
   InverseTextBox textTM33_1
End Sub

Private Sub textTM34_GotFocus()
   InverseTextBox textTM34
End Sub

Private Sub textTM35_GotFocus()
   InverseTextBox textTM35
End Sub

Private Sub textTM36_GotFocus()
   InverseTextBox textTM36
End Sub

Private Sub textTM37_GotFocus()
   InverseTextBox textTM37
End Sub

'Add By Sindy 2025/3/6
Private Sub textTM140_GotFocus()
   InverseTextBox textTM140
End Sub
Private Sub textTM141_GotFocus()
   InverseTextBox textTM141
End Sub
'2025/3/6 END

Private Sub textTM38_GotFocus()
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textTM38.IMEMode = 1
   OpenIme
   InverseTextBox textTM38
End Sub

Private Sub textTM39_GotFocus()
   InverseTextBox textTM39
End Sub

Private Sub textTM40_GotFocus()
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textTM40.IMEMode = 1
   OpenIme
   InverseTextBox textTM40
End Sub

Private Sub textTM41_GotFocus()
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textTM41.IMEMode = 1
   OpenIme
   InverseTextBox textTM41
End Sub

Private Sub textTM42_GotFocus()
   InverseTextBox textTM42
End Sub

Private Sub textTM43_GotFocus()
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textTM43.IMEMode = 1
   OpenIme
   InverseTextBox textTM43
End Sub

Private Sub textTM44_1_GotFocus()
   InverseTextBox textTM44_1
End Sub

Private Sub textTM45_GotFocus()
   InverseTextBox textTM45
End Sub

Private Sub textTM46_GotFocus()
   InverseTextBox textTM46
End Sub

Private Sub textTM47_GotFocus()
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textTM47.IMEMode = 1
   OpenIme
   InverseTextBox textTM47
End Sub

Private Sub textTM48_GotFocus()
   InverseTextBox textTM48
End Sub

Private Sub textTM49_GotFocus()
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textTM49.IMEMode = 1
   OpenIme
   InverseTextBox textTM49
End Sub

Private Sub textTM50_GotFocus()
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textTM50.IMEMode = 1
   OpenIme
   InverseTextBox textTM50
End Sub

Private Sub textTM51_GotFocus()
   InverseTextBox textTM51
End Sub

Private Sub textTM52_GotFocus()
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textTM52.IMEMode = 1
   OpenIme
   InverseTextBox textTM52
End Sub

Private Sub textTM53_GotFocus()
   InverseTextBox textTM53
End Sub

'Add By Sindy 2009/09/09
Private Sub textTM77_GotFocus()
   InverseTextBox textTM77
End Sub

Private Sub textTM54_1_GotFocus()
   InverseTextBox textTM54_1
End Sub

Private Sub textTM55_GotFocus()
   InverseTextBox textTM55
End Sub

Private Sub textTM56_1_GotFocus()
   InverseTextBox textTM56_1
End Sub

Private Sub textTM57_GotFocus()
   InverseTextBox textTM57
End Sub

Private Sub textTM58_GotFocus()
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textTM58.IMEMode = 1
   OpenIme
   'InverseTextBox textTM58
   textTM58.SetFocus
End Sub

Private Sub textTM65_GotFocus()
   InverseTextBox textTM65
End Sub

Private Sub textTM66_1_GotFocus()
   InverseTextBox textTM66_1
End Sub

Private Sub textTM67_GotFocus()
   InverseTextBox textTM67
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textTM67.IMEMode = 1
   OpenIme
End Sub

Private Sub textTM68_GotFocus()
   InverseTextBox textTM68
End Sub

'Add By Sindy 2009/09/09
Private Sub textTM126_GotFocus()
   InverseTextBox textTM126
End Sub

'Add By Sindy 2013/8/15
Private Sub textTM129_GotFocus()
   InverseTextBox textTM129
End Sub

Private Function CheckDataValid() As Boolean
Dim strTit As String
Dim strMsg As String
Dim nResponse
Dim strTM01 As String, strTM02 As String, strTM03 As String, strTM04 As String
Dim m_rs As New ADODB.Recordset 'Add By Sindy 2015/6/22
Dim strText As String 'Add By Sindy 2015/6/22

   CheckDataValid = False
   Is105OK = True  'add by sonia 2023/5/31
   
   Select Case m_EditMode
      Case 1, 2, 4:
         If IsEmptyText(textTM02_1) = True Then
            strTit = "檢核資料"
            strMsg = "請輸入正確的本所案號"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textTM02_1.SetFocus
            GoTo EXITSUB
         End If
         If Len(textTM04) = 1 Then
            strTit = "檢核資料"
            strMsg = "本所案號輸入不正確"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textTM04.SetFocus
            GoTo EXITSUB
         End If
      Case Else:
   End Select
   
   'add by nickc 2007/02/07 新增時，編號不可大於自動編號
   If m_EditMode = 1 Then
        If ClsPDChkCaseNum(textTM01, textTM02_1 & textTM02_2) Then
            GoTo EXITSUB
        End If
   End If
   
   'Add By Sindy 2015/6/22
   'Set m_rs = New ADODB.Recordset
   '以申請案號+申請國家檢查是否有重覆
   If Trim(textTM12) <> "" And Trim(textTM10_1) <> "" Then
      If m_rs.State = 1 Then m_rs.Close
      'Modified by Lydia 2020/12/01 +ChgSQL
      'modify by sonia 2022/11/25 區分TF
      'strExc(0) = "select tm01,tm02,tm03,tm04 from trademark" & _
      '            " where tm12='" & ChgSQL(textTM12) & "' and tm10='" & textTM10_1 & "'" & _
      '            " and tm01||tm02||tm03||tm04<>'" & textTM01 & textTM02_1 & textTM02_2 & textTM03 & textTM04 & "'"
      If textTM01 = "TF" Then
         strExc(0) = "select tm01,tm02,tm03,tm04 from trademark" & _
                     " where tm12='" & ChgSQL(textTM12) & "' and tm10='" & textTM10_1 & "'" & _
                     " and tm01||substr(tm02,1,5)<>'" & textTM01 & textTM02_1 & "'"
      Else
         strExc(0) = "select tm01,tm02,tm03,tm04 from trademark" & _
                     " where tm12='" & ChgSQL(textTM12) & "' and tm10='" & textTM10_1 & "'" & _
                     " and tm01||tm02||tm03||tm04<>'" & textTM01 & textTM02_1 & textTM02_2 & textTM03 & textTM04 & "'"
      End If
      'end 2022/11/25
      m_rs.CursorLocation = adUseClient
      m_rs.Open strExc(0), cnnConnection, adOpenStatic, adLockReadOnly
      strText = ""
      If m_rs.RecordCount > 0 Then
         m_rs.MoveFirst
         Do While Not m_rs.EOF
            strText = strText & "," & m_rs.Fields("tm01") & "-" & m_rs.Fields("tm02") & "-" & m_rs.Fields("tm03") & "-" & m_rs.Fields("tm04")
            m_rs.MoveNext
         Loop
         If strText <> "" Then
            strText = Mid(strText, 2)
            If MsgBox("此申請案號已存在(" & strText & "), 請確認是否仍要存檔？", vbInformation + vbYesNo + vbDefaultButton2, "注意！！") = vbNo Then
               GoTo EXITSUB
            End If
         End If
      End If
   End If
   '以審定號+申請國家檢查資料庫中相同者且目前准駁(TM16)欄為'1'者是否有重覆
   If Trim(textTM15) <> "" And Trim(textTM10_1) <> "" Then
      If m_rs.State = 1 Then m_rs.Close
      'modify by sonia 2022/11/25 區分TF
      'strExc(0) = "select tm01,tm02,tm03,tm04 from trademark" & _
      '            " where tm15='" & textTM15 & "' and tm10='" & textTM10_1 & "' and tm16='1'" & _
      '            " and tm01||tm02||tm03||tm04<>'" & textTM01 & textTM02_1 & textTM02_2 & textTM03 & textTM04 & "'"
      If textTM01 = "TF" Then
         strExc(0) = "select tm01,tm02,tm03,tm04 from trademark" & _
                     " where tm15='" & textTM15 & "' and tm10='" & textTM10_1 & "' and tm16='1'" & _
                     " and tm01||substr(tm02,1,5)<>'" & textTM01 & textTM02_1 & "'"
      Else
         strExc(0) = "select tm01,tm02,tm03,tm04 from trademark" & _
                     " where tm15='" & textTM15 & "' and tm10='" & textTM10_1 & "' and tm16='1'" & _
                     " and tm01||tm02||tm03||tm04<>'" & textTM01 & textTM02_1 & textTM02_2 & textTM03 & textTM04 & "'"
      End If
      'end 2022/11/25
      m_rs.CursorLocation = adUseClient
      m_rs.Open strExc(0), cnnConnection, adOpenStatic, adLockReadOnly
      strText = ""
      If m_rs.RecordCount > 0 Then
         m_rs.MoveFirst
         Do While Not m_rs.EOF
            strText = strText & "," & m_rs.Fields("tm01") & "-" & m_rs.Fields("tm02") & "-" & m_rs.Fields("tm03") & "-" & m_rs.Fields("tm04")
            m_rs.MoveNext
         Loop
         If strText <> "" Then
            strText = Mid(strText, 2)
            If MsgBox("此審定號已存在(" & strText & "), 請確認是否仍要存檔？", vbInformation + vbYesNo + vbDefaultButton2, "注意！！") = vbNo Then
               GoTo EXITSUB
            End If
         End If
      End If
   End If
   '2015/6/22 END
   
   'add by sonia 2022/10/12 補輸審定號時檢查
   If textTM28 = "申請" And textTM01 <> "CFT" And textTM15.Tag = "" And textTM15 <> "" Then
      If textTM16 = "" Then
         strTit = "檢核資料"
         strMsg = "有審定號，請輸入目前准駁！並同時確認卷宗性質是否正確！"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textTM16.SetFocus
         GoTo EXITSUB
      End If
   End If
   'end 2022/10/12
   
   Select Case m_EditMode
      Case 1, 2:
         ' 商標種類不可空白
         If IsEmptyText(textTM08_1) = True Then
            strTit = "檢核資料"
            strMsg = "請輸入商標種類"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            'Modified by Lydia 2023/11/16
            'textTM08_1.SetFocus
            cboTM08.SetFocus
            GoTo EXITSUB
         End If
         'Add By Sindy 2013/7/17
         '台灣案時, 檢查商標種類不可輸入2,4,5,6
         If Trim(textTM10_1) = "000" And _
            (Trim(textTM08_1) = "2" Or Trim(textTM08_1) = "4" Or Trim(textTM08_1) = "5" Or Trim(textTM08_1) = "6") Then
            strTit = "檢核資料"
            strMsg = "台灣案時, 商標種類不可輸入2,4,5,6！"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            'Modified by Lydia 2023/11/16
            'textTM08_1.SetFocus
            cboTM08.SetFocus
            GoTo EXITSUB
         End If
         '2013/7/17 End
         ' 申請國家不可為空白
         If IsEmptyText(textTM10_1) = True Then
            strTit = "檢核資料"
            strMsg = "請輸入申請國家"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textTM10_1.SetFocus
            GoTo EXITSUB
         End If
         ' 案件名稱不可空白
         If IsEmptyText(textTM05) = True Then
            strTit = "檢核資料"
            strMsg = "案件名稱不可空白"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textTM05.SetFocus
            GoTo EXITSUB
         End If
         'Modify By Cheng 2002/07/12
         '申請人及代理人不可同時空白
'         ' 申請人編號不可空白
'         If IsEmptyText(textTM23_1) = True Then
         If IsEmptyText(textTM23_1) = True And IsEmptyText(textTM44_1) = True Then
            strTit = "檢核資料"
            strMsg = "申請人及FC代理人不可同時空白"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textTM23_1.SetFocus
            GoTo EXITSUB
         End If
         
         'Add by Morgan 2007/5/10
         If Not ((textTM29.Text = "" And textTM30.Text = "" And textTM31_1.Text = "") Or (textTM29.Text <> "" And textTM30.Text <> "" And textTM31_1.Text <> "")) Then
            MsgBox "是否閉卷、閉卷日期、閉卷原因三個欄位須同時空白或有值！", vbExclamation
            GoTo EXITSUB
         End If
         'end 2007/5/10
         
         'add by nickc 2008/01/31 若是有專用期間，要檢查  專用權是否存在  不可以空白
         If Trim(textTM21) <> "" And Trim(textTM22) <> "" Then
            If Trim(textTM17) = "" Then
                MsgBox "有 專用期間，專用權是否存在 不可以空白！", vbExclamation
                textTM17.SetFocus
                textTM17_GotFocus
                GoTo EXITSUB
            End If
         End If
         
         'Modify By Sindy 2013/5/14 移為函數
         strTM01 = textTM01
         strTM02 = textTM02_1
         If textTM01 = "TF" Then: strTM02 = strTM02 & textTM02_2
         strTM03 = textTM03
         If IsEmptyText(strTM03) = True Then: strTM03 = "0"
         strTM04 = textTM04
         If IsEmptyText(strTM04) = True Then: strTM04 = "00"
         If strTM01 = "TF" And strTM03 & strTM04 = "000" Then
            '母案在存檔時執行
         ElseIf strTM01 = "TF" And strTM03 & strTM04 <> "000" Then
            '若是TF案件並且輸入的維護案號是子案時,要抓取其母案資料才能進行其檢核
            strExc(0) = "select TM01,TM02,TM03,TM04,TM10,TM29,TM11,TM14,TM20," & _
                        "TM21,TM22,decode(TM28,'1','申請','2','異議','3','評定','4','廢止') as TM28,TM58" & _
                        " from trademark where tm01='" & strTM01 & "' and tm02='" & strTM02 & "' and tm03='0' and tm04='00'"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               Call CheckDeadLine(strTM01, strTM02, strTM03, strTM04, textTM10_1, textTM29, textTM58, _
                                  IIf(Val("" & RsTemp.Fields("TM11")) > 0, TAIWANDATE("" & RsTemp.Fields("TM11")), ""), _
                                  IIf(Val("" & RsTemp.Fields("TM14")) > 0, TAIWANDATE("" & RsTemp.Fields("TM14")), ""), _
                                  IIf(Val("" & RsTemp.Fields("TM20")) > 0, TAIWANDATE("" & RsTemp.Fields("TM20")), ""), _
                                  IIf(Val("" & RsTemp.Fields("TM21")) > 0, TAIWANDATE("" & RsTemp.Fields("TM21")), ""), _
                                  IIf(Val("" & RsTemp.Fields("TM22")) > 0, TAIWANDATE("" & RsTemp.Fields("TM22")), ""), _
                                  "" & RsTemp.Fields("TM28"))
            End If
         Else
            Call CheckDeadLine(strTM01, strTM02, strTM03, strTM04, textTM10_1, textTM29, textTM58, _
                               textTM11, textTM14, textTM20, textTM21, textTM22, textTM28)
            If Is105OK = False Then GoTo EXITSUB 'add by sonia 2023/5/31
         End If
         '2013/5/14 End
         
         'Add By Sindy 2016/11/23
         If Trim(Me.Combo4.Text) <> "" Then
            '若輸入幣別就一定要選格式
            If Trim(Me.Combo5.Text) = "" Then
               ShowMsg "請款單列印幣別格式不可空白 !"
               Me.Combo5.SetFocus
               GoTo EXITSUB
            End If
            '請款幣別<>NTD時不可輸入1
            If Trim(Me.Combo4.Text) <> "NTD" And Me.Combo5.ListIndex = 1 Then
               ShowMsg "請款幣別<>NTD時，請款單列印幣別格式不可選純台幣 !"
               Me.Combo5.SetFocus
               GoTo EXITSUB
            End If
         End If
         '2016/11/23 ENd
      Case Else:
   End Select
   
   'Added by Lydia 2021/11/24 脫歐英國案檢查: 申請國家tm10='201'英國、卷宗性質tm28='1'申請、申請案號tm12或審定號tm15前5碼為'UK009'，若沒有建立歐盟239案的相關卷號時，則存檔時彈訊息"此為脫歐英國案，若歐盟案亦為本所案件，請建立相關卷號關聯，否則此英國案在計算結餘時會扣安全基金！"，但仍可存檔。
   If textTM01 = "CFT" And textTM10_1 = "201" And textTM28 = "申請" And (Left(textTM12, 5) = "UK009" Or Left(textTM15, 5) = "UK009") Then
      strExc(0) = "select cr05 as d01,cr06 as d02,cr07 as d03,cr08 as d04 from caserelation1,trademark where cr01='" & textTM01 & "' and cr02='" & textTM02_1 & textTM02_2 & "' and cr03='" & textTM03 & "' and cr04='" & textTM04 & "' and cr05='" & textTM01 & "' and cr05=tm01(+) and cr06=tm02(+) and cr07=tm03(+) and cr08=tm04(+) and tm10='239' " & _
                       "union select cr01 as d01,cr02 as d02,cr03 as d03,cr04 as d04 from caserelation1,trademark where cr05='" & textTM01 & "' and cr06='" & textTM02_1 & textTM02_2 & "' and cr07='" & textTM03 & "' and cr08='" & textTM04 & "' and cr01='" & textTM01 & "' and cr01=tm01(+) and cr02=tm02(+) and cr03=tm03(+) and cr04=tm04(+) and tm10='239' "
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 0 Then
          MsgBox "此為脫歐英國案，若歐盟案亦為本所案件，請建立相關卷號關聯，否則此英國案在計算結餘時會扣安全基金！", vbExclamation, "英國脫歐案管制"
      End If
   End If
   'end 2021/11/24
   
   'Added by Lydia 2025/09/12 TF基礎案號設定
   'Modified by Lydia 2025/10/23
   If textTM01 = "TF" And cmdTFBaseNo.Visible = True Then
      If textTM28 = "申請" Then
         'TF案未閉卷(無專用期 or 專用期未過)，卷宗性質為申請之母案案號
         If textTM29 = "" And (Trim(textTM22) = "" Or DBDATE(textTM22) > strSrvDate(1)) Then
            strExc(0) = Pub_GetField("TFBaseNo", "TFBN01='" & textTM01 & "' AND TFBN02='" & textTM02_1 & textTM02_2 & "' AND TFBN03='" & textTM03 & "' AND TFBN04='" & textTM04 & "'", "TFBN05")
            If strExc(0) <> "" Then
                cmdTFBaseNo.BackColor = &HC0FFC0
            Else
                cmdTFBaseNo.BackColor = &H8000000F
                If MsgBox("是否要設定TF基礎案號？", vbInformation + vbYesNo + vbDefaultButton1) = vbYes Then
                   tabCtrl.Tab = 6
                   Exit Function
                End If
            End If
         End If
      End If
   End If
   'end 2025/09/12
   
   CheckDataValid = True
EXITSUB:
   Set m_rs = Nothing 'Add By Sindy 2015/6/22
End Function

Private Sub CheckDeadLine(strTM01 As String, strTM02 As String, strTM03 As String, strTM04 As String, strTM10 As String, strTM29 As String, _
                          strTM58 As String, strTM11 As String, strTM14 As String, strTM20 As String, strTM21 As String, strTM22 As String, _
                          strTM28 As String, _
                          Optional ByRef strUpdTM58_716 = "", Optional ByRef strUpdTM58_102 = "", Optional ByRef strUpdTM58_105 = "")
'add by nickc 2008/02/05
Dim m_rs As New ADODB.Recordset
Dim m_rs2 As New ADODB.Recordset
Dim m_str As String
Dim ii As Integer
Dim str105DATE As String   '使用宣誓起算日  2011/9/28 ADD BY SONIA
   
   'add by nickc 2008/01/31 新增修改時才檢查
   IsCreate716 = False
   IsCreate102 = False
   IsCreate105 = False
   IsCreate105Before = False  'add by sonia 2021/9/24
   m_716CP06 = ""
   m_716CP07 = ""
   m_716NP08 = ""
   m_716NP09 = ""
   m_102CP06 = ""
   m_102CP07 = ""
   m_102NP08 = ""
   m_102NP09 = ""
   m_105CP06 = ""
   m_105CP07 = ""
   m_105NP08 = ""
   m_105NP09 = ""
   'add by sonia 2021/9/24
   m_105CP06Before = ""
   m_105CP07Before = ""
   m_105NP08Before = ""
   m_105NP09Before = ""
   'end 2021/9/24
   m_716Key = ""
   m_102Key = ""
   m_105Key = ""
   m_105KeyBefore = ""   'add by sonia 2021/9/24
   m_716tmpDate1 = ""
   m_716tmpDate2 = ""
   m_102tmpDate1 = ""
   m_102tmpDate2 = ""
   m_105tmpDate1 = ""
   m_105tmpDate2 = ""
   m_105tmpDate1Before = ""   'add by sonia 2021/9/24
   m_105tmpDate2Before = ""   'add by sonia 2021/9/24
   If m_EditMode = 1 Or m_EditMode = 2 Then
      '若已經收文但未掛期限  或  期限不同，或者  未收文未掛期限  或  已掛期限但期限不同者  顯示訊息告知正確之期限，提示要掛期限或是更新期限
      '使用者可以選擇是否要處裡
      '若要
      '若已收文未發文==>更新已收文期限
      '    未收文            ==>檢查是否已掛期限
      '                                                       是    ==>   更新期限
      '                                                       否    ==>   新增期限  總收文號為該案之最後收文A類收文號
      '第二期  卷宗為  申請  申請國家為台灣 申請日在 92/11/28 前(不含) 公告在 92/09/01 後(含)有專用期且未過期
      '2008/6/18 modify by sonia 申請日判斷錯誤,另加入申請日在 92/11/28以後(含)已公告者都要檢查
      'If (strTM01 = "T" Or strTM01 = "FCT") And strTM28 = "申請" And strTM10 = "000" And DBDATE(strTM11) > "20031128" And DBDATE(strTM14) >= "20030901" And strTM21 <> "" And strTM22 <> "" And DBDATE(strTM22) >= strSrvDate(1) And InStr(1, strTM58, "第二期") = 0 Then
      'Modify By Sindy 2013/3/12 若為閉卷案不掛期限
      'Add By Sindy 2015/3/20 因全期註冊費101/7/1開始實施,因此到了三年(104/7/1)後則無第二期註冊費的問題了,故不用再Run下段程式
      If Val(strSrvDate(1)) < 20150701 Then
         If (strTM01 = "T" Or strTM01 = "FCT") And _
            strTM28 = "申請" And strTM10 = "000" And _
            ((DBDATE(strTM11) < "20031128" And DBDATE(strTM14) >= "20030901") Or DBDATE(strTM11) >= "20031128") And _
            strTM21 <> "" And strTM22 <> "" And DBDATE(strTM22) >= strSrvDate(1) And _
            InStr(1, strTM58, "第二期") = 0 And _
            strTM29 <> "Y" Then
            
            m_716tmpDate2 = DBDATE(DateAdd("d", -1, DateAdd("yyyy", 3, ChangeWStringToWDateString(DBDATE(strTM21)))))
            'Modify By Sindy 2014/11/20 台灣案之本所期限設定
            If strTM10 = "000" And Val(strSrvDate(1)) >= 台灣案所限新規則啟用日 Then
               m_716tmpDate1 = PUB_GetOurDeadline(DBDATE(m_716tmpDate2))
            Else
            '2014/11/20 END
               m_716tmpDate1 = DBDATE(DateAdd("d", -2, ChangeWStringToWDateString(DBDATE(m_716tmpDate2))))
            End If
             m_716tmpDate1 = PUB_GetWorkDay1(m_716tmpDate1, True) 'Added by Lydia 2020/07/07 本所期限檢查：若本所期限非工作天則直接調整至最近的工作天
             
             '檢查有無收文 717
             Set m_rs = New ADODB.Recordset
             'edit by nickc 2008/02/29 秀玲說，第二期不用管三年內發文，也不用管有沒有發文
             'm_str = "select * from caseprogress where cp01='" & strTM01 & "' and cp02='" & strTM02 & "' and cp03='" & strTM03 & "' and cp04='" & strTM04 & "' and cp10='717' and cp57 is null  and (cp27 is null or cp27>=" & DBDATE(DateAdd("yyyy", -3, ChangeWStringToWDateString(strSrvDate(1)))) & ")  "
             '2008/7/11 MODIFY BY SONIA 全期也不用管
             'm_str = "select * from caseprogress where cp01='" & strTM01 & "' and cp02='" & strTM02 & "' and cp03='" & strTM03 & "' and cp04='" & strTM04 & "' and cp10='717' and cp57 is null  "
             m_str = "select * from caseprogress where cp01='" & strTM01 & "' and cp02='" & strTM02 & "' and cp03='" & strTM03 & "' and cp04='" & strTM04 & "' and cp10>='716' AND CP10<='717' and cp57 is null  "
             If m_rs.State = 1 Then m_rs.Close
             m_rs.CursorLocation = adUseClient
             m_rs.Open m_str, cnnConnection, adOpenStatic, adLockReadOnly
             If Not m_rs.EOF And Not m_rs.BOF Then
                 '已經有，檢查是否無期限或期限不同
                 m_716CP06 = m_716tmpDate1
                 m_716CP07 = m_716tmpDate2
                 If Val(CheckStr(m_rs.Fields("cp07"))) <> Val(m_716CP07) Or CheckStr(m_rs.Fields("cp07")) = "" Then
                     If CheckStr(m_rs.Fields("cp27")) <> "" Then
                         '2008/2/14 cancel by sonia FCT-25023,不必提醒
                         'MsgBox "第二期註冊費已發文，請自行處裡期限！", vbInformation, "注意！"
                     Else
                         If MsgBox("第二期註冊費已收文，" & IIf(CheckStr(m_rs.Fields("cp07")) = "", "但未掛期限，系統計算的期限為：" & ChangeWStringToTDateString(m_716CP07), "目前掛的期限為：" & ChangeWStringToTDateString(CheckStr(m_rs.Fields("cp07"))) & "，系統計算的期限為：" & ChangeWStringToTDateString(m_716CP07)) & "，是否更新？", vbInformation + vbYesNo, "注意！！") = vbYes Then
                             IsCreate716 = True
                             If Val(CheckStr(m_rs.Fields("cp07"))) <> Val(m_716CP07) Then
                                 m_716Key = CheckStr(m_rs.Fields("cp09"))
                             End If
                         Else
                             m_716CP07 = ""
                             m_716CP06 = ""
                         End If
                     End If
                 Else
                     '期限相同，不用變更
                     m_716CP07 = ""
                     m_716CP06 = ""
                 End If
             Else
                 m_716NP08 = m_716tmpDate1
                 m_716NP09 = m_716tmpDate2
                 '檢查有無掛期限
                 Set m_rs = New ADODB.Recordset
                 m_str = "select * from nextprogress where np02='" & strTM01 & "' and np03='" & strTM02 & "' and np04='" & strTM03 & "' and np05='" & strTM04 & "' and np07=716 and np06 is null "
                 If m_rs.State = 1 Then m_rs.Close
                 m_rs.CursorLocation = adUseClient
                 m_rs.Open m_str, cnnConnection, adOpenStatic, adLockReadOnly
                 If Not m_rs.EOF And Not m_rs.BOF Then
                     If Val(CheckStr(m_rs.Fields("np09"))) <> Val(m_716NP09) Or CheckStr(m_rs.Fields("np09")) = "" Then
                         If MsgBox("第二期註冊費" & IIf(CheckStr(m_rs.Fields("np09")) = "", "未掛期限，系統計算的期限為：" & ChangeWStringToTDateString(m_716NP09), "目前掛的期限為：" & ChangeWStringToTDateString(CheckStr(m_rs.Fields("np09"))) & "，系統計算的期限為：" & ChangeWStringToTDateString(m_716NP09)) & "，是否更新？", vbInformation + vbYesNo, "注意！！") = vbYes Then
                             IsCreate716 = True
                             If Val(CheckStr(m_rs.Fields("np09"))) <> Val(m_716NP09) Then
                                 m_716Key = CheckStr(m_rs.Fields("np01"))
                             End If
                         Else
                             m_716NP09 = ""
                             m_716NP08 = ""
                             If CheckStr(m_rs.Fields("np09")) = "" Then
                                 textTM58 = textTM58 & ";不管制第二期"
                                 If strUpdTM58_716 = "" Then strUpdTM58_716 = ";不管制第二期" 'Add By Sindy 2013/5/14
                             End If
                         End If
                     Else
                         '期限相同，不用變更
                         m_716NP09 = ""
                         m_716NP08 = ""
                     End If
                 Else
                         If Val(m_716NP09) < Val(strSrvDate(1)) Then    '已過期
                             If MsgBox("第二期註冊費已過期，是否仍要管制？", vbInformation + vbYesNo, "注意！！") = vbYes Then
                                 IsCreate716 = True
                             Else
                                 m_716NP09 = ""
                                 m_716NP08 = ""
                                 textTM58 = textTM58 & ";不管制第二期"
                                 If strUpdTM58_716 = "" Then strUpdTM58_716 = ";不管制第二期" 'Add By Sindy 2013/5/14
                             End If
                         Else
                             If MsgBox("是否要管制第二期？", vbInformation + vbYesNo, "注意！！") = vbYes Then
                                 IsCreate716 = True
                             Else
                                 m_716NP09 = ""
                                 m_716NP08 = ""
                                 textTM58 = textTM58 & ";不管制第二期"
                                 If strUpdTM58_716 = "" Then strUpdTM58_716 = ";不管制第二期" 'Add By Sindy 2013/5/14
                             End If
                         End If
                 End If
             End If
         End If
      End If '2015/3/20 END
      
     '延展 and 使用宣誓  卷宗為申請，專用期未過期
      'Modify By Sindy 2013/7/2 若為閉卷案不掛期限
      If strTM28 = "申請" And strTM21 <> "" And strTM22 <> "" And Val(DBDATE(strTM22)) >= Val(strSrvDate(1)) And _
         strTM29 <> "Y" Then
         'Modify By Sindy 2013/7/17 馬德里只有母案一筆需要檢查延展期限問題
         If (strTM01 <> "TF" Or (strTM01 = "TF" And Mid(strTM02, 6, 1) = "0" And strTM03 & strTM04 = "000")) Then
            m_102tmpDate2 = DBDATE(strTM22)
            If strTM10 = "000" Or strTM10 = "020" Then  '台灣、大陸為 - 2 天
              'Modify By Sindy 2014/11/20 台灣案之本所期限設定
              If strTM10 = "000" And Val(strSrvDate(1)) >= 台灣案所限新規則啟用日 Then
                 m_102tmpDate1 = PUB_GetOurDeadline(DBDATE(m_102tmpDate2))
              Else
              '2014/11/20 END
                 m_102tmpDate1 = DBDATE(DateAdd("d", -2, ChangeWStringToWDateString(DBDATE(m_102tmpDate2))))
              End If
            ElseIf strTM01 = "TF" Then    ' TF 為 - 1個月
                m_102tmpDate1 = DBDATE(DateAdd("m", -1, ChangeWStringToWDateString(DBDATE(m_102tmpDate2))))
            Else   '其他為 - 2個月
                m_102tmpDate1 = DBDATE(DateAdd("m", -2, ChangeWStringToWDateString(DBDATE(m_102tmpDate2))))
            End If
            m_102tmpDate1 = PUB_GetWorkDay1(m_102tmpDate1, True) 'Added by Lydia 2020/07/07 本所期限檢查：若本所期限非工作天則直接調整至最近的工作天
            '檢查有無收文 102
            Set m_rs = New ADODB.Recordset
            '2010/11/30 modify by sonia 抓延展未核准者FCT-016597
            'm_str = "select * from caseprogress where cp01='" & strTM01 & "' and cp02='" & strTM02 & "' and cp03='" & strTM03 & "' and cp04='" & strTM04 & "' and cp10='102' and cp57 is null  and (cp27 is null or cp27>=" & DBDATE(DateAdd("yyyy", -3, ChangeWStringToWDateString(strSrvDate(1)))) & ")  "
            'modify by sonia 2022/12/28 加AND (CP07 IS NULL OR CP07>" & (m_102tmpDate2 - 30000) & ")以免抓到前次延展資料 CFT-014018
            'm_str = "select * from caseprogress where cp01='" & strTM01 & "' and cp02='" & strTM02 & "' and cp03='" & strTM03 & "' and cp04='" & strTM04 & "' and cp10='102' and cp57 is null  and (cp27 is null or cp27>=" & DBDATE(DateAdd("yyyy", -3, ChangeWStringToWDateString(strSrvDate(1)))) & ") and cp24 is null "
            m_str = "select * from caseprogress where cp01='" & strTM01 & "' and cp02='" & strTM02 & "' and cp03='" & strTM03 & "' and cp04='" & strTM04 & "' and cp10='102' and cp57 is null  and (cp27 is null or cp27>=" & DBDATE(DateAdd("yyyy", -3, ChangeWStringToWDateString(strSrvDate(1)))) & ") and cp24 is null AND (CP07 IS NULL OR CP07>" & (m_102tmpDate2 - 30000) & ") "
            If m_rs.State = 1 Then m_rs.Close
            m_rs.CursorLocation = adUseClient
            m_rs.Open m_str, cnnConnection, adOpenStatic, adLockReadOnly
            If Not m_rs.EOF And Not m_rs.BOF Then
                '已經有，檢查是否無期限或期限不同
                m_102CP06 = m_102tmpDate1
                m_102CP07 = m_102tmpDate2
                If Val(CheckStr(m_rs.Fields("cp07"))) <> Val(m_102CP07) Or CheckStr(m_rs.Fields("cp07")) = "" Then
                    If CheckStr(m_rs.Fields("cp27")) <> "" Then
                        '2008/2/14 cancel by sonia FCT-25023第二期,不必提醒
                        'MsgBox "延展已發文，請自行處裡期限！", vbInformation, "注意！"
                    Else
                        If strTM04 = "00" Then 'Add By Sindy 2012/6/15 +if 若為00者才須提醒
                          If MsgBox("延展已收文，" & IIf(CheckStr(m_rs.Fields("cp07")) = "", "但未掛期限，系統計算的期限為：" & ChangeWStringToTDateString(m_102CP07), "目前掛的期限為：" & ChangeWStringToTDateString(CheckStr(m_rs.Fields("cp07"))) & "，系統計算的期限為：" & ChangeWStringToTDateString(m_102CP07)) & "，是否更新？", vbInformation + vbYesNo, "注意！！") = vbYes Then
                              IsCreate102 = True
                              If Val(CheckStr(m_rs.Fields("cp07"))) <> Val(m_102CP07) Then
                                  m_102Key = CheckStr(m_rs.Fields("cp09"))
                              End If
                          Else
                              m_102CP07 = ""
                              m_102CP06 = ""
                          End If
                        End If
                    End If
                Else
                    '期限相同，不用變更
                    m_102CP07 = ""
                    m_102CP06 = ""
                End If
            Else
                m_102NP08 = m_102tmpDate1
                m_102NP09 = m_102tmpDate2
                '檢查有無掛期限
                Set m_rs = New ADODB.Recordset
                m_str = "select * from nextprogress where np02='" & strTM01 & "' and np03='" & strTM02 & "' and np04='" & strTM03 & "' and np05='" & strTM04 & "' and np07=102 and np06 is null  "
                If m_rs.State = 1 Then m_rs.Close
                m_rs.CursorLocation = adUseClient
                m_rs.Open m_str, cnnConnection, adOpenStatic, adLockReadOnly
                If Not m_rs.EOF And Not m_rs.BOF Then
                    If Val(CheckStr(m_rs.Fields("np09"))) <> Val(m_102NP09) Or CheckStr(m_rs.Fields("np09")) = "" Then
                        If strTM04 = "00" Then 'Add By Sindy 2012/6/15 +if 若為00者才須提醒
                          If MsgBox("延展" & IIf(CheckStr(m_rs.Fields("np09")) = "", "未掛期限，系統計算的期限為：" & ChangeWStringToTDateString(m_102NP09), "目前掛的期限為：" & ChangeWStringToTDateString(CheckStr(m_rs.Fields("np09"))) & "，系統計算的期限為：" & ChangeWStringToTDateString(m_102NP09)) & "，是否更新？", vbInformation + vbYesNo, "注意！！") = vbYes Then
                              IsCreate102 = True
                              If Val(CheckStr(m_rs.Fields("np09"))) <> Val(m_102NP09) Then
                                  m_102Key = CheckStr(m_rs.Fields("np01"))
                              End If
                          Else
                              m_102NP09 = ""
                              m_102NP08 = ""
                              '2012/3/29 ADD BY SONIA
                              If CheckStr(m_rs.Fields("np09")) = "" Then
                                  textTM58 = textTM58 & ";不管制延展"
                                  If strUpdTM58_102 = "" Then strUpdTM58_102 = ";不管制延展" 'Add By Sindy 2013/5/14
                              End If
                              '2012/3/29 END
                          End If
                        End If
                    Else
                        '期限相同，不用變更
                        m_102NP09 = ""
                        m_102NP08 = ""
                    End If
                Else
                        If strTM04 = "00" Then 'Add By Sindy 2012/6/15 +if 若為00者才須提醒
                          If Val(m_102NP09) < Val(strSrvDate(1)) Then    '已過期
                              If MsgBox("延展已過期，是否仍要管制？", vbInformation + vbYesNo, "注意！！") = vbYes Then
                                  IsCreate102 = True
                              Else
                                  m_102NP09 = ""
                                  m_102NP08 = ""
                                  textTM58 = textTM58 & ";不管制延展"   '2012/3/29 ADD BY SONIA
                                  If strUpdTM58_102 = "" Then strUpdTM58_102 = ";不管制延展" 'Add By Sindy 2013/5/14
                              End If
                          Else
                              If MsgBox("是否要管制延展？", vbInformation + vbYesNo, "注意！！") = vbYes Then
                                  IsCreate102 = True
                              Else
                                  m_102NP09 = ""
                                  m_102NP08 = ""
                                  textTM58 = textTM58 & ";不管制延展"   '2012/3/29 ADD BY SONIA
                                  If strUpdTM58_102 = "" Then strUpdTM58_102 = ";不管制延展" 'Add By Sindy 2013/5/14
                              End If
                          End If
                        End If
                End If
            End If
         End If
         '使用宣誓
         'Modify By Sindy 2013/5/14 +TF
         'If strTM01 = "CFT" Then
         If strTM01 = "CFT" Or strTM01 = "TF" Then
         '2013/5/14 End
             '2009/10/27 modify by sonia 應判斷na38,CFT-011013
             'm_str = "Select NA39,NA38 From Nation Where NA01='" & strTM10 & "' AND NA39 IS NOT NULL "
             'modify by sonia 2023/9/15 +na13,na14,na78  CFT-023278
             m_str = "Select NA39,NA38,na13,na14,na78 From Nation Where NA01='" & strTM10 & "' AND NA38 IS NOT NULL "
             If m_rs2.State = 1 Then m_rs2.Close
              m_rs2.CursorLocation = adUseClient
             m_rs2.Open m_str, cnnConnection, adOpenStatic, adLockReadOnly
             ii = 0
             If m_rs2.RecordCount > 0 Then
                 'add by sonia 2023/5/31
                 If Val(strTM22) <> 0 And Val(strTM20) = 0 Then
                     ShowMsg "已發證案件要計算使用宣誓期限，請輸入發證日 !"
                     Me.textTM20.SetFocus
                     Is105OK = False
                     Exit Sub
                 End If
                 'end 2023/5/31
                 '2011/9/28 modify by sonia 改以發證日計算使用宣誓法定期限
                 'm_105tmpDate2 = DBDATE(Format(DateSerial(Val(DBYEAR(strTM21)) + Val(m_rs2.Fields(1).Value), Val(DBMONTH(strTM21)), Val(DBDAY(strTM21)))))
                 m_105tmpDate2 = DBDATE(Format(DateSerial(Val(DBYEAR(strTM20)) + Val(m_rs2.Fields(1).Value), Val(DBMONTH(strTM20)), Val(DBDAY(strTM20)))))
                 'add by sonia 2024/9/4 TF子案046柬埔寨改為子案公告日起算(TF-000833-1-03)
                 If strTM01 = "TF" And strTM10 = "046" Then
                     m_105tmpDate2 = DBDATE(Format(DateSerial(Val(DBYEAR(textTM14)) + Val(m_rs2.Fields(1).Value), Val(DBMONTH(textTM14)), Val(DBDAY(textTM14)))))
                 End If
                 'end 2024/9/4
                 
                 'add by sonia 2018/11/19  104墨西哥案為三年加三個月
                 'modify by sonia 2023/9/15 110海地案為五年加三個月
                 If strTM10 = "104" Or strTM10 = "110" Then
                    m_105tmpDate2 = CompDate(1, 3, m_105tmpDate2)
                 End If
                 'end 2018/11/19
                 If Val(m_105tmpDate2) >= Val(strSrvDate(1)) Then
                     If Val(m_105tmpDate2) >= Val(DBDATE(strTM22)) Then
                         m_105tmpDate2 = ""
                         m_105tmpDate1 = ""
                     Else
                         m_105tmpDate1 = DBDATE(DateAdd("m", -2, ChangeWStringToWDateString(DBDATE(m_105tmpDate2))))
                     End If
                 Else
                    If Not IsNull(m_rs2.Fields(0)) Then    '2009/10/27 add by sonia CFT-011013
                       '2011/9/28 add by sonia 改以發證日計算使用宣誓法定期限,但柬埔寨延展後以延展核准日計算
                       str105DATE = Val(strTM20)
                       If strTM10 = "046" Then
                          Set m_rs = New ADODB.Recordset
                          'Modified by Lydia 2017/12/21
                          'm_str = "select max(cp09||cp25) from caseprogress where cp01='" & strTM01 & "' and cp02='" & strTM02 & "' and cp03='" & strTM03 & "' and cp04='" & strTM04 & "' and cp10='102' and cp27 is not null and cp24='1' "
                          m_str = "select nvl(max(cp09||cp25),0) from caseprogress where cp01='" & strTM01 & "' and cp02='" & strTM02 & "' and cp03='" & strTM03 & "' and cp04='" & strTM04 & "' and cp10='102' and cp27 is not null and cp24='1' "
                          If m_rs.State = 1 Then m_rs.Close
                          m_rs.CursorLocation = adUseClient
                          m_rs.Open m_str, cnnConnection, adOpenStatic, adLockReadOnly
                          If Not m_rs.EOF And Not m_rs.BOF Then
                             If "" & m_rs.Fields(0) <> "0" Then  'Added by Lydia 2017/12/21
                                 If Val(Mid(m_rs.Fields(0), 10)) > 0 Then str105DATE = Val(Mid(m_rs.Fields(0), 10))
                             End If
                          End If
                       End If
ReDo:
                       'm_105tmpDate2 = DBDATE(Format(DateSerial(Val(DBYEAR(strTM21)) + Val(m_rs2.Fields(1).Value) + Val(m_rs2.Fields(0).Value * ii), Val(DBMONTH(strTM21)), Val(DBDAY(strTM21)))))
                       m_105tmpDate2 = DBDATE(Format(DateSerial(Val(DBYEAR(str105DATE)) + Val(m_rs2.Fields(1).Value) + Val(m_rs2.Fields(0).Value * ii), Val(DBMONTH(str105DATE)), Val(DBDAY(str105DATE)))))
                       '2011/9/28 end
                       If Val(m_105tmpDate2) < Val(strSrvDate(1)) Then
                          ii = ii + 1
                          GoTo ReDo
                       Else
                          If Val(m_105tmpDate2) >= Val(DBDATE(strTM22)) Then
                             'modify by sonia 2021/9/27 葉易雲於2021/8/27提出菲律賓2017/8/1新法延展核准後一年方再提出「延展使用宣誓」，故菲律賓不檢查專用期止日
                             'm_105tmpDate2 = ""
                             'm_105tmpDate1 = ""
                             If strTM10 <> "030" Then
                                m_105tmpDate2 = ""
                                m_105tmpDate1 = ""
                             Else
                                m_105tmpDate1 = DBDATE(DateAdd("m", -2, ChangeWStringToWDateString(DBDATE(m_105tmpDate2))))
                             End If
                             'end 2021/9/27
                          Else
                             m_105tmpDate1 = DBDATE(DateAdd("m", -2, ChangeWStringToWDateString(DBDATE(m_105tmpDate2))))
                          End If
                       End If
                    'add by sonia 2023/9/15 延展後之使用宣誓CFT-023278
                    ElseIf Not IsNull(m_rs2.Fields("na78")) Then
                       str105DATE = Val(strTM20)
ReDo105:
                       m_105tmpDate2 = DBDATE(Format(DateSerial(Val(DBYEAR(str105DATE)) + Val(m_rs2.Fields("na13").Value) + Val(m_rs2.Fields("na78").Value) + Val(m_rs2.Fields("na14").Value) * ii, Val(DBMONTH(str105DATE)), Val(DBDAY(str105DATE)))))
                       '110海地案為五年加三個月
                       If strTM10 = "110" Then
                          m_105tmpDate2 = CompDate(1, 3, m_105tmpDate2)
                       End If
                       If Val(m_105tmpDate2) < Val(strSrvDate(1)) Then
                          ii = ii + 1
                          GoTo ReDo105
                       Else
                          If Val(m_105tmpDate2) >= Val(DBDATE(strTM22)) Then
                             If strTM10 <> "030" Then
                                m_105tmpDate2 = ""
                                m_105tmpDate1 = ""
                             Else
                                m_105tmpDate1 = DBDATE(DateAdd("m", -2, ChangeWStringToWDateString(DBDATE(m_105tmpDate2))))
                             End If
                          Else
                             m_105tmpDate1 = DBDATE(DateAdd("m", -2, ChangeWStringToWDateString(DBDATE(m_105tmpDate2))))
                          End If
                       End If
                    'end 2023/9/15
                    End If     '2009/10/27 add by sonia
                 End If
             End If
             If m_105tmpDate1 <> "" And m_105tmpDate2 <> "" Then
                 m_105tmpDate1 = PUB_GetWorkDay1(m_105tmpDate1, True) 'Added by Lydia 2020/07/07 本所期限檢查：若本所期限非工作天則直接調整至最近的工作天
                 '檢查有無收文 105
                 Set m_rs = New ADODB.Recordset
                 'modify by sonia 2021/9/27 配合葉易雲於2021/8/27提出菲律賓2017/8/1新法延展核准後一年方再提出「延展使用宣誓」，發證後之第一次使用宣誓後要掛延展後一年的使用宣誓期限，故發文日不以系統日-3年判斷，改為以法定期限-3年，測試案件CFT-015425
                 'm_str = "select * from caseprogress where cp01='" & strTM01 & "' and cp02='" & strTM02 & "' and cp03='" & strTM03 & "' and cp04='" & strTM04 & "' and cp10='105' and cp57 is null  and (cp27 is null or cp27>=" & DBDATE(DateAdd("yyyy", -3, ChangeWStringToWDateString(strSrvDate(1)))) & ")  "
                 m_str = "select * from caseprogress where cp01='" & strTM01 & "' and cp02='" & strTM02 & "' and cp03='" & strTM03 & "' and cp04='" & strTM04 & "' and cp10='105' and cp57 is null  and (cp27 is null or cp27>=" & DBDATE(DateAdd("yyyy", -3, ChangeWStringToWDateString(m_105tmpDate2))) & ")  "
                 If m_rs.State = 1 Then m_rs.Close
                 m_rs.CursorLocation = adUseClient
                 m_rs.Open m_str, cnnConnection, adOpenStatic, adLockReadOnly
                 If Not m_rs.EOF And Not m_rs.BOF Then
                     '已經有，檢查是否無期限或期限不同
                     m_105CP06 = m_105tmpDate1
                     m_105CP07 = m_105tmpDate2
                     If Val(CheckStr(m_rs.Fields("cp07"))) <> Val(m_105CP07) Or CheckStr(m_rs.Fields("cp07")) = "" Then
                         If CheckStr(m_rs.Fields("cp27")) <> "" Then
                             '2008/2/14 cancel by sonia FCT-25023第二期,不必提醒
                             'MsgBox "使用宣誓已發文，請自行處裡期限！", vbInformation, "注意！"
                         Else
                             If MsgBox("使用宣誓已收文，" & IIf(CheckStr(m_rs.Fields("cp07")) = "", "但未掛期限，系統計算的期限為：" & ChangeWStringToTDateString(m_105CP07), "目前掛的期限為：" & ChangeWStringToTDateString(CheckStr(m_rs.Fields("cp07"))) & "，系統計算的期限為：" & ChangeWStringToTDateString(m_105CP07)) & "，是否更新？", vbInformation + vbYesNo, "注意！！") = vbYes Then
                                 IsCreate105 = True
                                 If Val(CheckStr(m_rs.Fields("cp07"))) <> Val(m_105CP07) Then
                                     m_105Key = CheckStr(m_rs.Fields("cp09"))
                                 End If
                             Else
                                 'modify by sonia 2016/12/13 CFT-018723 第三年使用宣誓接進來且已發證故要管制下次
                                 'm_105CP07 = ""
                                 'm_105CP06 = ""
                                 If MsgBox("您選擇不更新已收文使用宣誓的期限，是否要更新(或新增) " & ChangeWStringToTDateString(m_105CP07) & " 的下一程序期限？", vbInformation + vbYesNo, "注意！！") = vbYes Then
                                   GoTo Nextstep
                                 Else
                                   m_105CP07 = ""
                                   m_105CP06 = ""
                                 End If
                                 'end 2016/12/13
                             End If
                         End If
                     Else
                         '期限相同，不用變更
                         m_105CP07 = ""
                         m_105CP06 = ""
                     End If
                 Else
                 
Nextstep:   'add 2016/12/13
                     m_105NP08 = m_105tmpDate1
                     m_105NP09 = m_105tmpDate2
                     '檢查有無掛期限
                     Set m_rs = New ADODB.Recordset
                     m_str = "select * from nextprogress where np02='" & strTM01 & "' and np03='" & strTM02 & "' and np04='" & strTM03 & "' and np05='" & strTM04 & "' and np07=105 and np06 is null "
                     If m_rs.State = 1 Then m_rs.Close
                     m_rs.CursorLocation = adUseClient
                     m_rs.Open m_str, cnnConnection, adOpenStatic, adLockReadOnly
                     If Not m_rs.EOF And Not m_rs.BOF Then
                         If Val(CheckStr(m_rs.Fields("np09"))) <> Val(m_105NP09) Or CheckStr(m_rs.Fields("np09")) = "" Then
                             If MsgBox("使用宣誓" & IIf(CheckStr(m_rs.Fields("np09")) = "", "未掛期限，系統計算的期限為：" & ChangeWStringToTDateString(m_105NP09), "目前掛的期限為：" & ChangeWStringToTDateString(CheckStr(m_rs.Fields("np09"))) & "，系統計算的期限為：" & ChangeWStringToTDateString(m_105NP09)) & "，是否更新？", vbInformation + vbYesNo, "注意！！") = vbYes Then
                                 IsCreate105 = True
                                 If Val(CheckStr(m_rs.Fields("np09"))) <> Val(m_105NP09) Then
                                     m_105Key = CheckStr(m_rs.Fields("np01"))
                                 End If
                             Else
                                 m_105NP09 = ""
                                 m_105NP08 = ""
                                 '2012/3/29 ADD BY SNOIA
                                 If CheckStr(m_rs.Fields("np09")) = "" Then
                                    textTM58 = textTM58 & ";不管制使用宣誓"
                                    If strUpdTM58_105 = "" Then strUpdTM58_105 = ";不管制使用宣誓" 'Add By Sindy 2013/5/14
                                 End If
                                 '2012/3/29 END
                             End If
                         Else
                             '期限相同，不用變更
                             m_105NP09 = ""
                             m_105NP08 = ""
                         End If
                     Else
                             If Val(m_105NP09) < Val(strSrvDate(1)) Then    '已過期
                                 If MsgBox("使用宣誓已過期，是否仍要管制？", vbInformation + vbYesNo, "注意！！") = vbYes Then
                                     IsCreate105 = True
                                 Else
                                     m_105NP09 = ""
                                     m_105NP08 = ""
                                     textTM58 = textTM58 & ";不管制使用宣誓"   '2012/3/29 ADD BY SONIA
                                     If strUpdTM58_105 = "" Then strUpdTM58_105 = ";不管制使用宣誓" 'Add By Sindy 2013/5/14
                                 End If
                             Else
                                 If MsgBox("是否要管制使用宣誓？", vbInformation + vbYesNo, "注意！！") = vbYes Then
                                     IsCreate105 = True
                                 Else
                                     m_105NP09 = ""
                                     m_105NP08 = ""
                                     textTM58 = textTM58 & ";不管制使用宣誓"   '2012/3/29 ADD BY SONIA
                                     If strUpdTM58_105 = "" Then strUpdTM58_105 = ";不管制使用宣誓" 'Add By Sindy 2013/5/14
                                 End If
                             End If
                     End If
                 End If
             End If
         End If
      End If   'add by sonia 2020/12/29 菲律賓及波多黎各的申請日+3年未過期也要保留
      'Add By Sindy 2012/9/18 CFT申請國家為"030菲律賓"時,若無專用期限但有申請日時,加掛105使用宣誓期限
      'Modify By Sindy 2013/5/14 +TF
      'ElseIf strTM01 = "CFT" And strTM10 = "030" And strTM28 = "申請" And strTM21 = "" And strTM22 = "" And Val(DBDATE(strTM11)) >= 0 Then
      'Modify By Sindy 2013/7/2 若為閉卷案不掛期限
      'modify by sonia 2020/12/29 +波多黎各112, 菲律賓及波多黎各的申請日+3年未過期也要保留,不管有無專用期間
      'elseIf (strTM01 = "CFT" Or strTM01 = "TF") And (strTM10 = "030" Or strTM10 = "112") And strTM28 = "申請" And strTM21 = "" And strTM22 = "" And Val(DBDATE(strTM11)) >= 0 And _
         strTM29 <> "Y" Then
      If (strTM01 = "CFT" Or strTM01 = "TF") And (strTM10 = "030" Or strTM10 = "112") And strTM28 = "申請" And Val(DBDATE(strTM11)) >= 0 And strTM29 <> "Y" Then
      '2013/5/14 End
         'Add By Sindy 2013/1/3 過濾申請日為空白時
         m_105tmpDate2Before = "": m_105tmpDate1Before = ""
         If strTM11 <> "" Then
         '2013/1/3 End
            '法定期限
            m_105tmpDate2Before = DBDATE(DateAdd("yyyy", 3, ChangeWStringToWDateString(DBDATE(strTM11))))
            '本所期限
            m_105tmpDate1Before = DBDATE(DateAdd("m", -2, ChangeWStringToWDateString(DBDATE(m_105tmpDate2Before))))
         End If
         If m_105tmpDate1Before <> "" And m_105tmpDate2Before <> "" Then
             '檢查有無收文 105
             Set m_rs = New ADODB.Recordset
             m_str = "select * from caseprogress where cp01='" & strTM01 & "' and cp02='" & strTM02 & "' and cp03='" & strTM03 & "' and cp04='" & strTM04 & "' and cp10='105' and cp57 is null  and (cp27 is null or cp27>=" & DBDATE(DateAdd("yyyy", -3, ChangeWStringToWDateString(strSrvDate(1)))) & ")  "
             If m_rs.State = 1 Then m_rs.Close
             m_rs.CursorLocation = adUseClient
             m_rs.Open m_str, cnnConnection, adOpenStatic, adLockReadOnly
             If Not m_rs.EOF And Not m_rs.BOF Then
                 '已經有，檢查是否無期限或期限不同
                 m_105CP06Before = m_105tmpDate1Before
                 m_105CP07Before = m_105tmpDate2Before
                 If Val(CheckStr(m_rs.Fields("cp07"))) <> Val(m_105CP07Before) Or CheckStr(m_rs.Fields("cp07")) = "" Then
                     If CheckStr(m_rs.Fields("cp27")) <> "" Then
                         '2008/2/14 cancel by sonia
                         'MsgBox "使用宣誓已發文，請自行處裡期限！", vbInformation, "注意！"
                     Else
                         If MsgBox("使用宣誓已收文，" & IIf(CheckStr(m_rs.Fields("cp07")) = "", "但未掛期限，系統計算的期限為：" & ChangeWStringToTDateString(m_105CP07Before), "目前掛的期限為：" & ChangeWStringToTDateString(CheckStr(m_rs.Fields("cp07"))) & "，系統計算的期限為：" & ChangeWStringToTDateString(m_105CP07Before)) & "，是否更新？", vbInformation + vbYesNo, "注意！！") = vbYes Then
                             IsCreate105Before = True
                             If Val(CheckStr(m_rs.Fields("cp07"))) <> Val(m_105CP07Before) Then
                                 m_105KeyBefore = CheckStr(m_rs.Fields("cp09"))
                             End If
                         Else
                             m_105CP07Before = ""
                             m_105CP06Before = ""
                         End If
                     End If
                 Else
                     '期限相同，不用變更
                     m_105CP07Before = ""
                     m_105CP06Before = ""
                 End If
             Else
                 m_105NP08Before = m_105tmpDate1Before
                 m_105NP09Before = m_105tmpDate2Before
                 '檢查有無掛期限
                 Set m_rs = New ADODB.Recordset
                 'modify by sonia 2020/12/29 +np09條件
                 'm_str = "select * from nextprogress where np02='" & strTM01 & "' and np03='" & strTM02 & "' and np04='" & strTM03 & "' and np05='" & strTM04 & "' and np07=105 and np06 is null "
                 m_str = "select * from nextprogress where np02='" & strTM01 & "' and np03='" & strTM02 & "' and np04='" & strTM03 & "' and np05='" & strTM04 & "' and np07=105 and np06 is null and np09<=" & m_105tmpDate2Before + 10000
                 If m_rs.State = 1 Then m_rs.Close
                 m_rs.CursorLocation = adUseClient
                 m_rs.Open m_str, cnnConnection, adOpenStatic, adLockReadOnly
                 If Not m_rs.EOF And Not m_rs.BOF Then
                     If Val(CheckStr(m_rs.Fields("np09"))) <> Val(m_105NP09Before) Or CheckStr(m_rs.Fields("np09")) = "" Then
                         If MsgBox("註冊前使用宣誓" & IIf(CheckStr(m_rs.Fields("np09")) = "", "未掛期限，系統計算的期限為：" & ChangeWStringToTDateString(m_105NP09Before), "目前掛的期限為：" & ChangeWStringToTDateString(CheckStr(m_rs.Fields("np09"))) & "，系統計算的期限為：" & ChangeWStringToTDateString(m_105NP09Before)) & "，是否更新？", vbInformation + vbYesNo, "注意！！") = vbYes Then
                             IsCreate105Before = True
                             If Val(CheckStr(m_rs.Fields("np09"))) <> Val(m_105NP09Before) Then
                                 m_105KeyBefore = CheckStr(m_rs.Fields("np01"))
                             End If
                         Else
                             m_105NP09Before = ""
                             m_105NP08Before = ""
                             If CheckStr(m_rs.Fields("np09")) = "" Then
                                textTM58 = textTM58 & ";不管制註冊前使用宣誓"
                                If strUpdTM58_105 = "" Then strUpdTM58_105 = ";不管制註冊前使用宣誓" 'Add By Sindy 2013/5/14
                             End If
                         End If
                     Else
                         '期限相同，不用變更
                         m_105NP09Before = ""
                         m_105NP08Before = ""
                     End If
                 Else
                         If Val(m_105NP09Before) < Val(strSrvDate(1)) Then    '已過期
                             If MsgBox("註冊前使用宣誓已過期，是否仍要管制？", vbInformation + vbYesNo, "注意！！") = vbYes Then
                                 IsCreate105Before = True
                             Else
                                 m_105NP09Before = ""
                                 m_105NP08Before = ""
                                 textTM58 = textTM58 & ";不管制註冊前使用宣誓"
                                 If strUpdTM58_105 = "" Then strUpdTM58_105 = ";不管制註冊前使用宣誓" 'Add By Sindy 2013/5/14
                             End If
                         Else
                             If MsgBox("是否要管制註冊前使用宣誓？", vbInformation + vbYesNo, "注意！！") = vbYes Then
                                 IsCreate105Before = True
                             Else
                                 m_105NP09Before = ""
                                 m_105NP08Before = ""
                                 textTM58 = textTM58 & ";不管制註冊前使用宣誓"
                                 If strUpdTM58_105 = "" Then strUpdTM58_105 = ";不管制註冊前使用宣誓" 'Add By Sindy 2013/5/14
                             End If
                         End If
                 End If
             End If
         End If
      End If
   End If
End Sub

'Add By Cheng 2002/05/23
Private Function TxtValidate() As Boolean
Dim objTxt As Object
Dim ii As Integer
Dim Cancel As Boolean
'Add by Amy 2017/03/15
Dim bolData As Boolean, strApply As String, strMsg As String
Dim strMCTFNew(0) As String, strTmp(0) As String

   TxtValidate = False
   
   'Add By Sindy 2010/12/24
   If Me.textTM12.Enabled = True Then
      Cancel = False
      textTM12_Validate Cancel
      If Cancel = True Then
         textTM12.SetFocus
         Exit Function
      End If
   End If
   
   'Add By Cheng 2003/05/22
   '檢查本所案號
   If Me.textTM01.Enabled = True Then
       If Me.textTM01.Text = "" Then
           MsgBox "請輸入本所案號!!!", vbExclamation + vbOKOnly
           Me.tabCtrl.Tab = 0
           Me.textTM01.SetFocus
           textTM01_GotFocus
           Exit Function
       End If
   End If
   If Me.textTM02_1.Enabled = True Then
       If Me.textTM02_1.Text = "" Then
           MsgBox "請輸入本所案號!!!", vbExclamation + vbOKOnly
           Me.tabCtrl.Tab = 0
           Me.textTM02_1.SetFocus
           textTM02_1_GotFocus
           Exit Function
       End If
   End If
   If Me.textTM01.Enabled = True Then
      Cancel = False
      textTM01_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If Me.textTM05.Enabled = True Then
      Cancel = False
      textTM05_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   'Add By Sindy 2015/6/30
   If Me.textTM131.Enabled = True Then
      Cancel = False
      textTM131_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   '2015/6/30 END
   'Add By Sindy 2024/6/14
   If Trim(cboTM72.Text) <> "" Then '特殊商標
      If textTM137.Text = "" And textTM138.Text = "" And textTM139.Text = "" Then
         If MsgBox("特殊商標，商標描述不可空白!!!" & vbCrLf & "要繼續嗎?", vbInformation + vbYesNo + vbDefaultButton2, "注意！！") = vbNo Then
            Me.tabCtrl.Tab = 7
            Exit Function
         End If
      End If
   End If
   If Me.textTM137.Visible = True And Me.textTM137.Enabled = True Then
      Cancel = False
      textTM137_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   ElseIf Me.textTM137.Visible = False And Trim(cboTM72.Text) = "" Then '非特殊商標
      textTM137.Text = ""
   End If
   If Me.textTM138.Visible = True And Me.textTM138.Enabled = True Then
      Cancel = False
      textTM138_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   ElseIf Me.textTM138.Visible = False And Trim(cboTM72.Text) = "" Then '非特殊商標
      textTM138.Text = ""
   End If
   If Me.textTM139.Visible = True And Me.textTM139.Enabled = True Then
      Cancel = False
      textTM139_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   ElseIf Me.textTM139.Visible = False And Trim(cboTM72.Text) = "" Then '非特殊商標
      textTM139.Text = ""
   End If
   '2024/6/14 END
   If Me.textTM08_1.Enabled = True Then
      Cancel = False
      textTM08_1_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   'Added by Lydia 2023/11/16
   If Me.cboTM08.Enabled = True Then
      Cancel = False
      cboTM08_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   'end 2023/11/16
   
   If Me.textTM09.Enabled = True Then
      Cancel = False
      textTM09_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textTM10_1.Enabled = True Then
      Cancel = False
      textTM10_1_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textTM11.Enabled = True Then
      Cancel = False
      textTM11_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textTM13.Enabled = True Then
      Cancel = False
      textTM13_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textTM14.Enabled = True Then
      Cancel = False
      textTM14_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
      'Added by Lydia 2017/03/06 他所轉來案件在補輸商標案件基本資料時，由系統自動產生B類收文和催審期限。因為催審期限=申請日+審查天數，所以增加檢查無申請日不可存檔。
      'Modified by Lydia 2017/03/14 textTM14=>textTM11
      If textTM11.Text = "" And textTM01 <> "TF" And textTM16 = "" And textTM28 = "申請" Then
        '檢查進度檔若無申請101或分割308進度 , 則自動產生B類「申請」進度
        strExc(0) = "select cp05,cp09,cp10 from caseprogress where cp01='" & textTM01 & "' and cp02='" & textTM02_1 & textTM02_2 & "' and cp03='" & textTM03 & "' and cp04='" & textTM04 & "' and cp10 in ('101','308') order by cp05 "
        intI = 1
        Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
        If intI = 0 Then
           MsgBox "他所轉來案件在補輸商標案件基本資料時，請輸入申請日！", vbCritical
           Exit Function
        End If
      End If
      'end 2017/03/06
   End If
   
   If Me.textTM15.Enabled = True Then
      Cancel = False
      textTM15_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textTM16.Enabled = True Then
      Cancel = False
      textTM16_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textTM17.Enabled = True Then
      Cancel = False
      textTM17_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textTM18.Enabled = True Then
      Cancel = False
      textTM18_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textTM19.Enabled = True Then
      Cancel = False
      textTM19_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textTM20.Enabled = True Then
      Cancel = False
      textTM20_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textTM21.Enabled = True Then
      Cancel = False
      textTM21_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textTM22.Enabled = True Then
      Cancel = False
      textTM22_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textTM23_1.Enabled = True Then
      Cancel = False
      textTM23_1_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textTM24.Enabled = True Then
      Cancel = False
      textTM24_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textTM26.Enabled = True Then
      Cancel = False
      textTM26_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textTM29.Enabled = True Then
      Cancel = False
      textTM29_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textTM30.Enabled = True Then
      Cancel = False
      textTM30_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textTM31_1.Enabled = True Then
      Cancel = False
      textTM31_1_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textTM32.Enabled = True Then
      Cancel = False
      textTM32_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textTM33_1.Enabled = True Then
      Cancel = False
      textTM33_1_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   'Add By Sindy 2012/7/18
   If Me.textTM34.Enabled = True Then
      Cancel = False
      textTM34_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textTM38.Enabled = True Then
      Cancel = False
      textTM38_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textTM40.Enabled = True Then
      Cancel = False
      textTM40_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textTM41.Enabled = True Then
      Cancel = False
      textTM41_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textTM43.Enabled = True Then
      Cancel = False
      textTM43_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textTM44_1.Enabled = True Then
      Cancel = False
      textTM44_1_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   'Add by Amy 2017/03/15 T字頭申請國家為台灣,若FC代理人之管控智權人員為MCTF時,修改成不同組別不可存檔
   'Modify by Amy 拿掉 And textTM10_1 = "000" 申請國家為台灣之判斷 (X39289040 為MCTF但要收申請國家為大陸的案件)
   If m_EditMode = 2 And Left(textTM01, 1) = "T" And Trim(Len(textTM44_1)) > 0 And m_FieldList(43).fiOldData <> ChangeCustomerL(textTM44_1) Then
        strMsg = ""
        bolData = GetCusORFagentData(ChangeCustomerL(textTM44_1), "FA120", strMCTFNew())
        If Left(strMCTFNew(0), 4) = "MCTF" Then
            For ii = 0 To 4
                Select Case ii
                    Case 0
                        strApply = textTM23_1
                    Case 1
                        strApply = textTM78_1
                    Case 2
                        strApply = textTM79_1
                    Case 3
                        strApply = textTM80_1
                    Case 4
                        strApply = textTM81_1
                End Select
                If Len(Trim(strApply)) = 0 Then Exit For
                bolData = GetCusORFagentData(ChangeCustomerL(strApply), "CU13", strTmp())
                If strMCTFNew(0) <> strTmp(0) And Left(strTmp(0), 4) = "MCTF" Then
                    strMsg = strMsg & "申請人" & ii + 1 & "：" & strApply & " (" & strTmp(0) & ")" & "及"
                End If
            Next ii
            If strMsg <> MsgText(601) Then
                MsgBox Left(strMsg, Len(strMsg) - 1) & vbCrLf & "與代理人" & textTM44_1 & _
                            "商標管控智權人員(" & strMCTFNew(0) & ")不同，不可存檔！"
                Exit Function
            End If
        End If
   End If
   'end 2017/03/15
   
   If Me.textTM46.Enabled = True Then
      Cancel = False
      textTM46_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   'Add By Sindy 2013/12/13
   If Me.textTM130.Enabled = True Then
      Cancel = False
      textTM130_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textTM47.Enabled = True Then
      Cancel = False
      textTM47_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textTM49.Enabled = True Then
      Cancel = False
      textTM49_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textTM50.Enabled = True Then
      Cancel = False
      textTM50_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textTM52.Enabled = True Then
      Cancel = False
      textTM52_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textTM53.Enabled = True Then
      Cancel = False
      textTM53_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textTM54_1.Enabled = True Then
      Cancel = False
      textTM54_1_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textTM56_1.Enabled = True Then
      Cancel = False
      textTM56_1_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textTM57.Enabled = True Then
      Cancel = False
      textTM57_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textTM58.Enabled = True Then
      Cancel = False
      textTM58_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textTM66_1.Enabled = True Then
      Cancel = False
      textTM66_1_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textTM67.Enabled = True Then
      Cancel = False
      textTM67_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textTM68.Enabled = True Then
      Cancel = False
      textTM68_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If Me.textTM69_1.Enabled = True Then
      Cancel = False
      textTM69_1_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If Me.textTM70_1.Enabled = True Then
      Cancel = False
      textTM70_1_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If Me.textTM72_1.Enabled = True Then
      Cancel = False
      textTM72_1_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   'Added by Lydia 2023/11/16
   If Me.cboTM72.Enabled = True Then
      Cancel = False
      cboTM72_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   'end 2023/11/16
   
   'Add By Sindy 2012/7/18
   If Me.textTM76.Enabled = True Then
      Cancel = False
      textTM76_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   'Add By Sindy 2009/09/09
   If Me.textTM77.Enabled = True Then
      Cancel = False
      textTM77_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   'Add By Sindy 2012/7/18
   If Me.textTM78_1.Enabled = True Then
      Cancel = False
      textTM78_1_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If Me.textTM79_1.Enabled = True Then
      Cancel = False
      textTM79_1_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If Me.textTM80_1.Enabled = True Then
      Cancel = False
      textTM80_1_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If Me.textTM81_1.Enabled = True Then
      Cancel = False
      textTM81_1_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If Me.textTM82.Enabled = True Then
      Cancel = False
      textTM82_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If Me.textTM83.Enabled = True Then
      Cancel = False
      textTM83_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If Me.textTM84.Enabled = True Then
      Cancel = False
      textTM84_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If Me.textTM85.Enabled = True Then
      Cancel = False
      textTM85_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If Me.textTM90.Enabled = True Then
      Cancel = False
      textTM90_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If Me.textTM91.Enabled = True Then
      Cancel = False
      textTM91_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If Me.textTM92.Enabled = True Then
      Cancel = False
      textTM92_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If Me.textTM93.Enabled = True Then
      Cancel = False
      textTM93_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   '2012/7/18 End
   
   'Add By Sindy 2012/7/18
   If Me.textTM121.Enabled = True Then
      Cancel = False
      textTM121_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   'add by Toni 2008/10/21
   If Me.textTM122.Enabled = True Then
      Cancel = False
      textTM122_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   'end 2008/10/21
   
   If Me.textTM126.Enabled = True Then
      Cancel = False
      textTM126_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   'Add By Sindy 2013/8/15
   If Me.textTM129.Enabled = True Then
      Cancel = False
      textTM129_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   'Added by Morgan 2022/12/1
   If Me.textTM136.Enabled = True Then
      Cancel = False
      textTM136_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   'Add By Sindy 2016/11/23
   If Me.Combo4.Enabled = True Then
      Cancel = False
      Combo4_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   '2016/11/23 End
   
   'add by sonia 2023/3/28 專用期未過期但已閉卷,提醒不會管制期限
   If Val(DBDATE(textTM22)) >= Val(strSrvDate(1)) And textTM29 = "Y" Then
      MsgBox "專用期未過期但已閉卷，系統不會管制任何期限！"
   End If
   'end 2023/3/28
      
    'Added by Lydia 2021/11/29 檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
   If PUB_ChkUniText(Me, , True, "TextBox") = False Then
       Exit Function
   End If
   
   'Added by Lydia 2024/06/14 對申請人1~5的重複輸入檢查
   If Pub_ChkAppList(strExc(0), textTM23_1 & "," & textTM78_1 & "," & textTM79_1 & "," & textTM80_1 & "," & textTM81_1) = False Then
      Select Case Val(strExc(0))
         Case 1
            Me.tabCtrl.Tab = 2
            textTM23_1.SetFocus
            textTM23_1_GotFocus
         Case 2
            Me.tabCtrl.Tab = 2
            textTM78_1.SetFocus
            textTM78_1_GotFocus
         Case 3
            Me.tabCtrl.Tab = 2
            textTM79_1.SetFocus
            textTM79_1_GotFocus
         Case 4
            Me.tabCtrl.Tab = 3
            textTM80_1.SetFocus
            textTM80_1_GotFocus
         Case 5
            Me.tabCtrl.Tab = 3
            textTM81_1.SetFocus
            textTM81_1_GotFocus
      End Select
      Exit Function
   End If
   'end 2024/06/14
   
   'Added by Lydia 2024/06/13 檢查更新代理人／申請人狀態排除「不得代理」
   Dim iPage As Integer
   For ii = 1 To 6
      strExc(1) = ""
      Select Case ii
         Case 1 '申請人1
            strExc(1) = ChangeCustomerL(textTM23_1)
            strExc(2) = ChangeCustomerL(m_TM23)
            iPage = 2
         Case 2 '申請人2
            strExc(1) = ChangeCustomerL(textTM78_1)
            strExc(2) = ChangeCustomerL(m_TM78)
            iPage = 2
         Case 3 '申請人3
            strExc(1) = ChangeCustomerL(textTM79_1)
            strExc(2) = ChangeCustomerL(m_TM79)
            iPage = 2
         Case 4 '申請人4
            strExc(1) = ChangeCustomerL(textTM80_1)
            strExc(2) = ChangeCustomerL(m_TM80)
            iPage = 3
         Case 5 '申請人5
            strExc(1) = ChangeCustomerL(textTM81_1)
            strExc(2) = ChangeCustomerL(m_TM81)
            iPage = 3
         Case 6 '代理人
            strExc(1) = ChangeCustomerL(textTM44_1)
            strExc(2) = ChangeCustomerL(m_TM44)
            iPage = 4
      End Select
      If strExc(1) <> "" And strExc(1) <> strExc(2) Then
         If Left(strExc(1), 1) = "X" Then
            If GetCustomerAndState(strExc(1), strExc(3), , , , textTM01, strExc(8), False, Me.Name, textTM02_1 & IIf(textTM01 = "TF", textTM02_2, ""), textTM03, textTM04) = False Then
               Me.tabCtrl.Tab = iPage
               If ii = 1 Then
                  textTM23_1.SetFocus
                  textTM23_1_GotFocus
                  Exit Function
               ElseIf ii = 2 Then
                  textTM78_1.SetFocus
                  textTM78_1_GotFocus
                  Exit Function
               ElseIf ii = 3 Then
                  textTM79_1.SetFocus
                  textTM79_1_GotFocus
                  Exit Function
               ElseIf ii = 4 Then
                  textTM80_1.SetFocus
                  textTM80_1_GotFocus
                  Exit Function
               ElseIf ii = 5 Then
                  textTM81_1.SetFocus
                  textTM81_1_GotFocus
                  Exit Function
               End If
            End If
         Else
            If GetAgentAndState(strExc(1), strExc(3), , , , textTM01, strExc(8), False) = False Then
               Me.tabCtrl.Tab = iPage
               textTM44_1.SetFocus
               textTM44_1_GotFocus
               Exit Function
            End If
         End If
      End If
   Next
   'end 2024/06/13
   
   TxtValidate = True
End Function

'Add By Cheng 2003/08/18
'取得申請人名稱
Private Function GetTM23Name(strCode As String) As String
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset

   GetTM23Name = ""
   If strCode = "" Then Exit Function
   strCode = Left(strCode & "000000000", 9)
   If UCase(Left(strCode, 1)) = "X" Then
       '92.9.15 modify by sonia
       'strSQLA = "Select Nvl(CU04,Decode(CU05,Null,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)) From Customer Where CU01='" & Mid(strCode, 1, 8) & "' And CU02='0' "
       StrSQLa = "Select Nvl(CU04,Decode(CU05,Null,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)) From Customer Where CU01='" & Mid(strCode, 1, 8) & "' And CU02='" & Mid(strCode, 9, 1) & "' "
       '92.9.15 end
       rsA.CursorLocation = adUseClient
       rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
       If rsA.RecordCount > 0 Then
           GetTM23Name = "" & rsA.Fields(0).Value
       End If
       If rsA.State <> adStateClosed Then rsA.Close
       Set rsA = Nothing
   Else
       StrSQLa = "Select FA03 From Fagent Where FA01='" & Mid(strCode, 1, 8) & "' And FA02='" & Mid(strCode, 9, 1) & "' "
       rsA.CursorLocation = adUseClient
       rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
       If rsA.RecordCount > 0 Then
           '92.9.15 modify by sonia
           'strSQLA = "Select Nvl(CU04,Decode(CU05,Null,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)) From Customer Where CU01='" & Mid("" & rsA.Fields(0).Value, 1, 8) & "' And CU02='0' "
           StrSQLa = "Select Nvl(CU04,Decode(CU05,Null,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)) From Customer Where CU01='" & Mid("" & rsA.Fields(0).Value, 1, 8) & "' And CU02='" & Mid("" & rsA.Fields(0).Value, 9, 1) & "' "
           '92.9.15 end
           If rsA.State <> adStateClosed Then rsA.Close
           Set rsA = Nothing
           rsA.CursorLocation = adUseClient
           rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
           If rsA.RecordCount > 0 Then
               GetTM23Name = "" & rsA.Fields(0).Value
           End If
           If rsA.State <> adStateClosed Then rsA.Close
           Set rsA = Nothing
       End If
   End If
End Function

'add by nickc 2008/02/13 抓最後收文的A類收文號
Function GetLastA(strTM01 As String, strTM02_1 As String, strTM02_2 As String, strTM03 As String, strTM04 As String) As String
Dim m_rs As New ADODB.Recordset
Dim m_str As String
   
   GetLastA = ""
   'Modify By Sindy 2013/5/15 若為馬德里案則抓母案的最後收文號
   If strTM01 = "TF" Then
      'Modified by Lydia 2017/12/21
      'm_str = "select max(cp09) from caseprogress where cp01='" & strTM01 & "' and cp02='" & strTM02_1 & "0' " & _
              "and cp03='0' and cp04='00' and cp09<'B' " & _
              "and cp05 in(select max(cp05) from caseprogress where cp01='" & strTM01 & "' and cp02='" & strTM02_1 & "0' " & _
              "and cp03='0' and cp04='00' and cp09<'B')"
      m_str = "select nvl(max(cp09),'N') from caseprogress where cp01='" & strTM01 & "' and cp02='" & strTM02_1 & "0' " & _
              "and cp03='0' and cp04='00' and cp09<'B' " & _
              "and cp05 in(select nvl(max(cp05),0) from caseprogress where cp01='" & strTM01 & "' and cp02='" & strTM02_1 & "0' " & _
              "and cp03='0' and cp04='00' and cp09<'B')"
      If m_rs.State = 1 Then m_rs.Close
      m_rs.CursorLocation = adUseClient
      m_rs.Open m_str, cnnConnection, adOpenStatic, adLockReadOnly
      If Not m_rs.EOF And Not m_rs.BOF Then
          GetLastA = CheckStr(m_rs.Fields(0))
      'modify by sonia 2017/12/19
      'Else
      End If
      'Modified by Lydia 2017/12/21 + GetLastA = "N"
      If GetLastA = "" Or GetLastA = "N" Then
      'end 2017/12/19
         '抓最後收文的B類收文號
         'Modified by Lydia 2017/12/21
         'm_str = "select max(cp09) from caseprogress where cp01='" & strTM01 & "' and cp02='" & strTM02_1 & "0' " & _
                 "and cp03='0' and cp04='00' and cp09<'C' " & _
                 "and cp05 in(select max(cp05) from caseprogress where cp01='" & strTM01 & "' and cp02='" & strTM02_1 & "0' " & _
                 "and cp03='0' and cp04='00' and cp09<'C')"
         m_str = "select nvl(max(cp09),'N') from caseprogress where cp01='" & strTM01 & "' and cp02='" & strTM02_1 & "0' " & _
                 "and cp03='0' and cp04='00' and cp09<'C' " & _
                 "and cp05 in(select nvl(max(cp05),0) from caseprogress where cp01='" & strTM01 & "' and cp02='" & strTM02_1 & "0' " & _
                 "and cp03='0' and cp04='00' and cp09<'C')"
         If m_rs.State = 1 Then m_rs.Close
         m_rs.CursorLocation = adUseClient
         m_rs.Open m_str, cnnConnection, adOpenStatic, adLockReadOnly
         If Not m_rs.EOF And Not m_rs.BOF Then
             GetLastA = CheckStr(m_rs.Fields(0))
         'Modified by Lydia 2017/12/21
         'Else
         End If
         If GetLastA = "" Or GetLastA = "N" Then
         'end 2017/12/21
            MsgBox "系統抓不到最後收文的A、B類收文號，無法新增期限！"
         End If
      End If
   Else
      'Modified by Lydia 2017/12/21 在Client-Win7會出現"資料提供者或其他服務傳回E_Fail狀態" (ex. CFT-011222-1)
      'm_str = "select count(*) cnt,max(cp09) from caseprogress where cp01='" & strTM01 & "' and cp02='" & strTM02_1 & strTM02_2 & "' " & _
              "and cp03='" & strTM03 & "' and cp04='" & strTM04 & "' and cp09<'B' " & _
              "and cp05 in(select max(cp05) from caseprogress where cp01='" & strTM01 & "' and cp02='" & strTM02_1 & strTM02_2 & "' " & _
              "and cp03='" & strTM03 & "' and cp04='" & strTM04 & "' and cp09<'B')"
      m_str = "select nvl(max(cp09),'N') from caseprogress where cp01='" & strTM01 & "' and cp02='" & strTM02_1 & strTM02_2 & "' " & _
              "and cp03='" & strTM03 & "' and cp04='" & strTM04 & "' and cp09<'B' " & _
              "and cp05 in(select nvl(max(cp05),0) from caseprogress where cp01='" & strTM01 & "' and cp02='" & strTM02_1 & strTM02_2 & "' " & _
              "and cp03='" & strTM03 & "' and cp04='" & strTM04 & "' and cp09<'B')"
      If m_rs.State = 1 Then m_rs.Close
      m_rs.CursorLocation = adUseClient
      m_rs.Open m_str, cnnConnection, adOpenStatic, adLockReadOnly
      If Not m_rs.EOF And Not m_rs.BOF Then
          GetLastA = CheckStr(m_rs.Fields(0))
      'modify by sonia 2017/12/19 CFT-011222-1會抓不到
      'Else
      End If
      'Modified by Lydia 2017/12/21 + GetLastA = "N"
      If GetLastA = "" Or GetLastA = "N" Then
      'end 2017/12/19
         'Add by Sindy 2012/10/1 抓最後收文的B類收文號
         'Modified by Lydia 2017/12/21
         'm_str = "select max(cp09) from caseprogress where cp01='" & strTM01 & "' and cp02='" & strTM02_1 & strTM02_2 & "' " & _
                 "and cp03='" & strTM03 & "' and cp04='" & strTM04 & "' and cp09<'C' " & _
                 "and cp05 in(select max(cp05) from caseprogress where cp01='" & strTM01 & "' and cp02='" & strTM02_1 & strTM02_2 & "' " & _
                 "and cp03='" & strTM03 & "' and cp04='" & strTM04 & "' and cp09<'C')"
         m_str = "select nvl(max(cp09),'N') from caseprogress where cp01='" & strTM01 & "' and cp02='" & strTM02_1 & strTM02_2 & "' " & _
                 "and cp03='" & strTM03 & "' and cp04='" & strTM04 & "' and cp09<'C' " & _
                 "and cp05 in(select nvl(max(cp05),0) from caseprogress where cp01='" & strTM01 & "' and cp02='" & strTM02_1 & strTM02_2 & "' " & _
                 "and cp03='" & strTM03 & "' and cp04='" & strTM04 & "' and cp09<'C')"
         If m_rs.State = 1 Then m_rs.Close
         m_rs.CursorLocation = adUseClient
         m_rs.Open m_str, cnnConnection, adOpenStatic, adLockReadOnly
         If Not m_rs.EOF And Not m_rs.BOF Then
             GetLastA = CheckStr(m_rs.Fields(0))
         'Modified by Lydia 2017/12/21
         'Else
         End If
         If GetLastA = "" Or GetLastA = "N" Then
         'end 2017/12/21
            MsgBox "系統抓不到最後收文的A、B類收文號，無法新增期限！"
         End If
         '2012/10/1 End
      End If
   End If
End Function

'Add By Sindy 2013/12/13
Private Sub textTM130_GotFocus()
   InverseTextBox textTM130
End Sub

'Modified by Lydia 2021/11/29 改成Form 2.0; KeyAscii As Integer=>MSForms.ReturnInteger
Private Sub textTM130_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = UpperCase(KeyAscii)
End Sub
'特殊出名公司
Private Sub textTM130_Validate(Cancel As Boolean)
Dim strTit As String
Dim strMsg As String
Dim nResponse
   
   Cancel = False
   textTM130 = Trim(textTM130)
   If textTM130 <> "" And textTM130 <> "J" Then
      Cancel = True
      strTit = "檢核資料"
      strMsg = "請輸入J或空白"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textTM130_GotFocus
   End If
End Sub

'Added by Lydia 2017/06/14
Private Sub textTM39_Validate(Cancel As Boolean)
   Cancel = False
    If CheckLengthIsOK(textTM39, 35) = False Then
      Cancel = True
      textTM39_GotFocus
   End If
End Sub

Private Sub textTM42_Validate(Cancel As Boolean)
   Cancel = False
    If CheckLengthIsOK(textTM42, 35) = False Then
      Cancel = True
      textTM42_GotFocus
   End If
End Sub

'Added by Lydia 2023/11/16
Private Sub cboTM72_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub cboTM72_Validate(Cancel As Boolean)
Dim intX As Integer, intQ As Integer
   
   intX = -1
   If Trim(cboTM72.Text) <> "" And cboTM72.Tag <> cboTM72.Text Then
        For intQ = 0 To cboTM72.ListCount - 1
           If InStr(cboTM72.List(intQ), Trim(cboTM72.Text)) > 0 Then
              intX = intQ
              Exit For
           End If
        Next intQ
        If intX = -1 Then
           cboTM72.SetFocus
           cboTM72.Tag = cboTM72.Text
           Cancel = True
           Exit Sub
        Else
           cboTM72.ListIndex = intX
        End If
   End If
   textTM72_1 = Trim(Left(cboTM72, 1))
   cboTM72.Tag = cboTM72.Text
   Call SetTM72forCol(textTM71_1) 'Add By Sindy 2024/6/14
End Sub

Private Sub cboTM08_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub cboTM08_Validate(Cancel As Boolean)
Dim intX As Integer, intQ As Integer
   
   intX = -1
   If Trim(cboTM08.Text) <> "" And cboTM08.Tag <> cboTM08.Text Then
        For intQ = 0 To cboTM08.ListCount - 1
           If InStr(cboTM08.List(intQ), Trim(cboTM08.Text)) > 0 Then
              intX = intQ
              Exit For
           End If
        Next intQ
        If intX = -1 Then
            cboTM08.SetFocus
            cboTM08.Tag = cboTM08.Text
            Cancel = True
            Exit Sub
        Else
            cboTM08.ListIndex = intX
        End If
   End If
   textTM08_1 = Trim(Left(cboTM08.Text, 1))
   cboTM08.Tag = cboTM08.Text
End Sub
'end 2023/11/16

'Added by Lydia 2025/10/23
Private Sub cmdTFBaseNo_Click()
   
   If m_EditMode <> 0 And m_EditMode <> 4 Then
      MsgBox "基本檔" & IIf(m_EditMode = 1, "新增中", "修改中") & "不可執行！", vbInformation
      Exit Sub
   End If
   Call frm020509.SetParent(Me, textTM01 & textTM02_1 & textTM02_2 & textTM03 & textTM04, IIf(m_bUpdate = True, "U", "Q"))
   frm020509.Show vbModal
   
End Sub

