VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm050102_3 
   BorderStyle     =   1  '單線固定
   Caption         =   "發文（新申請案）"
   ClientHeight    =   6730
   ClientLeft      =   230
   ClientTop       =   1000
   ClientWidth     =   8530
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6730
   ScaleWidth      =   8530
   Begin TabDlg.SSTab SSTab1 
      Height          =   5385
      Left            =   90
      TabIndex        =   57
      Top             =   1260
      Width           =   8340
      _ExtentX        =   14711
      _ExtentY        =   9507
      _Version        =   393216
      Tabs            =   5
      TabsPerRow      =   7
      TabHeight       =   420
      TabCaption(0)   =   "基本資料"
      TabPicture(0)   =   "frm050102_3.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "chkChoose(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblCaseFees"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label12"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label9(0)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label9(2)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label9(3)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "lblPetition(4)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label3(0)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label4(0)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "lblPetition(2)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "lblPetition(3)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label6(2)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label7"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "lblPetition(0)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "lblPetition(1)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Label17"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Label11"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Label10"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Label16"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Label18(0)"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "lblNation"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "lblCaseProperty"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "Label37"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "Label23"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "Label3(2)"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "Label3(1)"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "Label4(2)"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "Label21"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "Label6(3)"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "Label4(1)"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "lblAgent"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "Label14(0)"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "Label15"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "Label19"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "lblAppNation(0)"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "Label13"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "Label20"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "Label6(4)"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "Label4(3)"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).Control(39)=   "Label3(18)"
      Tab(0).Control(39).Enabled=   0   'False
      Tab(0).Control(40)=   "lblAppNation(1)"
      Tab(0).Control(40).Enabled=   0   'False
      Tab(0).Control(41)=   "lblAppNation(2)"
      Tab(0).Control(41).Enabled=   0   'False
      Tab(0).Control(42)=   "lblAppNation(3)"
      Tab(0).Control(42).Enabled=   0   'False
      Tab(0).Control(43)=   "lblAppNation(4)"
      Tab(0).Control(43).Enabled=   0   'False
      Tab(0).Control(44)=   "Label4(4)"
      Tab(0).Control(44).Enabled=   0   'False
      Tab(0).Control(45)=   "Label3(20)"
      Tab(0).Control(45).Enabled=   0   'False
      Tab(0).Control(46)=   "lblCaseFee"
      Tab(0).Control(46).Enabled=   0   'False
      Tab(0).Control(47)=   "Label18(5)"
      Tab(0).Control(47).Enabled=   0   'False
      Tab(0).Control(48)=   "lblCP113(18)"
      Tab(0).Control(48).Enabled=   0   'False
      Tab(0).Control(49)=   "txtCaseField(8)"
      Tab(0).Control(49).Enabled=   0   'False
      Tab(0).Control(50)=   "txtCaseField(0)"
      Tab(0).Control(50).Enabled=   0   'False
      Tab(0).Control(51)=   "txtCaseField(1)"
      Tab(0).Control(51).Enabled=   0   'False
      Tab(0).Control(52)=   "txtCaseField(2)"
      Tab(0).Control(52).Enabled=   0   'False
      Tab(0).Control(53)=   "txtCaseField(7)"
      Tab(0).Control(53).Enabled=   0   'False
      Tab(0).Control(54)=   "txtCaseField(6)"
      Tab(0).Control(54).Enabled=   0   'False
      Tab(0).Control(55)=   "txtCaseField(4)"
      Tab(0).Control(55).Enabled=   0   'False
      Tab(0).Control(56)=   "txtCaseField(5)"
      Tab(0).Control(56).Enabled=   0   'False
      Tab(0).Control(57)=   "txtCaseField(3)"
      Tab(0).Control(57).Enabled=   0   'False
      Tab(0).Control(58)=   "txtCaseField(11)"
      Tab(0).Control(58).Enabled=   0   'False
      Tab(0).Control(59)=   "txtCaseField(10)"
      Tab(0).Control(59).Enabled=   0   'False
      Tab(0).Control(60)=   "txtCaseField(12)"
      Tab(0).Control(60).Enabled=   0   'False
      Tab(0).Control(61)=   "txtCaseField(9)"
      Tab(0).Control(61).Enabled=   0   'False
      Tab(0).Control(62)=   "txtCaseField(15)"
      Tab(0).Control(62).Enabled=   0   'False
      Tab(0).Control(63)=   "txtCaseField(14)"
      Tab(0).Control(63).Enabled=   0   'False
      Tab(0).Control(64)=   "txtCaseField(17)"
      Tab(0).Control(64).Enabled=   0   'False
      Tab(0).Control(65)=   "txtCaseField(16)"
      Tab(0).Control(65).Enabled=   0   'False
      Tab(0).Control(66)=   "txtCaseField(18)"
      Tab(0).Control(66).Enabled=   0   'False
      Tab(0).Control(67)=   "txtCaseField(13)"
      Tab(0).Control(67).Enabled=   0   'False
      Tab(0).Control(68)=   "txtCaseField(19)"
      Tab(0).Control(68).Enabled=   0   'False
      Tab(0).Control(69)=   "txtCaseField(20)"
      Tab(0).Control(69).Enabled=   0   'False
      Tab(0).Control(70)=   "txtCaseField(69)"
      Tab(0).Control(70).Enabled=   0   'False
      Tab(0).Control(71)=   "txtCaseField(38)"
      Tab(0).Control(71).Enabled=   0   'False
      Tab(0).Control(72)=   "chkChoose(3)"
      Tab(0).Control(72).Enabled=   0   'False
      Tab(0).Control(73)=   "chkChoose(1)"
      Tab(0).Control(73).Enabled=   0   'False
      Tab(0).Control(74)=   "chkChoose(2)"
      Tab(0).Control(74).Enabled=   0   'False
      Tab(0).Control(75)=   "chkChoose(4)"
      Tab(0).Control(75).Enabled=   0   'False
      Tab(0).Control(76)=   "chkChoose(5)"
      Tab(0).Control(76).Enabled=   0   'False
      Tab(0).Control(77)=   "chkChoose(6)"
      Tab(0).Control(77).Enabled=   0   'False
      Tab(0).Control(78)=   "cmdCountry"
      Tab(0).Control(78).Enabled=   0   'False
      Tab(0).Control(79)=   "cmdPriority"
      Tab(0).Control(79).Enabled=   0   'False
      Tab(0).Control(80)=   "optChoose(1)"
      Tab(0).Control(80).Enabled=   0   'False
      Tab(0).Control(81)=   "optChoose(0)"
      Tab(0).Control(81).Enabled=   0   'False
      Tab(0).Control(82)=   "cmdAddDeadLine"
      Tab(0).Control(82).Enabled=   0   'False
      Tab(0).Control(83)=   "cmdInventor"
      Tab(0).Control(83).Enabled=   0   'False
      Tab(0).Control(84)=   "Combo1"
      Tab(0).Control(84).Enabled=   0   'False
      Tab(0).Control(85)=   "txtAD(1)"
      Tab(0).Control(85).Enabled=   0   'False
      Tab(0).Control(86)=   "txtAD(3)"
      Tab(0).Control(86).Enabled=   0   'False
      Tab(0).Control(87)=   "txtAD(5)"
      Tab(0).Control(87).Enabled=   0   'False
      Tab(0).Control(88)=   "txtAD(2)"
      Tab(0).Control(88).Enabled=   0   'False
      Tab(0).Control(89)=   "txtAD(4)"
      Tab(0).Control(89).Enabled=   0   'False
      Tab(0).Control(90)=   "cmdSuggest"
      Tab(0).Control(90).Enabled=   0   'False
      Tab(0).Control(91)=   "optChoose(2)"
      Tab(0).Control(91).Enabled=   0   'False
      Tab(0).Control(92)=   "txtCP113"
      Tab(0).Control(92).Enabled=   0   'False
      Tab(0).ControlCount=   93
      TabCaption(1)   =   "備註"
      TabPicture(1)   =   "frm050102_3.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label2(1)"
      Tab(1).Control(1)=   "Label2(2)"
      Tab(1).Control(2)=   "txtCaseField(22)"
      Tab(1).Control(3)=   "txtCaseField(21)"
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "申請人地址"
      TabPicture(2)   =   "frm050102_3.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label3(3)"
      Tab(2).Control(1)=   "Label3(5)"
      Tab(2).Control(2)=   "Label3(6)"
      Tab(2).Control(3)=   "Label3(4)"
      Tab(2).Control(4)=   "Label3(7)"
      Tab(2).Control(5)=   "Label3(8)"
      Tab(2).Control(6)=   "Label3(9)"
      Tab(2).Control(7)=   "Label3(10)"
      Tab(2).Control(8)=   "Label3(11)"
      Tab(2).Control(9)=   "Label3(12)"
      Tab(2).Control(10)=   "Label3(13)"
      Tab(2).Control(11)=   "Label3(14)"
      Tab(2).Control(12)=   "Label3(15)"
      Tab(2).Control(13)=   "Label3(16)"
      Tab(2).Control(14)=   "Label3(17)"
      Tab(2).Control(15)=   "txtCaseField(23)"
      Tab(2).Control(16)=   "txtCaseField(24)"
      Tab(2).Control(17)=   "txtCaseField(25)"
      Tab(2).Control(18)=   "txtCaseField(26)"
      Tab(2).Control(19)=   "txtCaseField(27)"
      Tab(2).Control(20)=   "txtCaseField(28)"
      Tab(2).Control(21)=   "txtCaseField(29)"
      Tab(2).Control(22)=   "txtCaseField(30)"
      Tab(2).Control(23)=   "txtCaseField(31)"
      Tab(2).Control(24)=   "txtCaseField(32)"
      Tab(2).Control(25)=   "txtCaseField(33)"
      Tab(2).Control(26)=   "txtCaseField(34)"
      Tab(2).Control(27)=   "txtCaseField(35)"
      Tab(2).Control(28)=   "txtCaseField(36)"
      Tab(2).Control(29)=   "txtCaseField(37)"
      Tab(2).ControlCount=   30
      TabCaption(3)   =   "代表人1"
      TabPicture(3)   =   "frm050102_3.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label14(3)"
      Tab(3).Control(1)=   "Label5(35)"
      Tab(3).Control(2)=   "Label5(34)"
      Tab(3).Control(3)=   "Label5(33)"
      Tab(3).Control(4)=   "Label18(1)"
      Tab(3).Control(5)=   "Label14(2)"
      Tab(3).Control(6)=   "Label5(29)"
      Tab(3).Control(7)=   "Label5(28)"
      Tab(3).Control(8)=   "Label5(27)"
      Tab(3).Control(9)=   "Label5(26)"
      Tab(3).Control(10)=   "Label5(25)"
      Tab(3).Control(11)=   "Label5(24)"
      Tab(3).Control(12)=   "Label18(2)"
      Tab(3).Control(13)=   "Label14(1)"
      Tab(3).Control(14)=   "Label5(3)"
      Tab(3).Control(15)=   "Label5(4)"
      Tab(3).Control(16)=   "Label5(5)"
      Tab(3).Control(17)=   "Label5(6)"
      Tab(3).Control(18)=   "Label5(7)"
      Tab(3).Control(19)=   "Label5(8)"
      Tab(3).Control(20)=   "txtCaseField(53)"
      Tab(3).Control(21)=   "txtCaseField(52)"
      Tab(3).Control(22)=   "txtCaseField(50)"
      Tab(3).Control(23)=   "txtCaseField(41)"
      Tab(3).Control(24)=   "txtCaseField(40)"
      Tab(3).Control(25)=   "txtCaseField(49)"
      Tab(3).Control(26)=   "txtCaseField(39)"
      Tab(3).Control(27)=   "txtCaseField(44)"
      Tab(3).Control(28)=   "txtCaseField(43)"
      Tab(3).Control(29)=   "txtCaseField(42)"
      Tab(3).Control(30)=   "txtCaseField(47)"
      Tab(3).Control(31)=   "txtCaseField(46)"
      Tab(3).Control(32)=   "txtCaseField(45)"
      Tab(3).Control(33)=   "txtCaseField(48)"
      Tab(3).Control(34)=   "txtCaseField(51)"
      Tab(3).Control(35)=   "Combo2(0)"
      Tab(3).Control(36)=   "Combo2(1)"
      Tab(3).Control(37)=   "Combo2(2)"
      Tab(3).Control(38)=   "Combo2(3)"
      Tab(3).Control(39)=   "Combo2(4)"
      Tab(3).ControlCount=   40
      TabCaption(4)   =   "代表人2"
      TabPicture(4)   =   "frm050102_3.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Label14(4)"
      Tab(4).Control(1)=   "Label5(1)"
      Tab(4).Control(2)=   "Label5(2)"
      Tab(4).Control(3)=   "Label5(9)"
      Tab(4).Control(4)=   "Label18(3)"
      Tab(4).Control(5)=   "Label14(5)"
      Tab(4).Control(6)=   "Label5(10)"
      Tab(4).Control(7)=   "Label5(11)"
      Tab(4).Control(8)=   "Label5(12)"
      Tab(4).Control(9)=   "Label5(13)"
      Tab(4).Control(10)=   "Label5(14)"
      Tab(4).Control(11)=   "Label5(15)"
      Tab(4).Control(12)=   "Label18(4)"
      Tab(4).Control(13)=   "Label14(6)"
      Tab(4).Control(14)=   "Label5(16)"
      Tab(4).Control(15)=   "Label5(17)"
      Tab(4).Control(16)=   "Label5(18)"
      Tab(4).Control(17)=   "Label5(19)"
      Tab(4).Control(18)=   "Label5(20)"
      Tab(4).Control(19)=   "Label5(21)"
      Tab(4).Control(20)=   "txtCaseField(68)"
      Tab(4).Control(21)=   "txtCaseField(67)"
      Tab(4).Control(22)=   "txtCaseField(66)"
      Tab(4).Control(23)=   "txtCaseField(65)"
      Tab(4).Control(24)=   "txtCaseField(64)"
      Tab(4).Control(25)=   "txtCaseField(63)"
      Tab(4).Control(26)=   "txtCaseField(62)"
      Tab(4).Control(27)=   "txtCaseField(61)"
      Tab(4).Control(28)=   "txtCaseField(60)"
      Tab(4).Control(29)=   "txtCaseField(59)"
      Tab(4).Control(30)=   "txtCaseField(58)"
      Tab(4).Control(31)=   "txtCaseField(57)"
      Tab(4).Control(32)=   "txtCaseField(56)"
      Tab(4).Control(33)=   "txtCaseField(55)"
      Tab(4).Control(34)=   "txtCaseField(54)"
      Tab(4).Control(35)=   "Combo2(5)"
      Tab(4).Control(36)=   "Combo2(6)"
      Tab(4).Control(37)=   "Combo2(7)"
      Tab(4).Control(38)=   "Combo2(8)"
      Tab(4).Control(39)=   "Combo2(9)"
      Tab(4).ControlCount=   40
      Begin VB.TextBox txtCP113 
         Height          =   300
         Left            =   7470
         MaxLength       =   4
         TabIndex        =   15
         Top             =   2487
         Width           =   540
      End
      Begin VB.OptionButton optChoose 
         Caption         =   "微個體"
         Height          =   255
         Index           =   2
         Left            =   3720
         TabIndex        =   224
         Top             =   5034
         Width           =   1395
      End
      Begin VB.CommandButton cmdSuggest 
         Caption         =   "C"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   14.5
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   765
         TabIndex        =   219
         Top             =   3840
         Width           =   300
      End
      Begin VB.TextBox txtAD 
         Enabled         =   0   'False
         Height          =   300
         Index           =   4
         Left            =   1980
         MaxLength       =   1
         TabIndex        =   10
         Top             =   1920
         Width           =   240
      End
      Begin VB.TextBox txtAD 
         Enabled         =   0   'False
         Height          =   300
         Index           =   2
         Left            =   1980
         MaxLength       =   1
         TabIndex        =   6
         Top             =   1380
         Width           =   240
      End
      Begin VB.TextBox txtAD 
         Enabled         =   0   'False
         Height          =   300
         Index           =   5
         Left            =   1980
         MaxLength       =   1
         TabIndex        =   12
         Top             =   2190
         Width           =   240
      End
      Begin VB.TextBox txtAD 
         Enabled         =   0   'False
         Height          =   300
         Index           =   3
         Left            =   1980
         MaxLength       =   1
         TabIndex        =   8
         Top             =   1650
         Width           =   240
      End
      Begin VB.TextBox txtAD 
         Enabled         =   0   'False
         ForeColor       =   &H000000FF&
         Height          =   300
         Index           =   1
         Left            =   1980
         MaxLength       =   1
         TabIndex        =   4
         Top             =   1110
         Width           =   240
      End
      Begin VB.ComboBox Combo2 
         Height          =   300
         Index           =   9
         ItemData        =   "frm050102_3.frx":008C
         Left            =   -73560
         List            =   "frm050102_3.frx":008E
         Style           =   2  '單純下拉式
         TabIndex        =   207
         Top             =   4185
         Width           =   6135
      End
      Begin VB.ComboBox Combo2 
         Height          =   300
         Index           =   8
         ItemData        =   "frm050102_3.frx":0090
         Left            =   -73560
         List            =   "frm050102_3.frx":0092
         Style           =   2  '單純下拉式
         TabIndex        =   206
         Top             =   3225
         Width           =   6135
      End
      Begin VB.ComboBox Combo2 
         Height          =   300
         Index           =   7
         ItemData        =   "frm050102_3.frx":0094
         Left            =   -73560
         List            =   "frm050102_3.frx":0096
         Style           =   2  '單純下拉式
         TabIndex        =   205
         Top             =   2280
         Width           =   6135
      End
      Begin VB.ComboBox Combo2 
         Height          =   300
         Index           =   6
         ItemData        =   "frm050102_3.frx":0098
         Left            =   -73560
         List            =   "frm050102_3.frx":009A
         Style           =   2  '單純下拉式
         TabIndex        =   204
         Top             =   1305
         Width           =   6135
      End
      Begin VB.ComboBox Combo2 
         Height          =   300
         Index           =   5
         ItemData        =   "frm050102_3.frx":009C
         Left            =   -73560
         List            =   "frm050102_3.frx":009E
         Style           =   2  '單純下拉式
         TabIndex        =   203
         Top             =   345
         Width           =   6135
      End
      Begin VB.ComboBox Combo2 
         Height          =   300
         Index           =   4
         ItemData        =   "frm050102_3.frx":00A0
         Left            =   -73560
         List            =   "frm050102_3.frx":00A2
         Style           =   2  '單純下拉式
         TabIndex        =   167
         Top             =   4185
         Width           =   6135
      End
      Begin VB.ComboBox Combo2 
         Height          =   300
         Index           =   3
         ItemData        =   "frm050102_3.frx":00A4
         Left            =   -73560
         List            =   "frm050102_3.frx":00A6
         Style           =   2  '單純下拉式
         TabIndex        =   166
         Top             =   3225
         Width           =   6135
      End
      Begin VB.ComboBox Combo2 
         Height          =   300
         Index           =   2
         ItemData        =   "frm050102_3.frx":00A8
         Left            =   -73560
         List            =   "frm050102_3.frx":00AA
         Style           =   2  '單純下拉式
         TabIndex        =   165
         Top             =   2280
         Width           =   6135
      End
      Begin VB.ComboBox Combo2 
         Height          =   300
         Index           =   1
         ItemData        =   "frm050102_3.frx":00AC
         Left            =   -73560
         List            =   "frm050102_3.frx":00AE
         Style           =   2  '單純下拉式
         TabIndex        =   164
         Top             =   1305
         Width           =   6135
      End
      Begin VB.ComboBox Combo2 
         Height          =   300
         Index           =   0
         ItemData        =   "frm050102_3.frx":00B0
         Left            =   -73555
         List            =   "frm050102_3.frx":00B2
         Style           =   2  '單純下拉式
         TabIndex        =   163
         Top             =   345
         Width           =   6135
      End
      Begin VB.ComboBox Combo1 
         Height          =   300
         Left            =   1050
         TabIndex        =   27
         Top             =   3855
         Width           =   1485
      End
      Begin VB.CommandButton cmdInventor 
         Caption         =   "輸入(&I)"
         Height          =   270
         Left            =   4770
         TabIndex        =   36
         Top             =   4725
         Width           =   972
      End
      Begin VB.CommandButton cmdAddDeadLine 
         Caption         =   "輸入(&D)"
         Height          =   270
         Left            =   7125
         TabIndex        =   28
         Top             =   3855
         Width           =   870
      End
      Begin VB.OptionButton optChoose 
         Caption         =   "大個體"
         Height          =   255
         Index           =   0
         Left            =   1560
         TabIndex        =   37
         Top             =   5034
         Width           =   945
      End
      Begin VB.OptionButton optChoose 
         Caption         =   "小個體"
         Height          =   255
         Index           =   1
         Left            =   2580
         TabIndex        =   38
         Top             =   5034
         Width           =   1095
      End
      Begin VB.CommandButton cmdPriority 
         Caption         =   "輸入(&P)"
         Height          =   270
         Left            =   7125
         TabIndex        =   26
         Top             =   3570
         Width           =   870
      End
      Begin VB.CommandButton cmdCountry 
         Caption         =   "指定國家(&L)"
         Height          =   270
         Left            =   3180
         TabIndex        =   25
         Top             =   3570
         Width           =   1200
      End
      Begin MSForms.CheckBox chkChoose 
         Height          =   255
         Index           =   6
         Left            =   7020
         TabIndex        =   223
         Top             =   4800
         Width           =   1245
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         DisplayStyle    =   4
         Size            =   "2196;450"
         Value           =   "0"
         Caption         =   "一併提讓渡"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.CheckBox chkChoose 
         Height          =   255
         Index           =   5
         Left            =   6540
         TabIndex        =   142
         Top             =   5070
         Width           =   1785
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         DisplayStyle    =   4
         Size            =   "3149;450"
         Value           =   "0"
         Caption         =   "一併提檢索及實審"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.CheckBox chkChoose 
         Height          =   255
         Index           =   4
         Left            =   5850
         TabIndex        =   141
         Top             =   4800
         Width           =   1155
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         DisplayStyle    =   4
         Size            =   "2037;450"
         Value           =   "0"
         Caption         =   "一併提IDS"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.CheckBox chkChoose 
         Height          =   255
         Index           =   2
         Left            =   1050
         TabIndex        =   19
         Top             =   3000
         Width           =   765
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         DisplayStyle    =   4
         Size            =   "1349;450"
         Value           =   "0"
         Caption         =   "照片"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.CheckBox chkChoose 
         Height          =   255
         Index           =   1
         Left            =   135
         TabIndex        =   18
         Top             =   3000
         Width           =   735
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         DisplayStyle    =   4
         Size            =   "1296;450"
         Value           =   "0"
         Caption         =   "圖式"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.CheckBox chkChoose 
         Height          =   255
         Index           =   3
         Left            =   5190
         TabIndex        =   39
         Top             =   5070
         Width           =   1275
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         DisplayStyle    =   4
         Size            =   "2249;450"
         Value           =   "0"
         Caption         =   "一併提實審"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   300
         Index           =   38
         Left            =   2250
         TabIndex        =   21
         Top             =   3015
         Width           =   2415
         VariousPropertyBits=   671107097
         Size            =   "4260;529"
         FontName        =   "新細明體-ExtB"
         FontEffects     =   1073750016
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   300
         Index           =   69
         Left            =   4800
         TabIndex        =   14
         Top             =   2490
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
         Index           =   20
         Left            =   3735
         TabIndex        =   30
         Top             =   4140
         Width           =   870
         VariousPropertyBits=   671107099
         MaxLength       =   8
         Size            =   "1535;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   300
         Index           =   19
         Left            =   7065
         TabIndex        =   34
         Top             =   4440
         Width           =   375
         VariousPropertyBits=   671107099
         MaxLength       =   1
         Size            =   "661;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   300
         Index           =   54
         Left            =   -73560
         TabIndex        =   182
         Top             =   600
         Width           =   6135
         VariousPropertyBits=   671107099
         BackColor       =   -2147483644
         MaxLength       =   50
         Size            =   "10821;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   300
         Index           =   55
         Left            =   -73560
         TabIndex        =   181
         Top             =   825
         Width           =   6135
         VariousPropertyBits=   671107099
         BackColor       =   -2147483644
         MaxLength       =   80
         Size            =   "10821;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   300
         Index           =   56
         Left            =   -73560
         TabIndex        =   180
         Top             =   1050
         Width           =   6135
         VariousPropertyBits=   671107099
         BackColor       =   -2147483644
         MaxLength       =   40
         Size            =   "10821;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   300
         Index           =   57
         Left            =   -73560
         TabIndex        =   179
         Top             =   1560
         Width           =   6135
         VariousPropertyBits=   671107099
         BackColor       =   -2147483644
         MaxLength       =   50
         Size            =   "10821;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   300
         Index           =   58
         Left            =   -73560
         TabIndex        =   178
         Top             =   1800
         Width           =   6135
         VariousPropertyBits=   671107099
         BackColor       =   -2147483644
         MaxLength       =   80
         Size            =   "10821;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   300
         Index           =   59
         Left            =   -73560
         TabIndex        =   177
         Top             =   2025
         Width           =   6135
         VariousPropertyBits=   671107099
         BackColor       =   -2147483644
         MaxLength       =   40
         Size            =   "10821;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   300
         Index           =   60
         Left            =   -73560
         TabIndex        =   176
         Top             =   2535
         Width           =   6135
         VariousPropertyBits=   671107099
         BackColor       =   -2147483644
         MaxLength       =   50
         Size            =   "10821;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   300
         Index           =   61
         Left            =   -73560
         TabIndex        =   175
         Top             =   2760
         Width           =   6135
         VariousPropertyBits=   671107099
         BackColor       =   -2147483644
         MaxLength       =   80
         Size            =   "10821;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   300
         Index           =   62
         Left            =   -73560
         TabIndex        =   174
         Top             =   3000
         Width           =   6135
         VariousPropertyBits=   671107099
         BackColor       =   -2147483644
         MaxLength       =   40
         Size            =   "10821;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   300
         Index           =   63
         Left            =   -73560
         TabIndex        =   173
         Top             =   3480
         Width           =   6135
         VariousPropertyBits=   671107099
         BackColor       =   -2147483644
         MaxLength       =   50
         Size            =   "10821;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   300
         Index           =   64
         Left            =   -73560
         TabIndex        =   172
         Top             =   3705
         Width           =   6135
         VariousPropertyBits=   671107099
         BackColor       =   -2147483644
         MaxLength       =   80
         Size            =   "10821;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   300
         Index           =   65
         Left            =   -73560
         TabIndex        =   171
         Top             =   3930
         Width           =   6135
         VariousPropertyBits=   671107099
         BackColor       =   -2147483644
         MaxLength       =   40
         Size            =   "10821;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   300
         Index           =   66
         Left            =   -73560
         TabIndex        =   170
         Top             =   4440
         Width           =   6135
         VariousPropertyBits=   671107099
         BackColor       =   -2147483644
         MaxLength       =   50
         Size            =   "10821;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   300
         Index           =   67
         Left            =   -73560
         TabIndex        =   169
         Top             =   4665
         Width           =   6135
         VariousPropertyBits=   671107099
         BackColor       =   -2147483644
         MaxLength       =   80
         Size            =   "10821;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   300
         Index           =   68
         Left            =   -73560
         TabIndex        =   168
         Top             =   4890
         Width           =   6135
         VariousPropertyBits=   671107099
         BackColor       =   -2147483644
         MaxLength       =   40
         Size            =   "10821;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   300
         Index           =   51
         Left            =   -73560
         TabIndex        =   138
         Top             =   4440
         Width           =   6135
         VariousPropertyBits=   671107099
         BackColor       =   -2147483644
         MaxLength       =   50
         Size            =   "10821;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   300
         Index           =   48
         Left            =   -73560
         TabIndex        =   135
         Top             =   3480
         Width           =   6135
         VariousPropertyBits=   671107099
         BackColor       =   -2147483644
         MaxLength       =   50
         Size            =   "10821;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   300
         Index           =   45
         Left            =   -73560
         TabIndex        =   132
         Top             =   2535
         Width           =   6135
         VariousPropertyBits=   671107099
         BackColor       =   -2147483644
         MaxLength       =   50
         Size            =   "10821;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   300
         Index           =   46
         Left            =   -73560
         TabIndex        =   133
         Top             =   2760
         Width           =   6135
         VariousPropertyBits=   671107099
         BackColor       =   -2147483644
         MaxLength       =   80
         Size            =   "10821;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   300
         Index           =   47
         Left            =   -73560
         TabIndex        =   134
         Top             =   3000
         Width           =   6135
         VariousPropertyBits=   671107099
         BackColor       =   -2147483644
         MaxLength       =   40
         Size            =   "10821;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   300
         Index           =   42
         Left            =   -73560
         TabIndex        =   129
         Top             =   1560
         Width           =   6135
         VariousPropertyBits=   671107099
         BackColor       =   -2147483644
         MaxLength       =   50
         Size            =   "10821;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   300
         Index           =   43
         Left            =   -73560
         TabIndex        =   130
         Top             =   1800
         Width           =   6135
         VariousPropertyBits=   671107099
         BackColor       =   -2147483644
         MaxLength       =   80
         Size            =   "10821;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   300
         Index           =   44
         Left            =   -73560
         TabIndex        =   131
         Top             =   2025
         Width           =   6135
         VariousPropertyBits=   671107099
         BackColor       =   -2147483644
         MaxLength       =   40
         Size            =   "10821;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   300
         Index           =   39
         Left            =   -73560
         TabIndex        =   126
         Top             =   600
         Width           =   6135
         VariousPropertyBits=   671107099
         BackColor       =   -2147483644
         MaxLength       =   50
         Size            =   "10821;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   300
         Index           =   49
         Left            =   -73560
         TabIndex        =   136
         Top             =   3705
         Width           =   6135
         VariousPropertyBits=   671107099
         BackColor       =   -2147483644
         MaxLength       =   80
         Size            =   "10821;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   300
         Index           =   13
         Left            =   5295
         TabIndex        =   17
         Top             =   2760
         Width           =   375
         VariousPropertyBits=   671107099
         MaxLength       =   1
         Size            =   "661;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   300
         Index           =   37
         Left            =   -73200
         TabIndex        =   56
         Top             =   4140
         Width           =   6375
         VariousPropertyBits=   671107099
         MaxLength       =   80
         Size            =   "11245;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   300
         Index           =   36
         Left            =   -73200
         TabIndex        =   55
         Top             =   3870
         Width           =   6375
         VariousPropertyBits=   671107099
         MaxLength       =   185
         Size            =   "11245;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   300
         Index           =   35
         Left            =   -73200
         TabIndex        =   54
         Top             =   3585
         Width           =   6375
         VariousPropertyBits=   671107099
         MaxLength       =   80
         Size            =   "11245;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   300
         Index           =   34
         Left            =   -73200
         TabIndex        =   53
         Top             =   3330
         Width           =   6375
         VariousPropertyBits=   671107099
         MaxLength       =   80
         Size            =   "11245;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   300
         Index           =   33
         Left            =   -73200
         TabIndex        =   52
         Top             =   3030
         Width           =   6375
         VariousPropertyBits=   671107099
         MaxLength       =   185
         Size            =   "11245;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   300
         Index           =   32
         Left            =   -73200
         TabIndex        =   51
         Top             =   2760
         Width           =   6375
         VariousPropertyBits=   671107099
         MaxLength       =   80
         Size            =   "11245;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   300
         Index           =   31
         Left            =   -73200
         TabIndex        =   50
         Top             =   2490
         Width           =   6375
         VariousPropertyBits=   671107099
         MaxLength       =   80
         Size            =   "11245;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   300
         Index           =   30
         Left            =   -73200
         TabIndex        =   49
         Top             =   2205
         Width           =   6375
         VariousPropertyBits=   671107099
         MaxLength       =   185
         Size            =   "11245;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   300
         Index           =   29
         Left            =   -73200
         TabIndex        =   48
         Top             =   1935
         Width           =   6375
         VariousPropertyBits=   671107099
         MaxLength       =   80
         Size            =   "11245;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   300
         Index           =   28
         Left            =   -73200
         TabIndex        =   47
         Top             =   1650
         Width           =   6375
         VariousPropertyBits=   671107099
         MaxLength       =   80
         Size            =   "11245;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   300
         Index           =   27
         Left            =   -73200
         TabIndex        =   46
         Top             =   1365
         Width           =   6375
         VariousPropertyBits=   671107099
         MaxLength       =   185
         Size            =   "11245;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   300
         Index           =   26
         Left            =   -73200
         TabIndex        =   45
         Top             =   1110
         Width           =   6375
         VariousPropertyBits=   671107099
         MaxLength       =   80
         Size            =   "11245;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   300
         Index           =   18
         Left            =   2250
         TabIndex        =   35
         Top             =   4725
         Width           =   375
         VariousPropertyBits=   671107099
         MaxLength       =   1
         Size            =   "661;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   300
         Index           =   16
         Left            =   4155
         TabIndex        =   33
         Top             =   4440
         Width           =   375
         VariousPropertyBits=   671107099
         MaxLength       =   1
         Size            =   "661;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   300
         Index           =   17
         Left            =   1395
         TabIndex        =   32
         Top             =   4440
         Width           =   375
         VariousPropertyBits=   671107099
         MaxLength       =   1
         Size            =   "661;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   300
         Index           =   14
         Left            =   1395
         TabIndex        =   29
         Top             =   4140
         Width           =   870
         VariousPropertyBits=   671107099
         MaxLength       =   8
         Size            =   "1535;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   300
         Index           =   15
         Left            =   7125
         TabIndex        =   31
         Top             =   4140
         Width           =   870
         VariousPropertyBits=   671107099
         MaxLength       =   8
         Size            =   "1535;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   1815
         Index           =   21
         Left            =   -73800
         TabIndex        =   40
         Top             =   450
         Width           =   6975
         VariousPropertyBits=   -1467987941
         ScrollBars      =   2
         Size            =   "12303;3201"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   1470
         Index           =   22
         Left            =   -73800
         TabIndex        =   41
         Top             =   2280
         Width           =   6975
         VariousPropertyBits=   -1467987941
         ScrollBars      =   2
         Size            =   "12303;2593"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   300
         Index           =   25
         Left            =   -73200
         TabIndex        =   44
         Top             =   825
         Width           =   6375
         VariousPropertyBits=   671107099
         MaxLength       =   80
         Size            =   "11245;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   300
         Index           =   24
         Left            =   -73200
         TabIndex        =   43
         Top             =   555
         Width           =   6375
         VariousPropertyBits=   671107099
         MaxLength       =   185
         Size            =   "11245;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   300
         Index           =   23
         Left            =   -73200
         TabIndex        =   42
         Top             =   270
         Width           =   6375
         VariousPropertyBits=   671107099
         MaxLength       =   80
         Size            =   "11245;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   300
         Index           =   9
         Left            =   1455
         TabIndex        =   13
         Top             =   2490
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
         Index           =   12
         Left            =   1050
         TabIndex        =   24
         Top             =   3570
         Width           =   615
         VariousPropertyBits=   671107099
         MaxLength       =   3
         Size            =   "1085;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   300
         Index           =   10
         Left            =   1050
         TabIndex        =   22
         Top             =   3285
         Width           =   615
         VariousPropertyBits=   671107099
         MaxLength       =   4
         Size            =   "1085;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   300
         Index           =   11
         Left            =   7125
         TabIndex        =   23
         Top             =   3270
         Width           =   870
         VariousPropertyBits=   671107099
         MaxLength       =   8
         Size            =   "1535;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   300
         Index           =   3
         Left            =   885
         TabIndex        =   3
         Top             =   1110
         Width           =   1095
         VariousPropertyBits=   671107099
         MaxLength       =   9
         Size            =   "1931;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   300
         Index           =   5
         Left            =   885
         TabIndex        =   7
         Top             =   1650
         Width           =   1095
         VariousPropertyBits=   671107099
         MaxLength       =   9
         Size            =   "1931;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   300
         Index           =   4
         Left            =   885
         TabIndex        =   5
         Top             =   1380
         Width           =   1095
         VariousPropertyBits=   671107099
         MaxLength       =   9
         Size            =   "1931;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   300
         Index           =   6
         Left            =   885
         TabIndex        =   9
         Top             =   1920
         Width           =   1095
         VariousPropertyBits=   671107099
         MaxLength       =   9
         Size            =   "1931;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   300
         Index           =   7
         Left            =   885
         TabIndex        =   11
         Top             =   2190
         Width           =   1095
         VariousPropertyBits=   671107099
         MaxLength       =   9
         Size            =   "1931;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   300
         Index           =   2
         Left            =   1320
         TabIndex        =   2
         Top             =   810
         Width           =   6855
         VariousPropertyBits=   671107099
         MaxLength       =   160
         Size            =   "12091;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   300
         Index           =   1
         Left            =   1320
         TabIndex        =   1
         Top             =   570
         Width           =   6855
         VariousPropertyBits=   671107099
         MaxLength       =   250
         Size            =   "12091;529"
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
         Top             =   300
         Width           =   6855
         VariousPropertyBits=   671107099
         MaxLength       =   160
         Size            =   "12091;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   300
         Index           =   8
         Left            =   1575
         TabIndex        =   16
         Top             =   2760
         Width           =   375
         VariousPropertyBits=   671107099
         MaxLength       =   1
         Size            =   "661;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   300
         Index           =   40
         Left            =   -73560
         TabIndex        =   127
         Top             =   825
         Width           =   6135
         VariousPropertyBits=   671107099
         BackColor       =   -2147483644
         MaxLength       =   80
         Size            =   "10821;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   300
         Index           =   41
         Left            =   -73560
         TabIndex        =   128
         Top             =   1050
         Width           =   6135
         VariousPropertyBits=   671107099
         BackColor       =   -2147483644
         MaxLength       =   40
         Size            =   "10821;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   300
         Index           =   50
         Left            =   -73560
         TabIndex        =   137
         Top             =   3930
         Width           =   6135
         VariousPropertyBits=   671107099
         BackColor       =   -2147483644
         MaxLength       =   40
         Size            =   "10821;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   300
         Index           =   52
         Left            =   -73560
         TabIndex        =   139
         Top             =   4665
         Width           =   6135
         VariousPropertyBits=   671107099
         BackColor       =   -2147483644
         MaxLength       =   80
         Size            =   "10821;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   300
         Index           =   53
         Left            =   -73560
         TabIndex        =   140
         Top             =   4905
         Width           =   6135
         VariousPropertyBits=   671107099
         BackColor       =   -2147483644
         MaxLength       =   40
         Size            =   "10821;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label lblCP113 
         AutoSize        =   -1  'True
         Caption         =   "工作時數："
         Height          =   180
         Index           =   18
         Left            =   6570
         TabIndex        =   226
         Top             =   2535
         Width           =   900
      End
      Begin VB.Label Label18 
         Caption         =   "是否列印DHL：                (Y：DHL)"
         Height          =   255
         Index           =   5
         Left            =   3480
         TabIndex        =   225
         Top             =   2490
         Width           =   2835
      End
      Begin VB.Label lblCaseFee 
         Appearance      =   0  '平面
         AutoSize        =   -1  'True
         Caption         =   "@"
         BeginProperty Font 
            Name            =   "Webdings"
            Size            =   12
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   285
         Left            =   8010
         TabIndex        =   221
         Tag             =   "Y"
         Top             =   4110
         Width           =   255
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "指定提申日:"
         Height          =   180
         Index           =   20
         Left            =   2700
         TabIndex        =   220
         Top             =   4185
         Width           =   945
      End
      Begin VB.Label Label4 
         Caption         =   "是否印傳真封面：        （N:不印）"
         Height          =   255
         Index           =   4
         Left            =   5625
         TabIndex        =   218
         Top             =   4455
         Width           =   2685
      End
      Begin MSForms.Label lblAppNation 
         Height          =   255
         Index           =   4
         Left            =   6390
         TabIndex        =   217
         Top             =   2205
         Width           =   1785
         VariousPropertyBits=   27
         Size            =   "3149;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lblAppNation 
         Height          =   255
         Index           =   3
         Left            =   6390
         TabIndex        =   216
         Top             =   1950
         Width           =   1785
         VariousPropertyBits=   27
         Size            =   "3149;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lblAppNation 
         Height          =   255
         Index           =   2
         Left            =   6390
         TabIndex        =   215
         Top             =   1650
         Width           =   1785
         VariousPropertyBits=   27
         Size            =   "3149;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lblAppNation 
         Height          =   255
         Index           =   1
         Left            =   6390
         TabIndex        =   214
         Top             =   1380
         Width           =   1785
         VariousPropertyBits=   27
         Size            =   "3149;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label3 
         Caption         =   "國籍："
         Height          =   255
         Index           =   18
         Left            =   5760
         TabIndex        =   213
         Top             =   2205
         Width           =   630
      End
      Begin VB.Label Label4 
         Caption         =   "國籍："
         Height          =   255
         Index           =   3
         Left            =   5760
         TabIndex        =   212
         Top             =   1650
         Width           =   630
      End
      Begin VB.Label Label6 
         Caption         =   "國籍："
         Height          =   255
         Index           =   4
         Left            =   5760
         TabIndex        =   211
         Top             =   1380
         Width           =   630
      End
      Begin VB.Label Label20 
         Caption         =   "國籍："
         Height          =   255
         Left            =   5760
         TabIndex        =   210
         Top             =   1950
         Width           =   630
      End
      Begin VB.Label Label13 
         Caption         =   "國籍："
         Height          =   255
         Left            =   5760
         TabIndex        =   209
         Top             =   1110
         Width           =   630
      End
      Begin MSForms.Label lblAppNation 
         Height          =   255
         Index           =   0
         Left            =   6390
         TabIndex        =   208
         Top             =   1110
         Width           =   1785
         VariousPropertyBits=   27
         Size            =   "3149;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(日):"
         Height          =   180
         Index           =   21
         Left            =   -74055
         TabIndex        =   202
         Top             =   2055
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(英):"
         Height          =   180
         Index           =   20
         Left            =   -74055
         TabIndex        =   201
         Top             =   1815
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(中):"
         Height          =   180
         Index           =   19
         Left            =   -74055
         TabIndex        =   200
         Top             =   1575
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(日):"
         Height          =   180
         Index           =   18
         Left            =   -74055
         TabIndex        =   199
         Top             =   1095
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(英):"
         Height          =   180
         Index           =   17
         Left            =   -74055
         TabIndex        =   198
         Top             =   855
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(中):"
         Height          =   180
         Index           =   16
         Left            =   -74055
         TabIndex        =   197
         Top             =   615
         Width           =   345
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "代表人6"
         Height          =   180
         Index           =   6
         Left            =   -74415
         TabIndex        =   196
         Top             =   405
         Width           =   630
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "代表人7"
         Height          =   180
         Index           =   4
         Left            =   -74415
         TabIndex        =   195
         Top             =   1335
         Width           =   630
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(日):"
         Height          =   180
         Index           =   15
         Left            =   -74055
         TabIndex        =   194
         Top             =   3975
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(英):"
         Height          =   180
         Index           =   14
         Left            =   -74055
         TabIndex        =   193
         Top             =   3735
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(中):"
         Height          =   180
         Index           =   13
         Left            =   -74055
         TabIndex        =   192
         Top             =   3495
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(日):"
         Height          =   180
         Index           =   12
         Left            =   -74055
         TabIndex        =   191
         Top             =   3015
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(英):"
         Height          =   180
         Index           =   11
         Left            =   -74055
         TabIndex        =   190
         Top             =   2775
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(中):"
         Height          =   180
         Index           =   10
         Left            =   -74055
         TabIndex        =   189
         Top             =   2535
         Width           =   345
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "代表人8"
         Height          =   180
         Index           =   5
         Left            =   -74415
         TabIndex        =   188
         Top             =   2295
         Width           =   630
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "代表人9"
         Height          =   180
         Index           =   3
         Left            =   -74415
         TabIndex        =   187
         Top             =   3255
         Width           =   630
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(日):"
         Height          =   180
         Index           =   9
         Left            =   -74055
         TabIndex        =   186
         Top             =   4935
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(英):"
         Height          =   180
         Index           =   2
         Left            =   -74055
         TabIndex        =   185
         Top             =   4695
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(中):"
         Height          =   180
         Index           =   1
         Left            =   -74055
         TabIndex        =   184
         Top             =   4455
         Width           =   345
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "代表人10"
         Height          =   180
         Index           =   4
         Left            =   -74415
         TabIndex        =   183
         Top             =   4215
         Width           =   720
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(日):"
         Height          =   180
         Index           =   8
         Left            =   -74055
         TabIndex        =   162
         Top             =   2055
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(英):"
         Height          =   180
         Index           =   7
         Left            =   -74055
         TabIndex        =   161
         Top             =   1815
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(中):"
         Height          =   180
         Index           =   6
         Left            =   -74055
         TabIndex        =   160
         Top             =   1575
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(日):"
         Height          =   180
         Index           =   5
         Left            =   -74055
         TabIndex        =   159
         Top             =   1095
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(英):"
         Height          =   180
         Index           =   4
         Left            =   -74055
         TabIndex        =   158
         Top             =   855
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(中):"
         Height          =   180
         Index           =   3
         Left            =   -74055
         TabIndex        =   157
         Top             =   615
         Width           =   345
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "代表人1"
         Height          =   180
         Index           =   1
         Left            =   -74415
         TabIndex        =   156
         Top             =   405
         Width           =   630
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "代表人2"
         Height          =   180
         Index           =   2
         Left            =   -74415
         TabIndex        =   155
         Top             =   1335
         Width           =   630
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(日):"
         Height          =   180
         Index           =   24
         Left            =   -74055
         TabIndex        =   154
         Top             =   3975
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(英):"
         Height          =   180
         Index           =   25
         Left            =   -74055
         TabIndex        =   153
         Top             =   3735
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(中):"
         Height          =   180
         Index           =   26
         Left            =   -74055
         TabIndex        =   152
         Top             =   3495
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(日):"
         Height          =   180
         Index           =   27
         Left            =   -74055
         TabIndex        =   151
         Top             =   3015
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(英):"
         Height          =   180
         Index           =   28
         Left            =   -74055
         TabIndex        =   150
         Top             =   2775
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(中):"
         Height          =   180
         Index           =   29
         Left            =   -74055
         TabIndex        =   149
         Top             =   2535
         Width           =   345
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "代表人3"
         Height          =   180
         Index           =   2
         Left            =   -74415
         TabIndex        =   148
         Top             =   2295
         Width           =   630
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "代表人4"
         Height          =   180
         Index           =   1
         Left            =   -74415
         TabIndex        =   147
         Top             =   3255
         Width           =   630
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(日):"
         Height          =   180
         Index           =   33
         Left            =   -74055
         TabIndex        =   146
         Top             =   4935
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(英):"
         Height          =   180
         Index           =   34
         Left            =   -74055
         TabIndex        =   145
         Top             =   4695
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(中):"
         Height          =   180
         Index           =   35
         Left            =   -74055
         TabIndex        =   144
         Top             =   4455
         Width           =   345
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "代表人5"
         Height          =   180
         Index           =   3
         Left            =   -74415
         TabIndex        =   143
         Top             =   4215
         Width           =   630
      End
      Begin VB.Label Label19 
         Caption         =   "是否修改通知函內容：        （Y：Word）"
         Height          =   180
         Left            =   3465
         TabIndex        =   125
         Top             =   2805
         Width           =   3465
      End
      Begin VB.Label Label15 
         Caption         =   "補件期限:"
         Height          =   255
         Left            =   5955
         TabIndex        =   124
         Top             =   3885
         Width           =   975
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "發明人資料："
         Height          =   180
         Index           =   0
         Left            =   3660
         TabIndex        =   123
         Top             =   4770
         Width           =   1080
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "申請人地址（日）："
         Height          =   180
         Index           =   17
         Left            =   -74880
         TabIndex        =   114
         Top             =   4170
         Width           =   1620
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "申請人地址（英）："
         Height          =   180
         Index           =   16
         Left            =   -74880
         TabIndex        =   113
         Top             =   3885
         Width           =   1620
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "申請人地址（中）："
         Height          =   180
         Index           =   15
         Left            =   -74880
         TabIndex        =   112
         Top             =   3615
         Width           =   1620
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "申請人地址（日）："
         Height          =   180
         Index           =   14
         Left            =   -74880
         TabIndex        =   111
         Top             =   3330
         Width           =   1620
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "申請人地址（英）："
         Height          =   180
         Index           =   13
         Left            =   -74880
         TabIndex        =   110
         Top             =   3060
         Width           =   1620
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "申請人地址（中）："
         Height          =   180
         Index           =   12
         Left            =   -74880
         TabIndex        =   109
         Top             =   2790
         Width           =   1620
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "申請人地址（日）："
         Height          =   180
         Index           =   11
         Left            =   -74880
         TabIndex        =   108
         Top             =   2505
         Width           =   1620
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "申請人地址（英）："
         Height          =   180
         Index           =   10
         Left            =   -74880
         TabIndex        =   107
         Top             =   2235
         Width           =   1620
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "申請人地址（中）："
         Height          =   180
         Index           =   9
         Left            =   -74880
         TabIndex        =   106
         Top             =   1950
         Width           =   1620
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "申請人地址（日）："
         Height          =   180
         Index           =   8
         Left            =   -74880
         TabIndex        =   105
         Top             =   1680
         Width           =   1620
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "申請人地址（英）："
         Height          =   180
         Index           =   7
         Left            =   -74880
         TabIndex        =   104
         Top             =   1410
         Width           =   1620
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "申請人地址（中）："
         Height          =   180
         Index           =   4
         Left            =   -74880
         TabIndex        =   103
         Top             =   1125
         Width           =   1620
      End
      Begin MSForms.Label lblAgent 
         Height          =   255
         Left            =   2610
         TabIndex        =   102
         Top             =   3870
         Width           =   2670
         BackColor       =   -2147483634
         VariousPropertyBits=   27
         Size            =   "4710;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "申請人與發明人是否相同：        （Y/N）"
         Height          =   180
         Index           =   1
         Left            =   120
         TabIndex        =   101
         Top             =   4770
         Width           =   3165
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "（美、加、法國案）"
         Height          =   180
         Index           =   3
         Left            =   45
         TabIndex        =   100
         Top             =   5070
         Width           =   1620
      End
      Begin VB.Label Label21 
         Caption         =   "是否修改指示信：        （Y:WORD）"
         Height          =   255
         Left            =   2700
         TabIndex        =   99
         Top             =   4455
         Width           =   2895
      End
      Begin VB.Label Label4 
         Caption         =   "是否印指示信：        （N:不印）"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   98
         Top             =   4455
         Width           =   2550
      End
      Begin VB.Label Label3 
         Caption         =   "最終提申期限:"
         Height          =   180
         Index           =   1
         Left            =   120
         TabIndex        =   97
         Top             =   4185
         Width           =   1125
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "催審期限:"
         Height          =   180
         Index           =   2
         Left            =   5955
         TabIndex        =   96
         Top             =   4185
         Width           =   765
      End
      Begin VB.Label Label23 
         Caption         =   "代理人:"
         Height          =   255
         Left            =   120
         TabIndex        =   95
         Top             =   3885
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "進度備註："
         Height          =   255
         Index           =   2
         Left            =   -74880
         TabIndex        =   94
         Top             =   450
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "案件備註："
         Height          =   255
         Index           =   1
         Left            =   -74880
         TabIndex        =   93
         Top             =   2340
         Width           =   975
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "申請人地址（日）："
         Height          =   180
         Index           =   6
         Left            =   -74880
         TabIndex        =   92
         Top             =   840
         Width           =   1620
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "申請人地址（英）："
         Height          =   180
         Index           =   5
         Left            =   -74880
         TabIndex        =   91
         Top             =   555
         Width           =   1620
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "申請人地址（中）："
         Height          =   180
         Index           =   3
         Left            =   -74880
         TabIndex        =   90
         Top             =   300
         Width           =   1620
      End
      Begin VB.Label Label37 
         Caption         =   "優先權資料:"
         Height          =   255
         Left            =   5955
         TabIndex        =   89
         Top             =   3600
         Width           =   1095
      End
      Begin MSForms.Label lblCaseProperty 
         Height          =   255
         Left            =   1725
         TabIndex        =   88
         Top             =   3315
         Width           =   2280
         BackColor       =   -2147483634
         VariousPropertyBits=   27
         Size            =   "4022;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lblNation 
         Height          =   255
         Left            =   1725
         TabIndex        =   87
         Top             =   3600
         Width           =   1335
         BackColor       =   -2147483634
         VariousPropertyBits=   27
         Size            =   "2355;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label18 
         Caption         =   "是否列印TNT：                (Y：TNT)"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   86
         Top             =   2490
         Width           =   2835
      End
      Begin VB.Label Label16 
         Caption         =   "申請國家:"
         Height          =   255
         Left            =   120
         TabIndex        =   85
         Top             =   3570
         Width           =   975
      End
      Begin VB.Label Label10 
         Caption         =   "案件性質:"
         Height          =   255
         Left            =   120
         TabIndex        =   84
         Top             =   3285
         Width           =   975
      End
      Begin VB.Label Label11 
         Caption         =   "發文日:"
         Height          =   255
         Left            =   5955
         TabIndex        =   83
         Top             =   3315
         Width           =   735
      End
      Begin VB.Label Label17 
         Caption         =   "申請人："
         Height          =   255
         Left            =   120
         TabIndex        =   82
         Top             =   1110
         Width           =   855
      End
      Begin MSForms.Label lblPetition 
         Height          =   255
         Index           =   1
         Left            =   2235
         TabIndex        =   81
         Top             =   1380
         Width           =   3450
         VariousPropertyBits=   27
         Size            =   "6085;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lblPetition 
         Height          =   255
         Index           =   0
         Left            =   2235
         TabIndex        =   80
         Top             =   1110
         Width           =   3450
         VariousPropertyBits=   27
         Size            =   "6085;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label7 
         Caption         =   "申請人："
         Height          =   255
         Left            =   120
         TabIndex        =   79
         Top             =   1950
         Width           =   855
      End
      Begin VB.Label Label6 
         Caption         =   "申請人："
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   78
         Top             =   1380
         Width           =   855
      End
      Begin MSForms.Label lblPetition 
         Height          =   255
         Index           =   3
         Left            =   2235
         TabIndex        =   77
         Top             =   1920
         Width           =   3450
         VariousPropertyBits=   27
         Size            =   "6085;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lblPetition 
         Height          =   255
         Index           =   2
         Left            =   2235
         TabIndex        =   76
         Top             =   1650
         Width           =   3450
         VariousPropertyBits=   27
         Size            =   "6085;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label4 
         Caption         =   "申請人："
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   75
         Top             =   1650
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "申請人："
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   74
         Top             =   2205
         Width           =   855
      End
      Begin MSForms.Label lblPetition 
         Height          =   255
         Index           =   4
         Left            =   2235
         TabIndex        =   73
         Top             =   2190
         Width           =   3450
         VariousPropertyBits=   27
         Size            =   "6085;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label9 
         Caption         =   "專利名稱(日) :"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   72
         Top             =   855
         Width           =   1215
      End
      Begin VB.Label Label9 
         Caption         =   "專利名稱(英) :"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   71
         Top             =   570
         Width           =   1215
      End
      Begin VB.Label Label9 
         Caption         =   "專利名稱(中) :"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   70
         Top             =   300
         Width           =   1215
      End
      Begin VB.Label Label12 
         Caption         =   "是否列印通知函：        （N：不印）"
         Height          =   180
         Left            =   120
         TabIndex        =   69
         Top             =   2805
         Width           =   2895
      End
      Begin VB.Label lblCaseFees 
         BackColor       =   &H80000010&
         Height          =   255
         Left            =   8055
         TabIndex        =   222
         Top             =   4170
         Width           =   255
      End
      Begin MSForms.CheckBox chkChoose 
         CausesValidation=   0   'False
         Height          =   255
         Index           =   0
         Left            =   1980
         TabIndex        =   20
         Top             =   3030
         Width           =   3645
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         DisplayStyle    =   4
         Size            =   "6429;450"
         Value           =   "0"
         Caption         =   "                                                        說明書"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   345
      Index           =   2
      Left            =   7644
      TabIndex        =   62
      Top             =   0
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   345
      Index           =   0
      Left            =   5580
      TabIndex        =   60
      Top             =   0
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "回前畫面(&U)"
      CausesValidation=   0   'False
      Height          =   345
      Index           =   1
      Left            =   6420
      TabIndex        =   61
      Top             =   0
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "代理人新案統計(&A)"
      Height          =   345
      Index           =   4
      Left            =   2640
      TabIndex        =   58
      Top             =   0
      Width           =   1700
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "同時發文(&N)"
      Height          =   345
      Index           =   5
      Left            =   4368
      TabIndex        =   59
      Top             =   0
      Width           =   1200
   End
   Begin VB.Label lblTrademarkKind 
      Height          =   255
      Left            =   4170
      TabIndex        =   122
      Top             =   675
      Width           =   1005
   End
   Begin MSForms.Label lblSalesName 
      Height          =   255
      Left            =   4170
      TabIndex        =   121
      Top             =   375
      Width           =   4155
      VariousPropertyBits=   27
      Size            =   "7329;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label lblCaseField 
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   5
      Left            =   3660
      TabIndex        =   120
      Top             =   975
      Width           =   1095
   End
   Begin VB.Label lblCaseField 
      Height          =   255
      Index           =   4
      Left            =   3660
      TabIndex        =   119
      Top             =   375
      Width           =   525
   End
   Begin VB.Label lblCaseField 
      Height          =   255
      Index           =   3
      Left            =   3660
      TabIndex        =   118
      Top             =   675
      Width           =   405
   End
   Begin VB.Label lblCaseField 
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   2
      Left            =   1140
      TabIndex        =   117
      Top             =   975
      Width           =   1005
   End
   Begin VB.Label lblCaseField 
      Height          =   255
      Index           =   1
      Left            =   1140
      TabIndex        =   116
      Top             =   675
      Width           =   1515
   End
   Begin VB.Label lblCaseField 
      Height          =   255
      Index           =   0
      Left            =   1140
      TabIndex        =   115
      Top             =   375
      Width           =   1515
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "智權人員："
      Height          =   180
      Index           =   0
      Left            =   2700
      TabIndex        =   68
      Top             =   375
      Width           =   990
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "本所期限："
      Height          =   180
      Index           =   0
      Left            =   180
      TabIndex        =   67
      Top             =   975
      Width           =   900
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "專利種類："
      Height          =   180
      Index           =   0
      Left            =   2700
      TabIndex        =   66
      Top             =   675
      Width           =   900
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "本所案號："
      Height          =   180
      Index           =   0
      Left            =   180
      TabIndex        =   65
      Top             =   675
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "收文號："
      Height          =   180
      Left            =   180
      TabIndex        =   64
      Top             =   375
      Width           =   720
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "法定期限："
      Height          =   180
      Index           =   1
      Left            =   2700
      TabIndex        =   63
      Top             =   975
      Width           =   900
   End
End
Attribute VB_Name = "frm050102_3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2021/12/4 改成Form2.0 (txtCaseField,chkChoose,lblPetition...)
'Modified by Morgan 2021/12/7 chkChoose 改 Form2.0 後 Value 只能是 Trure 或 False
'Memo By Morgan 2012/12/11 智權人員欄已修改
'Memo by Morgan2010/12/28 申請案號欄已修改
'2010/12/3 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/11 日期欄已修改
Option Explicit

'此本所案號之系統類別，在ReadAllData中傳回真正的系統類別
Dim intCaseKind As Integer
'bolLeave判斷離開時，是否要彈出詢問視窗，回答Yes後改為True 跳下一畫面
Dim bolLeave As Boolean
'cp()存放CaseProgress,field()存放基本資料檔
Dim cp() As String, field() As String
'intLeaveKind離開時，是0:結束  1:回上一畫面
Dim intLeaveKind As Integer
'strCountry存放指定國家
Dim strCountry As String
'strPriority存放優先權
Dim strPriority1 As String, strPriority2 As String, strPriority3 As String, strPriority4 As String, strPriority5 As String
'strAddDeadline存放補件期限
Dim strAddDeadline1 As String, strAddDeadline2 As String, strAddDeadline3 As String
'strAddDeadline存放發明人資料
Dim strInventorNo As String
Dim intInventorCnt As Integer 'Add By Sindy 2014/11/13 發明人筆數

Dim SeekPrint As Integer, SeekPrintL As Integer
'Add By Cheng 2002/08/23
Dim m_strCust1 As String '申請人1
Dim m_strCust2 As String '申請人2
Dim m_strCust3 As String '申請人3
Dim m_strCust4 As String '申請人4
Dim m_strCust5 As String '申請人5
'92.1.12 add by sonia
Dim old_Entity As String   '原大小個體
Dim new_Entity As String   '新大小個體
'Dim m_strMailCP09 As String 'Add by Morgan 2006/7/31    '2010/1/22 CANCEL BY SONIA
'Add by Morgan 2007/10/19
Dim m_NoPicMailSub As String '無代表圖的通知信主旨
Dim m_bActived As Boolean
Dim skMail() As SeekMails     '2010/1/21 ADD BY SONIA
Dim m_iMultiDesign As Integer 'Add by Morgan 2010/4/2 集體設計數
Dim cm(5 To 8) As String 'Add by Amy 2013/08/06
Dim m_AccessCode As String 'Added by Morgan 2014/1/23 電子優先權文件存取碼
'Add by Amy 2014/04/23  for 台灣或大陸有關聯案使用
Dim bolFirstInventor As Boolean '是否第一次進發明人資料
Dim DefInventor As String, OrgInventor As String '預設關聯案發明人/本案發明人
Dim strPD06 As String, bolEPC_pri As Boolean  'Add by Lydia 2015/02/03 EPC新案發文，
Dim strExPD07 As String 'Added by Lydia 2022/04/18 EPC新案發同時有主張優先權，其基礎案的國家不須提供前案資料
'Add By Sindy 2018/1/8
Public m_strIR01 As String
Public m_strIR02 As String
Public m_strIR03 As String
Public m_strIR04 As String
'2018/1/8 END
Dim m_strAF01 As String, m_strLD18 As String 'Added by Morgan 2018/8/21
Dim m_bolEngLetter As Boolean, m_strSubject As String 'Added by Morgan 2018/9/6
Dim m_str416CP81 As String 'Added by Morgan 2019/4/29

Private Sub chkChoose_Click(Index As Integer)
Dim strTemp As String

If Index = 0 Then
   txtCaseField(38).Enabled = chkChoose(Index).Value
   If chkChoose(Index).Value Then
      '91.11.20 MODIFY BY SONIA
      'If objPublicData.GetNation(txtCaseField(12).Text, , strTemp) Then
      '   txtCaseField(38) = strTemp
      'End If
      Select Case txtCaseField(12)
         Case "011"
            txtCaseField(38) = "日文"
         Case "231"
            txtCaseField(38) = "德文"
         Case Else
            txtCaseField(38) = "英文"
      End Select
      '91.11.20 END
   Else
      txtCaseField(38) = ""
   End If
End If
End Sub

Private Sub cmdAddDeadLine_Click()
   ModifyAddDeadline strAddDeadline1, strAddDeadline2, strAddDeadline3
End Sub

Private Sub cmdInventor_Click()
Dim i As Integer, strPetition As String, varInventorNo As Variant

'Add By Cheng 2002/07/30
If intCaseKind <= "4" Then

   For i = 0 To 3
          strPetition = strPetition + txtCaseField(i + 3) + ","
   Next
   strPetition = strPetition + txtCaseField(i + 3)
   
   'Add by Morgan 2010/5/3 若有選過發明人且申請人有變更時提醒
   If PUB_ChkInventor(strInventorNo, strPetition) = False Then
      MsgBox "申請人已變更，請重新點選發明人資料！"
   End If
   'end 2010/5/3
   
   'Modify by Amy 2014/04/23
    If InStr(NewCasePtyList, txtCaseField(10)) > 0 And bolFirstInventor = True Then
        strInventorNo = DefInventor
        bolFirstInventor = False
    End If
    'end 2014/04/23
         
   ModifyInventor strPetition, strInventorNo
   
   'Modify By Sindy 2014/11/13
   If strSrvDate(1) < 專利發明人檔啟用日 Then
   '2014/11/13 END
      If strInventorNo = "" Then
         For i = 0 To 9
                field(60 + i) = ""
         Next
      Else
         varInventorNo = Split(strInventorNo, ",")
         For i = 0 To UBound(varInventorNo)
                field(60 + i) = varInventorNo(i)
         Next
         For i = i + 1 To 9
                field(60 + i) = ""
         Next
      End If
   End If
End If
End Sub

'Added by Morgan 2016/8/15 客戶函的例外欄位抽出來避免重複產生
Private Sub StartLetter1(ByVal ET01 As String, ByVal ET03 As String)
   Dim strTxt(1 To 15) As String, iStep As Integer, strTmp As String
   EndLetter ET01, cp(9), ET03, strUserNum
   
   iStep = 1
   
   strTmp = ""
   If chkChoose(1).Value Then strTmp = "圖式、"
   If chkChoose(2).Value Then strTmp = strTmp & "照片、"
   If chkChoose(0).Value Then strTmp = strTmp & txtCaseField(38) & "說明書"
   If Right(strTmp, 1) = "、" Then strTmp = Left(strTmp, Len(strTmp) - 1)
   '92.3.8 MODIFY BY SONIA
   If CheckStr(strTmp) <> "" Then
'Modify by Morgan 2004/211
'超過1個項目才印"各"字
'      strTxt(iStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         "VALUES ('" & ET01 & "','" & cp(9) & "','" & ET03 & "','" & strUserNum & _
         "','是否有說明書、圖式、申請收據','　　附送本案" & strTmp & "各一份，敬請查存。" & "')"
         
      strTxt(iStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         "VALUES ('" & ET01 & "','" & cp(9) & "','" & ET03 & "','" & strUserNum & _
         "','是否有說明書、圖式、申請收據','　　附送本案" & strTmp & IIf(InStr(strTmp, "、") > 0, "各", "") & "一份，敬請查存。" & "')"
         
      iStep = iStep + 1
   End If
   '92.3.8 END
   
    'Add By Cheng 2003/01/29
    '若指定的語文為中文或日文
    'Modify By Cheng 2003/02/17
    '不論指定的語文為中文或日文, 都抓案件中文名稱
'    If (ET03 = "96" Or ET03 = "97") And (Me.txtCaseField(38).Text = "中文" Or Me.txtCaseField(38).Text = "日文") Then
'        strTxt(iStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'           "VALUES ('" & ET01 & "','" & cp(9) & "','" & ET03 & "','" & strUserNum & _
'           "','專利名稱','" & IIf(Me.txtCaseField(38).Text = "中文", field(5), field(7)) & "')"
'        iStep = iStep + 1
'    End If
    If (ET03 = "96" Or ET03 = "97") And (Me.txtCaseField(38).Text = "中文" Or Me.txtCaseField(38).Text = "日文") Then
        strTxt(iStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
           "VALUES ('" & ET01 & "','" & cp(9) & "','" & ET03 & "','" & strUserNum & _
           "','專利名稱','" & IIf(Me.txtCaseField(38).Text = "中文", field(5), field(5)) & "')"
        iStep = iStep + 1
    End If
    
   
   'Add by Morgan 2011/4/14
   '一併辦理讓渡程序
   If chkChoose(6).Value Then
      strTxt(iStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            "VALUES ('" & ET01 & "','" & cp(9) & "','" & ET03 & "','" & strUserNum & _
            "','同時辦理事項','及辦理讓渡程序')"
      iStep = iStep + 1
   End If
   
   'Added by Morgan 2016/8/15
   If chkChoose(5).Value Then
      strTxt(iStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            "VALUES ('" & ET01 & "','" & cp(9) & "','" & ET03 & "','" & strUserNum & _
            "','一併提檢索及實審','♀')"
      iStep = iStep + 1
   ElseIf chkChoose(3).Value Then
      strTxt(iStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            "VALUES ('" & ET01 & "','" & cp(9) & "','" & ET03 & "','" & strUserNum & _
            "','一併提實審','♀')"
      iStep = iStep + 1
   End If
   
   'end 2016/8/15
   If Not ClsLawExecSQL(iStep - 1, strTxt) Then
      MsgBox "儲存例外欄位失敗，請洽系統管理員 !", vbCritical
   End If
   
End Sub
   
Private Sub StartLetter(ByVal ET01 As String, ByVal ET03 As String)
   'Modified by Lydia 2022/04/18 15=> 30
   Dim strTxt(1 To 30) As String, iStep As Integer, strTmp As String
   EndLetter ET01, cp(9), ET03, strUserNum
   
   'Add by Morgan 2009/5/13
   '指定提申
   If txtCaseField(20) <> "" Then
      strTxt(1) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            "VALUES ('" & ET01 & "','" & cp(9) & "','" & ET03 & "','" & strUserNum & _
            "','有無提申期限','on " & ChgEngDate(DBDATE(txtCaseField(20))) & "')"
   Else
   'end 2009/5/13
      If txtCaseField(14) <> "" Then
         strTxt(1) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            "VALUES ('" & ET01 & "','" & cp(9) & "','" & ET03 & "','" & strUserNum & _
            "','有無提申期限','no later than " & ChgEngDate(DBDATE(txtCaseField(14))) & "')"
      Else
         '韓國,法國,德國 指示提申期限
         'Modify by Morgan 2006/3/27 加土耳其235
         'Modify by Morgan 2006/5/17 加判斷非設計
         'Modify by Morgan 2006/6/26 加荷蘭207
         'Modify by Morgan 2006/9/22 改判斷非英語系國家 -- 禧佩
         'If field(8) <> "3" And InStr("012,203,231,235,207", txtCaseField(12)) > 0 Then
         'Modify by Morgan 2010/1/27 日本案非日文說明書
         'If field(8) <> "3" And InStr(NoneEngCountry, txtCaseField(12)) > 0 Then
         'Modify by Morgan 2010/4/15 日本改為非英語系國家
         'If field(8) <> "3" And (InStr(NoneEngCountry, txtCaseField(12)) > 0 Or (txtCaseField(12) = "011" And txtCaseField(38).Text <> "日文")) Then
         'Modify by Morgan 2010/11/4 +德文
         'If field(8) <> "3" And InStr(NoneEngCountry, txtCaseField(12)) > 0 And Not (txtCaseField(12) = "011" And txtCaseField(38).Text = "日文") Then
         'modify by sonia 2025/11/14 3 weeks改為4 weeks
         If field(8) <> "3" And InStr(NoneEngCountry, txtCaseField(12)) > 0 And Not (txtCaseField(12) = "011" And txtCaseField(38).Text = "日文") And Not (txtCaseField(12) = "231" And txtCaseField(38).Text = "德文") Then
            strTxt(1) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
               "VALUES ('" & ET01 & "','" & cp(9) & "','" & ET03 & "','" & strUserNum & _
               "','有無提申期限','within 4 weeks after receipt')"
         Else
            strTxt(1) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
               "VALUES ('" & ET01 & "','" & cp(9) & "','" & ET03 & "','" & strUserNum & _
               "','有無提申期限','upon receipt')"
         End If
      End If
   End If

   iStep = 2
   '92.1.12 add by sonia
   'Modify by Morgan 2006/9/21 加法國且申請日>=20050901
   'Modified by Morgan 2016/9/22 +印度040,菲律賓030--禧佩IsCustomerIndividual(m_TM23) =
   'modify by sonia 2017/11/21 印度040個人指示信不可印小個體CFP-029895
   'If txtCaseField(12) = "101" Or txtCaseField(12) = "102" Or (txtCaseField(12) = "203" And DBDATE(field(10)) >= "20050901") Or txtCaseField(12) = "040" Or txtCaseField(12) = "030" Then
   'Modified by Morgan 2024/7/30 +EPC微個體
   'Modified by Morgan 2025/5/12 +PCT--尚蓉
   If txtCaseField(12) = "101" Or txtCaseField(12) = "102" Or (txtCaseField(12) = "203" And (field(10) = "" Or DBDATE(field(10)) >= "20050901")) _
      Or (txtCaseField(12) = "040" And IsCustomerIndividual(ChangeCustomerL(txtCaseField(3))) = False) _
      Or txtCaseField(12) = "030" Or (txtCaseField(12) = "221" And field(179) = "2") Or txtCaseField(12) = "056" Then
      
      'Added by Morgan 2013/3/20
      'Modified by Morgan 2024/7/30 微個體改用中文判斷，因為已改存代碼，且不同國家可能會不同。Ex:EPC的微個體是2但美國的微個體是3
      'If optChoose(2).Value = True Then
      If (optChoose(0).Value = True And optChoose(0).Caption = "微個體") _
         Or (optChoose(1).Value = True And optChoose(1).Caption = "微個體") _
         Or (optChoose(2).Value = True And optChoose(2).Caption = "微個體") Then
         strTxt(iStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            "VALUES ('" & ET01 & "','" & cp(9) & "','" & ET03 & "','" & strUserNum & _
            "','微個體敘述','The applicant qualifies as Micro Entity. ')"
         iStep = iStep + 1
      
      'end 2024/7/30
         strTmp = "(Micro Entity)"
         'Added by Morgan 2013/4/10
         strTxt(iStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            "VALUES ('" & ET01 & "','" & cp(9) & "','" & ET03 & "','" & strUserNum & _
            "','微個體才印','♀')"
         iStep = iStep + 1
         'end 2013/4/10
      'end 2013/3/20
      ElseIf optChoose(0).Value = True Then
         strTmp = "(Large Entity)"
      ElseIf optChoose(1).Value = True Then
         strTmp = "(Small Entity)"
      End If
      
      strTxt(iStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         "VALUES ('" & ET01 & "','" & cp(9) & "','" & ET03 & "','" & strUserNum & _
         "','大小個體','" & strTmp & "')"
      iStep = iStep + 1
   End If
   '92.1.12 end

   'Add By Cheng 2002/07/30
   strTmp = ""
   If chkChoose(3).Value Then
      'Added by Morgan 2020/3/10
      If field(9) = "018" Then
         '馬來西亞:內容較多,單獨一行
         strTmp = vbCrLf & vbCrLf & "    The referenced application does not have any corresponding foreign applications in Australia, the United Kingdom, the United States of America, EPC, Japan and the Republic of Korea. Please simultaneously file a request for the normal substantive examination to the referenced application."
      Else
      'end 2020/3/10
         strTmp = "Please simultaneously file a request for substantive examination."
      End If
      
   'Add by Morgan 2004/7/28
   '若為日本發明時，未收文實審也要有敘述
   Else
      'Modify by Morgan 2004/10/18 加德國
      'If txtCaseField(12) = "011" And cp(10) = "101" Then
      '93.11.26 MODIFY BY SONIA
      'If (txtCaseField(12) = "011" Or txtCaseField(12) = "231") And cp(10) = "101" Then
      '   strTmp = "Please do not file a request for substantive examination."
      'End If
      'Modified by Morgan 2020/11/20 +分割 --禧佩 Ex:CFP-032039
      If cp(10) = "101" Or (field(8) = "1" And cp(10) = "307") Then '發明申請
         Select Case field(9)
            'Modified by Morgan 2017/7/3 俄羅斯5/2改國家代碼 "233"->"023"
            Case "040", "011", "012", "018", "030", "019", "042", "205", "223", "231", "201", "219", "234", "221", "253", "023", "118", "117", "102", "126", "114", "015"
               strTmp = "Please do not file a request for substantive examination."
            'Modified by Morgan 2025/5/28 +022汶萊
            Case "014", "022" '新加坡
               If chkChoose(5).Value = False Then strTmp = "Please do not file a request for substantive examination."
            Case Else
         End Select
      End If
      If cp(10) = "102" Or (field(8) = "2" And cp(10) = "307") Then  '新型申請
         Select Case field(9)
            'Modify by Morgan 2006/12/11 加韓國新型
            'Modified by Morgan 2016/1/13 +印尼新型
            Case "042", "219", "118", "117", "126", "114", "012", "017"
               strTmp = "Please do not file a request for substantive examination."
            Case Else
         End Select
      End If
      If cp(10) = "103" Or (field(8) = "3" And cp(10) = "307") Then '設計申請
         Select Case field(9)
            'Modified by Morgan 2025/7/21 越南設計實審為自動會提，實審費用後來已改含於申請費中--禧佩
            'Case "042", "117", "126"
            Case "117", "126"
               strTmp = "Please do not file a request for substantive examination."
            Case Else
         End Select
      End If
      '93.11.26 END
   End If
   If chkChoose(4).Value Then strTmp = strTmp & "Please file the Information Disclosure Statement, which includes : "
   '92.8.21 ADD BY SONIA
   If chkChoose(5).Value Then strTmp = "Please simultaneously file a request for combined Search and Examination Report. "
   '92.8.21 END
   If field(9) = "101" And chkChoose(4).Value = False Then strTmp = strTmp & "Please do not file Information Disclosure Statement without our instruction. "
   'Memo by Lydia 2022/04/18 EPC(221)新案指示信及提供前案資料控管之調整, 定稿別: CFP-01-000-B5
               '目前的指示信修改僅僅是針對EPC的新案指示信，提出EPC新案的同時不會要求提實審，所以已經在前段的段落裡告知代理人不要繳實審費，就是不要提實審的意思，可以不要再重複出現Please do not file a request for substantive examination.
                'IDS跟一併提檢索及實審，都不是EPC案會做的，本來也不會出現在EPC新案指示信裡，
   'end 2022/04/18
   If strTmp <> "" Then
      strTxt(iStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         "VALUES ('" & ET01 & "','" & cp(9) & "','" & ET03 & "','" & strUserNum & _
         "','有無實審','" & strTmp & "')"
      iStep = iStep + 1
   End If
    
   'Add by Morgan 2004/6/4
   Select Case txtCaseField(12)
      'Add by Morgan 2004/7/26
      Case "011"  '日本
         If txtCaseField(10) = "101" Or txtCaseField(10) = "103" Then '發明申請
            strTxt(iStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
               "VALUES ('" & ET01 & "','" & cp(9) & "','" & ET03 & "','" & strUserNum & _
               "','轉機關文件','When you receive the Official Action from the Japanese Patent Office, please simply forward it to us without any analysis and translation." & Chr(13) & "')"
            iStep = iStep + 1
         End If
      'Add by Morgan 2009/11/10
      Case "102" '加拿大
         If txtCaseField(18) <> "Y" Then   '發明人非申請人
            strTxt(iStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
               "VALUES ('" & ET01 & "','" & cp(9) & "','" & ET03 & "','" & strUserNum & _
               "','發明人讓與','Kindly note that the right of inventors passed to the applicant by virtue of an Assignment, and the date of the Assignment is " & ChgEngDate(DBDATE(txtCaseField(11))) & "." & Chr(13) & "')"
            iStep = iStep + 1
         End If
            
      Case "201" '英國
         'Modify by Morgan 2004/7/26
         '設計也要<發明人讓與>敘述
         'Modify by Morgan 2010/8/3 +307 分割
         If txtCaseField(10) = "101" Or txtCaseField(10) = "103" Or txtCaseField(10) = "307" Then '發明申請
            If txtCaseField(18) <> "Y" Then   '發明人非申請人
               strTxt(iStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                  "VALUES ('" & ET01 & "','" & cp(9) & "','" & ET03 & "','" & strUserNum & _
                  "','發明人讓與','Kindly note that the right of inventors passed to the applicant by virtue of an Assignment, and the date of the Assignment is " & ChgEngDate(DBDATE(txtCaseField(11))) & "." & Chr(13) & "')"
               iStep = iStep + 1
            End If
            
            'Modify by Morgan 2010/8/3 +307 分割
            If txtCaseField(10) = "101" Or txtCaseField(10) = "307" Then '發明申請
               If chkChoose(3).Value Then    '一併提實審
                  strTxt(iStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                     "VALUES ('" & ET01 & "','" & cp(9) & "','" & ET03 & "','" & strUserNum & _
                     "','檢索與實審',' Please simultaneously file a request for search and substantive examination.')"
                  iStep = iStep + 1
               Else
                  strTxt(iStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                     "VALUES ('" & ET01 & "','" & cp(9) & "','" & ET03 & "','" & strUserNum & _
                     "','檢索與實審',' Please file a request for search examination but do not file a request for substantive examination.')"
                  iStep = iStep + 1
               End If
            End If
         End If
      Case "203"  '法國
         If txtCaseField(10) = "101" Then '發明申請
            strTxt(iStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
               "VALUES ('" & ET01 & "','" & cp(9) & "','" & ET03 & "','" & strUserNum & _
               "','檢索報告',' Please simultaneously file a request for the search report.')"
            iStep = iStep + 1
         End If
         
      'Add by Morgan 2005/4/11
      Case "206" '奧地利
         strTxt(iStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            "VALUES ('" & ET01 & "','" & cp(9) & "','" & ET03 & "','" & strUserNum & _
            "','德文譯本',' We will send the German translation to you once upon we receive from our representative in Germany.')"
         iStep = iStep + 1
            
      'Add by Morgan 2007/10/4
      Case "208"  '盧森堡
         If txtCaseField(10) = "101" Then '發明申請
            strTxt(iStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
               "VALUES ('" & ET01 & "','" & cp(9) & "','" & ET03 & "','" & strUserNum & _
               "','檢索報告',' Please simultaneously file a request for a prior art search.')"
            iStep = iStep + 1
         End If
         
      Case "231"  '德國
         If Me.txtCaseField(38).Text = "德文" Then
            'Modify by Morgan 2009/11/12 不用了--甄妮
            'strTxt(iStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
               "VALUES ('" & ET01 & "','" & cp(9) & "','" & ET03 & "','" & strUserNum & _
               "','敘述',' Please file this application in Germany after editing the specification. Please send a Copy of the edited version of the specification back to us.')"
            'iStep = iStep + 1
            
            'Removed by Morgan 2024/7/11 刪除--禧佩
            'strTxt(iStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
               "VALUES ('" & ET01 & "','" & cp(9) & "','" & ET03 & "','" & strUserNum & _
               "','附件','" & IIf(ET03 = "29", "6", "5") & ". The disc')"
            'iStep = iStep + 1
            'end 2024/7/11
         Else
            'Modified by Morgan 2014/8/28 改內容(原為 Please simply file this application without amendment.)
            strTxt(iStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
               "VALUES ('" & ET01 & "','" & cp(9) & "','" & ET03 & "','" & strUserNum & _
               "','敘述',' This application is ready for filing without requiring further intensive review. We allow you to make minor changes with respects to formality or the local practices.')"
            iStep = iStep + 1
         End If
      'Add by Morgan 2006/8/10
      'Modified by Morgan 2014/6/13 +046柬埔寨
      'Modified by Morgan 2017/1/10 +042越南--禧佩
      Case "301", "046", "042" '南非
         strTxt(iStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            "VALUES ('" & ET01 & "','" & cp(9) & "','" & ET03 & "','" & strUserNum & _
            "','國際分類','International Classification:" & Chr(13) & "')"
         iStep = iStep + 1
         
      'Add by Morgan 2009/9/21
      Case "213" '葡萄牙
         strTxt(iStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            "VALUES ('" & ET01 & "','" & cp(9) & "','" & ET03 & "','" & strUserNum & _
            "','國籍欄位','Nationality:')"
         iStep = iStep + 1
   End Select
   
   'Add by Morgan 2004/12/30
   '新案(美洲國家案除外)指示信申請人非發明人加<發明人讓與>敘述,英國敘述不同(上面)
   'Modify by Morgan 2009/10/19 +加拿大102
   'If txtCaseField(18) <> "Y" And Left(txtCaseField(12), 1) <> "1" And txtCaseField(12) <> "201" Then
   'Modify by Morgan 2009/11/10 加拿大改和英國一樣要有日期
   'If txtCaseField(18) <> "Y" And txtCaseField(12) <> "201" And (txtCaseField(12) = "102" Or Left(txtCaseField(12), 1) <> "1") Then
   'Modified by Morgan 2017/11/14 assignment->agreement --玫音
   If txtCaseField(18) <> "Y" And txtCaseField(12) <> "201" And Left(txtCaseField(12), 1) <> "1" Then
      'Modified by Morgan 2017/12/1 EPC和德國加日期 --玫音
      If txtCaseField(12) = "231" Or txtCaseField(12) = "221" Then
         strTxt(iStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            "VALUES ('" & ET01 & "','" & cp(9) & "','" & ET03 & "','" & strUserNum & _
            "','發明人讓與','The applicant derives title to the invention from the inventor by agreement, and the date of the agreement is " & ChgEngDate(DBDATE(txtCaseField(11))) & "." & Chr(13) & "')"
      Else
         strTxt(iStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            "VALUES ('" & ET01 & "','" & cp(9) & "','" & ET03 & "','" & strUserNum & _
            "','發明人讓與','The applicant derives title to the invention from the inventor by agreement." & Chr(13) & "')"
      End If
      iStep = iStep + 1
   End If
   
   'Add by Morgan 2010/3/12 西班牙發明若有收文"申請檢索報告"421則指示信要帶
   If field(9) = "211" And cp(10) = "101" Then
      If PUB_ChkCPExist(cp, "421") Then
         strTxt(iStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            " VALUES ('" & ET01 & "','" & cp(9) & "','" & ET03 & "','" & strUserNum & "','有無檢索','♀')"
         iStep = iStep + 1
      End If
   End If
   
   'Add by Morgan 2011/10/24
   If field(9) = "101" Then
      strTxt(iStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         " VALUES ('" & ET01 & "','" & cp(9) & "','" & ET03 & "','" & strUserNum & "','美國案','♀')"
      iStep = iStep + 1
      
      'Added by Morgan 2013/4/18
      '美國發明申請101,CIP113,分割307
      If InStr("101,113,307", txtCaseField(10)) > 0 Then
         strExc(1) = ""
         If txtCaseField(10) = "307" Then
            strExc(0) = "select nvl(pd05,pa10) from divisioncase,patent,pridate where dc01='" & cp(1) & "' and dc02='" & cp(2) & "' and dc03='" & cp(3) & "' and dc04='" & cp(4) & "'" & _
               " and pa01(+)=dc05 and pa02(+)=dc06 and pa03(+)=dc07 and pa04(+)=dc08 and pd01(+)=dc05 and pd02(+)=dc06 and pd03(+)=dc07 and pd04(+)=dc08 and nvl(pd05,pa10)<20130316 and rownum<2"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               strExc(1) = "1"
            End If
         ElseIf txtCaseField(10) = "113" Then
            strExc(0) = "select nvl(pd05,pa10) from patent,pridate where pa01='" & cp(1) & "' and pa02='" & cp(2) & "' and pa03='0' and pa04='" & cp(4) & "'" & _
               " and pd01(+)=pa01 and pd02(+)=pa02 and pd03(+)=pa03 and pd04(+)=pa04 and nvl(pd05,pa10)<20130316 and rownum<2"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               strExc(1) = "1"
            End If
         Else
            strExc(0) = "select pd05 from patent,pridate where pa01='" & cp(1) & "' and pa02='" & cp(2) & "' and pa03='" & cp(3) & "' and pa04='" & cp(4) & "'" & _
               " and pd01(+)=pa01 and pd02(+)=pa02 and pd03(+)=pa03 and pd04(+)=pa04 and pd05<20130316 and rownum<2"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               strExc(1) = "1"
            End If
         End If
         If strExc(1) = "1" Then
            strTxt(iStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
               " VALUES ('" & ET01 & "','" & cp(9) & "','" & ET03 & "','" & strUserNum & "','FITF','♀')"
            iStep = iStep + 1
         End If
      End If
   End If
   
   'Add by Morgan 2010/4/2
   If m_iMultiDesign > 0 Then
      strTxt(iStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         " VALUES ('" & ET01 & "','" & cp(9) & "','" & ET03 & "','" & strUserNum & "','集體','Multiple ')"
      iStep = iStep + 1
      strExc(0) = "    First design: Fig. ~Fig. "
      For intI = 2 To m_iMultiDesign + 1
         strExc(0) = strExc(0) & vbCrLf & "    " & GetEngNum(intI) & " design: Fig. ~Fig. "
      Next
      strTxt(iStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         " VALUES ('" & ET01 & "','" & cp(9) & "','" & ET03 & "','" & strUserNum & "','集體說明','" & strExc(0) & "')"
      iStep = iStep + 1
   End If
   
   'Add by Morgan 2010/8/12
   '美國發明案若收-1時加主張暫時申請案優先權
   If field(9) = "101" And field(3) = "1" And cp(10) = "101" Then
      strTxt(iStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
               "VALUES ('" & ET01 & "','" & cp(9) & "','" & ET03 & "','" & strUserNum & _
               "','有暫時申請案不印','♀')"
      iStep = iStep + 1
      
      strExc(0) = "select pa10,pa11 from patent where pa01='" & cp(1) & "'" & _
         " and pa02='" & cp(2) & "' and pa03='0' and pa04='" & cp(4) & "'"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         strTxt(iStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
               "VALUES ('" & ET01 & "','" & cp(9) & "','" & ET03 & "','" & strUserNum & _
               "','暫時申請日','" & RsTemp("pa10") & "')"
         iStep = iStep + 1
         strTxt(iStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
               "VALUES ('" & ET01 & "','" & cp(9) & "','" & ET03 & "','" & strUserNum & _
               "','暫時申請號','" & RsTemp("pa11") & "')"
         iStep = iStep + 1
      End If
   End If
   'end 2010/8/12
   
   '2012/10/3 ADD BY SONIA X63219國立中正大學的新案指示信要加一段只能一次請款
   '2013/3/12 MODIFY BY SONIA 加入X6383801中國醫藥大學
   'MODIFY BY SONIA 2014/3/28 顏永堅3/25郵件所列大學及其關係企業都要加
   'modify by sonia 2017/4/27 X60149改不含關係企業,另婉莘於4/10要求加X69534彩豐精技,4/14雅娟要求X69011010
   'modify by sonia 2017/5/15陳德發及郭雅娟要求取消X69534彩豐精技
   'modify by sonia 2017/10/11 張詠翔要求取消X6014900
   'modify by sonia 2017/10/26 郭雅娟要求加X60738010國立清華大學
   'modify by sonia 2020/7/20 茹曣加X79919000
   'modify by sonia 2021/3/17 顧服組黃教威加5編號X69365010(長庚醫療財團法人嘉義長庚紀念醫院),X54243070(國立台灣大學),X80847020(財團法人國家實驗研究院),X83983000(盧彥蓓),X83984000(林致廷)
   If Left(ChangeCustomerL(txtCaseField(3)), 6) = "X44551" Or Left(ChangeCustomerL(txtCaseField(3)), 6) = "X62079" & _
      Left(ChangeCustomerL(txtCaseField(3)), 8) = "X6901101" Or Left(ChangeCustomerL(txtCaseField(3)), 8) = "X6073801" Or Left(ChangeCustomerL(txtCaseField(3)), 8) = "X7991900" & _
      Left(ChangeCustomerL(txtCaseField(3)), 6) = "X63219" Or Left(ChangeCustomerL(txtCaseField(3)), 6) = "X60498" & _
      Left(ChangeCustomerL(txtCaseField(3)), 6) = "X62319" Or Left(ChangeCustomerL(txtCaseField(3)), 6) = "X43988" & _
      Left(ChangeCustomerL(txtCaseField(3)), 6) = "X62702" Or Left(ChangeCustomerL(txtCaseField(3)), 6) = "X63838" & _
      Left(ChangeCustomerL(txtCaseField(3)), 8) = "X6936501" Or Left(ChangeCustomerL(txtCaseField(3)), 8) = "X5424307" Or Left(ChangeCustomerL(txtCaseField(3)), 8) = "X8084702" & _
      Left(ChangeCustomerL(txtCaseField(3)), 8) = "X8398300" Or Left(ChangeCustomerL(txtCaseField(3)), 8) = "X8398300" Then
      If InStr(NewCasePtyList, txtCaseField(10)) > 0 Then
         strTxt(iStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            "VALUES ('" & ET01 & "','" & cp(9) & "','" & ET03 & "','" & strUserNum & _
            "','中正大學一次請款','Due to the institutional limits on payment of our client, please list all the related charges for one procedure on one invoice and DO NOT separate them in different invoices. Please be aware of said policy when you handle all the future invoices for the referenced application." & Chr(13) & "')"
         iStep = iStep + 1
      End If
   End If
   '2012/10/3 END
   
   'Add By Sindy 2014/4/21
   If field(161) = "J" Then
      'Modified by Morgan 2014/8/1
      'strExc(0) = "The referenced application is entrusted to you in the name of our new firm, Tai E Intellectual Property Co., Ltd., operating from the same premises and with the same contact information and staff as Tai E International Patent and Law Office. Please issue all of your invoices for this application in the title of Tai E Intellectual Property Co., Ltd." & Chr(13)
      'Modified by Lydia 2016/02/19
      'strExc(0) = "The referenced application is entrusted to you in the name of our new firm, |#(粗體,底線)Tai E Intellectual Property Co., Ltd.#|, operating from the same premises and in association with Tai E International Patent and Law Office with the same contact information and staff. |#(底線)Please issue all of your invoices for this application in the title of Tai E Intellectual Property Co., Ltd#|." & vbCrLf
      strExc(0) = "The referenced application is entrusted to you in the name of our subsidiary company, |#(粗體,底線)Tai E Intellectual Property Co., Ltd.#|, operating from the same premises and in association with Tai E International Patent and Law Office with the same contact information and staff. |#(底線)Please issue all of your invoices for this application in the title of Tai E Intellectual Property Co., Ltd#|." & vbCrLf
      strTxt(iStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         "VALUES ('" & ET01 & "','" & cp(9) & "','" & ET03 & "','" & strUserNum & _
         "','開智權帳單提醒','" & ChgSQL(strExc(0)) & "')"
      'end 2014/8/1
      iStep = iStep + 1
   End If
   '2014/4/21 END
   
   'Added by Morgan 2014/1/23
   If m_AccessCode <> "" Then
      strTxt(iStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         "VALUES ('" & ET01 & "','" & cp(9) & "','" & ET03 & "','" & strUserNum & _
         "','存取碼','" & ChgSQL(m_AccessCode) & "')"
      iStep = iStep + 1
   End If
    'Add by Lydia 2015/02/03 EPC新案發文，同時有主張優先權=>優先權案為新型案,加註記。
    'Memo by Lydia 2020/06/08 優先權為新型案，且該所主張之優先權國家新型無實審制，加註記。
   If txtCaseField(12) = "221" And bolEPC_pri = True And strPD06 <> "" Then
      strTxt(iStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         "VALUES ('" & ET01 & "','" & cp(9) & "','" & ET03 & "','" & strUserNum & _
         "','EPC新型優先權案-案號','" & ChgSQL(strPD06) & "')"
      iStep = iStep + 1
   End If
   'Added by Lydia 2022/04/18 EPC新案發同時有主張優先權，其基礎案的國家不須提供前案資料
                   '調整設定：若優先權基礎案的國家為奧地利、丹麥、日本、中國大陸、韓國、西班牙、瑞典、瑞士、英國、美國，則新案發文時不在下一程序掛提供前案的期限(basUpdate. PUB_UpdExamDate)，並在新案指示信帶出相關段落。
   If txtCaseField(12) = "221" And strExPD07 <> "" Then
      strTxt(iStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         "VALUES ('" & ET01 & "','" & cp(9) & "','" & ET03 & "','" & strUserNum & "','EPC新案發文不須提供前案','♀')"
      iStep = iStep + 1
   End If
   'end 2022/04/18
   
   'edit by nickc 2007/02/05 不用 dll 了
   'If Not objLawDll.ExecSQL(iStep - 1, strTxt) Then
   If Not ClsLawExecSQL(iStep - 1, strTxt) Then
      MsgBox "儲存例外欄位失敗，請洽系統管理員 !", vbCritical
   End If
End Sub
'Add by Morgan 2010/4/2
'數字轉英文
Private Function GetEngNum(iNum As Integer) As String
   Select Case iNum
      Case 1
         GetEngNum = "First"
      Case 2
         GetEngNum = "Second"
      Case 3
         GetEngNum = "Third"
      Case Else
         GetEngNum = iNum & "th"
   End Select
End Function

Private Sub cmdOK_Click(Index As Integer)
Dim stLetter As String 'Add by Morgan 2004/9/27
Dim i As Integer
Dim bolChk As Boolean, strTmp As String
'Add By Sindy 2011/1/26
Dim strApplID As String
Dim rsAddrNotAlike As New ADODB.Recordset
'2011/1/26 End
   
   Select Case Index
      Case 0, 5 '確定, 同時發文
         'Modified by Lydia 2014/12/27 + (DHL列印)
          If txtCaseField(9) = "Y" And txtCaseField(69) = "Y" Then
            MsgBox "TNT及DHL只能擇一列印！"
            Exit Sub
          End If
         'Add by Morgan 2009/6/1
         If PUB_ChkFileNP(cp(9)) Then
            MsgBox "下一程序已有提申或收達期限，不可發文！"
            Exit Sub
         End If
         
         'Add by Amy 2014/04/23 +P案(台灣或大陸)若關聯案發明人與畫面上發明人不同時詢問是否更新
         If DefInventor <> "" And DefInventor <> strInventorNo Then
            If MsgBox("發明人與" & cm(5) & "-" & cm(6) & "(關聯案)不同,請確認是否要存檔??", vbYesNo + vbDefaultButton2) = vbNo Then
                Exit Sub
            End If
         End If
         'end 2014/04/23
   
         'Add by Morgan 2007/3/14 多國案若有其他相同案已核准、發證、公開時不可分案、發文、提申
         If cp(27) = "" And (txtCaseField(10) = "101" Or txtCaseField(10) = "102" Or txtCaseField(10) = "103") Then
            If PUB_SameCaseCheck(cp) = False Then
               Exit Sub
            End If
         End If
         'end 2007/3/14
         
         'Modified by Morgan 2016/9/22 +印度040,菲律賓030--禧佩
         If txtCaseField(12) = "101" Or txtCaseField(12) = "102" Or (txtCaseField(12) = "203" And DBDATE(field(10)) >= "20050901") Or txtCaseField(12) = "040" Or txtCaseField(12) = "030" Then
            If optChoose(0).Enabled Or optChoose(1).Enabled Or optChoose(2).Enabled Then 'Added by Morgan 2025/7/21
               'Modified by Morgan 2013/3/20 +微個體
               If Not optChoose(0).Value And Not optChoose(1).Value And Not optChoose(2).Value Then
                  'Modified by Morgan 2023/3/24
                  If optChoose(2).Enabled = True Then
                     MsgBox "請選擇" & optChoose(0).Caption & "、" & optChoose(1).Caption & "或" & optChoose(2).Caption & "資料 !", vbCritical
                  Else
                     MsgBox "請選擇" & optChoose(0).Caption & "或" & optChoose(1).Caption & "資料 !", vbCritical
                  End If
                  Exit Sub
               End If
            End If
         End If
         
         'Add by Morgan 2008/4/17
         If CheckCP44 = False Then
            Combo1.SetFocus
            Exit Sub
         End If
         Screen.MousePointer = vbHourglass
         For i = 0 To 37
            If i <> 13 Then
               'Add by Morgan 2004/5/26
               If i = 19 Or i = 20 Then
                  '跳過移掉的欄位
               ElseIf txtCaseField(i).Enabled Then
                  If CheckKeyIn(i) <> 1 Then
                     txtCaseField(i).SetFocus
                     txtCaseField_GotFocus (i)
                     Exit For
                  End If
               End If
            'Add By Cheng 2002/03/08
            Else
               If CheckKeyIn(i) <> 1 Then
                  Me.Combo1.SetFocus
                  Exit For
               End If
            End If
         Next
         If cmdCountry.Visible And strCountry = "" Then
            ShowMsg MsgText(9180)
         End If
         If i = 38 Then
            '2008/9/18 add by sonia
            'Modify by Amy 2013/08/06 +若cm5為空才顯示未輸入發明人資料是否輸入的訊息
            'If strInventorNo = "" And InStr(NewCasePtyList, txtCaseField(10)) > 0 Then
            If strInventorNo = "" And InStr(NewCasePtyList, txtCaseField(10)) > 0 And cm(5) = "" Then
               If MsgBox("未輸入發明人資料，請問是否要輸入？", vbCritical + vbYesNo + vbDefaultButton1) = vbYes Then
                  Screen.MousePointer = vbDefault
                  Exit Sub
               End If
            End If
            '2008/9/18 end
            
            'Add by Lydia 2015/02/03 EPC新案發文，同時有主張優先權
            strPD06 = "": bolEPC_pri = False
            strExPD07 = "" 'Added by Lydia 2022/04/18
            If txtCaseField(12) = "221" And InStr("101,102,103", txtCaseField(10)) > 0 Then
               'Modified by Lydia 2020/06/08
               'strExc(0) = " select * from pridate where pd01='" & field(1) & "' and pd02='" & field(2) & "' and pd03='" & field(3) & "' and pd04='" & field(4) & "' "
               strExc(0) = "select pd05,pd06,pd07,pd08,decode(pd08,'1',na26,'2',na28,'3',na30,null) 實審起算日 " & _
                                "from pridate,nation where pd01='" & field(1) & "' and pd02='" & field(2) & "' and pd03='" & field(3) & "' and pd04='" & field(4) & "' and pd07=na01(+) " & _
                                "order by pd08 "
               intI = 1
                Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                If intI = 1 Then
                    bolEPC_pri = True
                    RsTemp.MoveFirst
                    'Modified by Lydia 2020/06/08 若主張多筆優先權，且基礎案有發明案也有新型案，則在輸入案件提申時，下一程序掛「提供前案」，期限同實審。=>比照發明案
                    '2.若優先權為新型案，且該所主張之優先權國家新型無實審制，指示信請加一段如附件。(維持此項設定現行定稿)
                    'Do While Not RsTemp.EOF
                       'If "" & RsTemp!PD08 = "2" Then  '優先權案為新型案
                       '   strPD06 = Trim("" & RsTemp!pd06)
                       'End If
                    '   RsTemp.MoveNext
                    'Loop
                    If "" & RsTemp.Fields("pd08") = "2" And Val("" & RsTemp.Fields("實審起算日")) = 0 Then
                        strPD06 = Trim("" & RsTemp!pd06)
                    End If
                    'end 2020/06/08
                    'Added by Lydia 2022/04/18 EPC新案發同時有主張優先權，其基礎案的國家不須提供前案資料
                    Do While Not RsTemp.EOF
                        If "" & RsTemp.Fields("pd07") <> "" And InStr(cntEPC新案發文不須提供前案, "" & RsTemp.Fields("pd07")) > 0 Then
                            strExPD07 = strExPD07 & "," & RsTemp.Fields("pd07")
                        End If
                        RsTemp.MoveNext
                    Loop
                    If strExPD07 <> "" Then strExPD07 = Mid(strExPD07, 2)
                    'end 2022/04/18
                End If
            End If
            'end 2015/02/03
            
            'Added by Morgan 2014/1/23
            'Modified by Morgan 2014/2/25 設計不用
            'Modified by Morgan 2018/8/10 從存檔後移上來，否則會無法選擇取消去查資料 CFP-029613--玫音
            m_AccessCode = ""
            If txtCaseField(17) <> "N" And field(46) <> "Y" And txtCaseField(12) = "011" Then
               If strPriority1 <> "" And field(8) <> "3" Then
                  If ChkIsDASCountry(strPriority1) = True Then 'Added by Morgan 2020/10/7
                     Do
                        m_AccessCode = InputBox("請輸入電子優先權文件存取碼!!", , "存取碼")
                        If m_AccessCode = "" Then Screen.MousePointer = vbDefault: Exit Sub
                     Loop While (m_AccessCode = "存取碼")
                     
                  End If 'Added by Morgan 2020/10/7
               End If
            End If
            'end2014/1/23
            
            'Add By Cheng 2002/05/22
            '重新檢查欄位有效性
            If TxtValidate = False Then Screen.MousePointer = vbDefault: Exit Sub
            If SaveDatabase Then
               'Add by Morgan 2008/2/20 檢查代理人Email(需考慮可能為FF案件)
               PUB_CheckEMail cp(44), cp(116)
               If field(1) = "CFP" Then
                  PUB_CheckEMail field(75), field(144)
                  If field(145) <> "" Then
                     PUB_CheckEMail field(75), field(145)
                  End If
               Else
                  PUB_CheckEMail field(26), field(76)
                  If field(77) <> "" Then
                     PUB_CheckEMail field(26), field(77)
                  End If
               End If
               'end 2008/2/20
               
               'Add by Morgan 2006/7/31
               '2010/1/21 MODIFY BY SONIA 發E-Mail給承辦人
               'If m_strMailCP09 <> "" Then
               '   MailToPromoter m_strMailCP09
               'End If
               BatchMail
               '2010/1/22 END
               'end 2006/7/31
               
               'Add by Morgan 2007/10/22 若無代表圖則發Mail通知郭雅娟
               If m_NoPicMailSub <> "" Then
                  PUB_SendMail strUserNum, "79075", "", m_NoPicMailSub, " "
               End If
               
         
               strTmp = "30"
               If txtCaseField(17) <> "N" Then '指示信
                  If txtCaseField(16) = "Y" Then
                     bolChk = True
                  Else
                     bolChk = False
                  End If
                  
                  'Add by Morgan 2009/7/30
                  'PCT 進各國指示信改單一版面
                  If field(46) = "Y" And (txtCaseField(10) = "101" Or txtCaseField(10) = "102") Then
                     'Modify by Morgan 2009/10/21 PCT進日本時申請人印中文
                     If txtCaseField(12) = "011" Then
                        strTmp = "49"
                     Else
                        strTmp = "50"
                        If txtCaseField(12) = "221" Then strTmp = "48"   '2013/4/29 add by sonia 慧汶說EPC之PCT案的指示信 the national stage改為the reginal stage
                     End If
                  Else
                     
                     Select Case txtCaseField(12)
                        Case "011" '日本
                           'Remove by Morgan 2011/10/24 PCT案已改用統一版面,舊定稿刪除否則會增加維護負擔
                           'If field(46) = "Y" Then
                           '   strTmp = "32"   'PCT
                           'Else
                              strTmp = "01"
                              'Add by Morgan 2007/5/31
                              If txtCaseField(38) = "中文" Then
                                 strTmp = "B4"
                              'Add by Morgan 2010/1/19
                              ElseIf txtCaseField(38) = "英文" Then
                                 strTmp = "B9"
                              End If
                           'End If
                           
                        Case "012" '韓國
                           'Remove by Morgan 2011/10/24 PCT案已改用統一版面,舊定稿刪除否則會增加維護負擔
                           'If field(46) = "Y" Then
                           '   strTmp = "33"   'PCT
                           'Else
                              If txtCaseField(38) = "中文" Then
                                 strTmp = "02"   '中文送件
                              Else
                                 strTmp = "37"   '英文送件
                              End If
                           'End If
                        Case "014" '新加坡
                           strTmp = "03"
                           'Remove by Morgan 2011/10/24 PCT案已改用統一版面,舊定稿刪除否則會增加維護負擔
                           ''Add by Morgan 2008/8/22
                           'If field(46) = "Y" Then
                           '   strTmp = "B7" 'PCT
                           'End If
                        Case "015" '澳洲
                           strTmp = "04"
                        'Add by Morgan 2005/1/3
                        Case "016" '紐西蘭
                           strTmp = "38"
                        Case "017" '印尼
                           strTmp = "05"
                        Case "018" '馬來西亞
                           strTmp = "06"
                        Case "019" '泰國
                           strTmp = "07"
                        Case "021" '沙烏地阿拉伯
                           strTmp = "08"
                        
                        'Added by Morgan 2025/5/28
                        Case "022" '汶萊
                           If txtCaseField(10) = "101" Then
                              strTmp = "03"
                           End If
                           
                        Case "025" '伊朗
                           strTmp = "B3"
                        Case "030" '菲律賓
                           strTmp = "09"
                        Case "031" '斯里蘭卡
                           strTmp = "10"
                        Case "038" '巴基斯坦
                           strTmp = "11"
                        Case "040" '印度
                           strTmp = "12"
                        Case "042" '越南
                           strTmp = "13"
                        'Added by Morgan 2012/2/6
                        Case "056" 'PCT
                           If txtCaseField(10) = "109" Then
                              strTmp = "85"
                           End If
                           
                        Case "101" '美國
                           If txtCaseField(10) = "118" Then
                              strTmp = "86"
                           Else
                              strTmp = "35"
                              'Remove by Morgan 2011/10/24 PCT案已改用統一版面,舊定稿刪除否則會增加維護負擔
                              ''Add by Morgan 2008/8/22
                              'If field(46) = "Y" Then
                              '   strTmp = "B6" 'PCT
                              'End If
                           End If
                        Case "102" '加拿大
                           strTmp = "14"
                        Case "104" '墨西哥
                           strTmp = "15"
                        Case "117" '巴西
                           strTmp = "16"
                        'Add by Morgan 2007/1/17
                        Case "118" '阿根廷
                           strTmp = "36"
                        Case "201" '英國
                           strTmp = "17"
                           'Remove by Morgan 2011/10/24 PCT案已改用統一版面,舊定稿刪除否則會增加維護負擔
                           ''Add by Morgan 2008/8/22
                           'If field(46) = "Y" Then
                           '   strTmp = "B5" 'PCT
                           'End If
                        Case "203" '法國
                           strTmp = "18"
                        Case "204" '義大利
                           strTmp = "19"
                        'Add by Morgan 2005/4/28
                        Case "206" '奧地利
                           strTmp = "40"
                        Case "207" '荷蘭
                           strTmp = "20"
                        Case "210" '荷比盧
                           strTmp = "21"
                        Case "211" '西班牙
                           strTmp = "22"
                        Case "214" '瑞典
                           strTmp = "23"
                        Case "215" '挪威
                           strTmp = "24"
                        Case "216" '丹麥
                           strTmp = "25"
                        Case "217" '芬蘭
                           strTmp = "26"
                        Case "221" 'EPC
                           'Remove by Morgan 2011/10/24 PCT案已改用統一版面,舊定稿刪除否則會增加維護負擔
                           'If field(46) = "Y" Then
                           '   strTmp = "34"   'PCT
                           'Else
                              'Modified by Morgan 2021/11/3 與發文通用定稿重複，改B5
                              'strTmp = "27"
                              strTmp = "B5"
                              'end 2021/11/3
                           'End If
                        Case "223" '捷克
                           strTmp = "28"
                        Case "231" '德國
                           strTmp = "29"
                           'Add by Morgan 2006/4/26 勾德文說明書時定稿不同
                           If Me.txtCaseField(38).Text = "德文" Then
                              strTmp = "B2"
                           End If
                        'Add by Morgan 2005/6/22
                        'Modified by Morgan 2017/7/3 俄羅斯5/2改國家代碼 "233"->"023"
                        Case "023" '俄羅斯
                           strTmp = "41"
                        'Add by Morgan 2006/3/27
                        'Modified by Morgan 2017/6/21
                        'Case "253" '土耳其
                        Case "235" '土耳其
                        'end 2017/6/21
                           strTmp = "B1"
                        Case "239" '歐盟
                           strTmp = "31"
                        Case "301" '南非
                           strTmp = "30"
                     End Select
                     '無主張優先權
                     'Modified by Morgan 2013/4/19 CIP113,分割307優先權跟隨母案不必另外主張故指示信都不用帶--郭
                     'Memo by Morgan 2021/2/3 CIP113,分割307改也要帶優先權主張但無需附證明(定稿內控制)--郭
                     'If strPriority1 = "" Then
                     If strPriority1 = "" And (txtCaseField(10) <> "113" And txtCaseField(10) <> "307") Then
                        'Modify by Morgan 2005/4/28
                        '處理狀況已不夠用,<39的+50,<45的+6
                        'strTmp = Val(strTmp) + 50
                        If strTmp < "39" Then
                           strTmp = Val(strTmp) + 50
                        ElseIf strTmp < "45" Then
                           strTmp = Val(strTmp) + 6
                        'Add by Morgan 2006/3/27 B->C
                        ElseIf Left(strTmp, 1) = "B" Then
                           strTmp = "C" & Mid(strTmp, 2)
                        End If
                     End If
                     
                  End If
               
                  StartLetter "01", strTmp

                  'Add by Morgan 2004/9/27
                  '加是否印傳真封面選項
                  If txtCaseField(19) <> "N" Then
                     If Me.txtCaseField(16).Text = "Y" Then
                        NowPrint cp(9), "01", "99", False, strUserNum, , , True, stLetter, , , , , , , , , m_strAF01
                     Else
                        NowPrint cp(9), "01", "99", False, strUserNum, 0, , , , , , , , , , , , m_strAF01
                        stLetter = ""
                     End If
                     If m_strAF01 <> "" Then Sleep 1000 '等1秒以確保letterdemand不會發生dupe錯誤 Added by Morgan 2018/8/20
                  End If
                  NowPrint cp(9), "01", strTmp, IIf(Me.txtCaseField(16).Text = "Y", True, False), strUserNum, 0, stLetter, , , , , , , , , , , m_strAF01
                  
                  'Added by Morgan 2018/8/21 CFP電子化
                  If txtCaseField(16).Text = "Y" And m_strAF01 <> "" Then
                     frm1105_1.m_RecNo = m_strAF01
                     frm1105_1.m_PdfName = PUB_CaseNo2FileName(cp(1), cp(2), cp(3), cp(4)) & "." & txtCaseField(10) & ".DATA.PDF"
                     frm1105_1.Show
                     If txtCaseField(13).Text = "Y" Then
                        MsgBox "指示信編輯中，客戶函請至定稿維護修改！", vbExclamation
                        txtCaseField(13).Text = ""
                     End If
                  End If
                  'end 2018/8/21
                  
               'Added by Morgan 2018/9/6
               ElseIf m_bolEngLetter Then
                  PUB_SendOrderLetterP m_strAF01, m_strSubject
               'end 2018/9/6
               
               End If
            
               If txtCaseField(8) <> "N" Then '通知函
                  'Modify by Morgan 2009/7/22 有無優先權定稿合併
                  'If strPriority1 <> "" Then '有優先權
                     If txtCaseField(12) = 美國國家代號 Then '美國有優先權
                        strTmp = "91"
                     Else ' 一般有優先權
                        strTmp = "93"
                        '若有中文或日文副本
                        If Me.txtCaseField(38).Text = "中文" Or Me.txtCaseField(38).Text = "日文" Then
                            strTmp = "96"
                        End If
                     End If
                  'Else '無優先權
                  '   If txtCaseField(12) = 美國國家代號 Then '美國無優先權
                  '      strTmp = "92"
                  '   Else '一般無優先權
                  '      strTmp = "94"
                  '      '若有中文或日文副本
                  '      If Me.txtCaseField(38).Text = "中文" Or Me.txtCaseField(38).Text = "日文" Then
                  '          strTmp = "97"
                  '      End If
                  '   End If
                  'End If
                  '92.2.18 ADD BY SONIA
                  If field(46) = "Y" Then    'PCT
                     'If strPriority1 <> "" Then '有優先權
                        strTmp = "98"
                        If txtCaseField(12) = "221" Then strTmp = "97"   '2013/4/29 add by sonia 慧汶說EPC之PCT案發文定稿 EPC國家階段 改為 歐洲區域專利階段
                     'Else
                     '   strTmp = "95"           '無優先權
                     'End If
                  End If
                  '92.2.18 END
                  'end 2009/7/22
                  
                  'Romove by Morgan 2006/8/18 改用系統例外欄位"<客戶案號或專利種類/CFP>"替代
'                  'Add by Morgan 2004/6/28
'                  '若有客戶案件案號則本所案號印客戶案件案號不印申請國家專利種類
'                  If field(48) <> "" Then
'                     strTmp = "A" & Right(strTmp, 1)
'                  End If
'                  'End
                  'end 2006/8/18
                  
                  'Modified by Morgan 2016/8/15
                  'StartLetter "01", strTmp
                  StartLetter1 "01", strTmp
                  'end 2016/8/15
                  NowPrint cp(9), "01", strTmp, IIf(Me.txtCaseField(13).Text = "Y", True, False), strUserNum, 0, , , , , , , , , , , , m_strLD18
                  
                  'Added by Morgan 2018/8/21 CFP電子化
                  If txtCaseField(13).Text = "Y" And m_strLD18 <> "" Then
                     frm1105_1.m_RecNo = m_strLD18
                     frm1105_1.m_PdfName = PUB_CaseNo2FileName(cp(1), cp(2), cp(3), cp(4)) & "." & txtCaseField(10) & ".CUS.PDF"
                     frm1105_1.Show
                  End If
                  'end 2018/8/21
               End If
               
               '910816 Sieg 303
               'PrtGreenPaper  92.1.21 cancel by sonia 改用開窗定稿
               
               'Add By Sindy 2011/1/26 檢查相同國家若有舊案申請地址與客戶目前申請地址不同者
               strApplID = ""
               If Trim(txtCaseField(3)) <> "" Then '申請人1
                  If strApplID <> "" Then strApplID = strApplID & ","
                  strApplID = strApplID & "'" & Trim(txtCaseField(3)) & "'"
               End If
               If Trim(txtCaseField(4)) <> "" Then '申請人2
                  If strApplID <> "" Then strApplID = strApplID & ","
                  strApplID = strApplID & "'" & Trim(txtCaseField(4)) & "'"
               End If
               If Trim(txtCaseField(5)) <> "" Then '申請人3
                  If strApplID <> "" Then strApplID = strApplID & ","
                  strApplID = strApplID & "'" & Trim(txtCaseField(5)) & "'"
               End If
               If Trim(txtCaseField(6)) <> "" Then '申請人4
                  If strApplID <> "" Then strApplID = strApplID & ","
                  strApplID = strApplID & "'" & Trim(txtCaseField(6)) & "'"
               End If
               If Trim(txtCaseField(7)) <> "" Then '申請人5
                  If strApplID <> "" Then strApplID = strApplID & ","
                  strApplID = strApplID & "'" & Trim(txtCaseField(7)) & "'"
               End If
               If ChkOCaseAndCAddrNotAlike(strApplID, txtCaseField(12), cp(1), txtCaseField(10), rsAddrNotAlike, False) = True Then
                  Set frm880018.fmParent = Me
                  Set frm880018.RsTemp = rsAddrNotAlike
                  frm880018.m_Appl1 = Trim(Me.txtCaseField(3).Text)
                  frm880018.m_Appl2 = Trim(Me.txtCaseField(4).Text)
                  frm880018.m_Appl3 = Trim(Me.txtCaseField(5).Text)
                  frm880018.m_Appl4 = Trim(Me.txtCaseField(6).Text)
                  frm880018.m_Appl5 = Trim(Me.txtCaseField(7).Text)
                  frm880018.Show vbModal
               End If
               '2011/1/26 End
               
               bolLeave = True
               intLeaveKind = 1
               'Add By Cheng 2002/04/30
               '若有未發文資料顯示警告
               PUB_GetCPunIssueDatas "" & Me.lblCaseField(1).Caption
                'Add By Cheng 2003/11/27
                ' 發文回前畫面時
                Select Case Index
                   Case 0:
                        ' 90.07.12 modify by louis (回發文主畫面並清除畫面)
                        'Add By Sindy 2013/5/28
                        If frm050102_1.bolIsEMPFlow = True Then
                           intLeaveKind = 0
                           'Unload frm050102_1
                           frm090202_4.Show
                           frm090202_4.QueryData
                        '2013/5/28 End
                        'Add By Sindy 2018/1/8
                        ElseIf Me.m_strIR01 <> "" Then
                           intLeaveKind = 0
                           'Modify By Sindy 2022/5/20
                           'frm04010519.GoNext
                           Forms(0).Tmpfrm04010519.GoNext
                           Set Forms(0).Tmpfrm04010519 = Nothing
                           '2022/5/20 END
                        '2018/1/8 END
                        Else
                           frm050102_1.Show
                           frm050102_1.Clear
                        End If
                   Case 5:
                        '若尚有未發文資料
                        If PUB_ChkUnissueDatas(Me.lblCaseField(1).Caption) = True Then
                            ' 90.07.12 modify by louis (回發文主畫面並重新查詢)
                           'Add By Sindy 2013/5/28
                           If frm050102_1.bolIsEMPFlow = True Then
                              frm090202_4.QueryData
                           'End If
                           '2013/5/28 End
                           'Add By Sindy 2018/1/8
                           ElseIf Me.m_strIR01 <> "" Then
                              'intLeaveKind = 0
                              'Modify By Sindy 2022/5/20
                              'frm04010519.GoNext
                              Forms(0).Tmpfrm04010519.GoNext
                              Set Forms(0).Tmpfrm04010519 = Nothing
                              '2022/5/20 END
                           '2018/1/8 END
                           End If
                           frm050102_1.Show
                           frm050102_1.ReQuery
                        '若無未發文資料
                        Else
                            ' 90.07.12 modify by louis (回發文主畫面並清除畫面)
                           'Add By Sindy 2013/5/28
                           If frm050102_1.bolIsEMPFlow = True Then
                              intLeaveKind = 0
                              'Unload frm050102_1
                              frm090202_4.Show
                              frm090202_4.QueryData
                           '2013/5/28 End
                           'Add By Sindy 2018/1/8
                           ElseIf Me.m_strIR01 <> "" Then
                              intLeaveKind = 0
                              'Modify By Sindy 2022/5/20
                              'frm04010519.GoNext
                              Forms(0).Tmpfrm04010519.GoNext
                              Set Forms(0).Tmpfrm04010519 = Nothing
                              '2022/5/20 END
                           '2018/1/8 END
                           Else
                              frm050102_1.Show
                              frm050102_1.Clear
                           End If
                        End If
                End Select
                'End
               Unload Me
            '911202 nick
            Else
                MsgBox "存檔失敗, 請洽電腦中心人員!!!", vbExclamation + vbOKOnly
            End If
         End If
         Screen.MousePointer = vbDefault
      Case 1, 2
         'Add by Amy 2014/04/23 +P案(台灣或大陸)若為關聯基礎案,若發明人與關聯案發明人不同時詢問是否更新
         If DefInventor <> "" And DefInventor <> OrgInventor Then
            If MsgBox("是否更新關聯案的發明人資料至本案??", vbYesNo + vbDefaultButton2) = vbYes Then
                '更新發明人資訊
                CopyInventor cm(5) & cm(6) & cm(7) & cm(8), cp(1) & cp(2) & cp(3) & cp(4)
            End If
         End If
         'end 2014/04/23
         'Add By Sindy 2013/5/28
         If frm050102_1.bolIsEMPFlow = True Then
            intLeaveKind = 0
            'Unload frm050102_1
            frm090202_4.Show
            frm090202_4.QueryData
         '2013/5/28 End
         'Add By Sindy 2018/1/8
         ElseIf Me.m_strIR01 <> "" Then
            intLeaveKind = 0
            'Modify By Sindy 2022/5/20
            'frm04010519.GoNext
            Forms(0).Tmpfrm04010519.GoNext
            Set Forms(0).Tmpfrm04010519 = Nothing
            '2022/5/20 END
         '2018/1/8 END
         Else
            If Index = 2 Then
               intLeaveKind = 0
            Else
               intLeaveKind = 1
            End If
         End If
         Unload Me
      Case 4
         frm050201.Show
         frm050201.txt1(0).Text = txtCaseField(12)
         frm050201.txt1(1).Text = txtCaseField(12)
         frm050201.txt1(4).Text = "2"
         frm050201.Bol050102_3 = True
         Me.Hide
   End Select

End Sub

Private Function SaveDatabase() As Boolean
Dim i As Integer, j As Integer
Dim strTxt(1 To 30) As String, iStep As Integer
Dim strCusTemp As String, strTemp As String, strTemp1 As String
Dim strPromoteDate As String 'Add by Morgan 2006/7/31
Dim strCN(1 To 4) As String  'Add by Morgan 2007/10/23 國內案號
Dim Tmp As String            '2008/922 add by sonia
Dim bolLstInnerCase As Boolean '2010/1/21 是否最後一個國內案
Dim strNewCustNo As String
Dim strLetterJudge As String '指示信判發人/主旨 Added by Morgan 2018/8/20

 '911106 nick transation
 SaveDatabase = True
 On Error GoTo CheckingErr
 cnnConnection.BeginTrans
 
      'Add by Morgan 2004/9/23
      '設定客戶減免身分
      For i = 1 To 5
         If txtAD(i).Enabled = True Then
            '身分有變更才要做
            If txtAD(i).Tag <> txtAD(i).Text Then
               strSql = PUB_GetADSQL(txtCaseField(2 + i), txtCaseField(12).Text, txtAD(i).Text)
               cnnConnection.Execute strSql
            End If
         End If
      Next
  
   'Modify by Morgan 2005/11/18 調整
   '申請人1變更時要更新收據資料
   strNewCustNo = ChangeCustomerL(txtCaseField(3))
   If intCaseKind = 專利 Then
      strCusTemp = ChangeCustomerL(field(26))
   Else
      strCusTemp = ChangeCustomerL(field(8))
   End If
   If strNewCustNo <> strCusTemp Then
      '有開收據,更新acc0k0
      If cp(60) <> "" Then
         strExc(1) = field(1)
         strExc(2) = field(2)
         strExc(3) = field(3)
         strExc(4) = field(4)
         strExc(5) = cp(60)
         strExc(6) = strNewCustNo
         'edit by nickc 2007/02/05 不用 dll 了
         'If Not objLawDll.UpdAcc0k0(strExc(), True) Then
         If Not ClsLawUpdAcc0k0(strExc(), True) Then
            SaveDatabase = False
            cnnConnection.RollbackTrans
            txtCaseField(3).SetFocus
            Exit Function
         End If
      End If
   End If
   '2005/11/18 end
   
   field(5) = txtCaseField(0)
   field(6) = txtCaseField(1)
   field(7) = txtCaseField(2)
   cp(10) = txtCaseField(10)
   cp(27) = txtCaseField(11)
   field(9) = txtCaseField(12)
   'Remove by Morgan 2004/2/13
   '取消受讓人資料
   'cp(55) = txtCaseField(19)
   
   cp(64) = txtCaseField(21)
   If intCaseKind = 專利 Then
      
      Select Case txtCaseField(10)
         Case 異議_專
            field(23) = "2"
         Case 舉發
            field(23) = "3"
         Case Else
            field(23) = "1"
      End Select
      
      field(91) = txtCaseField(22)
      '92.1.12 add by sonia 大小個體
      'Modified by Morgan 2016/9/22 +印度040,菲律賓030--禧佩
      'Modified by Morgan 2023/3/24 條件同 ReadAllData
      'If txtCaseField(12) = "101" Or txtCaseField(12) = "102" Or (txtCaseField(12) = "203" And DBDATE(field(10)) >= "20050901") Or txtCaseField(12) = "040" Or txtCaseField(12) = "030" Then
      If InStr(CFP_ChkEntity, field(9)) > 0 Or field(179) <> "" Then
         If strSrvDate(1) >= PA179啟用日 Then
            If optChoose(0).Value = True Then
               field(179) = "1"
            ElseIf optChoose(1).Value = True Then
               field(179) = "2"
            ElseIf optChoose(2).Value = True Then
               field(179) = "3"
            End If
         Else
      'end 2023/3/24
            If optChoose(0).Value = True Then
               new_Entity = "大個體"
            ElseIf optChoose(1).Value = True Then
               new_Entity = "小個體"
            'Added by Morgan 2013/3/20
            ElseIf optChoose(2).Value = True Then
               new_Entity = "微個體"
            'end 2013/3/20
            End If
            If InStr(1, field(91), old_Entity, 1) > 0 Then
               field(91) = Replace(field(91), old_Entity, new_Entity, InStr(1, field(91), old_Entity, 1), , 1)
            Else
               If field(91) = "" Then
                  field(91) = new_Entity
               Else
                  field(91) = new_Entity & "，" & field(91)
               End If
            End If
         End If 'Added by Morgan 2023/3/24
      End If
      '92.1.12 end
      
      For i = 0 To 4
         field(i + 26) = txtCaseField(i + 3)
         For j = 0 To 2
            'Modify by Morgan 2005/11/18
            'field(j * 5 + i + 31) = txtCaseField(i * 3 + j + 23)
            field(j * 5 + i + 31) = Trim(txtCaseField(i * 3 + j + 23))
         Next
      Next
      
      '910816 Sieg 303
      For i = 79 To 84
         field(i) = txtCaseField(i - 40)
      Next
      
      For i = 109 To 132
         field(i) = txtCaseField(i - 64)
      Next
      
      Select Case cp(10)
         Case 發明申請, 追加申請
            field(8) = "1"
         Case 新型申請
            field(8) = "2"
         Case 設計申請, 聯合申請
            field(8) = "3"
      End Select
   
   Else
      field(26) = strNewCustNo
      '92.1.12 add by sonia 大小個體
      field(18) = txtCaseField(22)
      
      'Modify by Morgan 2006/9/21 加法國且申請日>=20050901
      'Modified by Morgan 2016/9/22 +印度040,菲律賓030--禧佩
'Removed by Morgan 2023/3/24
'      If txtCaseField(12) = "101" Or txtCaseField(12) = "102" Or (txtCaseField(12) = "203" And DBDATE(field(10)) >= "20050901") Or txtCaseField(12) = "040" Or txtCaseField(12) = "030" Then
'         If optChoose(0).Value = True Then
'            new_Entity = "大個體"
'         ElseIf optChoose(1).Value = True Then
'            new_Entity = "小個體"
'         'Added by Morgan 2013/3/20
'         ElseIf optChoose(2).Value = True Then
'            new_Entity = "微個體"
'         'end 2013/3/20
'         End If
'         If InStr(1, field(18), old_Entity, 1) > 0 Then
'            field(18) = Replace(field(18), old_Entity, new_Entity, InStr(1, field(18), old_Entity, 1), , 1)
'         Else
'            If field(18) = "" Then
'               field(18) = new_Entity
'            Else
'               field(18) = new_Entity & "，" & field(18)
'            End If
'         End If
'      End If
'end 2023/3/24
      '92.1.12 end
   End If
      
   field(157) = txtCaseField(18) 'Add by Morgan 2010/6/17

   If Mid(lblCaseField(0), 1, 1) = "B" Then
      cp(16) = 0
      cp(17) = 0
      cp(18) = 0
      cp(19) = 0
      cp(20) = "N"
      cp(26) = "N"
      cp(32) = "N"
   End If
   
   'Modify by Morgan 2008/2/14
   'cp(44) = Combo1
   intI = InStr(Combo1, "-")
   If intI > 0 Then
      cp(44) = Left(Combo1, intI - 1)
      cp(116) = Mid(Combo1, intI + 1)
   Else
      cp(44) = Combo1
      cp(116) = ""
   End If
   'end 2008/2/14
   cp(44) = ChangeCustomerL(cp(44))

   Select Case cp(10)
      Case 讓與, 專利權讓與, 授權, 變更, 補換發證書, 領證及繳年費, 年費, 申請英文證明, 實體審查, 申請優先權證明, 維持費
         cp(26) = "N"
   End Select
   
   'Modified by Morgan 2012/2/15 改呼叫共用函數
   'strExc(0) = "select cp45 from caseprogress where cp01=" + CNULL(cp(1)) + _
   '   " and cp02=" + CNULL(cp(2)) + " and cp03=" + CNULL(cp(3)) + _
   '   " and cp04=" + CNULL(cp(4)) + " and cp44=" + CNULL(cp(44)) + " order by cp27 desc"
   'intI = 1
   'Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
   'cp(45) = ""
   'If intI = 1 And Not IsNull(RsTemp.Fields("CP45")) Then cp(45) = RsTemp.Fields("CP45")
   If Not ClsPDGetCaseThatCode(cp) Then cp(45) = ""
   'end 2012/2/15
   
   cp(33) = ""
   cp(34) = ""
   strExc(0) = "select cf13,cf14 from casefee where cf01=" + CNULL(field(1)) + " and cf02=" + CNULL(field(9)) + _
      " and cf03=" + CNULL(cp(10)) + ""
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      If Not IsNull(RsTemp.Fields("cf13")) Then cp(33) = RsTemp.Fields("cf13")
      If Not IsNull(RsTemp.Fields("cf14")) Then cp(34) = RsTemp.Fields("cf14")
   End If
      
   cp(113) = txtCP113 'Added by Lydia 2021/05/25 工作時數
   
   strTxt(1) = GetCPSQL(cp())
   
   '91106 nick transation
   cnnConnection.Execute strTxt(1)
   
   If intCaseKind = 專利 Then
      strTxt(2) = GetPASQL(field())
   Else
      strTxt(2) = GetSPSQL(field())
   End If
   '91106 nick transation
   cnnConnection.Execute strTxt(2)
   
   iStep = 3
   
'Modify by Morgan 2009/8/18 改呼叫共用函式
   If Trim(txtCaseField(15)) <> "" Then
      PUB_UpdateChkResultDate txtCaseField(15), cp, cp(9), cp(10), cp(43)
   End If
'end 2009/8/18
   
   
   'Modify by Morgan 2009/5/12 +指定提申,最終提申(畫面輸入的提申期限或系統計算的最終期限),一般提申
   '指定提申(不可與一般提申或最終提申同時存在)
   If Trim(txtCaseField(20)) <> "" Then
      strExc(1) = DBDATE(txtCaseField(20))
      strExc(2) = PUB_GetWorkDay1(strExc(1), True)
      strSql = " insert into nextprogress a (np01,np02,np03,np04,np05,np07,np08,np09,np10,np22)" & _
         " values('" & cp(9) & "','" & cp(1) & "','" & cp(2) & "','" & cp(3) & "','" & cp(4) & "','995'" & _
         "," & strExc(2) & "," & strExc(1) & ",'" & strUserNum & "',GETNP22)"
      cnnConnection.Execute strSql, intI
   
   '其他有設定提申期限的才要管制
   'Modified by Morgan 2015/8/7 改呼叫共用
   Else
      '最終提申
      strExc(1) = DBDATE(cp(7))
      '畫面輸入
      If Trim(txtCaseField(14)) <> "" Then
         strExc(1) = DBDATE(txtCaseField(14))
      End If
      PUB_SetApplyDate cp(1), cp(2), cp(3), cp(4), strExc(1), cp(9), txtCaseField(10), txtCaseField(11), txtCaseField(12)
   'end 2015/8/7
   End If
    
    'Modify by Amy 2014/04/14 +strPriority5
    If ClsPDSavePriority(cp(), strPriority1, strPriority2, strPriority3, strPriority4, strPriority5) Then
    Else
        GoTo CheckingErr
    End If
   
    'edit by nickc 2007/02/02 不用 dll 了
    'If objPublicData.DeleteCountry(2, intCaseKind, cp()) Then
    If ClsPDDeleteCountry(2, intCaseKind, cp()) Then
        'Modify by Morgan 2006/4/7
        'If Not objPublicData.SaveCountry(2, intCaseKind, cp(1) & cp(2) & cp(3) & cp(4), cp(9), strCountry) Then
        If Not PUB_SaveCountry(2, intCaseKind, cp(1) & cp(2) & cp(3) & cp(4), cp(9), strCountry) Then
            GoTo CheckingErr
        End If
    Else
        GoTo CheckingErr
    End If

    If SaveAddDeadline(cp(), strAddDeadline1, strAddDeadline2, strAddDeadline3, , cp(13)) Then
    Else
        GoTo CheckingErr
    End If
    
   'Add by Morgan 2006/7/31
   '2010/1/22 MODIFY BY SONIA 國外案未發文之新申請程序重新更新齊備日並計算承辦期限並發Mail通知工程師
   'm_strMailCP09 = ""
   ReDim skMail(0) As SeekMails
   '2010/1/22 END

   strExc(0) = "SELECT CM01,CM02,CM03,CM04,NVL(PA05,NVL(PA06,PA07)) PA05 FROM CASEMAP,PATENT WHERE " & ChgCaseMap(cp(1) & cp(2) & cp(3) & cp(4), 0, 1) & " AND PA01(+)=CM01 AND PA02(+)=CM02 AND PA03(+)=CM03 AND PA04(+)=CM04 AND CM10='0' AND PA57 IS NULL"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      With RsTemp
         Do While Not .EOF
            strSql = "Select EP02,CP06,CP14 From CaseProgress,EngineerProgress WHERE " & ChgCaseprogress("" & .Fields(0) & .Fields(1) & .Fields(2) & .Fields(3)) & _
               " AND CP10 in (" & CaseMapOut & ") And CP27 Is Null AND  CP57 IS NULL and EP02=CP09 "
            '國外案為大陸外觀設計時則發Mail通知工程師並更新齊備日及承辦期限--郭
            'Modify by Morgan 2009/12/11 外觀設計已改由工程師承辦,若已有齊備日則無須再更新--郭2009/10/28
            'strSQL = strSQL & " AND (EP06 IS NULL OR EP06=0 OR EXISTS(SELECT * FROM PATENT WHERE PA01=CP01 AND PA02=CP02 AND PA03=CP03 AND PA04=CP04 AND PA08='3' AND PA09='020')) "
            strSql = strSql & " AND (EP06 IS NULL OR EP06=0) "
            'End
            CheckOC
            With adoRecordset
               .CursorLocation = adUseClient
               .Open strSql, cnnConnection
               If .RecordCount > 0 Then
                  '2010/1/21 ADD BY SONIA 若國外案有其他國內案未發文時不做
                  bolLstInnerCase = True
                  strExc(0) = "select cp01,cp02,cp03,cp04 from casemap,caseprogress" & _
                     " where cm10='0' and cm01='" & RsTemp.Fields("cm01") & "' and cm02='" & RsTemp.Fields("cm02") & "' and cm03='" & RsTemp.Fields("cm03") & "' and cm04='" & RsTemp.Fields("cm04") & "'" & _
                     " and not (cm05='" & cp(1) & "' and cm06='" & cp(2) & "' and cm07='" & cp(3) & "' and cm08='" & cp(4) & "')" & _
                     " and cp01(+)=cm05 and cp02(+)=cm06 and cp03(+)=cm07 and cp04(+)=cm08 AND CP10 IN (" & CaseMapIn & ") and cp27 is null and cp57 is null and rownum<2"
                  intI = 1
                  Set AdoRecordSet3 = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
                  If intI = 1 Then
                     bolLstInnerCase = False
                  End If
                  If bolLstInnerCase = True Then
                  '2010/1/21 END
                     '更新齊備日&發Mail(考慮修正及大陸外觀設計狀況)
                     strSql = "Update ENGINEERPROGRESS Set EP06=" & strSrvDate(1) & " Where EP02='" & adoRecordset.Fields(0).Value & "' "
                     cnnConnection.Execute strSql, intI
                     '2010/1/22 MODIFY BY SONIA
                     'm_strMailCP09 = m_strMailCP09 & adoRecordset.Fields(0).Value & ","
                     ReDim Preserve skMail(UBound(skMail) + 1) As SeekMails
                     skMail(UBound(skMail)).fiSender = strUserNum
                     skMail(UBound(skMail)).fiReceiver = adoRecordset.Fields("CP14").Value
                     skMail(UBound(skMail)).fiRecriverNo = ""
                     skMail(UBound(skMail)).fiSubject = "因關聯基礎案" & lblCaseField(1) & "的" & lblCaseProperty & "發文，上未齊備關聯案的文件齊備日及承辦期限！"
                     skMail(UBound(skMail)).fiContent = "未齊備關聯案：" & vbCrLf & vbCrLf & "本所案號：" & RsTemp("cm01") & "-" & RsTemp("cm02") & IIf(RsTemp("cm03") & RsTemp("cm04") = "000", "", "-" & RsTemp("cm03") & "-" & RsTemp("cm04")) & vbCrLf & "總收文號：" & adoRecordset.Fields(0).Value & vbCrLf & "案件名稱：" & RsTemp("PA05") & vbCrLf & "文件齊備日：" & ChangeTStringToTDateString(strSrvDate(2))
                     '2010/1/22 END
                     
                     If PUB_IfSetCP48() Then  'Add by Morgan 2010/10/1
                     
                        'Modify by Morgan 2007/10/11 承辦期限改呼叫共用函數計算
                        strSql = "Select cp01,cp10,pa09 From CaseProgress, Patent Where CP09='" & .Fields(0).Value & "' and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+)"
                        CheckOC2
                        With adoRecordset1
                           .CursorLocation = adUseClient
                           .Open strSql, cnnConnection
                           If .RecordCount > 0 Then
                              strPromoteDate = Pub_GetHandleDay(.Fields("cp01"), .Fields("pa09"), .Fields("cp10"), , "" & adoRecordset.Fields(1))
                              If strPromoteDate <> "" Then
                                 strSql = "Update CaseProgress Set CP48=" & CNULL(strPromoteDate) & " Where CP09='" & adoRecordset.Fields(0).Value & "' "
                                 cnnConnection.Execute strSql, intI
                              End If
                           End If
                        End With
                        
                     End If 'Add by Morgan 2010/10/1
                  End If
               End If
            End With
            .MoveNext
         Loop
      End With
   End If
   'end 2006/7/31
   
   'Add by Morgan 2007/2/9
   '日本新型發文預估2個月的核准日更新到其他多國新申請案的期限
   '德國新型發文預估3個月的核准日更新到其他多國新申請案的期限
   '澳洲新型發文預估3禮拜的核准日更新到其他多國新申請案的期限
   If (txtCaseField(12) = "011" Or txtCaseField(12) = "231" Or txtCaseField(12) = "015") And txtCaseField(10) = "102" Then
      'Modify by Morgan 2007/9/6 改呼叫共用程式
      PUB_UpdCP07byCP27 cp
   End If
   'End 2007/2/9
   
   'Add by Morgan 2007/10/22 96/11/1起
   '1. 若無代表圖則檢查國內&多國案是否有，有則複製否則發Mail通知郭雅娟
   '2. 當有圖則再複製到其他無圖之多國案
   m_NoPicMailSub = ""
   If Val(strSrvDate(1)) >= 20071101 And InStr(CaseMapOut, txtCaseField(10)) > 0 Then
      '檢查是否已有代表圖
      'Modify by Amy 2018/07/23 改寫至function for 彩色代表圖
'      strSql = "select 1 from imgbytefile where ibf01='" & field(1) & "' and ibf02='" & field(2) & "' and ibf03='" & field(3) & "' and ibf04='" & field(4) & "' and ibf05='1'"
'      intI = 1
'      Set RsTemp = ClsLawReadRstMsg(intI, strSql)
'      '無圖時,再檢查是否有國內案
'      If intI <> 1 Then
      If ChkImgByteFile(field(1), field(2), field(3), field(4)) = False Then
      'end 2018/07/23
         'Add by Morgan 2008/11/19 +分割案,CA申請抓母案
         '分割案
         If txtCaseField(10) = "307" Then
            strSql = "select dc05,dc06,dc07,dc08,ibf01 from divisioncase,imgbytefile where dc01='" & field(1) & "' and dc02='" & field(2) & "' and dc03='" & field(3) & "' and dc04='" & field(4) & "' and ibf01(+)=dc05 and ibf02(+)=dc06 and ibf03(+)=dc07 and ibf04(+)=dc08 and ibf05(+)='1' order by ibf01"
         'CA申請案
         ElseIf txtCaseField(10) = "122" Then
            strSql = "select pa01,pa02,pa03,pa04,ibf01 from patent,imgbytefile where pa01='" & field(1) & "' and pa02='" & field(2) & "' and pa03='0' and pa04='00' and ibf01(+)=pa01 and ibf02(+)=pa02 and ibf03(+)=pa03 and ibf04(+)=pa04 and ibf05(+)='1' order by ibf01"
         Else
         'end 2008/11/19
            strSql = "select cm05,cm06,cm07,cm08,ibf01 from casemap,imgbytefile where cm01='" & field(1) & "' and cm02='" & field(2) & "' and cm03='" & field(3) & "' and cm04='" & field(4) & "' and cm10='0' and ibf01(+)=cm05 and ibf02(+)=cm06 and ibf03(+)=cm07 and ibf04(+)=cm08 and ibf05(+)='1' order by ibf01"
         End If
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strSql)
         '有國內案
         If intI = 1 Then
            With RsTemp
               strCN(1) = .Fields(0)
               strCN(2) = .Fields(1)
               strCN(3) = .Fields(2)
               strCN(4) = .Fields(3)
               '有國內且無圖發Mail
               If IsNull(.Fields("ibf01")) Then
                  strExc(1) = field(1) & "-" & field(2) & IIf(field(3) & field(4) = "000", "", "-" & field(3) & "-" & field(4))
                  strExc(2) = strCN(1) & "-" & strCN(2) & IIf(strCN(3) & strCN(4) = "000", "", "-" & strCN(3) & "-" & strCN(4))
                  'Add by Morgan 2008/11/19 +分割案,CA申請抓母案
                  If txtCaseField(10) = "307" Then
                     m_NoPicMailSub = "分割案【" & strExc(1) & "】已發文但尚缺代表圖且母案【" & strExc(2) & "】亦無代表圖！"
                  ElseIf txtCaseField(10) = "122" Then
                     m_NoPicMailSub = "CA申請案【" & strExc(1) & "】已發文但尚缺代表圖且母案【" & strExc(2) & "】亦無代表圖！"
                  Else
                  'end 2008/11/19
                     m_NoPicMailSub = "【" & strExc(1) & "】已發文但尚缺代表圖且相關國內案【" & strExc(2) & "】亦無代表圖！"
                  End If
               Else
                  If PUB_CopyImgFile(strCN, field) = False Then
                     GoTo CheckingErr
                  End If
               End If
            End With
         '無國內案
         Else
            '若有其他多國有圖者
            strSql = "select cr05,cr06,cr07,cr08 from caserelation,imgbytefile where cr01='" & field(1) & "' and cr02='" & field(2) & "' and cr03='" & field(3) & "' and cr04='" & field(4) & "' and ibf01(+)=cr05 and ibf02(+)=cr06 and ibf03(+)=cr07 and ibf04(+)=cr08 and ibf05(+)='1' and ibf01 is not null"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strSql)
            If intI = 1 Then
               With RsTemp
                  strCN(1) = .Fields("cr05")
                  strCN(2) = .Fields("cr06")
                  strCN(3) = .Fields("cr07")
                  strCN(4) = .Fields("cr08")
                  If PUB_CopyImgFile(strCN, field) = False Then
                     GoTo CheckingErr
                  End If
               End With
            '有國外亦無圖則發Mail
            Else
               strExc(1) = field(1) & "-" & field(2) & IIf(field(3) & field(4) = "000", "", "-" & field(3) & "-" & field(4))
               m_NoPicMailSub = "【" & strExc(1) & "】已發文但尚缺代表圖！"
            End If
         End If
      End If

      '若有代表圖則複製到其他多國案
      If m_NoPicMailSub = "" Then
         '其他多國且無圖者
         strSql = "select cr05,cr06,cr07,cr08 from caserelation,imgbytefile where cr01='" & field(1) & "' and cr02='" & field(2) & "' and cr03='" & field(3) & "' and cr04='" & field(4) & "' and ibf01(+)=cr05 and ibf02(+)=cr06 and ibf03(+)=cr07 and ibf04(+)=cr08 and ibf05(+)='1' and ibf01 is null"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strSql)
         If intI = 1 Then
            With RsTemp
               Do While Not .EOF
                  strCN(1) = .Fields("cr05")
                  strCN(2) = .Fields("cr06")
                  strCN(3) = .Fields("cr07")
                  strCN(4) = .Fields("cr08")
                  If PUB_CopyImgFile(field, strCN) = False Then
                     GoTo CheckingErr
                  End If
                  .MoveNext
               Loop
            End With
         End If
      End If
      
   End If
   'end 2007/10/22
   
   'Add by Morgan 2009/7/3
   If cp(10) = "101" Or cp(10) = "102" Or cp(10) = "103" Then
      '實審期限為發文日起算者
      If ClsPDGetNationTaxEx(Val(field(8)) + 3, field(9), strTemp, strTemp1, , , False) = 0 Then
         If Val(strTemp) = 發文日 Then
            'Memo by Lydia 2022/04/18 調整設定：若優先權基礎案的國家為奧地利、丹麥、日本、中國大陸、韓國、西班牙、瑞典、瑞士、英國、美國，則新案發文時不在下一程序掛提供前案的期限，並在新案指示信帶出相關段落。
            PUB_UpdExamDate cp(1), cp(2), cp(3), cp(4), cp(9)
         End If
      End If
   End If
   
   PUB_SetArriveDate cp(9)  'Add by Morgan 2009/11/11
   
   'Add by Morgan 2010/4/2
   '母案發文後集體設計子案也要一併上發文
   m_iMultiDesign = 0
   If txtCaseField(10) = "103" And cp(3) = "0" Then
      'Modified by Morgan 2014/11/21 取消子案一併發文(只統計件數帶入指示信)--甄妮
      'strSql = "update caseprogress set (cp27,cp44,cp116)=(select b.cp27,b.cp44,b.cp116 from caseprogress b where b.cp09='" & cp(9) & "')" & _
         " where cp01='" & cp(1) & "' and cp02='" & cp(2) & "' and cp03<>'0' and cp04='" & cp(4) & "' and cp10='105' and cp57 is null"
      strSql = "update caseprogress set cp27=cp27" & _
         " where cp01='" & cp(1) & "' and cp02='" & cp(2) & "' and cp03<>'0' and cp04='" & cp(4) & "' and cp10='105' and cp57 is null"
      'end 2014/11/21
      cnnConnection.Execute strSql, intI
      m_iMultiDesign = intI
   End If
   
   PUB_944Inform cp(9)  'Add by Morgan 2011/7/27
   
   'Added by Morgan 2012/3/21
   'EPC發明要管制檢索報告來函期限=發文日+6個月
   'Modify by Amy 2013/05/27
   'EPC分割案發文下一程序要產生管制檢索報告的程序
   'If txtCaseField(12) = "221" And txtCaseField(10) = "101" Then
   If txtCaseField(12) = "221" And (txtCaseField(10) = "101" Or txtCaseField(10) = "307") Then
      'Modified by Morgan 2013/5/16 改+1年
      'strExc(1) = CompDate(1, 6, txtCaseField(11))
      strExc(1) = CompDate(0, 1, txtCaseField(11))
      strExc(2) = PUB_GetWorkDay1(strExc(1), True)
      strSql = " insert into nextprogress a (np01,np02,np03,np04,np05,np07,np08,np09,np10,np22)" & _
         " values('" & cp(9) & "','" & cp(1) & "','" & cp(2) & "','" & cp(3) & "','" & cp(4) & "','1209'" & _
         "," & strExc(2) & "," & strExc(1) & ",'" & strUserNum & "',GETNP22)"
      cnnConnection.Execute strSql, intI
   End If
   'end 2012/3/21
   
   
   
   'Modify By Sindy 2014/11/13
   If strSrvDate(1) >= 專利發明人檔啟用日 Then
      '更新發明人資訊
      UpdPatentInventor cp(1) & cp(2) & cp(3) & cp(4), strInventorNo
   Else
   '2014/11/13 END
      'Add by Amy 2013/08/06 CFP案發文時,發明人資料帶已鍵關聯之P案最新發明人資訊
      If cm(5) <> "" Then
         '更新發明人資訊
         'Modify by Amy 2014/04/23 改只save Patent
         'CopyInventor cm(5) & cm(6) & cm(7) & cm(8), cp(1) & cp(2) & cp(3) & cp(4)
         UpdPatentInventor cp(1) & cp(2) & cp(3) & cp(4), strInventorNo
      End If
      'end 2013/08/06
   End If
   
   'Add by Sindy 2018/1/8
   If m_strIR01 <> "" Then
      PUB_UpdateEMailRec m_strIR01, m_strIR02, m_strIR03, m_strIR04, "frm050102_1"
   End If
   '2018/1/8 END
   
   'Added by Morgan 2018/8/20 CFP電子化
   If strSrvDate(1) >= CFP指示信電子化啟用日 Then
      'Modified by Morgan 2018/9/6 +有工程師的指示信
      If txtCaseField(17) <> "N" Or m_bolEngLetter = True Then
         If m_bolEngLetter Then
            strLetterJudge = strUserNum
         Else
            strLetterJudge = PUB_GetLetterJudgeNew("2", field(1), txtCaseField(10), txtCaseField(12))
         End If
         m_strSubject = PUB_GetSubject(field(1), field(2), field(3), field(4), txtCaseField(10), field(11), cp(45), txtCaseField(12), field(46))
         PUB_AddAppForm cp(9), True, strLetterJudge, m_strSubject
         m_strAF01 = cp(9)
      End If
   End If
   If strSrvDate(1) >= CFP第一階段電子化啟用日 Then
      If txtCaseField(8) <> "N" Then
         strLetterJudge = PUB_GetLetterJudgeNew("1", field(1), txtCaseField(10), txtCaseField(12))
         PUB_AddLetterProgress cp(9), 0, True, strLetterJudge, False, field(26), txtCaseField(10), field(75)
         m_strLD18 = cp(9)
      End If
   End If
   'end 2018/8/20
      
   'Added by Morgan 2019/4/29 更新實審減免資格(日本發明案指示信會用,故需新更新)
   If m_str416CP81 <> "" Then
      strSql = "update caseprogress set cp81='" & m_str416CP81 & "' where cp01='" & cp(1) & "' and cp02='" & cp(2) & "' and cp03='" & cp(3) & "' and cp04='" & cp(4) & "' and cp10='416' AND CP57 IS NULL"
      cnnConnection.Execute strSql, intI
   End If
   'end 2019/4/29
   
   '911106 nick transation
   cnnConnection.CommitTrans
   
   ' 90.12.05 modify by louis 若案件國家收費表存在代理人收達天數則新增一筆收達的下一程序檔
   If SaveDatabase = True Then
   
'Modify by Morgan 2009/11/11 收達期限管控改呼叫公用函式(移到上面)
'      Dim strCF23 As String
'      Dim strNPDate As String
'      Dim strNPSerial As String
'
'      If IsExistCasefee(cp(1), field(9), cp(10), strCF23) Then
'        'Modify By Cheng 2003/09/01
''         strNPDate = DBDATE(Format(DateSerial(Val(DBYEAR(cp(27))), Val(DBMONTH(cp(27))), Val(DBDAY(cp(27))) + Val(strCF23))))
'         strNPDate = DBDATE(DateAdd("d", Val(strCF23), ChangeWStringToWDateString(DBDATE(cp(27)))))
'         strNPSerial = InsertNextProgress_997(cp(9), cp(1), cp(2), cp(3), cp(4), strNPDate)
'      End If
'end 2009/11/11
      
      ' 90.12.25  郵寄方式
      'Modified by Lydia 2014/12/27 + (DHL列印)
'      Select Case txtCaseField(9)
'      '92.1.12 MODIFYB BY SONIA
'      'Case "1", "3"
'      Case ""
'      '92.1.12 END
       If txtCaseField(9) = "" And txtCaseField(69) = "" Then
       
'Modified by Morgan 2018/8/31 不必印地址條 --慧汶
'            Screen.MousePointer = vbDefault
'            frm083014.Show
'            bolToEndByNick = False
''            frm083014.Hide
'            '***** 加入本所案號  91.08.07  nick
'            frm083014.Text1(6).Text = cp(1) & "-" & cp(2) & "-" & cp(3) & "-" & cp(4)
'            '******* end
'            frm083014.Text1(0).Text = strNewCustNo
'            frm083014.Text1(3).Text = "1"
'            frm083014.Text1(4).Text = "1"
'            frm083014.Text1(5).Text = "Y"
'            frm083014.Opt1(0).Value = True
'            frm083014.Opt1(0).Enabled = False
'            frm083014.Opt1(1).Enabled = False
'            frm083014.Opt1(2).Enabled = False
'            frm083014.Text1(0).Enabled = False
'            frm083014.Text1(1).Enabled = False
'            frm083014.Text1(2).Enabled = False
'            frm083014.Text1(3).Enabled = False
'            frm083014.Text1(4).Enabled = False
'            frm083014.Text1(5).Enabled = False
'            frm083014.Show
'            Do
'                DoEvents
'                If bolToEndByNick = True Then Exit Do
'            Loop Until Not frm083014.Visible
'            Unload frm083014
'            Me.Enabled = True
'      '92.1.12 MODIFY BY SONIA
'      'Case "2"
''      Case "Y"
'end 2018/8/31

       ElseIf txtCaseField(9) = "Y" Then 'TNT列印
      '92.1.12 END
            Screen.MousePointer = vbDefault
            frm060321.Show
            bolToEndByNick = False
            frm060321.Hide
            'Add By Cheng 2002/12/24
            '傳收文號
            frm060321.GetCP09 = cp(9)
            frm060321.txt1(0).Text = cp(1)
            frm060321.txt1(1).Text = cp(2)
            frm060321.txt1(2).Text = cp(3)
            frm060321.txt1(3).Text = cp(4)
            frm060321.txt1(0).Enabled = False
            frm060321.txt1(1).Enabled = False
            frm060321.txt1(2).Enabled = False
            frm060321.txt1(3).Enabled = False
            Me.Enabled = False
            frm060321.Show
            Do
                DoEvents
                If bolToEndByNick = True Then Exit Do
            Loop Until Not frm060321.Visible
            Unload frm060321
            Me.Enabled = True
            
       ElseIf txtCaseField(69) = "Y" Then  'DHL列印
            Screen.MousePointer = vbDefault
            frm060330.Show
            bolToEndByNick = False
            frm060330.Hide
            'frm060330.GetCP09 = cp(9) 'mark by Lydia 2022/03/28
            frm060330.txt1(0).Text = cp(1)
            frm060330.txt1(1).Text = cp(2)
            frm060330.txt1(2).Text = cp(3)
            frm060330.txt1(3).Text = cp(4)
            frm060330.txt1(0).Enabled = False
            frm060330.txt1(1).Enabled = False
            frm060330.txt1(2).Enabled = False
            frm060330.txt1(3).Enabled = False
            Me.Enabled = False
            frm060330.Show
            Do
                DoEvents
                If bolToEndByNick = True Then Exit Do
            Loop Until Not frm060330.Visible
            Unload frm060330
            Me.Enabled = True
       End If
'      Case Else
'      End Select
   End If
 '911106 nick transation
     Exit Function
CheckingErr:
    SaveDatabase = False
     cnnConnection.RollbackTrans
   
End Function
'補件期限存檔
Public Function SaveAddDeadline(ByRef cp() As String, ByRef strAddDeadline1 As String, ByRef strAddDeadline2 As String, ByRef strAddDeadline3 As String, Optional intWhere As Integer = 國外_CF, Optional ByRef strUserNum As String) As Boolean
Dim i As Integer, varAddDeadLineTemp1, varAddDeadLineTemp2, varAddDeadLineTemp3, strTemp As String
'edit by nickc 2007/02/06 不用 dll 了
'Dim objPublicData   As Object
'Set objPublicData = CreateObject("prjTaieDll.clsPublicData")
On Error GoTo ErrHnd
If strAddDeadline1 <> "" Then
   varAddDeadLineTemp1 = Split(strAddDeadline1, ",")
   varAddDeadLineTemp2 = Split(strAddDeadline2, ",")
   varAddDeadLineTemp3 = Split(strAddDeadline3, ",")
   If strAddDeadline3 = "" Then
      'Modify By Cheng 2002/03/06
'      If objPublicData.SaveNewNextProgressDatabase(intWhere, cp(), CStr(varAddDeadLineTemp1(i)), CStr(varAddDeadLineTemp2(i)), 補文件) = False Then GoTo Err
      If SaveNewNextProgressDatabase(intWhere, cp(), CStr(varAddDeadLineTemp1(i)), CStr(varAddDeadLineTemp2(i)), 補文件, , , , lblCaseField(4)) = False Then GoTo ErrHnd
   Else
      For i = 0 To UBound(varAddDeadLineTemp1)
         'Modify By Cheng 2002/03/06
'         If objPublicData.SaveNewNextProgressDatabase(intWhere, cp(), CStr(varAddDeadLineTemp1(i)), CStr(varAddDeadLineTemp2(i)), 補文件, CStr(varAddDeadLineTemp3(i))) = False Then GoTo Err
         If SaveNewNextProgressDatabase(intWhere, cp(), CStr(varAddDeadLineTemp1(i)), CStr(varAddDeadLineTemp2(i)), 補文件, CStr(varAddDeadLineTemp3(i)), , , lblCaseField(4)) = False Then GoTo ErrHnd
      Next
   End If
End If
SaveAddDeadline = True
'edit by nickc 2007/02/06 不用 dll 了
'Set objPublicData = Nothing
Exit Function
ErrHnd:
   SaveAddDeadline = False
End Function

'新增資料至NextProgress檔
Public Function SaveNewNextProgressDatabase( _
      ByRef intWhere As Integer, ByRef cp() As String, _
      ByVal strDate1 As String, ByVal StrDate2 As String, _
      ByRef strCaseProperty As String, _
      Optional strMemo As String, Optional strNumber As String, _
      Optional strPerson As String, Optional strUserNum As String) As Boolean
   Dim strSql As String, strCounter As String
   Dim adoRecord As New ADODB.Recordset
   Dim objPrtForm001 As New ClsPrtForm001
   
   On Error GoTo ErrHnd
   strDate1 = TransDate(strDate1, 2)
   StrDate2 = TransDate(StrDate2, 2)
   
   adoRecord.CursorLocation = adUseClient
   adoRecord.Open "select max(np22) from nextprogress", cnnConnection, adOpenStatic, adLockReadOnly   'edit by nickc 2005/08/04 原先是  動態開啟
   If IsNull(adoRecord.Fields(0).Value) Then
      strCounter = "1"
   Else
      strCounter = Val(adoRecord.Fields(0).Value) + 1
   End If
   'Add By Cheng 2002/04/11
   '取得欲新增的下一程序檔的序號
   g_dbl_NPSerialNo = Val(strCounter)
   
   'Modify By Cheng 2002/01/29
   '若為CFP案且下一程序為催審, 提申或收達時, 智權人員代號(NP10)以操作員代號寫入
   '若為CFP案且下一程序為補文件, 智權人員代號(NP10)以操作員代號寫入
'   If intWhere = 國外_CF And (strCaseProperty = 催審 Or strCaseProperty = 提申 Or strCaseProperty = 收達) Then
   If intWhere = 國外_CF And (strCaseProperty = 催審 Or strCaseProperty = 提申 Or strCaseProperty = 收達 Or strCaseProperty = 補文件) Then
      'Modify By Cheng 2002/09/25
'      strSQL = "insert into nextprogress (np01,np02,np03,np04,np05,np07,np08,np09,np10,np13,np14,np15,np22) values (" + _
'         CNULL(cp(9)) + "," + CNULL(cp(1)) + "," + CNULL(cp(2)) + "," + CNULL(cp(3)) + _
'         "," + CNULL(cp(4)) + "," + CNULL(strCaseProperty) + "," & Val(strDate1) & "," & Val(strDate2) & _
'         "," + CNULL(strUserNum) + "," + CNULL(strNumber) + "," + CNULL(strPerson) + "," + CNULL(strMemo) + "," & Val(strCounter) & ")"
      strSql = "insert into nextprogress (np01,np02,np03,np04,np05,np07,np08,np09,np10,np13,np14,np15,np22) values (" + _
         CNULL(cp(9)) + "," + CNULL(cp(1)) + "," + CNULL(cp(2)) + "," + CNULL(cp(3)) + _
         "," + CNULL(cp(4)) + "," + CNULL(strCaseProperty) + "," & Val(strDate1) & "," & Val(StrDate2) & _
         "," + CNULL(strUserNum) + "," + CNULL(strNumber) + "," + CNULL(ChgSQL(strPerson)) + "," + CNULL(strMemo) + "," & Val(strCounter) & ")"
   '其他的維持原狀
   Else
        'Modify By Cheng 2003/11/24
        '重抓智權人員
'      strSQL = "insert into nextprogress (np01,np02,np03,np04,np05,np07,np08,np09,np10,np13,np14,np15,np22) values (" + _
'         CNULL(cp(9)) + "," + CNULL(cp(1)) + "," + CNULL(cp(2)) + "," + CNULL(cp(3)) + _
'         "," + CNULL(cp(4)) + "," + CNULL(strCaseProperty) + "," & Val(strDate1) & "," & Val(strDate2) & _
'         "," + CNULL(cp(13)) + "," + CNULL(strNumber) + "," + CNULL(ChgSQL(strPerson)) + "," + CNULL(strMemo) + "," & Val(strCounter) & ")"
      strSql = "insert into nextprogress (np01,np02,np03,np04,np05,np07,np08,np09,np10,np13,np14,np15,np22) values (" + _
         CNULL(cp(9)) + "," + CNULL(cp(1)) + "," + CNULL(cp(2)) + "," + CNULL(cp(3)) + _
         "," + CNULL(cp(4)) + "," + CNULL(strCaseProperty) + "," & Val(strDate1) & "," & Val(StrDate2) & _
         "," + CNULL(PUB_GetAKindSalesNo(cp(1), cp(2), cp(3), cp(4))) + "," + CNULL(strNumber) + "," + CNULL(ChgSQL(strPerson)) + "," + CNULL(strMemo) + "," & Val(strCounter) & ")"
   End If
   cnnConnection.Execute strSql
   SaveNewNextProgressDatabase = True
   ' 90.12.18 modify by louis 列印接洽結案單
   '92.3.9 cancel by sonia
   'objPrtForm001.PrintForm strCounter, cp(1), cp(2), cp(3), cp(4)
   Exit Function
ErrHnd:
   ShowMsg MsgText(9138)
   SaveNewNextProgressDatabase = False
End Function
'Add by Morgan 2003/12/19
'更新代表人選單
Private Sub renewRepCombo(Optional ByVal iIdxFrom As Integer = 3, Optional ByVal iIdxTo As Integer = 7)
      
      Dim i As Integer, stSQL As String, j As Integer, iRtn As Integer, strValue As String
      Dim rsTmp As New ADODB.Recordset
      
      For i = 2 * (iIdxFrom - 3) To 2 * (iIdxTo - 3) + 1
         Combo2(i).Clear
         Combo2(i).AddItem ""
      Next

      For i = iIdxFrom To iIdxTo
         If txtCaseField(i) <> "" Then
            stSQL = "SELECT CU40,CU43,CU46,CU49,CU52,CU55 FROM CUSTOMER WHERE " & ChgCustomer(txtCaseField(i))
            iRtn = 1
            'edit by nickc 2007/02/05 不用 dll 了
            'Set rsTmp = objLawDll.ReadRstMsg(iRtn, stSQL)
            Set rsTmp = ClsLawReadRstMsg(iRtn, stSQL)
            If iRtn = 1 Then
               For j = 1 To 6
                  If IsNull(rsTmp.Fields(j - 1)) Then
                     strValue = ""
                  Else
                     strValue = "-" & rsTmp.Fields(j - 1)
                  End If
                  Combo2((i - 3) * 2).AddItem txtCaseField(i) & "-" & j & strValue
                  Combo2((i - 3) * 2 + 1).AddItem txtCaseField(i) & "-" & j & strValue
               Next
            End If
         End If
      Next
      
End Sub

Private Sub ReadAllData()
Dim rt As Boolean, i As Integer, strTemp As String, strTemp1 As String, j As Integer
Dim adoRecord As Object, strSameName As String
'Add By Sindy 2014/11/13
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
'2014/11/13 END

On Error GoTo ErrHnd
Screen.MousePointer = vbHourglass
'Modify by Morgan 2006/10/19 改不Call Dll
'If objPublicData.ReadAllData(frm050102_1.grdDataList.TextMatrix(frm050102_1.grdDataList.Row, 5), cp(), field(), intCaseKind, intPWhere) Then
ReDim cp(TF_CP) As String
cp(9) = frm050102_1.grdDataList.TextMatrix(frm050102_1.grdDataList.row, 5)
If PUB_ReadAllData(cp(), field(), intCaseKind, intPWhere) Then
'end 2006/10/19

   lblCaseField(0) = cp(9)
   lblCaseField(1) = cp(1) + " - " + cp(2) + _
      IIf(cp(4) = "00" And cp(3) = "0", "", " - " + cp(3)) + _
      IIf(cp(4) = "00", "", " - " + cp(4))
   lblCaseField(2) = TransDate(cp(6), 1)
   lblCaseField(4) = cp(13)
   lblCaseField(5) = TransDate(cp(7), 1)
   txtCaseField(0) = field(5)
   txtCaseField(1) = field(6)
   txtCaseField(2) = field(7)
   txtCaseField(10) = cp(10)
   txtCaseField(12) = field(9)
   'Remove by Morgan 2004/2/13
   '取消受讓人資料
   'txtCaseField(19) = cp(55)
   
   txtCaseField(21) = cp(64)
    'Modify By Cheng 2003/02/17
    '移至LostFocus
'   CheckKeyIn 10
    txtCaseField_LostFocus 10
   CheckKeyIn 12
   
   strInventorNo = ""
   If intCaseKind = 專利 Then
      txtCaseField(22) = field(91)
      lblCaseField(3) = field(8)
      '申請人
      For i = 0 To 4
         txtCaseField(i + 3) = field(i + 26)
         For j = 0 To 2
              txtCaseField(i * 3 + j + 23) = field(j * 5 + i + 31)
         Next
         CheckKeyIn i + 3
      Next
      'Add By Cheng 2002/08/23
      m_strCust1 = "" & Me.txtCaseField(3).Text
      m_strCust2 = "" & Me.txtCaseField(4).Text
      m_strCust3 = "" & Me.txtCaseField(5).Text
      m_strCust4 = "" & Me.txtCaseField(6).Text
      m_strCust5 = "" & Me.txtCaseField(7).Text
      
      'Added by Morgan 2012/3/31
      For intI = 3 To 7
         txtCaseField(intI).Tag = txtCaseField(intI).Text
      Next
      'end 2012/3/30
      
      '發明人
      'Modify By Sindy 2014/11/13
      If strSrvDate(1) >= 專利發明人檔啟用日 Then
         StrSQLa = "SELECT pi06 from PatentInventor where pi01=" + CNULL(field(1)) + " and pi02=" + CNULL(field(2)) + " and pi03=" + CNULL(field(3)) + " and pi04=" + CNULL(field(4)) & _
                   " order by pi05 asc"
         If rsA.State <> adStateClosed Then rsA.Close
         rsA.CursorLocation = adUseClient
         rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
         If rsA.RecordCount > 0 Then
            rsA.MoveFirst
            Do While Not rsA.EOF
               strInventorNo = strInventorNo & rsA.Fields("pi06") & ","
               rsA.MoveNext
            Loop
         End If
      Else
      '2014/11/13 END
         For i = 60 To 69
            If field(i) <> "" Then
               strInventorNo = strInventorNo & field(i) & ","
            End If
         Next
      End If
      If Right(strInventorNo, 1) = "," Then strInventorNo = Left(strInventorNo, Len(strInventorNo) - 1)
      
      'Morgan 2003/11/20
      Call renewRepCombo
      'Morgan 2003/11/20 -- end
      
      '910816 Sieg 303
      '代表人
      For i = 79 To 84
         txtCaseField(i - 40) = field(i)
      Next
      '代表人
      For i = 109 To 132
         txtCaseField(i - 64) = field(i)
      Next
      
      txtCaseField(18) = field(157) 'Add by Morgan 2010/6/17
   
   Else
      Me.txtCaseField(3).Text = field(8)
      CheckKeyIn 3
      Me.txtCaseField(4).Text = field(58)
      CheckKeyIn 4
      Me.txtCaseField(5).Text = field(59)
      CheckKeyIn 5
   End If
   
      
   '92.1.12 add by sonia
   'Modify by Morgan 2006/9/20 加法國
   'Modified by Morgan 2016/9/22 +印度040,菲律賓030--禧佩
   'Modified by Morgan 2023/3/24 國家改抓常數，收文有設定的也顯示
   'If field(9) = "101" Or field(9) = "102" Or field(9) = "203" Or field(9) = "040" Or field(9) = "030" Then
   '   'Added by Morgan 2013/3/20
   '   If field(9) = "101" Then
   '      optChoose(2).Enabled = True
   '   Else
   '      optChoose(2).Enabled = False
   '   End If
   '   'end 2013/3/20
   PUB_SetEntityOpt field(1), field(9), field(8), optChoose
   If InStr(CFP_ChkEntity, field(9)) > 0 Or field(179) <> "" Then
      If strSrvDate(1) >= PA179啟用日 Then
         If field(179) = "1" Then
            optChoose(0).Value = True
            old_Entity = optChoose(0).Caption
         ElseIf field(179) = "2" Then
            optChoose(1).Value = True
            old_Entity = optChoose(1).Caption
         ElseIf field(179) = "3" Then
            optChoose(2).Value = True
            old_Entity = optChoose(2).Caption
         Else
            old_Entity = ""
         End If
      Else
   'end 2023/3/24
         If InStr(1, field(91), "大個體", 1) > 0 Then
            optChoose(0).Value = True
            old_Entity = "大個體"
         ElseIf InStr(1, field(91), "小個體", 1) > 0 Then
            optChoose(1).Value = True
            old_Entity = "小個體"
         'Added by Morgan 2013/3/20
         ElseIf InStr(1, field(91), "微個體", 1) > 0 Then
            optChoose(2).Value = True
            old_Entity = "微個體"
         'end 2013/3/20
         Else
            old_Entity = ""
         End If
      End If 'Added by Morgan 2023/3/24
   End If
   '92.1.12 end
   
   Set adoRecord = CreateObject("ADODB.Recordset")
   'edit by nickc 2007/02/02 不用 dll 了
   'If objPublicData.SelectTable("select cp44 from caseprogress where cp01 = '" & cp(1) & "' and cp02 = '" & cp(2) & "' and cp03 = '" & cp(3) & "' and cp04 = '" & cp(4) & "'", adoRecord) Then
   '2007/4/23 MODIFY BY SONIA 加發文日降冪排序
   'If ClsPDSelectTable("select cp44 from caseprogress where cp01 = '" & cp(1) & "' and cp02 = '" & cp(2) & "' and cp03 = '" & cp(3) & "' and cp04 = '" & cp(4) & "'", adoRecord) Then
   'Modify by Morgan 2008/2/18 加聯絡人
   'Added by Lydia 2016/10/27 +新案有申請人指定國外代理人檔則預設
   If cp(31) = "Y" Then
      AddAgent Combo1, cp, , , , cp(9), field(9), field(26)
      If Combo1 <> "" Then CheckKeyIn 13
      
   Else '非新案照原本
        If ClsPDSelectTable("select cp44||decode(cp116,null,null,'-'||cp116) from caseprogress where cp01 = '" & cp(1) & "' and cp02 = '" & cp(2) & "' and cp03 = '" & cp(3) & "' and cp04 = '" & cp(4) & "' and cp09<'C' and cp44 is not null order by cp27 desc", adoRecord) Then
        '2007/4/23 END
           Do While adoRecord.EOF = False
              If IsNull(adoRecord.Fields(0).Value) = False Then
                 If strSameName <> adoRecord.Fields(0).Value Then
                    Combo1.AddItem adoRecord.Fields(0).Value
                    strSameName = adoRecord.Fields(0).Value
                 End If
              End If
              adoRecord.MoveNext
           Loop
           Combo1 = Combo1.List(0)
        End If
        
      'Added by Morgan 2023/10/30 已有設定時不必再重新設定(IDS分案會先設,且抓預設代理人時也會剔除)
      If cp(44) <> "" Then
         Combo1 = cp(44) & IIf(cp(116) <> "", "-" & cp(116), "")
         CheckKeyIn 13
      Else
      'end 2023/10/30
      
        'edit by nickc 2007/02/02 不用 dll 了
        'If objPublicData.GetCasePreAgent(cp(), strTemp) Then
        If ClsPDGetCasePreAgent(cp(), strTemp) Then
           Combo1 = strTemp
           CheckKeyIn 13
        End If
        
      End If 'Added by Morgan 2023/10/30
   End If
   'end 2016/10/27
   
   'Modify by Amy 2014/04/14 +strPriority5
   If ClsPDReadPriority(cp(), strPriority1, strPriority2, strPriority3, strPriority4, strPriority5) = False Then GoTo ErrHnd
   'Add By Cheng 2002/07/30
   'Me.txtCaseField(9).Text = "Y" 'Removed by Morgan 2018/8/22 取消--玫音
   Me.txtCaseField(16).Text = "Y"
   Me.chkChoose(1).Value = True
   'Modify by Morgan 2006/7/21 歐盟設計說明書不勾--禧佩
   'Modified by Morgan 2012/9/14 +英國,印度設計
   If Not (InStr("239,201,040", txtCaseField(12)) > 0 And txtCaseField(10) = "103") Then
      Me.chkChoose(0).Value = True
   End If
   ShowAudit
   
   'Added by Morgan 2019/1/22
   'EU的集體設計子案預設不出客戶函及指示信 -- 慧汶
   If cp(10) = "105" And cp(3) <> "0" Then
      txtCaseField(8) = "N"
      txtCaseField(17) = "N"
   End If
   'end 2019/1/22
Else
   bolLeave = True
   intLeaveKind = 1
   Unload Me
End If
ErrHnd:
ErrorMsg
Screen.MousePointer = vbDefault
End Sub

Private Sub cmdSuggest_Click()
   If txtCaseField(12) = "" Then
      MsgBox "請先輸入申請國家！"
   Else
      ShowSuggest
   End If
End Sub

Private Sub ShowSuggest(Optional bNoMsg As Boolean)
   Dim bCancel As Boolean
   If txtCaseField(12) <> "" And InStr(NewCasePtyList, txtCaseField(10)) > 0 And bNoMsg = False Then
      If PUB_ReadFTList(field(1), txtCaseField(12), RsTemp, txtCaseField(11)) = True Then
         Set frm880012.grdDataList.Recordset = RsTemp
         Set frm880012.fmParent = Me
         frm880012.Show vbModal
         If Me.Tag <> "" Then
            Combo1 = Me.Tag
            Combo1_Validate bCancel
         End If
      ElseIf bNoMsg = False Then
         MsgBox "該申請國無建議代理人！"
      End If
   End If
End Sub
'Add by Morgan 2008/4/17
'檢查給案量是否超過
Private Function CheckCP44() As Boolean
   Dim stDate As String, stFT05 As String, stDate1 As String, stDate2 As String, stYear As String
   Dim iPos As Integer, stCon As String, stVTB As String, stConCP As String
   Dim bolRtn As Boolean
   
   bolRtn = True
   If Combo1 <> "" And txtCaseField(12) <> "" And InStr(NewCasePtyList, txtCaseField(10)) > 0 Then
      stDate = strSrvDate(1)
      stYear = Left(stDate, 4)
      '下半年
      If Val(Mid(stDate, 5, 2)) > 6 Then
         stFT05 = "2"
         stDate1 = stYear & "0701"
         stDate2 = stYear & "1231"
      '上半年
      Else
         stFT05 = "1"
         stDate1 = stYear & "0101"
         stDate2 = stYear & "0630"
      End If
      stCon = " AND FT04=" & (stYear - 1911) & " and FT05='" & stFT05 & "'"
      stConCP = " and instr('" & NewCasePtyList & "',cp10)>0 and cp27>=" & stDate1 & " and cp27<=" & stDate2
      
      If InStr(Combo1, "-") > 0 Then
         iPos = InStr(Combo1, "-")
         stCon = stCon & " and FT01||FT02='" & ChangeCustomerL(Left(Combo1, iPos - 1)) & "' and FT03='" & Mid(Combo1, iPos + 1) & "'"
         stConCP = stConCP & " and CP44='" & ChangeCustomerL(Left(Combo1, iPos - 1)) & "' and CP116='" & Mid(Combo1, iPos + 1) & "'"
      Else
         stCon = stCon & " and FT01||FT02='" & ChangeCustomerL(Combo1) & "' AND FT03 IS NULL"
         stConCP = stConCP & " and CP44='" & ChangeCustomerL(Combo1) & "' and CP116 IS NULL"
      End If
      
      stVTB = "select nvl(count(*),0) Q1 from caseprogress where CP01='CFP'" & stConCP
      strExc(0) = "select FT07,Q1" & _
      " From fagenttarget,(" & stVTB & ") X" & _
      " where  FT06='CFP'" & stCon
      
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         If RsTemp.Fields(0) <= RsTemp.Fields(1) Then
            If MsgBox("已達該代理人目標給案量，是否繼續給案？", vbYesNo + vbDefaultButton2) = vbNo Then
               bolRtn = False
            End If
         End If
      End If
      
   End If
   CheckCP44 = bolRtn
End Function

Private Sub Combo1_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Combo1_Validate(Cancel As Boolean)
   Dim strNo As String, iPos As Integer
   If Combo1.Text <> "" Then
      If CheckKeyIn(13) = -1 Then
         Cancel = True
      End If
      
      'Add by Morgan 2004/9/14 檢查客戶/代理人是否不再使用
      If Cancel = False Then
         strNo = Combo1.Text
         'Add by Morgan 2008/2/18 加聯絡人判斷
         iPos = InStr(Combo1.Text, "-")
         If iPos > 0 Then
            strNo = Left(Combo1.Text, iPos - 1)
         End If
         'end 2008/2/18
         
         If PUB_CheckStatus(strNo) = False Then Cancel = True
         
      End If
      
      If Cancel Then
         Combo1.SetFocus
      Else
         
         'Modify by Morgan 2008/9/30 因程序反應存檔時提醒太晚，故從txtvalidate 搬來
         '2008/08/19 By Toni 有代理人備註秀備註
         'Modified by Morgan 2012/3/7 發文都要顯示代理人備註--甄妮
         'If InStr(NewCasePtyList, txtCaseField(10)) > 0 And Combo1.Text <> "" Then
         'Modified by Morgan 2012/3/30 顯示依一次後,當代理人有變更時才再顯示
         'If Combo1.Text <> "" Then
         If Combo1.Text <> "" And Combo1.Tag <> Combo1 Then
            'Modify by Morgan 2008/9/30 需考慮聯絡人
            'strExc(0) = "select FA29 from Fagent where " & ChgFagent(Combo1.Text) & " and FA29 is not null"
            strExc(0) = "select FA29 from Fagent where " & ChgFagent(strNo) & " and FA29 is not null"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               MsgBox "" & RsTemp(0), vbExclamation, "代理人備註"
            End If
         End If
         Combo1.Tag = Combo1 'Added by Morgan 2012/3/30
         'end 2008/9/30
      End If
   End If
End Sub

Private Sub Form_Activate()
   Dim ii As Integer
   
   If m_bActived = True Then Exit Sub 'Add by Morgan 2010/1/27 不加則 unload 時會不斷的觸發
   If Me.Enabled = True Then
      'txtCaseField(0).SetFocus
      '2008/5/1 add by sonia CIP(113),CPA(114),CA(122),分割(307)改預設母案發年費之最新代理人且不跑建議代理人的功能
      If cp(10) = "113" Or cp(10) = "114" Or cp(10) = "122" Or cp(10) = "307" Then
         m_bActived = True
         If cp(10) = "113" Or cp(10) = "114" Or cp(10) = "122" Then
            strExc(0) = "select cp44||decode(cp116,null,null,'-'||cp116) from caseprogress where cp01 = '" & cp(1) & "' and cp02 = '" & cp(2) & "' and cp03 = '" & cp(3) & "' and cp04 = '00' and cp09<'C' and cp10 not in ('605','606','607') and cp44 is not null order by cp27 desc"
         Else
            strExc(0) = "select cp44||decode(cp116,null,null,'-'||cp116) from caseprogress,DivisionCase where dc01 = '" & cp(1) & "' and dc02 = '" & cp(2) & "' and dc03 = '" & cp(3) & "' and dc04 = '" & cp(4) & "' and dc05=cp01(+) and dc06=cp02(+) and dc07=cp03(+) and dc08=cp04(+) and cp09<'C' and cp10 not in ('605','606','607') and cp44 is not null order by cp27 desc"
         End If
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
         End If
      End If
      '2008/5/1 end
      
      'Added by Morgan 2024/5/31
      '日本新申請案發文時檢查相同申請人
      If txtCaseField(12) = "011" And InStr(NewCasePtyList, txtCaseField(10)) > 0 Then
         strExc(1) = "and instr(pa26||pa27||pa28||pa29||pa30,'" & Left(ChangeCustomerL(field(26)), 8) & "')"
         For intI = 2 To 5
            If field(25 + intI) <> "" Then
               strExc(1) = strExc(1) & "+instr(pa26||pa27||pa28||pa29||pa30,'" & Left(ChangeCustomerL(field(25 + intI)), 8) & "')"
            End If
         Next
         strExc(1) = strExc(1) & ">0"
         
         strExc(0) = "select distinct '' C1,cp44,RTRIM(FA05||' '||FA63||' '||FA64||' '||FA65)||' '||FA06||' '||FA04 FName" & _
            " from patent,caseprogress a,fagent" & _
            " where pa57||pa108||pa24||pa22 is null and pa09='011' and pa01='CFP'" & strExc(1) & _
            " and cp01(+)=pa01 and cp02(+)=pa02 and cp03(+)=pa03 and cp04(+)=pa04 and cp44 is not null and cp158>0 and cp159=0" & _
            " and not exists(select * from caseprogress b where cp01=pa01 and cp02=pa02 and cp03=pa03 and cp04=pa04 and a.cp44 is not null and a.cp158>cp158 and cp159=0)" & _
            " and exists(select * from caseprogress b where cp01=pa01 and cp02=pa02 and cp03=pa03 and cp04=pa04 and cp10 in ('101','102','103','307'))" & _
            " and fa01(+)=substr(cp44,1,8) and fa02(+)=substr(cp44,9) order by 1"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            Set frm880012.grdDataList.Recordset = RsTemp
            Set frm880012.fmParent = Me
            frm880012.iTyp = 7
            frm880012.Show vbModal
            If Me.Tag <> "" Then
               Combo1 = Me.Tag
               Combo1_Validate False
            End If
         End If
      End If
      'end 2024/5/31
      
      
      'Add by Morgan 2008/4/17
      If Combo1 = "" And m_bActived = False Then
         m_bActived = True
         ShowSuggest False
      End If
      'Added by Morgan 2012/3/30
      For ii = 3 To 7
         If txtCaseField(ii) <> "" Then
            CustMemoAlert txtCaseField(ii)
         End If
      Next
      If Combo1 <> "" Then Combo1_Validate False
      'end 2012/3/30
      'add by sonia 2019/7/30 PCT進入國家階段提醒
      If InStr(NewCasePtyList, txtCaseField(10)) > 0 And field(46) = "Y" Then
         MsgBox "本案為PCT國際申請案進入國家階段案件，請確認申請文件之PCT相關資料是否正確無誤。"
      End If
      'end 2019/7/30
      
      m_bActived = True 'Add by Morgan 2010/1/27
   End If
   
End Sub
Private Sub Form_Load()
Dim i As Integer, j As Integer

    MoveFormToCenter Me
    
   'Add By Sindy 2018/1/8
   m_strIR01 = frm050102_1.m_strIR01
   m_strIR02 = frm050102_1.m_strIR02
   m_strIR03 = frm050102_1.m_strIR03
   m_strIR04 = frm050102_1.m_strIR04
   If m_strIR01 <> "" Then
      Me.Caption = Me.Caption & "（信件編號:" & m_strIR01 & "-" & m_strIR03 & "）"
   End If
   '2018/1/8 END
   SSTab1.Tab = 0 'Added by Lydia 2021/05/25
   
    bolLeave = False
    intLeaveKind = 1
    txtCaseField(11) = strSrvDate(2)
    ReadAllData
    lblCaseProperty.BackColor = &H8000000F
    lblNation.BackColor = &H8000000F
    lblAgent.BackColor = &H8000000F
    'Modify by Amy 2014/04/23 由cmdok(0)_click搬過來 有國內案者預設關聯案發明人並修改
    'Add by Amy 2013/08/06 +新案判斷是否有國內案且已發文(排除FCP)
    If InStr(NewCasePtyList, txtCaseField(10)) > 0 Then
        bolFirstInventor = True
        cm(5) = "": cm(6) = "": cm(7) = "": cm(8) = ""
        strExc(0) = "Select CM05,CM06,CM07,CM08 From CaseMap,Caseprogress Where CM01='" & cp(1) & "' AND CM02='" & cp(2) & "' AND CM03='" & cp(3) & "' AND CM04='" & cp(4) & "' " & _
                         "AND CM10='0' AND CM05<>'FCP' And  CM05=CP01 And CM06=CP02 And CM07=CP03 And CM08=CP04 And CP27>0 And InStr('" & NewCasePtyList & "',CP10)>0"
        intI = 1
        Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
        If intI = 1 Then
            cm(5) = RsTemp.Fields("CM05")
            cm(6) = RsTemp.Fields("CM06")
            cm(7) = RsTemp.Fields("CM07")
            cm(8) = RsTemp.Fields("CM08")
            OrgInventor = strInventorNo
            DefInventor = "Y"
            CopyInventor cm(5) & cm(6) & cm(7) & cm(8), cp(1) & cp(2) & cp(3) & cp(4), True, DefInventor
            MsgBox "國內案 " & cm(5) & "-" & cm(6) & IIf(cm(7) & cm(8) = "000", "", "-" & cm(7) & "-" & cm(8)) & " 已發文！發明人資訊將會複製到本案！", vbExclamation
        End If
    End If
    'end 2014/04/23
    If PUB_ChkFileNP(cp(9)) Then MsgBox "下一程序已有提申或收達期限，若為重新發文時需要先刪除後才可作業！" 'Add by Morgan 2009/6/1
    txtCaseField(19) = "N" '預設不印傳真封面--慧汶
    
    'Added by Lydia 2021/05/25
    txtCP113 = ""
    If cp(113) <> "" Then txtCP113 = cp(113)
    'end 2021/05/25
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If bolLeave = False Then
   If MsgBox("你並未存檔，確定離開嗎?", vbYesNo + vbCritical + vbDefaultButton2) = vbNo Then
      Cancel = 1
   End If
End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    '不印副本貼紙
'    Set Printer = Printers(SeekPrint)
'    Printer.Orientation = SeekPrintL
   PUB_SendMailCache 'Add by Morgan 2007/3/23
    If intLeaveKind = 1 Then
        frm050102_1.Show
    ElseIf intLeaveKind = 0 Then
        Unload frm050102_1
    End If
    ShowEditForm 'Added by Morgan 2018/8/22
    
    'Set frm050102_3 = Nothing 'Removed by Morgan 2021/12/10 form2.0會有問題，改在呼叫時清除記憶體變數
End Sub

Private Sub lblCaseField_Change(Index As Integer)
Dim strTemp As String, bolIsChina As Boolean

Select Case Index
   Case 3
      'edit by nickc 2007/02/02 不用 dll 了
      'If objPublicData.GetPatentTrademarkKind(專利, lblCaseField(Index), strTemp, , txtCaseField(12)) = 1 Then
      If ClsPDGetPatentTrademarkKind(專利, lblCaseField(Index), strTemp, , txtCaseField(12)) = 1 Then
         lblTrademarkKind = strTemp
      End If
   Case 4
      'edit by nickc 2007/02/02 不用 dll 了
      'If objPublicData.GetStaff(lblCaseField(Index), strTemp) Then
      If ClsPDGetStaff(lblCaseField(Index), strTemp) Then
         lblSalesName = strTemp
      Else
         lblSalesName = ""
      End If
End Select
End Sub

'Morgan 2003/11/20
Private Sub Combo2_Click(Index As Integer)

   Dim i As Integer, strTmp As String
   
   If Combo2(Index) = "" Then
      For i = 0 To 2
         txtCaseField(i + (Index + 1) * 3 + 36) = ""
      Next
      Exit Sub
   End If
   
   strTmp = Mid(Combo2(Index).Text, InStr(Combo2(Index).Text, "-") + 1, 1)
   strExc(1) = "CU" & 39 + (Val(strTmp) - 1) * 3 & ",CU" & 40 + (Val(strTmp) - 1) * 3 & ",CU" & 41 + (Val(strTmp) - 1) * 3
   strExc(0) = "SELECT " & strExc(1) & " FROM CUSTOMER WHERE " & ChgCustomer(Left(Combo2(Index).Text, InStr(Combo2(Index).Text, "-") - 1))
   
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      For i = 0 To 2
         
         If Not IsNull(RsTemp.Fields(i)) Then
            txtCaseField(i + (Index + 1) * 3 + 36) = RsTemp.Fields(i)
         Else
            txtCaseField(i + (Index + 1) * 3 + 36) = ""
         End If
         
      Next
   End If
End Sub
'Add by Morgan 2004/9/23
Private Sub txtAD_Change(Index As Integer)
   SetEntity
End Sub

'Add by Morgan 2004/9/23
Private Sub txtAD_GotFocus(Index As Integer)
   'edit by nickc 2007/06/06 切換輸入法改用API
   'txtAD(Index).IMEMode = 2
   CloseIme
   TextInverse txtAD(Index)
End Sub
'Add by Morgan 2004/9/23
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
      MsgBox "減免身分不可空白！"
      Cancel = True
   End If
End Sub

Private Sub txtCaseField_KeyPress(Index As Integer, KeyAscii As MSForms.ReturnInteger)
Select Case Index
             Case 3, 4, 5, 6, 7, 8, 12, 16, 17, 19
                       KeyAscii = UpperCase(KeyAscii)
             'Add By Cheng 2002/07/30
             'Modify by Morgan 2004/9/27 加19
             'Modified by Lydia 2014/12/27 + 69 (DHL列印)
             Case 9, 13, 19, 69
               KeyAscii = UpperCase(KeyAscii)
               If KeyAscii <> 8 And KeyAscii <> 89 Then
                  KeyAscii = 0
               End If
               
            Case 18
               KeyAscii = UpperCase(KeyAscii)
               If KeyAscii <> 8 And KeyAscii <> Asc("Y") And KeyAscii <> Asc("N") Then
                  KeyAscii = 0
               End If
End Select
End Sub
'Add by Morgan 2004/9/23
Private Sub SetEntity()
   Dim i As Integer
   optChoose(0).Value = False: optChoose(1).Value = False: optChoose(2).Value = False
   For i = 1 To 5
      If txtAD(i).Text = "N" Then
         optChoose(0).Value = True
         Exit For
      '只要有未設定減免身分的公司申請人則不預設大小個體
      ElseIf txtAD(i).Enabled = True And txtAD(i).Text = "" Then
         Exit For
      End If
   Next
   '若五個申請人檢查完都不是大個體則為小個體
   If optChoose(2).Enabled = False Then 'Added by Morgan 2013/3/20 不可選微個體時才預設
      If optChoose(0).Value = False And i = 6 Then optChoose(1).Value = True
   End If
   
End Sub

'Add by Morgan 2004/9/23
Private Sub SetAD(ByVal i As Integer)
   txtAD(i).Enabled = False
   txtAD(i).Tag = ""
   txtAD(i).Text = ""
   'Modify by Morgan 2006/9/21 加法國
   'Modified by Morgan 2016/9/22 +印度040,菲律賓030--禧佩
   'Modified by Morgan 2018/11/6 法國新案沒有申請日 Ex:CFP-030712
   'Modified by Morgan 2023/3/24
   'If txtCaseField(i + 2) <> "" And (txtCaseField(12) = "101" Or txtCaseField(12) = "102" Or (txtCaseField(12) = "203" And (field(10) = "" Or DBDATE(field(10)) >= "20050901")) Or txtCaseField(12) = "040" Or txtCaseField(12) = "030") Then
   If txtCaseField(i + 2) <> "" And InStr(CFP_ChkEntity, txtCaseField(12)) > 0 Then
   'end 2023/3/24
      txtAD(i).Text = PUB_GetAD03(txtCaseField(i + 2), txtCaseField(12).Text)
      txtAD(i).Tag = txtAD(i).Text
      txtAD(i).Enabled = True
   End If
End Sub

Private Sub txtCaseField_Change(Index As Integer)
Select Case Index
             Case 3, 4, 5, 6, 7
                        lblPetition(Index - 3).Caption = ""
                        SetAD Index - 2
             Case 10
                        lblCaseProperty = ""
             Case 12
                        If txtCaseField(Index) = EPC指定國家 Then
                           cmdCountry.Visible = True
                           'edit by nickc 2007/02/02 不用 dll 了
                           'objPublicData.ReadCountry intCaseKind, cp(), strCountry
                           ClsPDReadCountry intCaseKind, cp(), strCountry
                        Else
                           cmdCountry.Visible = False
                           strCountry = ""
                        End If
                        lblNation = ""
                        'Add by Morgan 2004/9/23
                        SetAD 1
                        SetAD 2
                        SetAD 3
                        SetAD 4
                        SetAD 5
                        '2004/9/23 end
             'Modified by Lydia 2014/12/27 + DHL
             Case 9
                       If txtCaseField(9) = "Y" Then
                          txtCaseField(69) = ""
                       End If
             Case 69
                       If txtCaseField(69) = "Y" Then
                          txtCaseField(9) = ""
                       End If
                        
End Select
End Sub

Private Sub txtCaseField_LostFocus(Index As Integer)
   Dim strTemp As String, strTemp1 As String, strCusTemp As String

   CloseIme
   
   Select Case Index
   Case 10 '案件性質
      'Modify by Morgan 2009/8/18 CheckKeyIn已有的功能不再重複
      'If ClsPDGetCaseProperty(cp(1), txtCaseField(Index), strTemp) Then
      '   lblCaseProperty.Caption = strTemp
      '   If ClsPDGetCaseDelayDay(cp(1), txtCaseField(12), txtCaseField(10), , strTemp, strTemp1) Then
      '      If strTemp <> "" And strTemp <> "0" And txtCaseField(11) <> "" Then
      '         strTemp = CompDate(2, Val(strTemp), TransDate(txtCaseField(11), 2))
      '         txtCaseField(15) = TransDate(strTemp, 1)
      '      Else
      '         txtCaseField(15) = ""
      '      End If
      '      If strTemp1 <> "" And txtCaseField(11) <> "" Then
      '         strTemp1 = CompDate(2, Val(strTemp1), TransDate(txtCaseField(11), 2))
      '         '92.2.18 CANCEL BY SONIA
      '         'txtCaseField(14) = TransDate(strTemp1, 1)
      '         txtCaseField(14) = ""
      '         '92.2.18 END
      '      Else
      '         txtCaseField(14) = ""
      '      End If
      '   Else
      '      txtCaseField(14) = ""
      '      txtCaseField(15) = ""
      '   End If
      'End If
      If txtCaseField(Index).Tag <> txtCaseField(Index) Then
         SetChkDate
         txtCaseField(Index).Tag = txtCaseField(Index)
      End If
      'end 2009/8/18
   Case 11 '發文日
      'Modify by Morgan 2009/8/18
      'txtCaseField_LostFocus 10
      If txtCaseField(Index).Tag <> txtCaseField(Index) Then
         SetChkDate
         txtCaseField(Index).Tag = txtCaseField(Index)
      End If
      'end 2009/8/18
   Case 12 '申請國家
      If intCaseKind <= "4" Then
         lblCaseField_Change 3
         'Modify by Morgan 2009/8/18
         'txtCaseField_LostFocus 10
         If txtCaseField(Index).Tag <> txtCaseField(Index) Then
            SetChkDate
            txtCaseField(Index).Tag = txtCaseField(Index)
         End If
         'end 2009/8/18
      End If
   End Select
End Sub

Private Sub txtCaseField_Validate(Index As Integer, Cancel As Boolean)
   Select Case Index
             Case 11
                        If CheckKeyIn(Index) <> -1 Then
                           CheckKeyIn 10
                        Else
                           Cancel = True
                        End If
             Case 12
                        If CheckKeyIn(Index) <> -1 Then
                           'Add By Cheng 2002/07/30
                           If intCaseKind <= "4" Then
                              
                              lblCaseField_Change 3
                              CheckKeyIn 10
                           
                           End If
                        Else
                           Cancel = True
                        End If
             'Modified by Lydia 2014/12/27 + 69 (DHL列印)
             Case 9, 13, 69
                        If CheckKeyIn(16) <> 1 Then
                           Cancel = True
                        End If
             Case 21, 22
                        cmdOK(0).Default = True
                        cmdOK(0).CausesValidation = True
             Case Else
                        If CheckKeyIn(Index) = -1 Then
                           Cancel = True
                        End If
   End Select
   ' 90.12.07 modify by louis (加檢查文字字串的長度)
   If CheckFieldLength(Index) = False Then
      Cancel = True
   End If
   'Add By Cheng 2002/08/23
   If Cancel = False Then
      Select Case Index
      Case 3
         If m_strCust1 <> Me.txtCaseField(Index).Text Then
            If Not PUB_EditCustOk(Me.lblCaseField(0).Caption, field(1), field(2), field(3), field(4)) Then Cancel = True
         End If
      Case 4
         If m_strCust2 <> Me.txtCaseField(Index).Text Then
            If Not PUB_EditCustOk(Me.lblCaseField(0).Caption, field(1), field(2), field(3), field(4)) Then Cancel = True
         End If
      Case 5
         If m_strCust3 <> Me.txtCaseField(Index).Text Then
            If Not PUB_EditCustOk(Me.lblCaseField(0).Caption, field(1), field(2), field(3), field(4)) Then Cancel = True
         End If
      Case 6
         If m_strCust4 <> Me.txtCaseField(Index).Text Then
            If Not PUB_EditCustOk(Me.lblCaseField(0).Caption, field(1), field(2), field(3), field(4)) Then Cancel = True
         End If
      Case 7
         If m_strCust5 <> Me.txtCaseField(Index).Text Then
            If Not PUB_EditCustOk(Me.lblCaseField(0).Caption, field(1), field(2), field(3), field(4)) Then Cancel = True
         End If
      End Select
   End If
   
   'Add by Morgan 2004/9/14 檢查客戶/代理人是否不再使用
   If Cancel = False And InStr("3,4,5,6,7", Format(Index)) > 0 Then
      If PUB_CheckStatus(txtCaseField(Index).Text) = False Then Cancel = True
      
      'Added by Morgan 2012/3/30
      If Cancel = False Then
         If txtCaseField(Index).Tag <> txtCaseField(Index).Text Then
            txtCaseField(Index).Tag = txtCaseField(Index).Text
            If txtCaseField(Index).Text <> "" Then
               CustMemoAlert txtCaseField(Index)
            End If
         End If
      End If
      'end 2012/3/30
   End If

   If Cancel Then txtCaseField_GotFocus (Index)
End Sub
Private Function CheckKeyIn(intIndex As Integer) As Integer
Dim strTemp As String, strTemp1 As String, strCusTemp As String, bolIsChina As Boolean
Dim i As Integer
'Add By Cheng 2002/12/18
Dim strRPSTTV
'Add by Morgan 2004/7/27
Dim stAppNation As String  '申請人國籍

CheckKeyIn = -1
Select Case intIndex
             Case 2
                        If txtCaseField(0) = "" And txtCaseField(1) = "" And txtCaseField(2) = "" Then
                           ShowMsg MsgText(1031)
                           intIndex = 0
                           CheckKeyIn = 0
                        Else
                           CheckKeyIn = 1
                        End If
             Case 3
                        strCusTemp = txtCaseField(intIndex)
                        'Modify by Morgan 2004/7/27
                        'If objPublicData.GetCustomerNameAndAddress(strCusTemp, strTemp, strExc(1), strExc(2), strExc(3)) Then
                        'Modify by Morgan 2011/1/11
                        'If PUB_GetCustomerNameAndAddress(strCusTemp, strTemp, strExc(1), strExc(2), strExc(3), stAppNation) Then
                        If ClsPDGetCustomerNameAndAddress(strCusTemp, strTemp, strExc(1), strExc(2), strExc(3), stAppNation) Then
                           lblAppNation(intIndex - 3).Caption = stAppNation
                           
                           txtCaseField(intIndex) = strCusTemp
                           lblPetition(intIndex - 3).Caption = strTemp
                           '顯示申請人地址
                           For i = 1 To 3
                              txtCaseField(22 + i) = strExc(i)
                           Next
                           
'Add by Morgan 2003/12/18
                           Call renewRepCombo(intIndex, intIndex)
                           
                           '910816 Sieg
                           If cp(60) <> "" And InStr(ChangeCustomerL(field(26)), ChangeCustomerL(strCusTemp)) = 0 Then
                              strExc(1) = field(1)
                              strExc(2) = field(2)
                              strExc(3) = field(3)
                              strExc(4) = field(4)
                              strExc(5) = cp(60)
                              strExc(6) = txtCaseField(intIndex)
                              strExc(7) = strTemp
                              'edit by nickc 2007/02/05 不用 dll 了
                              'If Not objLawDll.UpdAcc0k0(strExc()) Then
                              If Not ClsLawUpdAcc0k0(strExc()) Then
                                 lblPetition(intIndex - 3).Caption = ""
                              Else
                                 CheckKeyIn = 1
                              End If
                           Else
                              CheckKeyIn = 1
                           End If
                        End If
                        
             Case 4, 5, 6, 7
             
'Add by Morgan 2003/12/19

                        Call renewRepCombo(intIndex, intIndex)
'Add End 2003/12/19

                        '若未輸入申請人
                        If txtCaseField(intIndex) = "" Then
                           CheckKeyIn = 1
                           '清除申請人地址
                           For i = 1 To 3
                              txtCaseField(25 + i + (intIndex - 4) * 3) = ""
                           Next
                           
                           '清除代表人
                           For i = 1 To 6
                              txtCaseField(44 + i + (intIndex - 4) * 6) = ""
                           Next
                           'Modify end 2004/1/6
                           
                           lblPetition(intIndex - 3).Caption = ""
                        '若有輸入申請人
                        Else
                           strCusTemp = txtCaseField(intIndex)
                           'Modify by Morgan 2004/7/27
                           'If objPublicData.GetCustomerNameAndAddress(strCusTemp, strTemp, strExc(1), strExc(2), strExc(3)) Then
                           'Modify by Morgan 2011/1/11
                           'If PUB_GetCustomerNameAndAddress(strCusTemp, strTemp, strExc(1), strExc(2), strExc(3), stAppNation) Then
                           If ClsPDGetCustomerNameAndAddress(strCusTemp, strTemp, strExc(1), strExc(2), strExc(3), stAppNation) Then
                              lblAppNation(intIndex - 3).Caption = stAppNation
                              
                              txtCaseField(intIndex) = strCusTemp
                              lblPetition(intIndex - 3).Caption = strTemp
                              '取得申請人地址
                              For i = 1 To 3
                                 txtCaseField(25 + i + (intIndex - 4) * 3) = strExc(i)
                              Next

                              CheckKeyIn = 1
                           End If
                        End If
             Case 8, 17
                        If txtCaseField(intIndex) = "" Or txtCaseField(intIndex) = "N" Then
                           CheckKeyIn = 1
                        Else
                           ShowMsg MsgText(1038)
                        End If
             Case 9
                        If txtCaseField(intIndex) = "" Then
                           CheckKeyIn = 1
                        Else
                           If txtCaseField(intIndex).Text <> "Y" Then
                              ShowMsg MsgText(1054)
                           Else
                              CheckKeyIn = 1
                           End If
                        End If
             Case 10
                        'Modify By Cheng 2003/02/17
                        '移至Lost_Focus
                       If txtCaseField(12) = 大陸國家代號 Then bolIsChina = True Else bolIsChina = False
                       'edit by nickc 2007/02/02 不用 dll 了
                       'If objPublicData.GetCaseProperty(cp(1), txtCaseField(intIndex), strTemp, bolIsChina) Then
                       If ClsPDGetCaseProperty(cp(1), txtCaseField(intIndex), strTemp, bolIsChina) Then
                           lblCaseProperty.Caption = strTemp
                           CheckKeyIn = 1
                        End If
             Case 11
                         If CheckIsTaiwanDate(txtCaseField(intIndex).Text) Then
                            CheckKeyIn = 1
                        End If
             Case 12
                        'edit by nickc 2007/02/02 不用 dll 了
                        'If objPublicData.GetNation(txtCaseField(intIndex).Text, strTemp) Then
                        If ClsPDGetNation(txtCaseField(intIndex).Text, strTemp) Then
                           lblNation.Caption = strTemp
                           CheckKeyIn = 1
                        End If
             Case 13 '代理人
                        lblAgent.Caption = ""
                        If Combo1.Text = "" Then
                           MsgBox "代理人欄不可空白!!!", vbExclamation
                        Else
                           strCusTemp = Combo1
                           'Add by Morgan 2008/2/14 加判斷是否為聯絡人
                           If InStr(strCusTemp, "-") > 0 Then
                              If ClsPDGetContact(strCusTemp, strTemp) Then
                                 Combo1 = strCusTemp
                                 lblAgent.Caption = strTemp
                                 CheckKeyIn = 1
                              End If
                           Else
                              'edit by nickc 2007/02/02 不用 dll 了
                              'If objPublicData.GetAgent(strCusTemp, strTemp) Then
                              If ClsPDGetAgent(strCusTemp, strTemp) Then
                                 Combo1 = strCusTemp
                                 lblAgent.Caption = strTemp
                                 CheckKeyIn = 1
                              End If
                           End If
                        End If
             'Modify by Morgan 2009/5/12 +20
             Case 14, 15, 20
                         If txtCaseField(intIndex) = "" Then
                            CheckKeyIn = 1
                         Else
                            If CheckIsTaiwanDate(txtCaseField(intIndex).Text) Then
                               CheckKeyIn = 1
                            End If
                        End If
             Case 16
                        If txtCaseField(intIndex) = "" Or txtCaseField(intIndex) = "Y" Then
                           CheckKeyIn = 1
                        Else
                           ShowMsg MsgText(9174)
                        End If
             Case 18
                        If txtCaseField(intIndex) = "N" Or txtCaseField(intIndex) = "Y" Or txtCaseField(intIndex) = "" Then
                           CheckKeyIn = 1
                        Else
                           ShowMsg MsgText(9177)
                        End If
             Case 38
                        If chkChoose(0).Value And txtCaseField(intIndex) = "" Then
                           ShowMsg MsgText(1055)
                           CheckKeyIn = -1
                        Else
                           CheckKeyIn = 1
                        End If
             Case Else
                        CheckKeyIn = 1
End Select
End Function
Private Sub txtCaseField_GotFocus(Index As Integer)
If Index = 3 Then
   If txtCaseField(0) = "" And txtCaseField(1) = "" And txtCaseField(2) = "" Then
      txtCaseField(0).SetFocus
      Exit Sub
   End If
End If
txtCaseField(Index).SelStart = 0
txtCaseField(Index).SelLength = Len(txtCaseField(Index).Text)
'儲存未修改前之值至Tag中,供再確認時使用
txtCaseField(Index).Tag = txtCaseField(Index)
Select Case Index
   Case 21, 22
      cmdOK(0).Default = False
      cmdOK(0).CausesValidation = False
End Select
'Add By Cheng 2002/07/30
Select Case Index
Case 23, 25, 26, 28, 29, 31, 32, 34, 35, 37 '申請人地址(中,日)
   'edit by nickc 2007/06/06 切換輸入法改用API
   'Me.txtCaseField(Index).IMEMode = 1
   OpenIme
Case 24, 27, 30, 33, 36 '申請人地址(英)
   'edit by nickc 2007/06/06 切換輸入法改用API
   'Me.txtCaseField(Index).IMEMode = 2
   CloseIme
End Select
End Sub
Private Sub cmdCountry_Click()
ModifyAssignCountry strCountry
End Sub
Private Sub cmdPriority_Click()
'Modify by Amy 2014/04/14 +strPriority5
ModifyPriority strPriority1, strPriority2, strPriority3, field(8), , field(1) & field(2) & field(3) & field(4), field(9), , strPriority4, strPriority5
End Sub

Private Function CheckFieldLength(ByVal nIndex) As Boolean
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   CheckFieldLength = True
   
   'Modified by Morgan 2012/11/22 欄位長度有調整改抓 MaxLength 屬性
   If txtCaseField(nIndex).MaxLength > 0 Then
      If StrLength(txtCaseField(nIndex)) > txtCaseField(nIndex).MaxLength Then
         CheckFieldLength = False
      End If
   End If
   
   If CheckFieldLength = False Then
      strTit = "檢核資料"
      strMsg = "輸入的資料內容太長"
      nResponse = MsgBox(strMsg, vbOKOnly + vbCritical, strTit)
   End If
End Function

'Add By Cheng 2002/05/22
Private Function TxtValidate() As Boolean
Dim objTxt
Dim ii As Integer
Dim txtADN As Boolean 'Add by Amy 2013/06/06
Dim strNP07 As String
Dim Cancel As Boolean
Dim arrInv
txtADN = False 'Add by Amy 2013/06/06
TxtValidate = False

   'Added by Morgan 2021/12/6 檢查畫面輸入欄位是否含有Unicode文字
   If PUB_ChkUniText(Me, , True, "TextBox") = False Then
       Exit Function
   End If
   
   'Add by Morgan 2010/5/24
   '檢查相關案是否有保密審查未准
   If PUB_Exists430NotPassed(cp) Then
      Exit Function
   End If
   
   'add by nickc 2008/05/01
   If IsDebt(field(9), cp(9)) Then
        MsgBox "未收款且無 預定收款日 請轉告智權同仁！！", vbOKOnly, "警告！禁止發文！"
        Exit Function
   End If
For Each objTxt In Me.txtCaseField
   If objTxt.Enabled = True Then
      Cancel = False
      txtCaseField_Validate objTxt.Index, Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
Next

'Add by Morgan 2004/9/14
If Combo1.Enabled = True Then
   Cancel = False
   Combo1_Validate Cancel
   If Cancel = True Then
      Combo1.SetFocus
      Exit Function
   End If
End If

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
      'Add by Amy 2013/06/06
      If txtAD(ii).Text = "N" Then
         txtADN = True
      End If
      'end 2013/06/06
   End If
Next

'Add by Amy 2013/06/06
'Modified by Morgan 2024/12/10 個體別順序會因國家有所不同,且客戶設定目前只設定是否可減免,故只能用大個體中文判斷(非大個體都是可減免)
'If (txtADN And (optChoose(1).Value = True Or optChoose(2).Value = True)) Or (txtADN = False And optChoose(0).Value = True) Then
'   If MsgBox("減免身份與個體別不一致，是否要繼續發文？", vbYesNo + vbDefaultButton2) = vbNo Then
strExc(8) = ""
If optChoose(0).Value = True Then
   strExc(8) = optChoose(0).Caption
ElseIf optChoose(1).Value = True Then
   strExc(8) = optChoose(1).Caption
ElseIf optChoose(2).Value = True Then
   strExc(8) = optChoose(2).Caption
End If
strExc(9) = ""
If txtADN Then strExc(9) = "大個體"
If (strExc(9) = "大個體" Or strExc(8) = "大個體") And strExc(9) <> strExc(8) Then
'end 2024/12/10
   If MsgBox("申請人減免身份與案件個體別不一致，是否要繼續發文？", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then
      Exit Function
   End If
End If
'end 2013/0606
   


'Add by Morgan 2007/8/27  PCT進國家階段若有年費未繳且未收文時發Mail給慧汶並新增下一程序年費期限
If field(46) = "Y" And field(10) <> "" Then
   If PUB_CheckAnnuity(field(8), field(9), DBDATE(field(10)), strNP07) = True Then
      strExc(0) = "select * from caseprogress where cp01='" & field(1) & "' and cp02='" & field(2) & "'" & _
         " and cp03='" & field(3) & "' and cp04='" & field(4) & "' and cp10='" & strNP07 & "' and cp57 is null"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 0 Then
         If MsgBox("本案為PCT發明案且需繳交年費而未收文，是否要繼續發文？", vbYesNo + vbDefaultButton2) = vbNo Then
            Exit Function
         End If
      End If
   End If
End If
'end 2007/8/27
               
   'Add by Morgan 2009/5/12
   If txtCaseField(20) <> "" And txtCaseField(14) <> "" Then
      MsgBox "已有指定提申日不可再輸入最終提申期限！"
      txtCaseField(14).SetFocus
      Exit Function
   End If
   If cp(7) <> "" Then
      If txtCaseField(14) <> "" Then
         If Val(DBDATE(txtCaseField(14))) > Val(DBDATE(cp(7))) Then
            MsgBox "最終提申期限不可晚於法定期限！"
            txtCaseField(14).SetFocus
            Exit Function
         End If
      ElseIf txtCaseField(20) <> "" Then
         If Val(DBDATE(txtCaseField(20))) > Val(DBDATE(cp(7))) Then
            MsgBox "指定提申日不可晚於法定期限！"
            txtCaseField(20).SetFocus
            Exit Function
         End If
      End If
   End If
   'end 2009/5/12
   
   'Add by Morgan 2010/5/3 申請人有變更且未重新點選發明人資料時清除原發明人資料
   'Modify by Morgan 2010/6/24 申請案才要
   If InStr(NewCasePtyList, txtCaseField(10)) > 0 Then
      strExc(1) = txtCaseField(3)
      For intI = 1 To 4
         If txtCaseField(intI + 3) <> "" Then
            strExc(1) = strExc(1) & "," & txtCaseField(intI + 3)
         End If
      Next intI
      'Modify By Sindy 2014/11/13
      If strSrvDate(1) >= 專利發明人檔啟用日 Then
         Call PUB_ChkInventor(strInventorNo, strExc(1), True)
      Else
      '2014/11/13 END
         If PUB_ChkInventor(strInventorNo, strExc(1), True) = False Then
            arrInv = Split(strInventorNo, ",")
            For intI = 0 To 9
               If intI <= UBound(arrInv) Then
                  field(60 + intI) = arrInv(intI)
               Else
                  field(60 + intI) = ""
               End If
            Next
         End If
      End If
   End If
   'end 2010/5/3
   
   'Add By Sindy 2014/11/13
   intInventorCnt = 0 '發明人筆數
   arrInv = Split(strInventorNo, ",")
   If UBound(arrInv) >= 0 Then
      intInventorCnt = UBound(arrInv) + 1
   End If
   '2014/11/13 END
   'Add by Morgan 2010/6/17
   '一人申請時(沒有第二申請人)
   If txtCaseField(4) = "" Then
      '發明人不只一人(有第二發明人)
      'Modify By Sindy 2014/11/13
      'If field(61) <> "" Then
      If intInventorCnt > 1 Then
      '2014/11/13 END
         If txtCaseField(18) = "Y" Then
            strExc(0) = MsgBox("申請人與發明人數不同，【申請人與發明人是否相同】欄位是否改為【N】？", vbYesNoCancel + vbDefaultButton3)
            If strExc(0) = vbCancel Then
               txtCaseField(18).SetFocus
               Exit Function
            ElseIf strExc(0) = vbYes Then
               txtCaseField(18) = "N"
            End If
         End If
      Else
         '一人申請時檢查若為個人且發明人同名時設定為 "Y"
         strExc(1) = ""
         strExc(2) = "" 'id
         strExc(3) = "" '中文名
         strExc(0) = "select cu15,cu11,cu04 from customer where cu01='" & Left(txtCaseField(3) & "000", 8) & "' and cu02='" & Mid(txtCaseField(3) & "000", 9, 1) & "'"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            strExc(1) = "" & RsTemp("cu15")
            strExc(2) = "" & RsTemp("cu11")
            strExc(3) = "" & RsTemp("cu04")
            '非個人
            If strExc(1) <> "0" Then
               If txtCaseField(18) = "Y" Then
                  strExc(0) = MsgBox("申請人不為個人，【申請人與發明人是否相同】欄位是否改為【N】？", vbYesNoCancel + vbDefaultButton3)
                  If strExc(0) = vbCancel Then
                     txtCaseField(18).SetFocus
                     Exit Function
                  ElseIf strExc(0) = vbYes Then
                     txtCaseField(18) = "N"
                  End If
               End If
            '個人
            Else
               If txtCaseField(18) <> "Y" Then
                  strExc(4) = "" 'id
                  strExc(5) = "" '中文名
                  'Modify By Sindy 2014/11/13
                  'strExc(0) = "select in03,in04 from inventor where in01='" & Left(field(60), 8) & "' and in02='" & Mid(field(60), 9) & "'"
                  If intInventorCnt > 0 Then
                     strExc(0) = "select in03,in04 from inventor where in01='" & Left(arrInv(0), 8) & "' and in02='" & Mid(arrInv(0), 9) & "'"
                  '2014/11/13 END
                     intI = 1
                     Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                     If intI = 1 Then
                        strExc(4) = "" & RsTemp("in03")
                        strExc(5) = "" & RsTemp("in04")
                        'id相同 或 無id但名稱相同
                        If ((strExc(2) <> "" And strExc(2) = strExc(4)) Or (strExc(2) = "" And strExc(3) = strExc(5))) Then
                           strExc(0) = MsgBox("申請人與發明人相同，【申請人與發明人是否相同】欄位是否改為【Y】？", vbYesNoCancel + vbDefaultButton3)
                           If strExc(0) = vbCancel Then
                              txtCaseField(18).SetFocus
                              Exit Function
                           ElseIf strExc(0) = vbYes Then
                              txtCaseField(18) = "Y"
                           End If
                        End If
                     End If
                  End If
               End If
            End If
         End If
      End If
   End If
   '韓國案若未設定時提醒
   If txtCaseField(12) = "012" And txtCaseField(18) = "" Then
      If txtCaseField(18) = "" Then
         MsgBox "韓國申請案必須設定【申請人與發明人是否相同】欄位！"
         txtCaseField(18).SetFocus
         Exit Function
      End If
   End If
   
   'Add by Morgan 2011/2/23
   '檢查申請人順序
   If txtCaseField(3).Enabled = True Then
      If (txtCaseField(4) <> "" And txtCaseField(3) = "") Or _
         (txtCaseField(5) <> "" And txtCaseField(4) = "") Or _
         (txtCaseField(6) <> "" And txtCaseField(5) = "") Or _
         (txtCaseField(7) <> "" And txtCaseField(6) = "") Then
         MsgBox "請依順序輸入申請人！"
         If txtCaseField(3) = "" Then txtCaseField(3).SetFocus: Exit Function
         If txtCaseField(4) = "" Then txtCaseField(4).SetFocus: Exit Function
         If txtCaseField(5) = "" Then txtCaseField(5).SetFocus: Exit Function
         If txtCaseField(6) = "" Then txtCaseField(6).SetFocus: Exit Function
         Exit Function
      End If
   End If
   
   'Added by Morgan 2013/3/20
   'Modified by Morgan 2023/9/11 美國才要--玫音
   'If OptChoose(2).Value = True Then
   If optChoose(2).Value = True And field(9) = "101" Then
   'end 2023/9/11
      For intI = 0 To 4
         If field(26 + intI) <> "" Then
            'Modified by Morgan 2021/4/16 發文要以發文件數計算
            If PUB_CheckMicroEntity(field(26 + intI), 1, 2, field(1) & field(2) & field(3) & field(4)) = False Then
               If SSTab1.Tab <> 0 Then SSTab1.Tab = 0
               Exit Function
            End If
         End If
      Next
      
      'Modify By Sindy 2014/11/13
      If strSrvDate(1) >= 專利發明人檔啟用日 Then
         If intInventorCnt > 0 Then
            For intI = 0 To intInventorCnt - 1
               'Modified by Morgan 2021/4/16 發文要以發文件數計算
               If PUB_CheckMicroEntity(arrInv(intI), 3, 2, field(1) & field(2) & field(3) & field(4)) = False Then
                  If SSTab1.Tab <> 0 Then SSTab1.Tab = 0
                  Exit Function
               End If
            Next intI
         End If
      Else
      '2014/11/13 END
         For intI = 0 To 9
            If field(60 + intI) <> "" Then
               'Modified by Morgan 2021/4/16 發文要以發文件數計算
               If PUB_CheckMicroEntity(field(60 + intI), 3, 2, field(1) & field(2) & field(3) & field(4)) = False Then
                  If SSTab1.Tab <> 0 Then SSTab1.Tab = 0
                  Exit Function
               End If
            End If
         Next intI
      End If
   End If
   'end 2013/3/20
   
'Added by Morgan 2018/9/6
'若系統不出指示信時判斷是否有工程師的指示信要寄送
m_bolEngLetter = False
If txtCaseField(17) = "N" Then
   If PUB_EngLtrChk(cp(9), txtCaseField(11).Text, m_bolEngLetter) = False Then
      Exit Function
   End If
End If
'end 2018/9/6

'Added by Morgan 2018/9/12 CFP電子化-接洽單檢查
If strSrvDate(1) >= CFP第一階段電子化啟用日 Then
   If cp(9) < "B" And Left(cp(12), 1) <> "F" Then
      If PUB_CheckPDF3(cp(1), cp(2), cp(3), cp(4)) = False Then
         Exit Function
      End If
   End If
End If
'end 2018/9/12

'Added by Morgan 2019/4/23
'日本實審發文要設定減免身分
m_str416CP81 = ""
'Modified by Morgan 2019/7/3 +分割307
If field(9) = "011" And (cp(10) = "101" Or cp(10) = "307") And chkChoose(3).Value Then
   For ii = 1 To 5
      If field(25 + ii) <> "" Then
         strExc(1) = PUB_GetAD03(field(25 + ii), "011")
         If strExc(1) = "" Then
            'Modified by Morgan 2019/6/19 改詢問是否不可減免,若是則系統自動設定--禧佩
            'MsgBox "申請人 " & field(25 + ii) & " " & GetCustomerNameAndState(field(25 + ii)) & " 尚未設定減免身分不可發文！", vbCritical, "日本實審發文減免身分檢查"
            'Exit Function
            If MsgBox("申請人【" & field(25 + ii) & " " & GetCustomerNameAndState(field(25 + ii)) & "】尚未設定減免身分！" & vbCrLf & vbCrLf & "是否要設定為【不可減免】？", vbYesNo + vbDefaultButton2 + vbExclamation, "日本實審發文減免身分檢查") = vbYes Then
               PUB_SetNoDisc field(25 + ii), field(9)
               m_str416CP81 = "N"
            Else
               Exit Function
            End If
            'end 2019/6/19
         ElseIf m_str416CP81 <> "N" Then
            m_str416CP81 = strExc(1)
         End If
      End If
   Next
End If
'end 2019/4/23

'Added by Lydia 2021/05/25 ACS智財顧問專業分配比例管制：有相關卷號(CaseRelation1)為ACS且曾有收文智財顧問112
If Pub_ChkACS112isNull(field(1), field(2), field(3), field(4), txtCP113) = True Then
      txtCP113.SetFocus
      txtCP113_GotFocus
      Exit Function
End If
'end 2021/05/25
   
TxtValidate = True
End Function

'Add By Cheng 2002/07/30
Private Sub ShowAudit()
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset

   'Added by Morgan 2020/3/10 馬來西亞判斷若無澳洲.英國.美國.EPC.日本.韓國及PCT相對應多國案時可一併提實審
   Dim bol018 As Boolean
   
   bol018 = False
   If field(9) = "018" Then
      strExc(0) = "select * from caserelation,patent" & _
         " where cr01='" & cp(1) & "' and cr02='" & cp(2) & "' and cr03='" & cp(3) & "' and cr04='" & cp(4) & "'" & _
         " and pa01(+)=cr05 and pa02(+)=cr06 and pa03(+)=cr07 and pa04(+)=cr08" & _
         " and instr('015,201,101,221,011,012,056',pa09)>0"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         bol018 = True
      End If
   End If
   'end 2020/3/10

   'Modify by Morgan 2004/9/9 加智利 126
   'Modified by Morgan 2020/3/10 馬來西亞有相對應多國案時不可一併提實審
   'If field(9) = "018" Or field(9) = "019" Or field(9) = "221" Or field(9) = "126" Then
   If (field(9) = "018" And bol018 = True) Or field(9) = "019" Or field(9) = "221" Or field(9) = "126" Then
   'end 2020/3/10
      Me.chkChoose(3).Enabled = False
      Me.chkChoose(3).Value = False
   Else
      Me.chkChoose(3).Enabled = True
      '92.1.21 modify by sonia
      'strSQLA = "Select * From CASEPROGRESS WHERE " & ChgCaseprogress(cp(1) & cp(2) & cp(3) & cp(4)) & " AND CP10='416' AND CP27 IS NULL AND CP57 IS NULL "
      StrSQLa = "Select * From CASEPROGRESS WHERE " & ChgCaseprogress(cp(1) & cp(2) & cp(3) & cp(4)) & " AND CP10='416' AND CP57 IS NULL "
      '92.1.21 end
      rsA.CursorLocation = adUseClient
      rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
      If rsA.RecordCount > 0 Then
         Me.chkChoose(3).Enabled = True
         Me.chkChoose(3).Value = True
      End If
      If rsA.State <> adStateClosed Then rsA.Close
      Set rsA = Nothing
   End If

   '92.1.22 add by sonia
   Me.chkChoose(4).Enabled = True
   'Modify by Morgan 2006/5/30 IDS的案件性質為214
   'StrSQLa = "Select * From CASEPROGRESS WHERE " & ChgCaseprogress(cp(1) & cp(2) & cp(3) & cp(4)) & " AND CP10='207' AND CP57 IS NULL "
   StrSQLa = "Select * From CASEPROGRESS WHERE " & ChgCaseprogress(cp(1) & cp(2) & cp(3) & cp(4)) & " AND CP10='214' AND CP57 IS NULL "
   rsA.CursorLocation = adUseClient
   rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
   If rsA.RecordCount > 0 Then
      Me.chkChoose(4).Enabled = True
      Me.chkChoose(4).Value = True
   End If
   If rsA.State <> adStateClosed Then rsA.Close
   Set rsA = Nothing
   '92.1.22 end
   '92.8.21 add by sonia
   Me.chkChoose(5).Enabled = True
   StrSQLa = "Select * From CASEPROGRESS WHERE " & ChgCaseprogress(cp(1) & cp(2) & cp(3) & cp(4)) & " AND CP10='427' AND CP57 IS NULL "
   rsA.CursorLocation = adUseClient
   rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
   If rsA.RecordCount > 0 Then
      Me.chkChoose(5).Enabled = True
      Me.chkChoose(5).Value = True
   End If
   If rsA.State <> adStateClosed Then rsA.Close
   Set rsA = Nothing
   '92.8.21 end
   
   'Add by Morgan 2011/4/14
   '美國申請程序(101,103,118,307)發文時若有讓渡當日或未發文時預設一併送件
   chkChoose(6).Enabled = False
   If field(9) = "101" And (cp(10) = "101" Or cp(10) = "103" Or cp(10) = "113" Or cp(10) = "307") Then
      chkChoose(6).Enabled = True
      strExc(0) = "select cp09 from caseprogress where " & ChgCaseprogress(cp(1) & cp(2) & cp(3) & cp(4)) & _
         " AND CP10='701' AND CP57 IS NULL and (cp27 is null or cp27=" & strSrvDate(1) & ")"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         chkChoose(6).Value = True
      End If
   End If
   
End Sub

'Add by Morgan 2006/7/31
'發E-Mail給承辦人
Private Sub MailToPromoter(strMailCp09 As String)
   Dim arrMailCP09
   Dim ii As Integer
   If strMailCp09 <> "" Then
      arrMailCP09 = Split(strMailCp09, ",")
      For ii = LBound(arrMailCP09) To UBound(arrMailCP09)
         If arrMailCP09(ii) <> "" Then
            strExc(0) = "Select * From CaseProgress Where CP09='" & arrMailCP09(ii) & "' And CP14 Is Not Null "
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               If "" & RsTemp("CP14").Value <> "" Then
                  PUB_SendMail strUserNum, RsTemp("CP14").Value, "", RsTemp("CP01").Value & "-" & RsTemp("CP02").Value & "-" & RsTemp("CP03").Value & "-" & RsTemp("CP04").Value & "已上文件齊備日及承辦期限!(因為" & cp(1) & "-" & cp(2) & "-" & cp(3) & "-" & cp(4) & "已送件)", " "
               End If
            End If
         End If
      Next
   End If
End Sub

'Add by Morgan 2009/8/18
Private Sub lblCaseFee_Click()
   If txtCaseField(10) <> "" And txtCaseField(11) <> "" And txtCaseField(12) <> "" Then
      frm12040102_2.txtCF(1) = cp(1)
      frm12040102_2.txtCF(2) = txtCaseField(12)
      frm12040102_2.txtCF(3) = txtCaseField(10)
      frm12040102_2.Show vbModal
      SetChkDate
   End If
End Sub

Private Sub lblCaseFee_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   PUB_LabelMouseDown lblCaseFee, lblCaseFees
End Sub

Private Sub lblCaseFee_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   PUB_LabelMouseUp lblCaseFee, lblCaseFees
End Sub

Private Sub SetChkDate()
   If txtCaseField(10) <> "" And txtCaseField(11) <> "" And txtCaseField(12) <> "" Then
      PUB_SetChkResultDate field(1), txtCaseField(12), txtCaseField(10), txtCaseField(11), txtCaseField(15), cp, field(8)
   End If
End Sub

'2010/1/22 ADD BY SONIA 批次發Mail
Private Sub BatchMail()
   Dim i As Integer
   For i = 1 To UBound(skMail)
        PUB_SendMail skMail(i).fiSender, skMail(i).fiReceiver, skMail(i).fiRecriverNo, skMail(i).fiSubject, skMail(i).fiContent
   Next i
   ReDim skMail(0) As SeekMails
End Sub
'2010/1/2 END

'Added by Morgan 2012/3/30
'客戶備註提醒
Private Sub CustMemoAlert(pCustNo As String)
   If cp(31) = "Y" Then
      strExc(0) = "select cu01||cu02 C1,nvl(nvl(cu04,cu06),cu05) C2,cu79 from customer where cu01||cu02='" & ChangeCustomerL(pCustNo) & "' and cu79 is not null"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         MsgBox RsTemp("cu79"), vbExclamation, "申請人備註 (" & RsTemp("C1") & " " & RsTemp("C2") & ")"
      End If
   End If
End Sub

'Add by Morgan 2008/2/13
'Move by Lydia 2018/01/15 從basQuery搬來
Public Function PUB_ReadFTList(ByVal p_Sys As String, ByVal p_Cty As String, ByRef adoRst As ADODB.Recordset, Optional ByVal p_Date As String) As Boolean
   Dim stYear As String, stDate1 As String, stDate2 As String, stFT05 As String
   Dim intR As Integer, stSQL As String
   Dim stVTB As String, stCon As String
      
   If p_Date = "" Then
      p_Date = strSrvDate(1)
   Else
      p_Date = DBDATE(p_Date)
   End If
   stYear = Left(p_Date, 4)
   '下半年
   If Val(Mid(p_Date, 5, 2)) > 6 Then
      stFT05 = "2"
      stDate1 = stYear & "0701" '"0601" Modify By Sindy 2013/5/20
      stDate2 = stYear & "1231"
   '上半年
   Else
      stFT05 = "1"
      stDate1 = stYear & "0101"
      stDate2 = stYear & "0630"
   End If
   
   stCon = " AND FT04=" & (stYear - 1911) & " and FT05='" & stFT05 & "' and FT06='" & p_Sys & "'"
   '申請國家為歐盟時帶出所有歐洲代理人
   If p_Cty = "239" Then
      stCon = stCon & " and substr(fa10,1,1)='2'"
   Else
      stCon = stCon & " and substr(fa10,1,3)='" & p_Cty & "'"
   End If
   
   '已給案量統計
   '2009/9/11 modify by sonia 因NewCasePtyList加入105但不統計給案量故加入cp10<>'105'條件
   stVTB = "select FT01||FT02||FT03 Cx1,nvl(count(*),0) Q1 From fagenttarget, fagent, caseprogress" & _
      " where FA01(+)=FT01 AND FA02(+)=FT02" & stCon & _
      " and cp44(+)=FT01||FT02 and cp44||cp116=FT01||FT02||FT03 and cp01(+)=FT06" & _
      " and instr('" & NewCasePtyList & "',cp10)>0 and cp10<>'105'" & _
      " and cp04='00' and cp27>=" & stDate1 & " and cp27<=" & stDate2 & _
      " group by FT01,FT02,FT03"

   stSQL = "select '' C1,FT01||FT02||decode(FT03,null,'','-'||FT03) C2" & _
      ",decode(FT03,null,RTRIM(FA05||' '||FA63||' '||FA64||' '||FA65)||' '||FA06||' '||FA04" & _
      ",PCC03||' '||PCC04||' '||PCC05) C3,nvl(Q1,0) C4,FT07 C5,round(100*nvl(Q1,0)/FT07)||'%' C6,round(100*nvl(Q1,0)/FT07) C7" & _
      " From fagenttarget, fagent, PotCustCont, (" & stVTB & ") X" & _
      " where FA01(+)=FT01 AND FA02(+)=FT02 AND PCC01(+)=FT01 AND PCC02(+)=FT03" & stCon & _
      " and Cx1(+)=FT01||FT02||FT03 and nvl(Q1,0)<FT07" & _
      " order by 7,4,5,2,3"
      
   intR = 1
   Set adoRst = ClsLawReadRstMsg(intR, stSQL)
   If intR = 1 Then
      PUB_ReadFTList = True
   End If
End Function

'Added by Morgan 2020/10/7
'和日本有DAS合作的國家，可不須提交優先權證明文件。CFP新案發文若有主張其中國家的優先權時要輸入存取碼。
Private Function ChkIsDASCountry(pCountryList As String) As Boolean
   Dim arrCountry() As String
   Dim ii As Integer, strCountry As String
   
   strCountry = Pub_GetSpecMan("CFP日本DAS合作國家")
   arrCountry() = Split(pCountryList, "，")
   For ii = LBound(arrCountry) To UBound(arrCountry)
      If arrCountry(ii) <> "" Then
         If InStr(strCountry, arrCountry(ii)) > 0 Then
            ChkIsDASCountry = True
            Exit For
         End If
      End If
   Next
End Function

'Added by Lydia 2021/05/25
Private Sub txtCP113_GotFocus()
   TextInverse txtCP113
End Sub

'Added by Lydia 2021/05/25
Private Sub txtCP113_Validate(Cancel As Boolean)
   If txtCP113 <> "" Then
      If Not IsNumeric(txtCP113) Then
         MsgBox "請輸入數字！", vbExclamation
         txtCP113.SetFocus
         txtCP113_GotFocus
         Cancel = True
         Exit Sub
      End If
   End If
End Sub
