VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm140401 
   BorderStyle     =   1  '單線固定
   Caption         =   "客戶基本資料維護"
   ClientHeight    =   6480
   ClientLeft      =   420
   ClientTop       =   4416
   ClientWidth     =   9156
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6480
   ScaleWidth      =   9156
   Begin TabDlg.SSTab tabCustomer 
      Height          =   5484
      Left            =   96
      TabIndex        =   150
      Top             =   960
      Width           =   8988
      _ExtentX        =   15854
      _ExtentY        =   9673
      _Version        =   393216
      Tabs            =   8
      Tab             =   2
      TabsPerRow      =   10
      TabHeight       =   420
      TabCaption(0)   =   "基本"
      TabPicture(0)   =   "frm140401.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label1(5)"
      Tab(0).Control(1)=   "Label30(0)"
      Tab(0).Control(2)=   "Label1(1)"
      Tab(0).Control(3)=   "Label1(2)"
      Tab(0).Control(4)=   "Label1(3)"
      Tab(0).Control(5)=   "Label1(4)"
      Tab(0).Control(6)=   "Label1(6)"
      Tab(0).Control(7)=   "Label1(7)"
      Tab(0).Control(8)=   "Label1(8)"
      Tab(0).Control(9)=   "Label1(9)"
      Tab(0).Control(10)=   "Label1(10)"
      Tab(0).Control(11)=   "Label1(11)"
      Tab(0).Control(12)=   "Label1(12)"
      Tab(0).Control(13)=   "Label1(13)"
      Tab(0).Control(14)=   "Label1(14)"
      Tab(0).Control(15)=   "Label1(15)"
      Tab(0).Control(16)=   "Label1(16)"
      Tab(0).Control(17)=   "Label1(17)"
      Tab(0).Control(18)=   "Label1(18)"
      Tab(0).Control(19)=   "Label30(1)"
      Tab(0).Control(20)=   "Label30(2)"
      Tab(0).Control(21)=   "Label30(3)"
      Tab(0).Control(22)=   "Label50"
      Tab(0).Control(23)=   "Label48"
      Tab(0).Control(24)=   "Label1(20)"
      Tab(0).Control(25)=   "Label1(22)"
      Tab(0).Control(26)=   "Label1(23)"
      Tab(0).Control(27)=   "Label1(24)"
      Tab(0).Control(28)=   "Label1(26)"
      Tab(0).Control(29)=   "Label1(29)"
      Tab(0).Control(30)=   "Label1(30)"
      Tab(0).Control(31)=   "lblCU143"
      Tab(0).Control(32)=   "Label1(19)"
      Tab(0).Control(33)=   "textCU04"
      Tab(0).Control(34)=   "textCU06"
      Tab(0).Control(35)=   "textCU07"
      Tab(0).Control(36)=   "textCU180"
      Tab(0).Control(37)=   "cboContact"
      Tab(0).Control(38)=   "lstDeveloper"
      Tab(0).Control(39)=   "textCU90"
      Tab(0).Control(40)=   "textCU89"
      Tab(0).Control(41)=   "textCU88"
      Tab(0).Control(42)=   "textCU05"
      Tab(0).Control(43)=   "LblCU144"
      Tab(0).Control(44)=   "textCU03"
      Tab(0).Control(45)=   "textCU09"
      Tab(0).Control(46)=   "textCU10"
      Tab(0).Control(47)=   "textCU11"
      Tab(0).Control(48)=   "textCU12"
      Tab(0).Control(49)=   "textCU13"
      Tab(0).Control(50)=   "textCU14"
      Tab(0).Control(51)=   "textCU32"
      Tab(0).Control(52)=   "textCU33"
      Tab(0).Control(53)=   "textCU34"
      Tab(0).Control(54)=   "textCU35"
      Tab(0).Control(55)=   "textCU64"
      Tab(0).Control(56)=   "textCU111"
      Tab(0).Control(57)=   "cboStatus"
      Tab(0).Control(58)=   "textCU132"
      Tab(0).Control(59)=   "textCU145"
      Tab(0).Control(60)=   "textCU153"
      Tab(0).Control(61)=   "Frame1"
      Tab(0).Control(62)=   "textCU144"
      Tab(0).Control(63)=   "ChkID"
      Tab(0).ControlCount=   64
      TabCaption(1)   =   "通訊"
      TabPicture(1)   =   "frm140401.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label63(0)"
      Tab(1).Control(1)=   "Label63(1)"
      Tab(1).Control(2)=   "Label63(2)"
      Tab(1).Control(3)=   "Label63(3)"
      Tab(1).Control(4)=   "Label63(4)"
      Tab(1).Control(5)=   "Label63(5)"
      Tab(1).Control(6)=   "Label63(6)"
      Tab(1).Control(7)=   "Label63(7)"
      Tab(1).Control(8)=   "Label63(8)"
      Tab(1).Control(9)=   "Label63(9)"
      Tab(1).Control(10)=   "Label63(10)"
      Tab(1).Control(11)=   "Label63(11)"
      Tab(1).Control(12)=   "Label63(12)"
      Tab(1).Control(13)=   "Label63(13)"
      Tab(1).Control(14)=   "Label63(14)"
      Tab(1).Control(15)=   "Label63(15)"
      Tab(1).Control(16)=   "Label63(16)"
      Tab(1).Control(17)=   "Label63(17)"
      Tab(1).Control(18)=   "Label63(18)"
      Tab(1).Control(19)=   "Label63(19)"
      Tab(1).Control(20)=   "Label63(20)"
      Tab(1).Control(21)=   "textCU58"
      Tab(1).Control(22)=   "textCU59"
      Tab(1).Control(23)=   "textCU60"
      Tab(1).Control(24)=   "textCU61"
      Tab(1).Control(25)=   "textCU62"
      Tab(1).Control(26)=   "textCU63"
      Tab(1).Control(27)=   "textCU91"
      Tab(1).Control(28)=   "textCU92"
      Tab(1).Control(29)=   "textCU93"
      Tab(1).Control(30)=   "textCU114"
      Tab(1).Control(31)=   "Label30(17)"
      Tab(1).Control(32)=   "Label1(33)"
      Tab(1).Control(33)=   "textCU16"
      Tab(1).Control(34)=   "textCU17"
      Tab(1).Control(35)=   "textCU18"
      Tab(1).Control(36)=   "textCU19"
      Tab(1).Control(37)=   "textCU20"
      Tab(1).Control(38)=   "textCU21"
      Tab(1).Control(39)=   "textCU22"
      Tab(1).Control(40)=   "textCU115"
      Tab(1).Control(41)=   "textCU116"
      Tab(1).Control(42)=   "textCU117"
      Tab(1).Control(43)=   "textCU118"
      Tab(1).ControlCount=   44
      TabCaption(2)   =   "地址"
      TabPicture(2)   =   "frm140401.frx":0038
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Label41(18)"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Label41(19)"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Label41(20)"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Label41(21)"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "Label41(22)"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "Label41(23)"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "Label41(24)"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "Label41(25)"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "Label41(26)"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).Control(9)=   "Label41(27)"
      Tab(2).Control(9).Enabled=   0   'False
      Tab(2).Control(10)=   "Label41(28)"
      Tab(2).Control(10).Enabled=   0   'False
      Tab(2).Control(11)=   "Label30(4)"
      Tab(2).Control(11).Enabled=   0   'False
      Tab(2).Control(12)=   "Label41(31)"
      Tab(2).Control(12).Enabled=   0   'False
      Tab(2).Control(13)=   "Label41(32)"
      Tab(2).Control(13).Enabled=   0   'False
      Tab(2).Control(14)=   "Label41(33)"
      Tab(2).Control(14).Enabled=   0   'False
      Tab(2).Control(15)=   "Label41(34)"
      Tab(2).Control(15).Enabled=   0   'False
      Tab(2).Control(16)=   "Label41(35)"
      Tab(2).Control(16).Enabled=   0   'False
      Tab(2).Control(17)=   "Label41(36)"
      Tab(2).Control(17).Enabled=   0   'False
      Tab(2).Control(18)=   "Label41(37)"
      Tab(2).Control(18).Enabled=   0   'False
      Tab(2).Control(19)=   "textCU23"
      Tab(2).Control(19).Enabled=   0   'False
      Tab(2).Control(20)=   "textCU29"
      Tab(2).Control(20).Enabled=   0   'False
      Tab(2).Control(21)=   "textCU31"
      Tab(2).Control(21).Enabled=   0   'False
      Tab(2).Control(22)=   "textCU24"
      Tab(2).Control(22).Enabled=   0   'False
      Tab(2).Control(23)=   "textCU25"
      Tab(2).Control(23).Enabled=   0   'False
      Tab(2).Control(24)=   "textCU26"
      Tab(2).Control(24).Enabled=   0   'False
      Tab(2).Control(25)=   "textCU27"
      Tab(2).Control(25).Enabled=   0   'False
      Tab(2).Control(26)=   "textCU28"
      Tab(2).Control(26).Enabled=   0   'False
      Tab(2).Control(27)=   "textCU65"
      Tab(2).Control(27).Enabled=   0   'False
      Tab(2).Control(28)=   "textCU66"
      Tab(2).Control(28).Enabled=   0   'False
      Tab(2).Control(29)=   "textCU67"
      Tab(2).Control(29).Enabled=   0   'False
      Tab(2).Control(30)=   "textCU68"
      Tab(2).Control(30).Enabled=   0   'False
      Tab(2).Control(31)=   "textCU69"
      Tab(2).Control(31).Enabled=   0   'False
      Tab(2).Control(32)=   "textCU102"
      Tab(2).Control(32).Enabled=   0   'False
      Tab(2).Control(33)=   "Label41(39)"
      Tab(2).Control(33).Enabled=   0   'False
      Tab(2).Control(34)=   "textCU191"
      Tab(2).Control(34).Enabled=   0   'False
      Tab(2).Control(35)=   "Label41(40)"
      Tab(2).Control(35).Enabled=   0   'False
      Tab(2).Control(36)=   "textCU87"
      Tab(2).Control(36).Enabled=   0   'False
      Tab(2).Control(37)=   "textCU30"
      Tab(2).Control(37).Enabled=   0   'False
      Tab(2).Control(38)=   "textCU112"
      Tab(2).Control(38).Enabled=   0   'False
      Tab(2).Control(39)=   "cmdTW(0)"
      Tab(2).Control(39).Enabled=   0   'False
      Tab(2).Control(40)=   "cmdTW(1)"
      Tab(2).Control(40).Enabled=   0   'False
      Tab(2).Control(41)=   "cmdSearchZip(0)"
      Tab(2).Control(41).Enabled=   0   'False
      Tab(2).Control(42)=   "cmdSearchZip(1)"
      Tab(2).Control(42).Enabled=   0   'False
      Tab(2).ControlCount=   43
      TabCaption(3)   =   "代表人"
      TabPicture(3)   =   "frm140401.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label41(0)"
      Tab(3).Control(1)=   "Label41(1)"
      Tab(3).Control(2)=   "Label41(2)"
      Tab(3).Control(3)=   "Label41(3)"
      Tab(3).Control(4)=   "Label41(4)"
      Tab(3).Control(5)=   "Label41(5)"
      Tab(3).Control(6)=   "Label41(6)"
      Tab(3).Control(7)=   "Label41(7)"
      Tab(3).Control(8)=   "Label41(8)"
      Tab(3).Control(9)=   "Label41(9)"
      Tab(3).Control(10)=   "Label41(10)"
      Tab(3).Control(11)=   "Label41(11)"
      Tab(3).Control(12)=   "Label41(12)"
      Tab(3).Control(13)=   "Label41(13)"
      Tab(3).Control(14)=   "Label41(14)"
      Tab(3).Control(15)=   "Label41(15)"
      Tab(3).Control(16)=   "Label41(16)"
      Tab(3).Control(17)=   "Label41(17)"
      Tab(3).Control(18)=   "Label41(29)"
      Tab(3).Control(19)=   "Label41(30)"
      Tab(3).Control(20)=   "Label41(38)"
      Tab(3).Control(21)=   "textCU39"
      Tab(3).Control(22)=   "textCU40"
      Tab(3).Control(23)=   "textCU41"
      Tab(3).Control(24)=   "textCU42"
      Tab(3).Control(25)=   "textCU43"
      Tab(3).Control(26)=   "textCU44"
      Tab(3).Control(27)=   "textCU45"
      Tab(3).Control(28)=   "textCU46"
      Tab(3).Control(29)=   "textCU47"
      Tab(3).Control(30)=   "textCU48"
      Tab(3).Control(31)=   "textCU49"
      Tab(3).Control(32)=   "textCU50"
      Tab(3).Control(33)=   "textCU51"
      Tab(3).Control(34)=   "textCU52"
      Tab(3).Control(35)=   "textCU53"
      Tab(3).Control(36)=   "textCU54"
      Tab(3).Control(37)=   "textCU55"
      Tab(3).Control(38)=   "textCU56"
      Tab(3).Control(39)=   "textCU104"
      Tab(3).Control(40)=   "textCU125"
      Tab(3).Control(41)=   "textCU103"
      Tab(3).ControlCount=   42
      TabCaption(4)   =   "專利"
      TabPicture(4)   =   "frm140401.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Label80(3)"
      Tab(4).Control(1)=   "Label80(4)"
      Tab(4).Control(2)=   "Label80(7)"
      Tab(4).Control(3)=   "Label80(8)"
      Tab(4).Control(4)=   "Label80(10)"
      Tab(4).Control(5)=   "Label11(0)"
      Tab(4).Control(6)=   "Label80(12)"
      Tab(4).Control(7)=   "Label80(13)"
      Tab(4).Control(8)=   "Label80(15)"
      Tab(4).Control(9)=   "Label30(6)"
      Tab(4).Control(10)=   "Label30(8)"
      Tab(4).Control(11)=   "Label30(9)"
      Tab(4).Control(12)=   "Label30(12)"
      Tab(4).Control(13)=   "Label80(17)"
      Tab(4).Control(14)=   "Label30(13)"
      Tab(4).Control(15)=   "Label80(18)"
      Tab(4).Control(16)=   "Label80(14)"
      Tab(4).Control(17)=   "Label80(20)"
      Tab(4).Control(18)=   "Label67(4)"
      Tab(4).Control(19)=   "Label67(0)"
      Tab(4).Control(20)=   "Label67(1)"
      Tab(4).Control(21)=   "lblCU(130)"
      Tab(4).Control(22)=   "lblCU(131)"
      Tab(4).Control(23)=   "Label70"
      Tab(4).Control(24)=   "Label69"
      Tab(4).Control(25)=   "Label38"
      Tab(4).Control(26)=   "Label1(27)"
      Tab(4).Control(27)=   "Label80(29)"
      Tab(4).Control(28)=   "Label49"
      Tab(4).Control(29)=   "Label6"
      Tab(4).Control(30)=   "Label55"
      Tab(4).Control(31)=   "Label1(21)"
      Tab(4).Control(32)=   "Label1(156)"
      Tab(4).Control(33)=   "Label80(11)"
      Tab(4).Control(34)=   "textCU78"
      Tab(4).Control(35)=   "textCU113"
      Tab(4).Control(36)=   "Label1(31)"
      Tab(4).Control(37)=   "Label8"
      Tab(4).Control(38)=   "textCU97"
      Tab(4).Control(39)=   "textCU72"
      Tab(4).Control(40)=   "textCU73"
      Tab(4).Control(41)=   "textCU74"
      Tab(4).Control(42)=   "textCU75"
      Tab(4).Control(43)=   "textCU105"
      Tab(4).Control(44)=   "textCU106"
      Tab(4).Control(45)=   "textCU77"
      Tab(4).Control(46)=   "txtCU(124)"
      Tab(4).Control(47)=   "txtCU(137)"
      Tab(4).Control(48)=   "txtCU(135)"
      Tab(4).Control(49)=   "txtCU(133)"
      Tab(4).Control(50)=   "txtCU(130)"
      Tab(4).Control(51)=   "txtCU(131)"
      Tab(4).Control(52)=   "textCU36"
      Tab(4).Control(53)=   "textCU37"
      Tab(4).Control(54)=   "textCU38"
      Tab(4).Control(55)=   "textCU96"
      Tab(4).Control(56)=   "textCU57"
      Tab(4).Control(57)=   "Combo3(0)"
      Tab(4).Control(58)=   "Combo2(0)"
      Tab(4).Control(59)=   "txtCU(174)"
      Tab(4).Control(60)=   "textCU123"
      Tab(4).Control(61)=   "textCU122"
      Tab(4).Control(62)=   "txtCU(177)"
      Tab(4).Control(63)=   "txtCU(189)"
      Tab(4).Control(64)=   "Combo4"
      Tab(4).Control(65)=   "txtCU(202)"
      Tab(4).ControlCount=   66
      TabCaption(5)   =   "商標"
      TabPicture(5)   =   "frm140401.frx":008C
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "textCU204"
      Tab(5).Control(1)=   "textCU203"
      Tab(5).Control(2)=   "textCU205"
      Tab(5).Control(3)=   "textCU149"
      Tab(5).Control(4)=   "Combo3(1)"
      Tab(5).Control(5)=   "Combo2(1)"
      Tab(5).Control(6)=   "txtCU(190)"
      Tab(5).Control(7)=   "textCU139"
      Tab(5).Control(8)=   "txtCU(126)"
      Tab(5).Control(9)=   "txtCU(138)"
      Tab(5).Control(10)=   "txtCU(136)"
      Tab(5).Control(11)=   "txtCU(134)"
      Tab(5).Control(12)=   "textCU107"
      Tab(5).Control(13)=   "textCU108"
      Tab(5).Control(14)=   "textCU109"
      Tab(5).Control(15)=   "textCU146"
      Tab(5).Control(16)=   "textCU147"
      Tab(5).Control(17)=   "textCU151"
      Tab(5).Control(18)=   "textCU152"
      Tab(5).Control(19)=   "textCU98"
      Tab(5).Control(20)=   "textCU99"
      Tab(5).Control(21)=   "TextCu128"
      Tab(5).Control(22)=   "textCU100"
      Tab(5).Control(23)=   "Label13"
      Tab(5).Control(24)=   "Label10"
      Tab(5).Control(25)=   "Label9"
      Tab(5).Control(26)=   "Label1(32)"
      Tab(5).Control(27)=   "textCU150"
      Tab(5).Control(28)=   "Label7"
      Tab(5).Control(29)=   "Label1(28)"
      Tab(5).Control(30)=   "Label5"
      Tab(5).Control(31)=   "Label80(19)"
      Tab(5).Control(32)=   "Label67(5)"
      Tab(5).Control(33)=   "Label67(2)"
      Tab(5).Control(34)=   "Label67(3)"
      Tab(5).Control(35)=   "Label4"
      Tab(5).Control(36)=   "Label3"
      Tab(5).Control(37)=   "Label2"
      Tab(5).Control(38)=   "Label1(25)"
      Tab(5).Control(39)=   "Label80(28)"
      Tab(5).Control(40)=   "Label80(27)"
      Tab(5).Control(41)=   "Label11(1)"
      Tab(5).Control(42)=   "Label30(16)"
      Tab(5).Control(43)=   "Label30(15)"
      Tab(5).Control(44)=   "Label80(25)"
      Tab(5).Control(45)=   "Label30(14)"
      Tab(5).Control(46)=   "Label80(24)"
      Tab(5).Control(47)=   "Label80(23)"
      Tab(5).Control(48)=   "Label80(5)"
      Tab(5).Control(49)=   "Label80(6)"
      Tab(5).Control(50)=   "Label30(10)"
      Tab(5).Control(51)=   "Label30(11)"
      Tab(5).Control(52)=   "Label80(21)"
      Tab(5).Control(53)=   "Label80(16)"
      Tab(5).ControlCount=   54
      TabCaption(6)   =   "其他"
      TabPicture(6)   =   "frm140401.frx":00A8
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "Label30(7)"
      Tab(6).Control(1)=   "Label30(5)"
      Tab(6).Control(2)=   "Label80(9)"
      Tab(6).Control(3)=   "Label80(2)"
      Tab(6).Control(4)=   "Label80(1)"
      Tab(6).Control(5)=   "Label80(0)"
      Tab(6).Control(6)=   "Label80(22)"
      Tab(6).Control(7)=   "lblComp(0)"
      Tab(6).Control(8)=   "lblComp(1)"
      Tab(6).Control(9)=   "lblComp(2)"
      Tab(6).Control(10)=   "lblComp(3)"
      Tab(6).Control(11)=   "lblComp(4)"
      Tab(6).Control(12)=   "lblComp(5)"
      Tab(6).Control(13)=   "lblCU16X(5)"
      Tab(6).Control(14)=   "lblCU16X(4)"
      Tab(6).Control(15)=   "lblCU16X(3)"
      Tab(6).Control(16)=   "lblCU16X(2)"
      Tab(6).Control(17)=   "lblCU16X(1)"
      Tab(6).Control(18)=   "lblCU16X(0)"
      Tab(6).Control(19)=   "textCU95"
      Tab(6).Control(20)=   "textCU70"
      Tab(6).Control(21)=   "Frame2"
      Tab(6).Control(22)=   "textCU94"
      Tab(6).Control(23)=   "textCU71"
      Tab(6).Control(24)=   "txtCU(141)"
      Tab(6).Control(25)=   "Frame1K"
      Tab(6).ControlCount=   26
      TabCaption(7)   =   "參考備註"
      TabPicture(7)   =   "frm140401.frx":00C4
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "textCU79"
      Tab(7).Control(1)=   "cmdIns"
      Tab(7).ControlCount=   2
      Begin VB.TextBox textCU204 
         Height          =   270
         Left            =   -67110
         MaxLength       =   2
         TabIndex        =   124
         Top             =   1290
         Width           =   320
      End
      Begin VB.TextBox textCU203 
         Height          =   270
         Left            =   -68880
         MaxLength       =   2
         TabIndex        =   123
         Top             =   1290
         Width           =   320
      End
      Begin VB.TextBox textCU205 
         Height          =   270
         Left            =   -69720
         MaxLength       =   7
         TabIndex        =   126
         Top             =   1620
         Width           =   1215
      End
      Begin VB.TextBox textCU149 
         Height          =   270
         Left            =   -67080
         MaxLength       =   1
         TabIndex        =   133
         Top             =   2610
         Width           =   255
      End
      Begin VB.ComboBox Combo3 
         Height          =   276
         Index           =   1
         ItemData        =   "frm140401.frx":00E0
         Left            =   -70650
         List            =   "frm140401.frx":00F3
         Style           =   2  '單純下拉式
         TabIndex        =   132
         Top             =   2610
         Width           =   1470
      End
      Begin VB.ComboBox Combo2 
         Height          =   276
         Index           =   1
         ItemData        =   "frm140401.frx":0127
         Left            =   -73650
         List            =   "frm140401.frx":0129
         Style           =   2  '單純下拉式
         TabIndex        =   131
         Top             =   2610
         Width           =   990
      End
      Begin VB.TextBox txtCU 
         Height          =   270
         Index           =   202
         Left            =   -67056
         MaxLength       =   1
         TabIndex        =   109
         Top             =   4050
         Width           =   255
      End
      Begin VB.Frame Frame1K 
         Height          =   280
         Left            =   -71550
         TabIndex        =   349
         Top             =   1530
         Width           =   4930
         Begin VB.CheckBox Chk1K 
            Caption         =   "月帳單"
            Height          =   180
            Index           =   2
            Left            =   3840
            TabIndex        =   148
            Top             =   60
            Width           =   1030
         End
         Begin VB.CheckBox Chk1K 
            Caption         =   "上傳平台"
            Height          =   180
            Index           =   1
            Left            =   2790
            TabIndex        =   147
            Top             =   60
            Width           =   1030
         End
         Begin VB.CheckBox Chk1K 
            Caption         =   "帳單另寄"
            Height          =   180
            Index           =   0
            Left            =   1740
            TabIndex        =   146
            Top             =   60
            Width           =   1030
         End
         Begin VB.Label Label80 
            AutoSize        =   -1  'True
            Caption         =   "請款單寄送類型："
            Height          =   180
            Index           =   26
            Left            =   150
            TabIndex        =   350
            Top             =   60
            Width           =   1440
         End
      End
      Begin VB.CheckBox ChkID 
         Caption         =   "不提供ID"
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   -71030
         TabIndex        =   347
         Top             =   3240
         Width           =   1000
      End
      Begin VB.ComboBox Combo4 
         Height          =   260
         ItemData        =   "frm140401.frx":012B
         Left            =   -69360
         List            =   "frm140401.frx":013B
         TabIndex        =   341
         Top             =   870
         Visible         =   0   'False
         Width           =   2235
      End
      Begin VB.TextBox txtCU 
         Height          =   270
         Index           =   190
         Left            =   -72780
         MaxLength       =   1
         TabIndex        =   140
         Top             =   4920
         Width           =   255
      End
      Begin VB.TextBox txtCU 
         Height          =   270
         Index           =   189
         Left            =   -68430
         MaxLength       =   1
         TabIndex        =   91
         Top             =   330
         Width           =   255
      End
      Begin VB.TextBox txtCU 
         Height          =   270
         Index           =   177
         Left            =   -73260
         MaxLength       =   1
         TabIndex        =   95
         Top             =   900
         Width           =   255
      End
      Begin VB.TextBox textCU122 
         Height          =   270
         Left            =   -66960
         MaxLength       =   1
         TabIndex        =   331
         Top             =   2070
         Width           =   255
      End
      Begin VB.TextBox textCU123 
         Height          =   270
         Left            =   -70920
         MaxLength       =   1
         TabIndex        =   330
         Top             =   2070
         Width           =   255
      End
      Begin VB.TextBox txtCU 
         Height          =   270
         Index           =   174
         Left            =   -66960
         MaxLength       =   1
         TabIndex        =   329
         Top             =   2355
         Width           =   255
      End
      Begin VB.CommandButton cmdIns 
         Caption         =   "各項指示"
         Height          =   300
         Left            =   -74880
         TabIndex        =   155
         Top             =   360
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CommandButton cmdSearchZip 
         Caption         =   "郵遞區號查詢"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   8.4
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   3240
         TabIndex        =   321
         Top             =   2040
         Width           =   1160
      End
      Begin VB.CommandButton cmdSearchZip 
         Caption         =   "郵遞區號查詢"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   8.4
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   2600
         TabIndex        =   320
         Top             =   870
         Width           =   1160
      End
      Begin VB.CommandButton cmdTW 
         Caption         =   "臺灣地址格式"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   8.4
            Charset         =   136
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   50
         TabIndex        =   319
         Top             =   1680
         Width           =   1160
      End
      Begin VB.CommandButton cmdTW 
         Caption         =   "臺灣地址格式"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   8.4
            Charset         =   136
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   50
         TabIndex        =   318
         Top             =   520
         Width           =   1160
      End
      Begin VB.TextBox textCU144 
         Height          =   270
         Left            =   -67950
         MaxLength       =   1
         TabIndex        =   25
         Top             =   3753
         Width           =   330
      End
      Begin VB.TextBox textCU139 
         Height          =   270
         Left            =   -68580
         MaxLength       =   1
         TabIndex        =   139
         Top             =   4590
         Width           =   330
      End
      Begin VB.ComboBox Combo2 
         Height          =   260
         Index           =   0
         ItemData        =   "frm140401.frx":0170
         Left            =   -73770
         List            =   "frm140401.frx":0172
         Style           =   2  '單純下拉式
         TabIndex        =   96
         Top             =   1200
         Width           =   990
      End
      Begin VB.ComboBox Combo3 
         Height          =   260
         Index           =   0
         ItemData        =   "frm140401.frx":0174
         Left            =   -70650
         List            =   "frm140401.frx":0187
         Style           =   2  '單純下拉式
         TabIndex        =   97
         Top             =   1200
         Width           =   1470
      End
      Begin VB.TextBox textCU57 
         Height          =   270
         Left            =   -73260
         MaxLength       =   8
         TabIndex        =   101
         Top             =   2355
         Width           =   1095
      End
      Begin VB.TextBox textCU96 
         Height          =   270
         Left            =   -73260
         MaxLength       =   8
         TabIndex        =   102
         Top             =   2625
         Width           =   1095
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  '沒有框線
         Height          =   255
         Left            =   -74850
         TabIndex        =   302
         Top             =   2901
         Width           =   4005
         Begin VB.OptionButton optCustomer 
            Caption         =   "公司"
            Height          =   252
            Index           =   1
            Left            =   810
            TabIndex        =   14
            Top             =   0
            Width           =   732
         End
         Begin VB.OptionButton optCustomer 
            Caption         =   "個人"
            Height          =   252
            Index           =   0
            Left            =   0
            TabIndex        =   13
            Top             =   0
            Value           =   -1  'True
            Width           =   735
         End
         Begin VB.OptionButton optCustomer 
            Caption         =   "學校"
            Height          =   252
            Index           =   2
            Left            =   1770
            TabIndex        =   15
            Top             =   0
            Width           =   732
         End
         Begin VB.OptionButton optCustomer 
            Caption         =   "特殊機構"
            Height          =   252
            Index           =   3
            Left            =   2700
            TabIndex        =   16
            Top             =   0
            Width           =   1155
         End
      End
      Begin VB.TextBox textCU153 
         Height          =   270
         Left            =   -70350
         MaxLength       =   1
         TabIndex        =   24
         Top             =   3753
         Width           =   330
      End
      Begin VB.TextBox txtCU 
         Height          =   270
         Index           =   126
         Left            =   -73140
         MaxLength       =   1
         TabIndex        =   129
         Top             =   2280
         Width           =   380
      End
      Begin VB.TextBox txtCU 
         Height          =   270
         Index           =   138
         Left            =   -68580
         MaxLength       =   1
         TabIndex        =   130
         Top             =   2250
         Width           =   380
      End
      Begin VB.TextBox txtCU 
         Height          =   270
         Index           =   136
         Left            =   -69090
         MaxLength       =   1
         TabIndex        =   128
         Top             =   1970
         Width           =   372
      End
      Begin VB.TextBox txtCU 
         Height          =   270
         Index           =   134
         Left            =   -73380
         MaxLength       =   1
         TabIndex        =   127
         Top             =   1970
         Width           =   372
      End
      Begin VB.TextBox textCU107 
         Height          =   270
         Left            =   -73380
         MaxLength       =   2
         TabIndex        =   121
         Top             =   1290
         Width           =   320
      End
      Begin VB.TextBox textCU108 
         Height          =   270
         Left            =   -70980
         MaxLength       =   2
         TabIndex        =   122
         Top             =   1290
         Width           =   320
      End
      Begin VB.TextBox textCU109 
         Height          =   270
         Left            =   -72930
         MaxLength       =   7
         TabIndex        =   125
         Top             =   1620
         Width           =   1215
      End
      Begin VB.TextBox textCU146 
         Height          =   270
         Left            =   -72780
         MaxLength       =   30
         TabIndex        =   138
         Top             =   4590
         Width           =   2772
      End
      Begin VB.TextBox textCU147 
         Height          =   270
         Left            =   -72780
         MaxLength       =   8
         TabIndex        =   135
         Top             =   3750
         Width           =   1095
      End
      Begin VB.TextBox textCU151 
         Height          =   270
         Left            =   -72780
         MaxLength       =   8
         TabIndex        =   136
         Top             =   4050
         Width           =   1095
      End
      Begin VB.TextBox textCU152 
         Height          =   270
         Left            =   -72780
         MaxLength       =   8
         TabIndex        =   137
         Top             =   4320
         Width           =   1095
      End
      Begin VB.TextBox txtCU 
         Height          =   270
         Index           =   141
         Left            =   -72810
         MaxLength       =   1
         TabIndex        =   145
         Top             =   1530
         Width           =   255
      End
      Begin VB.TextBox textCU71 
         Height          =   270
         Left            =   -73290
         MaxLength       =   8
         TabIndex        =   141
         Top             =   420
         Width           =   1095
      End
      Begin VB.TextBox textCU94 
         Height          =   270
         Left            =   -73290
         MaxLength       =   8
         TabIndex        =   143
         Top             =   960
         Width           =   1095
      End
      Begin VB.TextBox textCU98 
         Height          =   270
         Left            =   -73380
         MaxLength       =   8
         TabIndex        =   119
         Top             =   690
         Width           =   1095
      End
      Begin VB.TextBox textCU99 
         Height          =   270
         Left            =   -73380
         MaxLength       =   8
         TabIndex        =   120
         Top             =   960
         Width           =   1095
      End
      Begin VB.TextBox TextCu128 
         Height          =   270
         Left            =   -73000
         MaxLength       =   1
         TabIndex        =   118
         Top             =   420
         Width           =   255
      End
      Begin VB.TextBox textCU100 
         Height          =   270
         Left            =   -73380
         MaxLength       =   1
         TabIndex        =   117
         Top             =   270
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.TextBox textCU38 
         Height          =   270
         Left            =   -72990
         MaxLength       =   7
         TabIndex        =   112
         Top             =   4680
         Width           =   1215
      End
      Begin VB.TextBox textCU37 
         Height          =   270
         Left            =   -70188
         MaxLength       =   2
         TabIndex        =   111
         Top             =   4350
         Width           =   615
      End
      Begin VB.TextBox textCU36 
         Height          =   270
         Left            =   -73560
         MaxLength       =   2
         TabIndex        =   110
         Top             =   4350
         Width           =   615
      End
      Begin VB.TextBox txtCU 
         Height          =   270
         Index           =   131
         Left            =   -70512
         MaxLength       =   2
         TabIndex        =   108
         Top             =   4050
         Width           =   645
      End
      Begin VB.TextBox txtCU 
         Height          =   270
         Index           =   130
         Left            =   -73560
         MaxLength       =   2
         TabIndex        =   107
         Top             =   4050
         Width           =   600
      End
      Begin VB.TextBox txtCU 
         Height          =   270
         Index           =   133
         Left            =   -69330
         MaxLength       =   1
         TabIndex        =   113
         Top             =   4680
         Width           =   372
      End
      Begin VB.TextBox txtCU 
         Height          =   270
         Index           =   135
         Left            =   -66990
         MaxLength       =   1
         TabIndex        =   114
         Top             =   4680
         Width           =   372
      End
      Begin VB.TextBox txtCU 
         Height          =   270
         Index           =   137
         Left            =   -68745
         MaxLength       =   1
         TabIndex        =   116
         Top             =   4980
         Width           =   255
      End
      Begin VB.TextBox txtCU 
         Height          =   270
         Index           =   124
         Left            =   -73305
         MaxLength       =   1
         TabIndex        =   115
         Top             =   4980
         Width           =   255
      End
      Begin VB.TextBox textCU145 
         Height          =   270
         Left            =   -67260
         MaxLength       =   1
         TabIndex        =   23
         Top             =   3483
         Width           =   330
      End
      Begin VB.TextBox textCU132 
         Height          =   270
         Left            =   -70110
         MaxLength       =   1
         TabIndex        =   22
         Top             =   3483
         Width           =   330
      End
      Begin VB.TextBox textCU118 
         Height          =   270
         Left            =   -73365
         MaxLength       =   50
         TabIndex        =   40
         Top             =   1530
         Width           =   2985
      End
      Begin VB.TextBox textCU117 
         Height          =   270
         Left            =   -69300
         MaxLength       =   50
         TabIndex        =   39
         Top             =   1260
         Width           =   2985
      End
      Begin VB.TextBox textCU116 
         Height          =   270
         Left            =   -73365
         MaxLength       =   50
         TabIndex        =   38
         Top             =   1260
         Width           =   2985
      End
      Begin VB.TextBox textCU115 
         Height          =   270
         Left            =   -69300
         Locked          =   -1  'True
         TabIndex        =   37
         Top             =   960
         Width           =   2985
      End
      Begin VB.ComboBox cboStatus 
         Height          =   260
         ItemData        =   "frm140401.frx":01BB
         Left            =   -73800
         List            =   "frm140401.frx":01C5
         TabIndex        =   31
         Text            =   "cboStatus"
         Top             =   4575
         Width           =   2055
      End
      Begin VB.TextBox textCU112 
         Height          =   300
         Left            =   1860
         MaxLength       =   10
         TabIndex        =   57
         Top             =   2025
         Width           =   1335
      End
      Begin VB.TextBox textCU111 
         Height          =   270
         Left            =   -67080
         MaxLength       =   1
         TabIndex        =   28
         Top             =   4023
         Width           =   255
      End
      Begin VB.TextBox textCU77 
         Height          =   270
         Left            =   -67050
         MaxLength       =   1
         TabIndex        =   98
         Top             =   1230
         Width           =   255
      End
      Begin VB.TextBox textCU106 
         Height          =   270
         Left            =   -72930
         MaxLength       =   8
         TabIndex        =   106
         Top             =   3720
         Width           =   1095
      End
      Begin VB.TextBox textCU105 
         Height          =   270
         Left            =   -72930
         MaxLength       =   8
         TabIndex        =   105
         Top             =   3450
         Width           =   1095
      End
      Begin VB.TextBox textCU103 
         Height          =   996
         Left            =   -67920
         MaxLength       =   70
         MultiLine       =   -1  'True
         ScrollBars      =   2  '垂直捲軸
         TabIndex        =   88
         Top             =   600
         Width           =   1770
      End
      Begin VB.TextBox textCU75 
         Height          =   270
         Left            =   -73260
         MaxLength       =   1
         TabIndex        =   93
         Top             =   600
         Width           =   255
      End
      Begin VB.TextBox textCU74 
         Height          =   270
         Left            =   -69690
         MaxLength       =   1
         TabIndex        =   94
         Top             =   630
         Width           =   255
      End
      Begin VB.TextBox textCU73 
         Height          =   270
         Left            =   -72720
         MaxLength       =   1
         TabIndex        =   92
         Top             =   360
         Width           =   255
      End
      Begin VB.TextBox textCU72 
         Height          =   270
         Left            =   -73740
         MaxLength       =   1
         TabIndex        =   100
         Top             =   2070
         Width           =   255
      End
      Begin VB.TextBox textCU64 
         Height          =   270
         Left            =   -73800
         MaxLength       =   1
         TabIndex        =   26
         Top             =   4023
         Width           =   375
      End
      Begin VB.TextBox textCU35 
         Height          =   270
         Left            =   -73800
         MaxLength       =   2
         TabIndex        =   20
         Top             =   3483
         Width           =   615
      End
      Begin VB.TextBox textCU34 
         Height          =   270
         Left            =   -72450
         MaxLength       =   8
         TabIndex        =   21
         Top             =   3483
         Width           =   1005
      End
      Begin VB.TextBox textCU33 
         Height          =   270
         Left            =   -69120
         MaxLength       =   30
         TabIndex        =   30
         Top             =   4293
         Width           =   2772
      End
      Begin VB.TextBox textCU32 
         Height          =   270
         Left            =   -73320
         MaxLength       =   1
         TabIndex        =   29
         Top             =   4293
         Width           =   375
      End
      Begin VB.TextBox textCU30 
         Height          =   270
         Left            =   1200
         MaxLength       =   10
         TabIndex        =   54
         Top             =   890
         Width           =   1335
      End
      Begin VB.TextBox textCU87 
         Height          =   270
         Left            =   1200
         MaxLength       =   3
         TabIndex        =   55
         Top             =   1170
         Width           =   615
      End
      Begin VB.TextBox textCU22 
         Height          =   270
         Left            =   -73185
         MaxLength       =   20
         TabIndex        =   41
         Top             =   1830
         Width           =   2565
      End
      Begin VB.TextBox textCU21 
         Height          =   270
         Left            =   -69180
         MaxLength       =   50
         TabIndex        =   42
         Top             =   1830
         Width           =   2475
      End
      Begin VB.TextBox textCU20 
         Height          =   270
         Left            =   -73365
         MaxLength       =   50
         TabIndex        =   36
         Top             =   960
         Width           =   2985
      End
      Begin VB.TextBox textCU19 
         Height          =   270
         Left            =   -69525
         MaxLength       =   20
         TabIndex        =   35
         Top             =   660
         Width           =   2085
      End
      Begin VB.TextBox textCU18 
         Height          =   270
         Left            =   -74085
         MaxLength       =   20
         TabIndex        =   34
         Top             =   660
         Width           =   2085
      End
      Begin VB.TextBox textCU17 
         Height          =   270
         Left            =   -69525
         MaxLength       =   20
         TabIndex        =   33
         Top             =   360
         Width           =   2085
      End
      Begin VB.TextBox textCU16 
         Height          =   270
         Left            =   -74085
         MaxLength       =   20
         TabIndex        =   32
         Top             =   360
         Width           =   2085
      End
      Begin VB.TextBox textCU14 
         Height          =   270
         Left            =   -69120
         MaxLength       =   7
         TabIndex        =   27
         Top             =   4023
         Width           =   975
      End
      Begin VB.TextBox textCU13 
         Height          =   270
         Left            =   -74000
         MaxLength       =   6
         TabIndex        =   11
         Top             =   2595
         Width           =   855
      End
      Begin VB.TextBox textCU12 
         Height          =   270
         Left            =   -68800
         MaxLength       =   3
         TabIndex        =   12
         Top             =   2595
         Width           =   735
      End
      Begin VB.TextBox textCU11 
         Height          =   270
         Left            =   -73020
         MaxLength       =   18
         TabIndex        =   18
         Top             =   3195
         Width           =   2000
      End
      Begin VB.TextBox textCU10 
         Height          =   270
         Left            =   -74000
         MaxLength       =   4
         TabIndex        =   9
         Top             =   2346
         Width           =   855
      End
      Begin VB.TextBox textCU09 
         Height          =   270
         Left            =   -68800
         MaxLength       =   2
         TabIndex        =   10
         Top             =   2346
         Width           =   615
      End
      Begin VB.TextBox textCU03 
         Height          =   270
         Left            =   -73320
         MaxLength       =   8
         TabIndex        =   2
         Top             =   270
         Width           =   990
      End
      Begin VB.TextBox textCU97 
         Height          =   270
         Left            =   -73260
         MaxLength       =   8
         TabIndex        =   104
         Top             =   3165
         Width           =   1095
      End
      Begin VB.Frame Frame2 
         Height          =   1764
         Left            =   -74904
         TabIndex        =   352
         Top             =   3600
         Width           =   8796
         Begin VB.CheckBox ChkCU186 
            Caption         =   "勾選讀取回條"
            Height          =   180
            Index           =   2
            Left            =   3240
            TabIndex        =   363
            Top             =   1440
            Width           =   1632
         End
         Begin VB.CheckBox ChkCU186 
            Caption         =   "不寄官方收據"
            Height          =   180
            Index           =   1
            Left            =   1560
            TabIndex        =   362
            Top             =   1440
            Width           =   1536
         End
         Begin VB.TextBox textCU176 
            Height          =   285
            Left            =   1560
            MaxLength       =   500
            TabIndex        =   361
            Top             =   192
            Width           =   7104
         End
         Begin VB.TextBox textCU185 
            Height          =   285
            Left            =   1560
            MaxLength       =   500
            TabIndex        =   360
            Top             =   492
            Width           =   7104
         End
         Begin VB.TextBox textCU187 
            Height          =   285
            Left            =   1560
            MaxLength       =   500
            TabIndex        =   359
            Top             =   792
            Width           =   7104
         End
         Begin VB.TextBox textCU188 
            Height          =   285
            Left            =   1560
            MaxLength       =   500
            TabIndex        =   358
            Top             =   1080
            Width           =   7104
         End
         Begin VB.Label Label63 
            AutoSize        =   -1  'True
            Caption         =   "全E化客戶"
            Height          =   180
            Index           =   21
            Left            =   144
            TabIndex        =   364
            Top             =   0
            Width           =   828
         End
         Begin VB.Label Label63 
            Alignment       =   1  '靠右對齊
            AutoSize        =   -1  'True
            Caption         =   "指定信箱(正本)："
            Height          =   180
            Index           =   22
            Left            =   144
            TabIndex        =   357
            Top             =   264
            Width           =   1380
         End
         Begin VB.Label Label63 
            AutoSize        =   -1  'True
            Caption         =   "(副本)："
            Height          =   180
            Index           =   23
            Left            =   864
            TabIndex        =   356
            Top             =   504
            Width           =   660
         End
         Begin VB.Label Label63 
            Alignment       =   1  '靠右對齊
            AutoSize        =   -1  'True
            Caption         =   "商標信箱(正本)："
            Height          =   180
            Index           =   24
            Left            =   144
            TabIndex        =   355
            Top             =   768
            Width           =   1380
         End
         Begin VB.Label Label63 
            Alignment       =   1  '靠右對齊
            AutoSize        =   -1  'True
            Caption         =   "(副本)："
            Height          =   180
            Index           =   25
            Left            =   864
            TabIndex        =   354
            Top             =   1116
            Width           =   660
         End
         Begin VB.Label Label63 
            AutoSize        =   -1  'True
            Caption         =   "特殊設定："
            Height          =   180
            Index           =   26
            Left            =   624
            TabIndex        =   353
            Top             =   1440
            Width           =   900
         End
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "延展折扣：        ％"
         Height          =   180
         Left            =   -68010
         TabIndex        =   367
         Top             =   1340
         Width           =   1480
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "繳註冊費折扣：        ％"
         Height          =   180
         Left            =   -70140
         TabIndex        =   366
         Top             =   1340
         Width           =   1840
      End
      Begin VB.Label Label9 
         Caption         =   "商標全部折扣終止日："
         Height          =   260
         Left            =   -71550
         TabIndex        =   365
         Top             =   1650
         Width           =   1850
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "專利不得請雜費：        (Y:是)"
         Height          =   180
         Left            =   -68472
         TabIndex        =   351
         Top             =   4032
         Width           =   2292
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "（請輸 姓名+職稱）"
         ForeColor       =   &H00FF0000&
         Height          =   180
         Index           =   40
         Left            =   7200
         TabIndex        =   348
         Top             =   2100
         Width           =   1584
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "顧問專用信箱："
         Height          =   180
         Index           =   33
         Left            =   -70110
         TabIndex        =   346
         Top             =   1560
         Width           =   1260
      End
      Begin MSForms.Label Label30 
         Height          =   288
         Index           =   17
         Left            =   -68832
         TabIndex        =   345
         Top             =   1560
         Width           =   2676
         VariousPropertyBits=   27
         Caption         =   "LblFM2"
         Size            =   "4720;508"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label LblCU144 
         AutoSize        =   -1  'True
         Caption         =   "(N:不開發票)"
         Height          =   180
         Left            =   -67590
         TabIndex        =   344
         Top             =   3800
         Width           =   1490
      End
      Begin MSForms.TextBox textCU191 
         Height          =   300
         Left            =   6090
         TabIndex        =   343
         Top             =   2025
         Width           =   1095
         VariousPropertyBits=   679493659
         Size            =   "1931;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "跨所同意主管："
         ForeColor       =   &H00FF0000&
         Height          =   180
         Index           =   39
         Left            =   4800
         TabIndex        =   342
         Top             =   2100
         Width           =   1260
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "台灣案商標註冊證形式：        (1:電子 2:紙本)"
         Height          =   180
         Index           =   32
         Left            =   -74760
         TabIndex        =   340
         Top             =   4950
         Width           =   3495
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "台灣案專利證書形式：        (1:電子 2:紙本)"
         Height          =   180
         Index           =   31
         Left            =   -70230
         TabIndex        =   339
         Top             =   360
         Width           =   3315
      End
      Begin MSForms.TextBox textCU102 
         Height          =   315
         Left            =   4995
         TabIndex        =   63
         Top             =   2955
         Width           =   3630
         VariousPropertyBits=   671105051
         MaxLength       =   30
         Size            =   "6403;556"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCU69 
         Height          =   315
         Left            =   1245
         TabIndex        =   69
         Top             =   4950
         Width           =   3735
         VariousPropertyBits=   671105051
         MaxLength       =   30
         Size            =   "6588;556"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCU68 
         Height          =   315
         Left            =   1245
         TabIndex        =   68
         Top             =   4650
         Width           =   3735
         VariousPropertyBits=   671105051
         MaxLength       =   30
         Size            =   "6588;556"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCU67 
         Height          =   315
         Left            =   1245
         TabIndex        =   67
         Top             =   4350
         Width           =   3735
         VariousPropertyBits=   671105051
         MaxLength       =   30
         Size            =   "6588;556"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCU66 
         Height          =   315
         Left            =   1245
         TabIndex        =   66
         Top             =   4050
         Width           =   3735
         VariousPropertyBits=   671105051
         MaxLength       =   30
         Size            =   "6588;556"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCU65 
         Height          =   315
         Left            =   1245
         TabIndex        =   65
         Top             =   3750
         Width           =   3735
         VariousPropertyBits=   671105051
         MaxLength       =   30
         Size            =   "6588;556"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCU28 
         Height          =   315
         Left            =   1245
         TabIndex        =   62
         Top             =   2955
         Width           =   3555
         VariousPropertyBits=   671105051
         MaxLength       =   30
         Size            =   "6271;556"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCU27 
         Height          =   315
         Left            =   4995
         TabIndex        =   61
         Top             =   2655
         Width           =   3630
         VariousPropertyBits=   671105051
         MaxLength       =   30
         Size            =   "6403;556"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCU26 
         Height          =   315
         Left            =   1245
         TabIndex        =   60
         Top             =   2655
         Width           =   3555
         VariousPropertyBits=   671105051
         MaxLength       =   30
         Size            =   "6271;556"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCU25 
         Height          =   315
         Left            =   4995
         TabIndex        =   59
         Top             =   2355
         Width           =   3630
         VariousPropertyBits=   671105051
         MaxLength       =   30
         Size            =   "6403;556"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCU24 
         Height          =   315
         Left            =   1245
         TabIndex        =   58
         Top             =   2355
         Width           =   3555
         VariousPropertyBits=   671105051
         MaxLength       =   30
         Size            =   "6271;556"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCU05 
         Height          =   300
         Left            =   -73320
         TabIndex        =   4
         Top             =   831
         Width           =   5415
         VariousPropertyBits=   671105051
         MaxLength       =   30
         Size            =   "9551;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCU88 
         Height          =   300
         Left            =   -73320
         TabIndex        =   5
         Top             =   1137
         Width           =   5415
         VariousPropertyBits=   671105051
         MaxLength       =   30
         Size            =   "9551;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCU89 
         Height          =   300
         Left            =   -73320
         TabIndex        =   6
         Top             =   1443
         Width           =   5415
         VariousPropertyBits=   671105051
         MaxLength       =   30
         Size            =   "9551;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCU90 
         Height          =   300
         Left            =   -73320
         TabIndex        =   7
         Top             =   1749
         Width           =   5415
         VariousPropertyBits=   671105051
         MaxLength       =   30
         Size            =   "9551;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.ListBox lstDeveloper 
         Height          =   315
         Left            =   -67650
         TabIndex        =   338
         TabStop         =   0   'False
         Top             =   1020
         Width           =   1305
         VariousPropertyBits=   746586139
         ScrollBars      =   2
         DisplayStyle    =   2
         Size            =   "2302;556"
         MatchEntry      =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.ComboBox cboContact 
         Height          =   300
         Left            =   -68800
         TabIndex        =   19
         Top             =   3180
         Width           =   1716
         VariousPropertyBits=   679495707
         DisplayStyle    =   3
         Size            =   "3016;529"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCU180 
         Height          =   465
         Left            =   -73800
         TabIndex        =   336
         Top             =   4920
         Width           =   7485
         VariousPropertyBits=   -1466941413
         MaxLength       =   200
         ScrollBars      =   2
         Size            =   "13203;811"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCU79 
         Height          =   4545
         Left            =   -74880
         TabIndex        =   157
         Top             =   720
         Width           =   8745
         VariousPropertyBits=   -1466941413
         ScrollBars      =   2
         Size            =   "15425;8017"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCU113 
         Height          =   285
         Left            =   -73260
         TabIndex        =   103
         Top             =   2880
         Width           =   5820
         VariousPropertyBits=   671105051
         MaxLength       =   50
         Size            =   "10266;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCU150 
         Height          =   840
         Left            =   -73560
         TabIndex        =   134
         Top             =   2910
         Width           =   7410
         VariousPropertyBits=   -1466941413
         MaxLength       =   180
         ScrollBars      =   2
         Size            =   "13070;1482"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCU70 
         Height          =   285
         Left            =   -73290
         TabIndex        =   142
         Top             =   690
         Width           =   6165
         VariousPropertyBits=   671105051
         MaxLength       =   35
         Size            =   "10874;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCU95 
         Height          =   285
         Left            =   -73290
         TabIndex        =   144
         Top             =   1230
         Width           =   6165
         VariousPropertyBits=   671105051
         MaxLength       =   35
         Size            =   "10874;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCU125 
         Height          =   990
         Left            =   -67920
         TabIndex        =   90
         Top             =   3165
         Width           =   1770
         VariousPropertyBits=   -1466941413
         MaxLength       =   2000
         ScrollBars      =   2
         Size            =   "3122;1746"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCU114 
         Height          =   285
         Left            =   -72975
         TabIndex        =   49
         Top             =   3930
         Width           =   6675
         VariousPropertyBits=   671105051
         MaxLength       =   60
         Size            =   "11774;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCU104 
         Height          =   990
         Left            =   -67920
         TabIndex        =   89
         Top             =   1890
         Width           =   1770
         VariousPropertyBits=   -1466941413
         MaxLength       =   70
         ScrollBars      =   2
         Size            =   "3122;1746"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCU93 
         Height          =   285
         Left            =   -72960
         TabIndex        =   52
         Top             =   4830
         Width           =   2805
         VariousPropertyBits=   671105051
         MaxLength       =   20
         Size            =   "4948;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCU92 
         Height          =   285
         Left            =   -72960
         TabIndex        =   51
         Top             =   4530
         Width           =   4005
         VariousPropertyBits=   671105051
         MaxLength       =   35
         Size            =   "7064;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCU91 
         Height          =   285
         Left            =   -72960
         TabIndex        =   50
         Top             =   4230
         Width           =   2085
         VariousPropertyBits=   671105051
         MaxLength       =   10
         Size            =   "3678;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCU78 
         Height          =   510
         Left            =   -73680
         TabIndex        =   99
         Top             =   1530
         Width           =   7530
         VariousPropertyBits=   -1466941413
         MaxLength       =   180
         ScrollBars      =   2
         Size            =   "13282;900"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCU63 
         Height          =   285
         Left            =   -73275
         TabIndex        =   48
         Top             =   3630
         Width           =   6675
         VariousPropertyBits=   671105051
         Size            =   "11774;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCU62 
         Height          =   285
         Left            =   -73275
         TabIndex        =   47
         Top             =   3330
         Width           =   4005
         VariousPropertyBits=   671105051
         Size            =   "7064;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCU61 
         Height          =   285
         Left            =   -73275
         TabIndex        =   46
         Top             =   3030
         Width           =   3630
         VariousPropertyBits=   671105051
         Size            =   "6403;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCU60 
         Height          =   285
         Left            =   -73275
         TabIndex        =   45
         Top             =   2730
         Width           =   6675
         VariousPropertyBits=   671105051
         Size            =   "11774;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCU59 
         Height          =   285
         Left            =   -73275
         TabIndex        =   44
         Top             =   2430
         Width           =   4005
         VariousPropertyBits=   671105051
         Size            =   "7064;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCU58 
         Height          =   285
         Left            =   -73275
         TabIndex        =   43
         Top             =   2130
         Width           =   3630
         VariousPropertyBits=   671105051
         Size            =   "6403;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCU56 
         Height          =   285
         Left            =   -73425
         TabIndex        =   87
         Top             =   5040
         Width           =   5415
         VariousPropertyBits=   671105051
         MaxLength       =   40
         Size            =   "9551;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCU55 
         Height          =   285
         Left            =   -73425
         TabIndex        =   86
         Top             =   4740
         Width           =   5415
         VariousPropertyBits=   671105051
         MaxLength       =   60
         Size            =   "9551;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCU54 
         Height          =   285
         Left            =   -73425
         TabIndex        =   85
         Top             =   4455
         Width           =   5415
         VariousPropertyBits=   671105051
         MaxLength       =   40
         Size            =   "9551;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCU53 
         Height          =   285
         Left            =   -73425
         TabIndex        =   84
         Top             =   4185
         Width           =   5415
         VariousPropertyBits=   671105051
         MaxLength       =   40
         Size            =   "9551;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCU52 
         Height          =   285
         Left            =   -73425
         TabIndex        =   83
         Top             =   3900
         Width           =   5415
         VariousPropertyBits=   671105051
         MaxLength       =   60
         Size            =   "9551;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCU51 
         Height          =   285
         Left            =   -73425
         TabIndex        =   82
         Top             =   3630
         Width           =   5415
         VariousPropertyBits=   671105051
         MaxLength       =   40
         Size            =   "9551;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCU50 
         Height          =   285
         Left            =   -73425
         TabIndex        =   81
         Top             =   3360
         Width           =   5415
         VariousPropertyBits=   671105051
         MaxLength       =   40
         Size            =   "9551;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCU49 
         Height          =   285
         Left            =   -73425
         TabIndex        =   80
         Top             =   3075
         Width           =   5415
         VariousPropertyBits=   671105051
         MaxLength       =   60
         Size            =   "9551;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCU48 
         Height          =   285
         Left            =   -73425
         TabIndex        =   79
         Top             =   2805
         Width           =   5415
         VariousPropertyBits=   671105051
         MaxLength       =   40
         Size            =   "9551;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCU47 
         Height          =   285
         Left            =   -73425
         TabIndex        =   78
         Top             =   2520
         Width           =   5415
         VariousPropertyBits=   671105051
         MaxLength       =   40
         Size            =   "9551;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCU46 
         Height          =   285
         Left            =   -73425
         TabIndex        =   77
         Top             =   2250
         Width           =   5415
         VariousPropertyBits=   671105051
         MaxLength       =   60
         Size            =   "9551;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCU45 
         Height          =   285
         Left            =   -73425
         TabIndex        =   76
         Top             =   1980
         Width           =   5415
         VariousPropertyBits=   671105051
         MaxLength       =   40
         Size            =   "9551;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCU44 
         Height          =   285
         Left            =   -73425
         TabIndex        =   75
         Top             =   1695
         Width           =   5415
         VariousPropertyBits=   671105051
         MaxLength       =   40
         Size            =   "9551;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCU43 
         Height          =   285
         Left            =   -73425
         TabIndex        =   74
         Top             =   1425
         Width           =   5415
         VariousPropertyBits=   671105051
         MaxLength       =   60
         Size            =   "9551;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCU42 
         Height          =   285
         Left            =   -73425
         TabIndex        =   73
         Top             =   1140
         Width           =   5415
         VariousPropertyBits=   671105051
         MaxLength       =   40
         Size            =   "9551;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCU41 
         Height          =   285
         Left            =   -73425
         TabIndex        =   72
         Top             =   870
         Width           =   5415
         VariousPropertyBits=   671105051
         MaxLength       =   40
         Size            =   "9551;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCU40 
         Height          =   285
         Left            =   -73425
         TabIndex        =   71
         Top             =   600
         Width           =   5415
         VariousPropertyBits=   671105051
         MaxLength       =   60
         Size            =   "9551;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCU39 
         Height          =   285
         Left            =   -73425
         TabIndex        =   70
         Top             =   315
         Width           =   5415
         VariousPropertyBits=   671105051
         MaxLength       =   40
         Size            =   "9551;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCU31 
         Height          =   510
         Left            =   1200
         TabIndex        =   53
         Top             =   330
         Width           =   7665
         VariousPropertyBits=   -1466941413
         MaxLength       =   80
         ScrollBars      =   2
         Size            =   "13520;900"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCU29 
         Height          =   510
         Left            =   1245
         TabIndex        =   64
         Top             =   3255
         Width           =   7665
         VariousPropertyBits=   -1466941413
         MaxLength       =   80
         ScrollBars      =   2
         Size            =   "13520;900"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCU23 
         Height          =   510
         Left            =   1200
         TabIndex        =   56
         Top             =   1485
         Width           =   7665
         VariousPropertyBits=   -1466941413
         MaxLength       =   80
         ScrollBars      =   2
         Size            =   "13520;900"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCU07 
         Height          =   288
         Left            =   -68800
         TabIndex        =   17
         Top             =   2880
         Width           =   2700
         VariousPropertyBits=   671105051
         MaxLength       =   30
         Size            =   "4762;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCU06 
         Height          =   285
         Left            =   -73320
         TabIndex        =   8
         Top             =   2070
         Width           =   6975
         VariousPropertyBits=   671105051
         MaxLength       =   79
         Size            =   "12303;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCU04 
         Height          =   300
         Left            =   -73320
         TabIndex        =   3
         Top             =   525
         Width           =   6975
         VariousPropertyBits=   671105051
         MaxLength       =   80
         Size            =   "12303;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "狀態備註："
         Height          =   210
         Index           =   19
         Left            =   -74880
         TabIndex        =   337
         Top             =   4920
         Width           =   900
      End
      Begin VB.Label Label80 
         AutoSize        =   -1  'True
         Caption         =   "FCP實審自動代繳：         (Y：自動代繳）"
         Height          =   180
         Index           =   11
         Left            =   -74880
         TabIndex        =   335
         Top             =   930
         Width           =   3225
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "FCP是否核對已准專利：  　  (N:否)"
         Height          =   180
         Index           =   156
         Left            =   -68880
         TabIndex        =   334
         Top             =   2115
         Width           =   2745
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "FCP重新委任註記：      ( Y：已辦  N：不辦 )"
         Height          =   180
         Index           =   21
         Left            =   -72450
         TabIndex        =   333
         Top             =   2115
         Width           =   3450
      End
      Begin VB.Label Label55 
         AutoSize        =   -1  'True
         Caption         =   "FCP是否電子送件：                (Y:是)"
         Height          =   180
         Left            =   -68880
         TabIndex        =   332
         Top             =   2445
         Width           =   2745
      End
      Begin VB.Label lblCU143 
         Alignment       =   2  '置中對齊
         BackColor       =   &H8000000E&
         Height          =   195
         Left            =   -73080
         TabIndex        =   328
         Top             =   3788
         Width           =   615
      End
      Begin VB.Label lblCU16X 
         Alignment       =   2  '置中對齊
         BackColor       =   &H8000000E&
         Height          =   195
         Index           =   0
         Left            =   -72195
         TabIndex        =   327
         Top             =   1935
         Width           =   330
      End
      Begin VB.Label lblCU16X 
         Alignment       =   2  '置中對齊
         BackColor       =   &H8000000E&
         Height          =   195
         Index           =   1
         Left            =   -72195
         TabIndex        =   326
         Top             =   2190
         Width           =   330
      End
      Begin VB.Label lblCU16X 
         Alignment       =   2  '置中對齊
         BackColor       =   &H8000000E&
         Height          =   195
         Index           =   2
         Left            =   -72195
         TabIndex        =   325
         Top             =   2445
         Width           =   330
      End
      Begin VB.Label lblCU16X 
         Alignment       =   2  '置中對齊
         BackColor       =   &H8000000E&
         Height          =   195
         Index           =   3
         Left            =   -72195
         TabIndex        =   324
         Top             =   2700
         Width           =   330
      End
      Begin VB.Label lblCU16X 
         Alignment       =   2  '置中對齊
         BackColor       =   &H8000000E&
         Height          =   195
         Index           =   4
         Left            =   -72195
         TabIndex        =   323
         Top             =   2955
         Width           =   330
      End
      Begin VB.Label lblCU16X 
         Alignment       =   2  '置中對齊
         BackColor       =   &H8000000E&
         Height          =   195
         Index           =   5
         Left            =   -72195
         TabIndex        =   322
         Top             =   3210
         Width           =   330
      End
      Begin VB.Label lblComp 
         AutoSize        =   -1  'True
         Caption         =   "其他案預設收據公司別-非台灣：         (1：專利商標 2：專利法律 J：台一智權)"
         Height          =   180
         Index           =   5
         Left            =   -74760
         TabIndex        =   317
         Top             =   3210
         Width           =   6135
      End
      Begin VB.Label lblComp 
         AutoSize        =   -1  'True
         Caption         =   "其他案預設收據公司別-台灣：             (1：專利商標 2：專利法律)"
         Height          =   180
         Index           =   4
         Left            =   -74760
         TabIndex        =   316
         Top             =   2955
         Width           =   5130
      End
      Begin VB.Label lblComp 
         AutoSize        =   -1  'True
         Caption         =   "商標案預設收據公司別-非台灣：         (1：專利商標 2：專利法律 J：台一智權)"
         Height          =   180
         Index           =   3
         Left            =   -74760
         TabIndex        =   315
         Top             =   2700
         Width           =   6135
      End
      Begin VB.Label lblComp 
         AutoSize        =   -1  'True
         Caption         =   "商標案預設收據公司別-台灣：             (1：專利商標 2：專利法律)"
         Height          =   180
         Index           =   2
         Left            =   -74760
         TabIndex        =   314
         Top             =   2445
         Width           =   5130
      End
      Begin VB.Label lblComp 
         AutoSize        =   -1  'True
         Caption         =   "專利案預設收據公司別-非台灣：         (1：專利商標 2：專利法律 J：台一智權)"
         Height          =   180
         Index           =   1
         Left            =   -74760
         TabIndex        =   313
         Top             =   2190
         Width           =   6135
      End
      Begin VB.Label lblComp 
         AutoSize        =   -1  'True
         Caption         =   "專利案預設收據公司別-台灣：             (1：專利商標 2：專利法律)"
         Height          =   180
         Index           =   0
         Left            =   -74760
         TabIndex        =   312
         Top             =   1920
         Width           =   5130
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "(僅提醒)"
         Height          =   180
         Left            =   -74370
         TabIndex        =   311
         Top             =   3150
         Width           =   660
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "(僅提醒)"
         Height          =   180
         Left            =   -74490
         TabIndex        =   310
         Top             =   1740
         Width           =   660
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "特殊發票："
         Height          =   180
         Index           =   30
         Left            =   -68820
         TabIndex        =   309
         Top             =   3800
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "預定收款日放寬月數："
         Height          =   180
         Index           =   29
         Left            =   -74880
         TabIndex        =   308
         Top             =   3795
         Width           =   1800
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "不催延展：        (Y:不催)"
         Height          =   180
         Index           =   28
         Left            =   -69480
         TabIndex        =   307
         Top             =   4650
         Width           =   1905
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "商標請款單列印幣別格式"
         Height          =   180
         Left            =   -72660
         TabIndex        =   306
         Top             =   2640
         Width           =   1980
      End
      Begin VB.Label Label49 
         AutoSize        =   -1  'True
         Caption         =   "專利請款單列印幣別格式"
         Height          =   180
         Left            =   -72690
         TabIndex        =   305
         Top             =   1260
         Width           =   1980
      End
      Begin VB.Label Label80 
         AutoSize        =   -1  'True
         Caption         =   "年費聯絡人："
         Height          =   180
         Index           =   29
         Left            =   -74880
         TabIndex        =   304
         Top             =   2925
         Width           =   1080
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "FCP提申急件預設組別："
         Height          =   180
         Index           =   27
         Left            =   -71280
         TabIndex        =   303
         Top             =   930
         Visible         =   0   'False
         Width           =   1920
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "是否寄發顧問電子報：       (Y:寄/N:不寄)"
         Height          =   180
         Index           =   26
         Left            =   -72120
         TabIndex        =   301
         Top             =   3800
         Width           =   3150
      End
      Begin VB.Label Label80 
         AutoSize        =   -1  'True
         Caption         =   "商標以 EMail 通知：       （Y：是   D：僅D/N）"
         Height          =   180
         Index           =   19
         Left            =   -74760
         TabIndex        =   300
         Top             =   2330
         Width           =   3720
      End
      Begin VB.Label Label67 
         AutoSize        =   -1  'True
         Caption         =   "商標 Email 同時寄紙本：         (Y:是)"
         Height          =   180
         Index           =   5
         Left            =   -70530
         TabIndex        =   299
         Top             =   2330
         Width           =   2870
      End
      Begin VB.Label Label67 
         AutoSize        =   -1  'True
         Caption         =   "商標定稿份數："
         Height          =   180
         Index           =   2
         Left            =   -74760
         TabIndex        =   298
         Top             =   2010
         Width           =   1260
      End
      Begin VB.Label Label67 
         AutoSize        =   -1  'True
         Caption         =   "商標請款單份數："
         Height          =   180
         Index           =   3
         Left            =   -70530
         TabIndex        =   297
         Top             =   2010
         Width           =   1440
      End
      Begin VB.Label Label4 
         Caption         =   "商標全部折扣起始日："
         Height          =   260
         Left            =   -74760
         TabIndex        =   296
         Top             =   1650
         Width           =   1850
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "商標申請/翻譯折扣：        ％"
         Height          =   180
         Left            =   -72660
         TabIndex        =   295
         Top             =   1340
         Width           =   2400
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "商標全部折扣：          ％"
         Height          =   180
         Left            =   -74760
         TabIndex        =   294
         Top             =   1340
         Width           =   2120
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "客戶彼所商標財務編號："
         Height          =   180
         Index           =   25
         Left            =   -74760
         TabIndex        =   293
         Top             =   4650
         Width           =   1980
      End
      Begin VB.Label Label80 
         AutoSize        =   -1  'True
         Caption         =   "商標固定請款對象："
         Height          =   180
         Index           =   28
         Left            =   -74760
         TabIndex        =   292
         Top             =   3795
         Width           =   1620
      End
      Begin VB.Label Label80 
         AutoSize        =   -1  'True
         Caption         =   "商標D/N備註："
         Height          =   180
         Index           =   27
         Left            =   -74760
         TabIndex        =   291
         Top             =   2940
         Width           =   1190
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "商標請款幣別"
         Height          =   180
         Index           =   1
         Left            =   -74760
         TabIndex        =   290
         Top             =   2640
         Width           =   1080
      End
      Begin MSForms.Label Label30 
         Height          =   285
         Index           =   16
         Left            =   -71670
         TabIndex        =   289
         Top             =   3750
         Width           =   5370
         VariousPropertyBits=   27
         Caption         =   "LblFM2"
         Size            =   "9472;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label30 
         Height          =   285
         Index           =   15
         Left            =   -71670
         TabIndex        =   288
         Top             =   4050
         Width           =   5370
         VariousPropertyBits=   27
         Caption         =   "LblFM2"
         Size            =   "9472;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label80 
         AutoSize        =   -1  'True
         Caption         =   "商標D/N固定列印對象："
         Height          =   180
         Index           =   25
         Left            =   -74760
         TabIndex        =   287
         Top             =   4080
         Width           =   1905
      End
      Begin MSForms.Label Label30 
         Height          =   285
         Index           =   14
         Left            =   -71670
         TabIndex        =   286
         Top             =   4320
         Width           =   5370
         VariousPropertyBits=   27
         Caption         =   "LblFM2"
         Size            =   "9472;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label80 
         AutoSize        =   -1  'True
         Caption         =   "延展D/N列印對象："
         Height          =   180
         Index           =   24
         Left            =   -74760
         TabIndex        =   285
         Top             =   4350
         Width           =   1545
      End
      Begin VB.Label Label80 
         AutoSize        =   -1  'True
         Caption         =   "商標D/N是否列印申請人：       (Y：印)"
         Height          =   180
         Index           =   23
         Left            =   -69180
         TabIndex        =   284
         Top             =   2640
         Width           =   3060
      End
      Begin VB.Label Label80 
         AutoSize        =   -1  'True
         Caption         =   "是否用LEDES電子帳單：     （Y：是）"
         Height          =   180
         Index           =   22
         Left            =   -74760
         TabIndex        =   283
         Top             =   1575
         Width           =   3030
      End
      Begin VB.Label Label80 
         AutoSize        =   -1  'True
         Caption         =   "副本聯絡人："
         Height          =   180
         Index           =   0
         Left            =   -74760
         TabIndex        =   282
         Top             =   720
         Width           =   1080
      End
      Begin VB.Label Label80 
         AutoSize        =   -1  'True
         Caption         =   "實體副本收受人："
         Height          =   180
         Index           =   1
         Left            =   -74760
         TabIndex        =   281
         Top             =   990
         Width           =   1440
      End
      Begin VB.Label Label80 
         AutoSize        =   -1  'True
         Caption         =   "實體副本聯絡人："
         Height          =   180
         Index           =   2
         Left            =   -74760
         TabIndex        =   280
         Top             =   1260
         Width           =   1440
      End
      Begin VB.Label Label80 
         AutoSize        =   -1  'True
         Caption         =   "副本收受人："
         Height          =   180
         Index           =   9
         Left            =   -74760
         TabIndex        =   279
         Top             =   450
         Width           =   1080
      End
      Begin MSForms.Label Label30 
         Height          =   285
         Index           =   5
         Left            =   -72120
         TabIndex        =   278
         Top             =   450
         Width           =   5700
         VariousPropertyBits=   27
         Caption         =   "LblFM2"
         Size            =   "10054;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label30 
         Height          =   285
         Index           =   7
         Left            =   -72120
         TabIndex        =   277
         Top             =   990
         Width           =   5700
         VariousPropertyBits=   27
         Caption         =   "LblFM2"
         Size            =   "10054;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label80 
         AutoSize        =   -1  'True
         Caption         =   "延展代理人："
         Height          =   180
         Index           =   5
         Left            =   -74760
         TabIndex        =   276
         Top             =   720
         Width           =   1080
      End
      Begin VB.Label Label80 
         AutoSize        =   -1  'True
         Caption         =   "延展請款對象："
         Height          =   180
         Index           =   6
         Left            =   -74760
         TabIndex        =   275
         Top             =   990
         Width           =   1260
      End
      Begin MSForms.Label Label30 
         Height          =   290
         Index           =   10
         Left            =   -72210
         TabIndex        =   274
         Top             =   690
         Width           =   5370
         VariousPropertyBits=   27
         Caption         =   "LblFM2"
         Size            =   "9472;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label30 
         Height          =   290
         Index           =   11
         Left            =   -72210
         TabIndex        =   273
         Top             =   960
         Width           =   5370
         VariousPropertyBits=   27
         Caption         =   "LblFM2"
         Size            =   "9472;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label80 
         AutoSize        =   -1  'True
         Caption         =   "FCT註冊費自動代繳：        (Y:自動代繳)"
         Height          =   180
         Index           =   21
         Left            =   -74760
         TabIndex        =   272
         Top             =   440
         Width           =   3120
      End
      Begin VB.Label Label80 
         AutoSize        =   -1  'True
         Caption         =   "延展單筆不跑：        （Y：單筆不跑）"
         Height          =   180
         Index           =   16
         Left            =   -74760
         TabIndex        =   271
         Top             =   320
         Visible         =   0   'False
         Width           =   3000
      End
      Begin VB.Label Label38 
         AutoSize        =   -1  'True
         Caption         =   "專利全部折扣：                 ％"
         Height          =   180
         Left            =   -74880
         TabIndex        =   270
         Top             =   4350
         Width           =   2205
      End
      Begin VB.Label Label69 
         AutoSize        =   -1  'True
         Caption         =   "專利申請/翻譯折扣：                 ％"
         Height          =   180
         Left            =   -71880
         TabIndex        =   269
         Top             =   4356
         Width           =   2712
      End
      Begin VB.Label Label70 
         Caption         =   "專利全部折扣起始日："
         Height          =   255
         Left            =   -74880
         TabIndex        =   268
         Top             =   4680
         Width           =   1845
      End
      Begin VB.Label lblCU 
         AutoSize        =   -1  'True
         Caption         =   "專利年費折扣：                   ％"
         Height          =   180
         Index           =   131
         Left            =   -71880
         TabIndex        =   267
         Top             =   4056
         Width           =   2436
      End
      Begin VB.Label lblCU 
         Caption         =   "專利領證折扣：                 ％"
         Height          =   180
         Index           =   130
         Left            =   -74880
         TabIndex        =   266
         Top             =   4050
         Width           =   2295
      End
      Begin VB.Label Label67 
         AutoSize        =   -1  'True
         Caption         =   "專利請款單份數："
         Height          =   180
         Index           =   1
         Left            =   -68475
         TabIndex        =   265
         Top             =   4680
         Width           =   1440
      End
      Begin VB.Label Label67 
         AutoSize        =   -1  'True
         Caption         =   "專利定稿份數： "
         Height          =   180
         Index           =   0
         Left            =   -70680
         TabIndex        =   264
         Top             =   4680
         Width           =   1305
      End
      Begin VB.Label Label67 
         AutoSize        =   -1  'True
         Caption         =   "專利 Email 同時寄紙本：       (Y:是)"
         Height          =   180
         Index           =   4
         Left            =   -70680
         TabIndex        =   263
         Top             =   5025
         Width           =   2715
      End
      Begin VB.Label Label80 
         AutoSize        =   -1  'True
         Caption         =   "專利以 EMail 通知：      （Y：是   D：僅D/N）"
         Height          =   180
         Index           =   20
         Left            =   -74880
         TabIndex        =   262
         Top             =   5025
         Width           =   3600
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "是否寄發專利雙週報：       (N:不寄)"
         Height          =   180
         Index           =   24
         Left            =   -69015
         TabIndex        =   261
         Top             =   3525
         Width           =   2760
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "業務備註："
         Height          =   180
         Index           =   38
         Left            =   -67920
         TabIndex        =   260
         Top             =   2940
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "是否寄電子報：       (N:不寄)"
         Height          =   180
         Index           =   23
         Left            =   -71340
         TabIndex        =   259
         Top             =   3525
         Width           =   2220
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "開發人員："
         Height          =   180
         Index           =   22
         Left            =   -67650
         TabIndex        =   258
         Top             =   840
         Width           =   900
      End
      Begin VB.Label Label63 
         AutoSize        =   -1  'True
         Caption         =   "(其他3)："
         Height          =   180
         Index           =   20
         Left            =   -74205
         TabIndex        =   257
         Top             =   1530
         Width           =   750
      End
      Begin VB.Label Label63 
         AutoSize        =   -1  'True
         Caption         =   "(其他2)："
         Height          =   180
         Index           =   19
         Left            =   -70065
         TabIndex        =   256
         Top             =   1260
         Width           =   750
      End
      Begin VB.Label Label63 
         AutoSize        =   -1  'True
         Caption         =   "(其他1)："
         Height          =   180
         Index           =   18
         Left            =   -74205
         TabIndex        =   255
         Top             =   1260
         Width           =   750
      End
      Begin VB.Label Label63 
         AutoSize        =   -1  'True
         Caption         =   "(財務)："
         Height          =   180
         Index           =   17
         Left            =   -69975
         TabIndex        =   254
         Top             =   990
         Width           =   660
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "客戶狀態除 業務自行處理 外, 其餘都不列印在客戶名冊中 !!"
         ForeColor       =   &H000000FF&
         Height          =   180
         Index           =   20
         Left            =   -71700
         TabIndex        =   253
         Top             =   4650
         Width           =   4665
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "6"
         Height          =   180
         Index           =   37
         Left            =   4860
         TabIndex        =   252
         Top             =   2955
         Width           =   90
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "4"
         Height          =   180
         Index           =   36
         Left            =   4860
         TabIndex        =   251
         Top             =   2655
         Width           =   90
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "2"
         Height          =   180
         Index           =   35
         Left            =   4860
         TabIndex        =   250
         Top             =   2360
         Width           =   90
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "5"
         Height          =   180
         Index           =   34
         Left            =   1125
         TabIndex        =   249
         Top             =   2955
         Width           =   90
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "3"
         Height          =   180
         Index           =   33
         Left            =   1125
         TabIndex        =   248
         Top             =   2655
         Width           =   90
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "1"
         Height          =   180
         Index           =   32
         Left            =   1125
         TabIndex        =   247
         Top             =   2360
         Width           =   90
      End
      Begin VB.Label Label63 
         AutoSize        =   -1  'True
         Caption         =   "聯絡人部門（日）："
         Height          =   180
         Index           =   16
         Left            =   -74700
         TabIndex        =   246
         Top             =   3930
         Width           =   1620
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "中文地址郵遞區號："
         Height          =   180
         Index           =   31
         Left            =   135
         TabIndex        =   245
         Top             =   2100
         Width           =   1620
      End
      Begin VB.Label Label48 
         AutoSize        =   -1  'True
         Caption         =   "呆帳記錄："
         Height          =   180
         Left            =   -68010
         TabIndex        =   244
         Top             =   4065
         Width           =   900
      End
      Begin VB.Label Label50 
         AutoSize        =   -1  'True
         Caption         =   "(Y:是)"
         Height          =   180
         Left            =   -66810
         TabIndex        =   243
         Top             =   4065
         Width           =   465
      End
      Begin VB.Label Label80 
         AutoSize        =   -1  'True
         Caption         =   "專利D/N是否列印申請人：       (Y：印)"
         Height          =   180
         Index           =   14
         Left            =   -69150
         TabIndex        =   242
         Top             =   1260
         Width           =   3000
      End
      Begin VB.Label Label80 
         AutoSize        =   -1  'True
         Caption         =   "年費D/N列印對象："
         Height          =   180
         Index           =   18
         Left            =   -74880
         TabIndex        =   241
         Top             =   3765
         Width           =   1545
      End
      Begin MSForms.Label Label30 
         Height          =   285
         Index           =   13
         Left            =   -71730
         TabIndex        =   240
         Top             =   3765
         Width           =   5460
         VariousPropertyBits=   27
         Caption         =   "LblFM2"
         Size            =   "9631;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label80 
         AutoSize        =   -1  'True
         Caption         =   "專利D/N固定列印對象："
         Height          =   180
         Index           =   17
         Left            =   -74880
         TabIndex        =   239
         Top             =   3495
         Width           =   1905
      End
      Begin MSForms.Label Label30 
         Height          =   285
         Index           =   12
         Left            =   -71730
         TabIndex        =   238
         Top             =   3495
         Width           =   5460
         VariousPropertyBits=   27
         Caption         =   "LblFM2"
         Size            =   "9631;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "定稿中文名稱："
         Height          =   180
         Index           =   30
         Left            =   -67920
         TabIndex        =   237
         Top             =   1650
         Width           =   1260
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "公司負責人英文名稱："
         ForeColor       =   &H000000C0&
         Height          =   180
         Index           =   29
         Left            =   -67920
         TabIndex        =   236
         Top             =   360
         Width           =   1800
      End
      Begin MSForms.Label Label30 
         Height          =   285
         Index           =   9
         Left            =   -72060
         TabIndex        =   233
         Top             =   3210
         Width           =   5460
         VariousPropertyBits=   27
         Caption         =   "LblFM2"
         Size            =   "9631;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label30 
         Height          =   285
         Index           =   8
         Left            =   -72060
         TabIndex        =   232
         Top             =   2670
         Width           =   5460
         VariousPropertyBits=   27
         Caption         =   "LblFM2"
         Size            =   "9631;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label30 
         Height          =   285
         Index           =   6
         Left            =   -72060
         TabIndex        =   231
         Top             =   2400
         Width           =   3060
         VariousPropertyBits=   27
         Caption         =   "LblFM2"
         Size            =   "5397;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label30 
         Height          =   285
         Index           =   4
         Left            =   1875
         TabIndex        =   230
         Top             =   1170
         Width           =   1380
         VariousPropertyBits=   27
         Caption         =   "LblFM2"
         Size            =   "2434;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label30 
         Height          =   288
         Index           =   3
         Left            =   -67944
         TabIndex        =   229
         Top             =   2640
         Width           =   1716
         VariousPropertyBits=   27
         Caption         =   "LblFM2"
         Size            =   "3016;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label30 
         Height          =   285
         Index           =   2
         Left            =   -73095
         TabIndex        =   228
         Top             =   2640
         Width           =   1710
         VariousPropertyBits=   27
         Caption         =   "LblFM2"
         Size            =   "3016;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label30 
         Height          =   288
         Index           =   1
         Left            =   -67944
         TabIndex        =   227
         Top             =   2376
         Width           =   1716
         VariousPropertyBits=   27
         Caption         =   "LblFM2"
         Size            =   "3016;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label80 
         AutoSize        =   -1  'True
         Caption         =   "FCP年費自動代繳：     （Y：自動代繳 / N：寄證書後年費不續辦)"
         Height          =   180
         Index           =   15
         Left            =   -71280
         TabIndex        =   226
         Top             =   630
         Width           =   5100
      End
      Begin VB.Label Label80 
         AutoSize        =   -1  'True
         Caption         =   "FCP年費通知函單筆不跑：       （Y：單筆不跑）"
         Height          =   180
         Index           =   13
         Left            =   -74880
         TabIndex        =   225
         Top             =   360
         Width           =   3795
      End
      Begin VB.Label Label80 
         AutoSize        =   -1  'True
         Caption         =   "FCP領證自動代繳：         (Y：自動代繳）"
         Height          =   180
         Index           =   12
         Left            =   -74880
         TabIndex        =   224
         Top             =   630
         Width           =   3225
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "專利請款幣別"
         Height          =   180
         Index           =   0
         Left            =   -74880
         TabIndex        =   223
         Top             =   1260
         Width           =   1080
      End
      Begin VB.Label Label80 
         AutoSize        =   -1  'True
         Caption         =   "專利D/N備註："
         Height          =   180
         Index           =   10
         Left            =   -74880
         TabIndex        =   222
         Top             =   1530
         Width           =   1185
      End
      Begin VB.Label Label80 
         AutoSize        =   -1  'True
         Caption         =   "專利固定請款對象："
         Height          =   180
         Index           =   8
         Left            =   -74880
         TabIndex        =   221
         Top             =   2400
         Width           =   1620
      End
      Begin VB.Label Label80 
         AutoSize        =   -1  'True
         Caption         =   "收款後辦案：       （Y：先收）"
         Height          =   180
         Index           =   7
         Left            =   -74880
         TabIndex        =   220
         Top             =   2115
         Width           =   2415
      End
      Begin VB.Label Label63 
         AutoSize        =   -1  'True
         Caption         =   "MOBILE PHONE："
         Height          =   180
         Index           =   15
         Left            =   -74700
         TabIndex        =   219
         Top             =   1830
         Width           =   1440
      End
      Begin VB.Label Label63 
         AutoSize        =   -1  'True
         Caption         =   "LINE ID："
         Height          =   180
         Index           =   14
         Left            =   -70110
         TabIndex        =   218
         Top             =   1830
         Width           =   795
      End
      Begin VB.Label Label63 
         AutoSize        =   -1  'True
         Caption         =   "FAX2："
         Height          =   180
         Index           =   13
         Left            =   -70155
         TabIndex        =   217
         Top             =   660
         Width           =   600
      End
      Begin VB.Label Label63 
         AutoSize        =   -1  'True
         Caption         =   "TEL2："
         Height          =   180
         Index           =   12
         Left            =   -70155
         TabIndex        =   216
         Top             =   360
         Width           =   585
      End
      Begin VB.Label Label63 
         AutoSize        =   -1  'True
         Caption         =   "E-Mail(代表)："
         Height          =   180
         Index           =   11
         Left            =   -74700
         TabIndex        =   215
         Top             =   960
         Width           =   1140
      End
      Begin VB.Label Label63 
         AutoSize        =   -1  'True
         Caption         =   "FAX1："
         Height          =   180
         Index           =   10
         Left            =   -74700
         TabIndex        =   214
         Top             =   660
         Width           =   600
      End
      Begin VB.Label Label63 
         AutoSize        =   -1  'True
         Caption         =   "TEL1："
         Height          =   180
         Index           =   9
         Left            =   -74700
         TabIndex        =   213
         Top             =   360
         Width           =   585
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "客戶狀態："
         Height          =   180
         Index           =   18
         Left            =   -74880
         TabIndex        =   212
         Top             =   4590
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "是否寄台一雜誌：           （N:不寄）"
         Height          =   180
         Index           =   17
         Left            =   -74880
         TabIndex        =   211
         Top             =   4335
         Width           =   2820
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "定稿語文：            （1:中文   2:英文   3:日文）"
         Height          =   180
         Index           =   16
         Left            =   -74880
         TabIndex        =   210
         Top             =   4065
         Width           =   3555
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "每月收款日："
         Height          =   180
         Index           =   15
         Left            =   -74880
         TabIndex        =   209
         Top             =   3525
         Width           =   1080
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "客戶彼所專利財務編號／分所編號："
         Height          =   180
         Index           =   14
         Left            =   -72000
         TabIndex        =   208
         Top             =   4335
         Width           =   2880
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "開發日期："
         Height          =   180
         Index           =   13
         Left            =   -70050
         TabIndex        =   207
         Top             =   4065
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "行業別："
         Height          =   180
         Index           =   12
         Left            =   -73155
         TabIndex        =   206
         Top             =   3525
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "預設接洽人："
         Height          =   180
         Index           =   11
         Left            =   -69920
         TabIndex        =   205
         Top             =   3240
         Width           =   1080
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "身分證字號/統一編號："
         Height          =   180
         Index           =   10
         Left            =   -74880
         TabIndex        =   204
         Top             =   3237
         Width           =   1845
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "業務區別："
         Height          =   180
         Index           =   9
         Left            =   -69740
         TabIndex        =   203
         Top             =   2628
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "客戶來源："
         Height          =   180
         Index           =   8
         Left            =   -69740
         TabIndex        =   202
         Top             =   2388
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "公司負責人："
         Height          =   180
         Index           =   7
         Left            =   -69920
         TabIndex        =   201
         Top             =   2940
         Width           =   1080
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "智權人員："
         Height          =   180
         Index           =   6
         Left            =   -74880
         TabIndex        =   200
         Top             =   2637
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "客戶名稱（日）："
         Height          =   180
         Index           =   4
         Left            =   -74880
         TabIndex        =   198
         Top             =   2070
         Width           =   1440
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "代理人編號："
         Height          =   180
         Index           =   3
         Left            =   -74880
         TabIndex        =   197
         Top             =   312
         Width           =   1080
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "日文地址："
         Height          =   180
         Index           =   28
         Left            =   135
         TabIndex        =   196
         Top             =   3255
         Width           =   900
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "POB5："
         Height          =   180
         Index           =   27
         Left            =   135
         TabIndex        =   195
         Top             =   4950
         Width           =   600
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "POB4："
         Height          =   180
         Index           =   26
         Left            =   135
         TabIndex        =   194
         Top             =   4650
         Width           =   600
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "POB3："
         Height          =   180
         Index           =   25
         Left            =   135
         TabIndex        =   193
         Top             =   4350
         Width           =   600
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "POB2："
         Height          =   180
         Index           =   24
         Left            =   135
         TabIndex        =   192
         Top             =   4050
         Width           =   600
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "POB1："
         Height          =   180
         Index           =   23
         Left            =   135
         TabIndex        =   191
         Top             =   3750
         Width           =   600
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "地址國籍：                                                    (國內聯絡地址 或 國外英文地址 國籍)"
         Height          =   180
         Index           =   22
         Left            =   135
         TabIndex        =   190
         Top             =   1170
         Width           =   6195
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "英文地址："
         Height          =   180
         Index           =   21
         Left            =   135
         TabIndex        =   189
         Top             =   2360
         Width           =   900
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "中文地址："
         Height          =   180
         Index           =   20
         Left            =   135
         TabIndex        =   188
         Top             =   1480
         Width           =   900
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "郵遞區號："
         Height          =   180
         Index           =   19
         Left            =   135
         TabIndex        =   187
         Top             =   890
         Width           =   900
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "聯絡地址："
         Height          =   180
         Index           =   18
         Left            =   135
         TabIndex        =   186
         Top             =   345
         Width           =   900
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "代表人6（日）："
         Height          =   180
         Index           =   17
         Left            =   -74850
         TabIndex        =   185
         Top             =   5040
         Width           =   1350
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "代表人6（英）："
         Height          =   180
         Index           =   16
         Left            =   -74850
         TabIndex        =   184
         Top             =   4740
         Width           =   1350
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "代表人6（中）："
         Height          =   180
         Index           =   15
         Left            =   -74850
         TabIndex        =   183
         Top             =   4455
         Width           =   1350
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "代表人5（日）："
         Height          =   180
         Index           =   14
         Left            =   -74850
         TabIndex        =   182
         Top             =   4185
         Width           =   1350
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "代表人5（英）："
         Height          =   180
         Index           =   13
         Left            =   -74850
         TabIndex        =   181
         Top             =   3900
         Width           =   1350
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "代表人5（中）："
         Height          =   180
         Index           =   12
         Left            =   -74850
         TabIndex        =   180
         Top             =   3630
         Width           =   1350
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "代表人4（日）："
         Height          =   180
         Index           =   11
         Left            =   -74850
         TabIndex        =   179
         Top             =   3360
         Width           =   1350
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "代表人4（英）："
         Height          =   180
         Index           =   10
         Left            =   -74850
         TabIndex        =   178
         Top             =   3075
         Width           =   1350
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "代表人4（中）："
         Height          =   180
         Index           =   9
         Left            =   -74850
         TabIndex        =   177
         Top             =   2805
         Width           =   1350
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "代表人3（日）："
         Height          =   180
         Index           =   8
         Left            =   -74850
         TabIndex        =   176
         Top             =   2520
         Width           =   1350
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "代表人3（英）："
         Height          =   180
         Index           =   7
         Left            =   -74850
         TabIndex        =   175
         Top             =   2250
         Width           =   1350
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "代表人3（中）："
         Height          =   180
         Index           =   6
         Left            =   -74850
         TabIndex        =   174
         Top             =   1980
         Width           =   1350
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "代表人2（日）："
         Height          =   180
         Index           =   5
         Left            =   -74850
         TabIndex        =   173
         Top             =   1695
         Width           =   1350
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "代表人2（英）："
         Height          =   180
         Index           =   4
         Left            =   -74850
         TabIndex        =   172
         Top             =   1425
         Width           =   1350
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "代表人2（中）："
         Height          =   180
         Index           =   3
         Left            =   -74850
         TabIndex        =   171
         Top             =   1140
         Width           =   1350
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "代表人1（日）："
         Height          =   180
         Index           =   2
         Left            =   -74850
         TabIndex        =   170
         Top             =   870
         Width           =   1350
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "代表人1（英）："
         Height          =   180
         Index           =   1
         Left            =   -74850
         TabIndex        =   169
         Top             =   600
         Width           =   1350
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "客戶名稱（英）："
         Height          =   180
         Index           =   2
         Left            =   -74880
         TabIndex        =   168
         Top             =   891
         Width           =   1440
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "客戶名稱（中）："
         Height          =   180
         Index           =   1
         Left            =   -74880
         TabIndex        =   167
         Top             =   585
         Width           =   1440
      End
      Begin VB.Label Label80 
         AutoSize        =   -1  'True
         Caption         =   "年費請款對象："
         Height          =   180
         Index           =   4
         Left            =   -74880
         TabIndex        =   166
         Top             =   3210
         Width           =   1260
      End
      Begin VB.Label Label80 
         AutoSize        =   -1  'True
         Caption         =   "年費代理人："
         Height          =   180
         Index           =   3
         Left            =   -74880
         TabIndex        =   165
         Top             =   2670
         Width           =   1080
      End
      Begin VB.Label Label63 
         AutoSize        =   -1  'True
         Caption         =   "實體聯絡人（日）："
         Height          =   180
         Index           =   8
         Left            =   -74700
         TabIndex        =   164
         Top             =   4830
         Width           =   1620
      End
      Begin VB.Label Label63 
         AutoSize        =   -1  'True
         Caption         =   "實體聯絡人（英）："
         Height          =   180
         Index           =   7
         Left            =   -74700
         TabIndex        =   163
         Top             =   4530
         Width           =   1620
      End
      Begin VB.Label Label63 
         AutoSize        =   -1  'True
         Caption         =   "實體聯絡人（中）："
         Height          =   180
         Index           =   6
         Left            =   -74700
         TabIndex        =   162
         Top             =   4230
         Width           =   1620
      End
      Begin VB.Label Label63 
         AutoSize        =   -1  'True
         Caption         =   "聯絡人2（日）："
         Height          =   180
         Index           =   5
         Left            =   -74700
         TabIndex        =   161
         Top             =   3630
         Width           =   1350
      End
      Begin VB.Label Label63 
         AutoSize        =   -1  'True
         Caption         =   "聯絡人2（英）："
         Height          =   180
         Index           =   4
         Left            =   -74700
         TabIndex        =   160
         Top             =   3330
         Width           =   1350
      End
      Begin VB.Label Label63 
         AutoSize        =   -1  'True
         Caption         =   "聯絡人2（中）："
         Height          =   180
         Index           =   3
         Left            =   -74700
         TabIndex        =   159
         Top             =   3030
         Width           =   1350
      End
      Begin VB.Label Label63 
         AutoSize        =   -1  'True
         Caption         =   "聯絡人1（日）："
         Height          =   180
         Index           =   2
         Left            =   -74700
         TabIndex        =   158
         Top             =   2730
         Width           =   1350
      End
      Begin VB.Label Label63 
         AutoSize        =   -1  'True
         Caption         =   "聯絡人1（英）："
         Height          =   180
         Index           =   1
         Left            =   -74700
         TabIndex        =   156
         Top             =   2430
         Width           =   1350
      End
      Begin MSForms.Label Label30 
         Height          =   285
         Index           =   0
         Left            =   -73095
         TabIndex        =   153
         Top             =   2370
         Width           =   1140
         VariousPropertyBits=   27
         Caption         =   "LblFM2"
         Size            =   "2011;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "代表人1（中）："
         Height          =   180
         Index           =   0
         Left            =   -74856
         TabIndex        =   152
         Top             =   312
         Width           =   1356
      End
      Begin VB.Label Label63 
         AutoSize        =   -1  'True
         Caption         =   "聯絡人1（中）："
         Height          =   180
         Index           =   0
         Left            =   -74700
         TabIndex        =   151
         Top             =   2130
         Width           =   1350
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "客戶國籍：                                              (申請地址國籍)"
         Height          =   180
         Index           =   5
         Left            =   -74880
         TabIndex        =   199
         Top             =   2370
         Width           =   4170
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7440
      Top             =   0
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
            Picture         =   "frm140401.frx":01E5
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm140401.frx":0501
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm140401.frx":081D
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm140401.frx":09F9
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm140401.frx":0D15
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm140401.frx":1031
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm140401.frx":134D
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm140401.frx":1669
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm140401.frx":1985
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm140401.frx":1CA1
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm140401.frx":1FBD
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.TextBox textCU02 
      Height          =   276
      Left            =   2136
      MaxLength       =   1
      TabIndex        =   1
      Top             =   645
      Width           =   255
   End
   Begin VB.TextBox textCU01 
      Height          =   276
      Left            =   1056
      MaxLength       =   8
      TabIndex        =   0
      Top             =   645
      Width           =   1092
   End
   Begin VB.TextBox textCU15 
      Height          =   264
      Left            =   8616
      MaxLength       =   8
      TabIndex        =   154
      Top             =   660
      Visible         =   0   'False
      Width           =   375
   End
   Begin MSComctlLib.Toolbar TBar1 
      Align           =   1  '對齊表單上方
      Height          =   570
      Left            =   0
      TabIndex        =   149
      Top             =   0
      Width           =   9160
      _ExtentX        =   16150
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
   End
   Begin MSForms.TextBox textCUID 
      Height          =   285
      Left            =   2550
      TabIndex        =   234
      TabStop         =   0   'False
      Top             =   645
      Width           =   6045
      VariousPropertyBits=   671105051
      MaxLength       =   30
      Size            =   "7223;529"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      Caption         =   "客戶編號："
      Height          =   210
      Index           =   0
      Left            =   90
      TabIndex        =   235
      Top             =   675
      Width           =   900
   End
End
Attribute VB_Name = "frm140401"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/12/07 Form2.0已修改
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/25 員工編號欄已修改
'Memo By Sindy 2010/7/26 日期欄已修改
'2007/8/30 modify by sonia 取消 cu101欄,cu80客戶狀態改為勾選
'edit by nickc 2006/10/19-2006/11/01 修改
'  把 addnew 方式修改成像代理人那樣，避免造字的錯誤
'2006/11/01 11:48 修改完成
Option Explicit
'Modify By Cheng 2003/09/23
'Const iTotal = 103
'Const iTotal = 105
'Const iTotal = 107
'edit by nickc 2005/12/06
'Const iTotal = 108 '(109-1)因為陣列是由0開始
'edit by nickc 2006/10/24
'Const iTotal = 111 '(112-1)因為陣列是由0開始

Dim RcMain As New ADODB.Recordset, RsAdo As New ADODB.Recordset
'edit by nickc 2006/10/25
'Dim TmpField(0 To iTotal) As String,
'Dim ActionEdit As Integer
'Dim Bmk As Variant
' 變數宣告區
Dim m_EditMode As Integer
Dim m_SubMode As Integer
' 90.07.13 modify by louis (執行各項功能的權限)
Dim m_bInsert As Boolean
Dim m_bUpdate As Boolean
Dim m_bDelete As Boolean
Dim m_bQuery As Boolean
'edit by nickc 2006/10/26
'Add By Cheng 2002/03/12
'Dim m_blnNoErrorMsg As Boolean '更新時, 是否有錯誤訊息
'Dim m_intTab As Integer '頁籤
'Dim m_blnFormLoad As Boolean
Dim m_CU01 As String
Dim m_CU02 As String
'add by nickc 2006/03/16
Dim m_CU03 As String
'add by nickc 2006/10/26
' 宣告欄位內容結構
Private Type FIELDITEM
   fiName As String
   fiOldData As String
   fiNewData As String
   fiType As Integer
End Type
Dim m_FieldList() As FIELDITEM
' 第一筆資料的本所案號
Dim m_FirstKEY(2) As String
' 最後一筆資料的本所案號
Dim m_LastKEY(2) As String
' 目前正在顯示的本所案號
Dim m_CurrKEY(2) As String
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
Dim SeekNewCu32 As String
Dim SeekOldCu32 As String
Dim m_fa76 As String
Dim old_fa76 As String     '2008/7/17 add by sonia
Dim m_CU10 As String, m_CU31 As String  '2008/9/4 add by sonia
Dim m_Txt As Object 'Add by Morgan 2008/11/13
Dim i As Integer 'Add By Sindy 2013/1/17
Dim m_CU07 As String 'Add By Sindy 2013/7/2
Dim m_CU15 As String 'Add By Sindy 2013/12/12
Dim bolChkCU15 As Boolean  'Add by Amy 2015/10/20
'Added by Lydia 2018/10/24
Dim m_PrevForm As Form '前一畫面
Dim m_PrevNo As String '傳入客戶編號
Dim m_Tuser As String 'Added by Lydia 2019/02/14 創新業務部預設收文人員
Dim bolMsg As Boolean  'Added by Lydia 2019/04/17 是否彈過提示
Dim strCra02 As String, strCra08 As String, strCra21 As String 'Add by Amy 2022/09/14
Public m_Cra04 As String  'Add by Amy 2022/10/03 前畫面關係企業中文
Dim bolAddFinish As Boolean 'Add by Amy 2023/02/14 新客戶建檔已完成
Public m_Crl49JCmp As String 'Add by Amy 2023/05/16 出名公司對應欄位
Private Const stNotModStatus As String = "解散,廢止,撤銷,死亡" 'Add by Amy 2023/06/06
Dim strIDRepeat As String 'Add by Amy 2023/09/01 身份證/統編重覆 發信內容


'Added by Lydia 2018/10/24 傳入前一畫面
Public Sub SetParent(ByVal pFM As Form, ByVal pNo As String)
    Set m_PrevForm = pFM
    'Modify by Amy 2022/09/14 前畫面進來進新增
    m_PrevNo = pNo
    If Left(m_PrevNo, 3) <> "Add" Then
        m_PrevNo = ChangeCustomerL(pNo)
    End If
End Sub
'end 2018/10/24

'Add by Morgan 2008/7/30
Private Sub cboContact_GotFocus()
   OpenIme
End Sub

'Add by Morgan 2008/7/30
Private Sub cboContact_Validate(Cancel As Boolean)
   If m_EditMode = 1 And cboContact.Locked = False And cboContact.Text <> "" Then
      If cboContact.Text = textCU04 Then
         Cancel = True
         MsgBox "若國內接洽人與客戶中文名稱相同, 請不要輸國內接洽人! ", vbCritical + vbOKOnly, "檢核資料"
      ElseIf Not CheckLengthIsOK(cboContact, 10) Then
         Cancel = True
      End If
   End If
End Sub

Private Sub cboStatus_KeyPress(KeyAscii As Integer)
    'Add by Amy 2015/08/24 只限M51可以自行輸入,其他人只能下拉
    If m_EditMode <> 1 And m_EditMode <> 2 Then Exit Sub
    If Pub_StrUserSt03 <> "M51" Then KeyAscii = 0
End Sub

'Added by Lydia 2016/11/24 各項指示
Private Sub cmdIns_Click()
   If Me.textCU01.Text = "" Then
      MsgBox "請輸入客戶編號", vbInformation
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
   
   frm12040159.SetParent "E", Trim(Me.textCU01.Text & Me.textCU02.Text), Me
   frm12040159.Show
End Sub
'Add by Amy 2015/09/10
Private Sub cmdSearchZip_Click(Index As Integer)
    Dim stBackField As String, stText As String
    
    If Index = 0 Then
        stBackField = "textCU30"
        stText = textCU31
    Else
        stBackField = "textCU112"
        stText = textCU23
    End If
    
    Call frm100134.SetParent(Me)
    Me.Hide
    frm100134.BFormZip = stBackField
    frm100134.BFormStatus = m_EditMode
    If stText <> MsgText(601) Then
        frm100134.GetStreet stText, 2
    End If
    frm100134.Show
End Sub

Private Sub cmdTW_Click(Index As Integer)
    'Moidfy by Amy 2020/04/10 +if  因X5044001 地址:臺中大里工業區仁化路 彈二次強制表單 會錯
    If PUB_CheckFormExist("frm100135") = True Then Exit Sub
    frm100135.Show vbModal
End Sub
'end 2015/08/24

'Add By Sindy 2013/1/17
Private Sub Combo2_Click(Index As Integer)
   Call GetCurrType(Index)
End Sub

'Add By Sindy 2013/1/17
Private Sub Combo2_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

'Add By Sindy 2013/1/17
Private Sub Combo2_Validate(Index As Integer, Cancel As Boolean)
   If Combo2(Index) = MsgText(601) Then
      Combo2(Index).Tag = Combo2(Index).Text
      Combo3(Index).ListIndex = 0
      Combo3(Index).Enabled = False
      Exit Sub
   End If
   If ExistCheck("acc1y0", "a1y01", Combo2(Index), Label11(Index)) = False Then
      Cancel = True
      Combo2(Index).SetFocus
   End If
   If Combo2(Index) <> "USD" Then
      If ExistCheck("DebitNoteRate", "DNR01", Combo2(Index), Label11(Index) & "匯率") = False Then
         Cancel = True
         Combo2(Index).SetFocus
         Exit Sub
      End If
   End If
   Call GetCurrType(Index)
End Sub

'Add By Sindy 2013/1/17
Private Sub GetCurrType(Index As Integer)
Dim intType As Integer
   
   If Combo2(Index) = MsgText(601) Then
      Combo2(Index).Tag = Combo2(Index).Text
      Combo3(Index).ListIndex = 0
      Combo3(Index).Enabled = False
      Exit Sub
   End If
   '若更改請款幣別
   If Me.Combo2(Index).Text <> Me.Combo2(Index).Tag Then
      Me.Combo2(Index).Tag = Me.Combo2(Index).Text
      '請款幣別變更要重新預設列印幣別
      '台幣
      If Me.Combo2(Index).Text = "NTD" Then
         intType = 1 '純台幣
      '人民幣
      ElseIf Me.Combo2(Index).Text = "RMB" Then
         intType = 4 '外幣+美金合計
      '其他幣別
      Else
         intType = 2 '台幣+外幣合計
      End If
      Combo3(Index).ListIndex = intType
      '若為台幣時則格式欄位鎖住不可修改
      If Me.Combo2(Index).Text = "NTD" Then
         Combo3(Index).Enabled = False
      Else
         Combo3(Index).Enabled = True
      End If
   End If
End Sub

Private Sub Form_Initialize()
   'add by nickc 2006/10/24
   ReDim m_FieldList(TF_CU) As FIELDITEM
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
         If m_EditMode = 0 Then
            OnAction KeyCode
         Else
            OnAction vbKeyF10
         End If
   End Select
End Sub

'add by nickc 2006/11/10 Enter 事件，等於存檔，做完取消，不然 form 內其他物件有寫 keycode 或是 keyascii 事件的話，也會做到
Private Sub Form_KeyPress(KeyAscii As Integer)
   'Add By Sindy 2014/8/29 當focus在備註欄時按enter鍵維持換行功能而不是存檔功能
   If KeyAscii = 13 And UCase(Me.ActiveControl.Name) = UCase("textCU79") Then
      'If Me.ActiveControl.Index = 15 Then
         Exit Sub
      'End If
   End If
   '2014/8/29 END
   Select Case KeyAscii
      Case 13:
         If m_EditMode <> 0 Then
            KeyAscii = 0
            OnAction vbKeyF9
         End If
   End Select
End Sub

Private Sub Form_Load()
'92.8.28 Add By SONIA
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
'92.8.28 END
   
    'edit by nickc 2006/10/26
    'm_blnFormLoad = True
   
   ' 90.07.13 modify by louis (取得使用者執行各項功能的權限)
   m_bInsert = IsUserHasRightOfFunction("frm140401", strAdd, False)
   m_bUpdate = IsUserHasRightOfFunction("frm140401", strEdit, False)
   m_bDelete = IsUserHasRightOfFunction("frm140401", strDel, False)
   m_bQuery = IsUserHasRightOfFunction("frm140401", strFind, False)
   
   lstDeveloper.Height = 780
   lstDeveloper.Width = 1300
   
   textCUID.BackColor = &H8000000F
   
   'Added by Lydia 2016/09/10 設定代表人中文名稱和英文名稱長度
   textCU39.MaxLength = Pub_MaxCEL10
   textCU40.MaxLength = Pub_MaxCEL11
   textCU42.MaxLength = Pub_MaxCEL10
   textCU43.MaxLength = Pub_MaxCEL11
   textCU45.MaxLength = Pub_MaxCEL10
   textCU46.MaxLength = Pub_MaxCEL11
   textCU48.MaxLength = Pub_MaxCEL10
   textCU49.MaxLength = Pub_MaxCEL11
   textCU51.MaxLength = Pub_MaxCEL10
   textCU52.MaxLength = Pub_MaxCEL11
   textCU54.MaxLength = Pub_MaxCEL10
   textCU55.MaxLength = Pub_MaxCEL11
   'end 2016/09/10
   
   MoveFormToCenter Me
   
   InitialField
   
   'Add By Sindy 2013/1/17
   '抓有輸入過匯率的請款幣別
   For i = 0 To 1
      Combo2(i).Clear
      Combo2(i).AddItem ""
      Combo2(i).AddItem "USD"
      If RsTemp.State <> adStateClosed Then RsTemp.Close
      RsTemp.CursorLocation = adUseClient
      RsTemp.Open "select distinct DNR01 from DebitNoteRate order by DNR01 asc", adoTaie, adOpenStatic, adLockReadOnly
      Do While RsTemp.EOF = False
         Combo2(i).AddItem RsTemp.Fields("DNR01").Value
         RsTemp.MoveNext
      Loop
      RsTemp.Close
   Next i
   '2013/1/17 End
   
   RefreshRange
   'Modify by Amy 2022/09/14
   If Left(m_PrevNo, 3) = "Add" Then
   ElseIf m_PrevNo = "" Then 'Added by Lydia 2018/10/24 判斷是否有傳入客戶編號
        ShowFirstRecord
   'Added by Lydia 2018/10/24 有傳入客戶編號
   Else
        ShowCurrRecord Mid(m_PrevNo, 1, 8), Mid(m_PrevNo, 9, 1)
   End If
   'end 2018/10/24
   
   UpdateToolbarState
   SetCtrlReadOnly True
   
   '2007/8/30 add by sonia 加客戶狀態的下拉選單
   Me.cboStatus.Clear
   Me.cboStatus.AddItem ""
   'Me.cboStatus.AddItem "刪址" 'cancel by sonia 2019/7/19
   'Me.cboStatus.AddItem "倒閉" 'Cancel by Amy 2019/08/27
   Me.cboStatus.AddItem "遷移不明"
   Me.cboStatus.AddItem "解散"
   Me.cboStatus.AddItem "廢止"
   Me.cboStatus.AddItem "撤銷"
   Me.cboStatus.AddItem "停業"
   Me.cboStatus.AddItem "死亡" 'Modify by Amy 2019/08/27 原:往生
   'modify by sonia 2023/10/30 杜協理說其他及業務自行處理，依異動表規則只能檔案室改
   'Me.cboStatus.AddItem "其他"
   'Me.cboStatus.AddItem "業務自行處理"
   'Modify by Amy 2024/09/03 原:Pub_StrUserSt03 = "M15" ,B0001 張耀文 無此下拉(st03=M11)
   If Pub_StrUserSt03 = "M51" Or Left(Pub_strUserST05, 1) = "F" Then
      Me.cboStatus.AddItem "其他"
      Me.cboStatus.AddItem "業務自行處理"
   End If
   'end 2023/10/30
   '2007/8/30 end
   'Add by Amy 2021/11/29 加客戶狀態的下拉選單,cboStatus_Validate也要加
   Me.cboStatus.AddItem "國內同業"
   If Pub_StrUserSt03 = "M51" Then
        Me.cboStatus.AddItem "設為對造"
        Me.cboStatus.AddItem "解除對造"
   End If
   
   'Modify by Amy 2022/09/14
   If Left(m_PrevNo, 3) = "Add" Then
       OnAction vbKeyF2
       m_EditMode = 0 'Add by Amy 2022/12/16 避免寫入接洽單有的資料觸發一些Validate,故先設成0
       ShowConsultRecApp
       m_EditMode = 1 'Add by Amy 2022/12/16 ex:日本國籍會彈「此非國外部客戶，定稿語文是否確定不為中文」
       SetLock
   ElseIf m_PrevNo = "" Then 'Added by Lydia 2018/10/24 判斷是否有傳入客戶編號
       OnAction vbKeyF4
   End If

   Me.tabCustomer.Tab = 0
   '93.8.5 ADD BY SONIA '定稿中文名稱僅限電腦中心人員可修改)
   StrSQLa = "SELECT ST03 FROM STAFF WHERE ST01 = '" & Trim(strUserNum) & "'"
   Set rsA = New ADODB.Recordset
   rsA.CursorLocation = adUseClient
   rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
   If rsA.RecordCount > 0 Then
      If rsA.Fields("ST03").Value <> "M51" Then
         textCU104.Enabled = False
      Else
         textCU104.Enabled = True
         Me.cboStatus.AddItem "不再使用"      '2007/8/30 add by sonia
         Me.cboStatus.AddItem "不得代理"      '2012/10/11 add by sonia
         Me.cboStatus.AddItem "不得代理專利"  '2013/9/3 add by sonia
         Me.cboStatus.AddItem "不得代理商標"  '2013/9/3 add by sonia
         Me.cboStatus.AddItem "宣告破產" 'Add by Amy 2013/10/29 此增加cbostatus_validate 也要加
      End If
   Else
      textCU104.Enabled = False
   End If
   '93.8.5 END
   
   'Add By Sindy 2013/12/18
   If Pub_StrUserSt03 = "M51" Then
      Label1(30).Visible = True
      textCU144.Visible = True
      LblCU144.Visible = True 'Add By Sindy 2023/9/4
   Else
      Label1(30).Visible = False
      textCU144.Visible = False
      LblCU144.Visible = False 'Add By Sindy 2023/9/4
   End If
   '2013/12/18 END
   'Add by Amy 2015/09/10
   If 案件預設收據公司別啟用日 >= Val(strSrvDate(1)) Then
        cmdSearchZip(0).Visible = False
        cmdSearchZip(1).Visible = False
        cmdTW(0).Visible = False
        cmdTW(1).Visible = False
        For i = 0 To 5
            lblCU16X(i).Visible = False
        Next i
    Else
        For i = 0 To 5
            lblCU16X(i).BackColor = &H8000000F
        Next i
   End If
   lblCU143.BackColor = &H8000000F
   'end 2015/09/10
   
   Frame1K.BorderStyle = 0 'Add By Sindy 2025/1/6
   
   'Added by Lydia 2020/03/31 事務所合併日起台灣案取消(1:專利商標 2:專利法律) 的標題，非台灣案改標題為(J:智權公司 空白:系統預設)。
   If strSrvDate(1) >= 事務所合併日 Then
       For i = 0 To 5
          Select Case i
              Case 0, 2, 4  '台灣案:CU160,CU162,CU164
                  'Modifed by Lydia 2021/07/13 debug-統一改標題為(J:智權公司 空白:系統預設)
                  'lblComp(intI).Visible = False
                  'lblCU16X(intI).Visible = False
                  lblComp(i).Caption = Replace(lblComp(i).Caption, "1：專利商標 2：專利法律", "J：智權公司 空白:系統預設")
                  'end 2021/07/13
              Case 1, 3, 5  '非台灣案:CU161,CU163,CU165
                  lblComp(i).Caption = Replace(lblComp(i).Caption, "1：專利商標 2：專利法律 J：台一智權", "J：智權公司 空白:系統預設")
          End Select
       Next
   End If
   'end 2020/03/30
      
   'Added by Lydia 2020/05/05 各項指示：顯示按鈕
   If strSrvDate(1) >= 各項指示啟用日 Then
      cmdIns.Visible = True
   Else
      cmdIns.Visible = False
      textCU79.Top = 360
      textCU79.Height = 4725
   End If
   'end 2020/05/05
   
   'Added by Lydia 2023/03/03 外專新案認領
   If strSrvDate(1) >= 外專新案認領啟用日 Then
      Label1(27).Visible = True
      Combo4.Visible = True
   End If
   
   'Add by Amy 2024/01/22 國外潛在客戶維護轉號存檔切至此畫面-陳金蓮
   If m_PrevForm Is Nothing = False Then
      If UCase(m_PrevForm.Name) = "FRM140402" Then
         OnAction vbKeyF3
      End If
   End If
   'Memo by Amy 2024/03/28 偶爾會有殘影,改成直接設定
   'Add by Amy 2024/03/08 隱藏延展單筆不跑,將FCT註冊費自動代繳移位
'   Label80(21).Left =4470
'   TextCu128.Left = 6240
End Sub

Private Sub Form_Unload(Cancel As Integer)
   'Modify by Amy 2022/09/14 回傳客戶編號
   If Left(m_PrevNo, 3) = "Add" And bolAddFinish = True Then
       m_PrevForm.TxtCRA(strCra02) = textCU01
       If textCU01 <> MsgText(601) Then
            m_PrevForm.TxtCRA(strCra02).Locked = True
       End If
       m_PrevForm.CmdAddCus.Tag = ""
       m_PrevForm.Show
   'Add by Amy 2024/01/22  國外潛在客戶維護轉號存檔切至此畫面-陳金蓮
   Else
      If m_PrevForm Is Nothing = False Then
         If UCase(m_PrevForm.Name) = "FRM140402" Then
             m_PrevForm.AfterTransfer
         End If
         m_PrevForm.Show 'Added by Lydia 2024/02/22
      End If
   End If
   m_Cra04 = "" 'Add by Amy 2022/10/03
   m_Crl49JCmp = "" 'Add by Amy 2023/05/16
   Set frm140401 = Nothing
End Sub

Private Sub optCustomer_Validate(Index As Integer, Cancel As Boolean)
   'edit by nick 2004/10/06
   'If m_editmode = 4 Then Exit Sub
   'add by nickc 2008/01/23 若不是修改狀態，將會出不去
   If m_EditMode <> 1 And m_EditMode <> 2 Then Exit Sub
   'add by nick 2004/09/14
   If textCU10.Text < "010" And IsEmpty(textCU04.Text) = False And IsEmpty(textCU10.Text) = False Then
       If GetTextLength(textCU04.Text) <= 6 Then
           'Modify by Amy 2015/10/20 +請確認是否為個人
           'If optCustomer(0).Value = False Then
           If m_EditMode = 1 Or (m_EditMode = 2 And (optCustomer(0).Tag = "False" Or optCustomer(0).Value = False)) Then
                If MsgBox("客戶名稱長度小於6碼, 是否修改為個人？", vbExclamation + vbYesNo) = vbYes Then
                    'ShowMsg "必須是個人 !"
                    optCustomer(0).Value = True 'Add By Sindy 2012/6/12
                    'Cancel = True
                    textCU07.Locked = False 'Added by Lydia 2023/01/17 暫時修改
                End If
                bolChkCU15 = True
                textCU01.SetFocus '未加會一直彈是否修改為個人
           End If
           'end 2015/10/20
       Else
            'Modify By Sindy 2012/5/24
'           If GetTextLength(textCU04.Text) >= 12 Then
'               If optCustomer(1).Value = False Then
'                   ShowMsg "必須是公司 !"
'                   Cancel = True
'               End If
'           End If
            'Ex.X1234904(李忠義．李孟育)長度14為個人時，通知電腦中心調整
            If GetTextLength(textCU04.Text) >= 12 Then
               If optCustomer(0).Value = True Then
                   If m_EditMode <> 1 And m_EditMode <> 2 Then Exit Sub
                   ShowMsg "不可以是個人 !"
                   optCustomer(0).Value = False
                   Cancel = True
               End If
            End If
            '2012/5/24 End
       End If
   End If
End Sub

Private Sub Tbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
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

Private Sub ShowMsg(ByVal St As String)
   MsgBox St, vbInformation
End Sub

Private Function ChgType(ByVal Sty As Integer, ByVal txt As String) As String
Dim strTmp As String, strTmp1 As String
Dim bolMsgOnly As Boolean 'Add by Amy 2015/09/09
   
   Select Case Sty
      Case 0
         'edit by nickc 2007/02/02 不用 dll 了
         'If objPublicData.GetCaseProperty("P", txt, strTmp) = True Then
         If ClsPDGetCaseProperty("P", txt, strTmp) = True Then
            ChgType = strTmp
         Else
            ChgType = ""
         End If
      Case 1, 5
         'Modify By Cheng 2002/07/29
         '只檢查智權人員代號是否存在, 不管是否仍在職
'         If objPublicData.GetStaff(txt, strTmp, strTmp1) = True Then
         'Modify byAmy 2015/09/09 新增時若輸離職人員只彈訊息,但可Save
         bolMsgOnly = False
         If m_EditMode = 1 Then bolMsgOnly = True
         If PUB_GetStaffNameDept(txt, strTmp, strTmp1, bolMsgOnly, bolMsgOnly) = True Then
            ChgType = strTmp & "," & strTmp1
         Else
            ChgType = ""
         End If
      Case 2
         'edit by nickc 2007/02/02 不用 dll 了
         'If objPublicData.GetNation(txt, strTmp) = True Then
         If ClsPDGetNation(txt, strTmp) = True Then
            ChgType = strTmp
         Else
            ChgType = ""
         End If
      Case 3
         'edit by nickc 2007/02/02 不用 dll 了
         'If objPublicData.GetCaseSource(txt, strTmp) = True Then
         If ClsPDGetCaseSource(txt, strTmp) = True Then
            ChgType = strTmp
         Else
            ChgType = ""
         End If
      Case 4
         'edit by nickc 2007/02/05 不用 dll 了
         'If objLawDll.LawGetName(txt, strTmp) = True Then
         If ClsLawLawGetName(txt, strTmp) = True Then
            ChgType = strTmp
         Else
            ChgType = ""
         End If
   End Select
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
   
   If IsNull(rsSrcTmp.Fields("CU81")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("CU81")) = False Then
         strCName = GetStaffName(rsSrcTmp.Fields("CU81"), True)
      End If
   End If
   If IsNull(rsSrcTmp.Fields("CU82")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("CU82")) = False Then
         strTemp = TAIWANDATE(rsSrcTmp.Fields("CU82"))
         strCDate = Format(strTemp, "###/##/##")
      End If
   End If
   If IsNull(rsSrcTmp.Fields("CU83")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("CU83")) = False Then
         strTemp = rsSrcTmp.Fields("CU83")
         strCTime = Format(strTemp, "##:##")
      End If
   End If
   If IsNull(rsSrcTmp.Fields("CU84")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("CU84")) = False Then
         strUName = GetStaffName(rsSrcTmp.Fields("CU84"), True)
      End If
   End If
   If IsNull(rsSrcTmp.Fields("CU85")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("CU85")) = False Then
         strTemp = TAIWANDATE(rsSrcTmp.Fields("CU85"))
         strUDate = Format(strTemp, "###/##/##")
      End If
   End If
   If IsNull(rsSrcTmp.Fields("CU86")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("CU86")) = False Then
         strTemp = rsSrcTmp.Fields("CU86")
         strUTime = Format(strTemp, "##:##")
      End If
   End If
   
   ' 設定CUID中的文字
   textCUID = "CREATE : " & strCName & " " & _
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
Dim strTmp As String    '2008/9/4 add by sonia

   TxtValidate = False
   
   'Add by Sindy 2021/12/07 檢查畫面上的物件是否含有Unicode文字
   If PUB_ChkUniText(Me, True, True) = False Then
      Exit Function
   End If
   
   'Add by Morgan 2009/10/16
   For Each objTxt In Me.txtCU
        If objTxt.Enabled = True Then
           Cancel = False
           txtCU_Validate objTxt.Index, Cancel
           If Cancel = True Then
              Exit Function
           End If
        End If
   Next
   'end 2009/10/16
   
   If Me.textCU01.Enabled = True Then
      Cancel = False
      textCU01_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textCU03.Enabled = True Then
      Cancel = False
      textCU03_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If Me.textCU04.Enabled = True Then
      Cancel = False
      textCU04_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If Me.textCU06.Enabled = True Then
      Cancel = False
      textCU06_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If Me.textCU09.Enabled = True Then
      Cancel = False
      textCU09_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If Me.textCU10.Enabled = True Then
      '2008/9/4 add by sonia
      'Modified by Lydia 2021/02/26 新增時判斷客戶國籍和地址國籍不一致,彈訊息提醒; ex. X83829000於新增時修改國籍
      'If textCU10.Text <> m_CU10 And m_CU10 <> "" And textCU87 <> textCU10 Then
      'Modify by Amy 2022/05/25 判斷國籍前3碼-陳金蓮
      'Modify by Amy 2022/06/06 秀玲:客戶國籍有修改(不用判斷前3碼),且與地址國籍前3碼不同彈訊息,原判斷 m_CU10 值於textCU10_Validate觸發,若改完跳離textCU10.Text = m_CU10,條件不會成立(改完按Enter才會成立)
      'If (textCU10.Text <> m_CU10 And m_CU10 <> "" And Left(textCU87, 3) <> Left(textCU10, 3)) Or (m_EditMode = 1 And Left(textCU87, 3) <> Left(textCU10, 3)) Then
      'Modify by Amy 2025/03/06  接洽單判斷textCU10.Tag,避免一些不需彈的訊息被觸發,並修改訊息
      '              ex:接洽單 1130021562,客戶國籍為香港,存檔時因m_FieldList(9).fiOldData =空,彈「修改客戶國籍,地址國籍是否同時修改...」,而修改成客戶國籍為台灣
      '              ex:直接新增時,客戶國籍為013,申請與聯絡地址為台灣,聯絡地址跳離開會改地址國籍,存檔檢查會彈此訊息,沒修改客戶國籍訊息會誤解,故改訊息
      'If textCU10.Text <> m_FieldList(9).fiOldData And Left(textCU10.Text, 3) <> Left(textCU87, 3) Then
      If (m_EditMode = 1 And Left(textCU10.Text, 3) <> Left(textCU87, 3)) _
        Or ((m_EditMode = 2 And textCU10.Text <> m_FieldList(9).fiOldData) Or (Left(m_PrevNo, 3) = "Add" And textCU10 <> textCU10.Tag)) And (Left(textCU10.Text, 3) <> Left(textCU87, 3)) Then
         'Modify by Amy 2022/07/11 +顯示空白 字樣-陳金蓮 ex:X5507001
         strTmp = Replace(Replace(Trim(textCU31), " ", ""), "　", "")
         If strTmp = MsgText(601) Then strTmp = "空白"
         'strTmp = "修改客戶國籍，地址國籍是否同時修改 ? " & vbCrLf & "聯絡地址為" & strTmp & "，地址國籍為" & textCU87
         strTmp = "客戶國籍為" & Label30(0) & "(" & textCU10 & ") 與地址國籍不同是否改成客戶國籍 ? " & vbCrLf & _
                            "聯絡地址為" & strTmp & "，地址國籍為" & Label30(4) & "(" & textCU87 & ")" & vbCrLf & _
                            "　是:地址國籍改成客戶國籍" & Label30(0) & "(" & textCU10 & ") ,存檔" & vbCrLf & _
                            "　否:地址國籍維持" & Label30(4) & "(" & textCU87 & ") ,存檔" & vbCrLf & _
                            "取消:不存檔,回前畫面"
         'end 2022/07/11
         ii = MsgBox(strTmp, vbYesNoCancel + vbCritical)
         If ii = vbCancel Then
            Exit Function
         ElseIf ii = vbYes Then
            textCU87 = textCU10.Text
         End If
      End If
      '2008/9/4 end
      Cancel = False
      textCU10_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   'Add By Sindy 2014/7/22
   If optCustomer(0).Value = False And optCustomer(1).Value = False _
      And optCustomer(2).Value = False And optCustomer(3).Value = False Then
      MsgBox "請點選個人或公司 !", vbExclamation
      Exit Function
   End If
   '2014/7/22 END
   'Add By Sindy 2012/5/24
   For ii = 0 To 3
      'Modify by Amy 2015/10/20 +bolChkCU15判斷
      If optCustomer(ii).Value = True And bolChkCU15 = False Then
         Cancel = False
         Call optCustomer_Validate(ii, Cancel)
         If Cancel = True Then
             Exit Function
         End If
      End If
   Next ii
   '2012/5/24 End
   
   If Me.textCU07.Enabled = True Then
      'Add by Amy 2023/04/28 從textCU07_validate搬過來,避免無法按取消
      If textCU07 = "" And optCustomer(1).Value = True And textCU10 < "010" Then
       tabCustomer.Tab = 0 'Add by Morgan 2005/3/4
       ShowMsg "客戶為公司, 公司負責人欄不可空白 !"
       Cancel = True
       textCU07.SetFocus
       textCU07_GotFocus
       Exit Function
     End If
     'end 2023/04/28
      Cancel = False
      textCU07_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   'Modify by Amy 2024/03/28 改Form2.0後,欄位檢查不符合彈訊息後,衍生欄位無窮切換問題,造成無法跳出檢查
   'ex:X37084 改統編後無法存檔(偶爾發生) Morgan:Form2.0改後Validate事件可能會無限觸發,導致無法跳離
'   If Me.textCU11.Enabled = True Then
'      Cancel = False
'      textCU11_Validate Cancel
'      If Cancel = True Then
'         Exit Function
'      End If
'   End If
   'Modify by Amy 2024/04/23 +個人或公司
   strExc(3) = "0"
   If optCustomer(0).Value = True Then
      strExc(3) = "1"
   ElseIf optCustomer(1).Value = True Then
      strExc(3) = "2"
   End If
   If Me.textCU11.Enabled = True And Trim(Me.textCU11) <> MsgText(601) Then
      'If Pub_CheckIDAll(0, Me.Name, textCU11, textCU10) = False Then
      If Pub_CheckIDAll(0, Me.Name, textCU11, textCU10, , , Val(strExc(3))) = False Then
         tabCustomer.Tab = 0
         'textCU11_GotFocus '避免一直彈訊息無法存檔不設
   'end 2024/04/23
         Exit Function
      End If
   End If
   'end 2024/03/28
   
   'Mark by Amy Add by Amy 2015/09/09 中文地址及聯絡地址不檢查 ex:新北市新莊區Zip(248XX)會被蓋掉成242(因多個zip)
'   If Me.textCU23.Enabled = True Then
'      Cancel = False
'      textCU23_Validate Cancel
'      If Cancel = True Then
'         Exit Function
'      End If
'   End If
   If Me.textCU29.Enabled = True Then
      Cancel = False
      textCU29_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   '2008/9/4 add by sonia
   If Me.textCU87.Enabled = True Then
      If Me.textCU30.Text <> "" And Me.textCU87.Text > "010" And Me.textCU64.Text = "1" Then
         strTmp = "聯絡地址有郵遞區號，請確認地址國籍是否為" & textCU87 & " ? "
         If MsgBox(strTmp, vbYesNo + vbCritical) = vbNo Then
            Cancel = False
            Me.textCU87.SetFocus
            tabCustomer.Tab = 2
            Exit Function
         End If
      End If
      Cancel = False
      textCU87_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   '2008/9/4 end
   If Me.textCU30.Enabled = True Then
      Cancel = False
      textCU30_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   'Mark by Amy Add by Amy 2015/09/09 中文地址及聯絡地址不檢查 ex:新北市新莊區Zip(248XX)會被蓋掉成242(因多個zip)
'   If Me.textCU31.Enabled = True Then
'      Cancel = False
'      textCU31_Validate Cancel
'      If Cancel = True Then
'         Exit Function
'      End If
'   End If
   If Me.textCU38.Enabled = True Then
      Cancel = False
      textCU38_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If Me.textCU39.Enabled = True Then
      Cancel = False
      textCU39_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If Me.textCU41.Enabled = True Then
      Cancel = False
      textCU41_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If Me.textCU42.Enabled = True Then
      Cancel = False
      textCU42_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If Me.textCU44.Enabled = True Then
      Cancel = False
      textCU44_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If Me.textCU54.Enabled = True Then
      Cancel = False
      textCU54_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If Me.textCU56.Enabled = True Then
      Cancel = False
      textCU56_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If Me.textCU58.Enabled = True Then
      Cancel = False
      textCU58_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If Me.textCU61.Enabled = True Then
      Cancel = False
      textCU61_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If Me.textCU71.Enabled = True Then
      Cancel = False
      textCU71_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If Me.textCU72.Enabled = True Then
      Cancel = False
      textCU72_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   'Add By Cheng 2003/11/17
   If Me.textCU73.Enabled = True Then
      Cancel = False
      textCU73_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If Me.textCU74.Enabled = True Then
      Cancel = False
      textCU74_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If Me.textCU75.Enabled = True Then
      Cancel = False
      textCU75_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
'   'add by nickc 2005/12/02
'   If Me.textCU76.Enabled = True Then
'      Cancel = False
'      textCU76_Validate Cancel
'      If Cancel = True Then
'         Exit Function
'      End If
'   End If
'
'   'Add By Sindy 2011/3/4
'   If Me.textCU148.Enabled = True Then
'      Cancel = False
'      textCU148_Validate Cancel
'      If Cancel = True Then
'         Exit Function
'      End If
'   End If
   
   'Add by Morgan 2008/7/30
   If cboContact.Locked = False Then
      Cancel = False
      cboContact_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   'add by nickc 2008/03/12
   If textCU03.Enabled Then
       If Trim(textCU03.Text) = "" And Trim(textCU03.Tag) <> "" And m_EditMode <> 1 Then
           m_fa76 = ""
           '2008/7/17 MODIFY BY SONIA
           'Do While Not m_fa76 = "A" And Not m_fa76 = "C"
           '    m_fa76 = UCase(InputBox("取消代理人編號與客戶編號的關連，代理人的性質應為??(A、C)"))
           old_fa76 = PUB_GetFAgentFA76(textCU03 & String(9 - Len(textCU03), "0"))
           Do While Not m_fa76 = "A" And Not m_fa76 = "B" And Not m_fa76 = "C"
              m_fa76 = UCase(InputBox("取消代理人編號與客戶編號的關連，代理人的性質應為??(A、B、C)，原代理人性質為" & old_fa76))
           Loop
       ElseIf Trim(textCU03.Text) <> "" And Trim(textCU03.Tag) <> "" And Trim(textCU03.Text) <> Trim(textCU03.Tag) Then
           MsgBox "注意：此代理人性質將會被改成  B:公司直接委辦 ！", vbInformation, "請注意！"
       End If
   End If
   'End
   
   'Add by Morgan 2010/7/16
   '新增關係企業時檢查母號特定欄位設定並提醒
   If m_EditMode = 1 And Mid(textCU01, 7) > "00" Then
      strExc(0) = "select * from customer where cu01='" & Left(textCU01, 6) & "00' and cu02='0'"
      strExc(1) = ""
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         With RsTemp
            If textCU73 = "" And Not IsNull(.Fields("cu73")) Then
               strExc(1) = strExc(1) & vbCrLf & vbTab & "FCP年費通知函單筆不跑：" & .Fields("cu73")
            End If
            If textCU75 = "" And Not IsNull(.Fields("cu75")) Then
               strExc(1) = strExc(1) & vbCrLf & vbTab & "FCP領證費自動代繳：" & .Fields("cu75")
            End If
'            If textCU76 = "" And Not IsNull(.Fields("cu76")) Then
'               strExc(1) = strExc(1) & vbCrLf & vbTab & "專利D/N幣別：" & .Fields("cu76") & "（U：美金 N：台幣 R：人民幣）"
'            End If
'            'Add By Sindy 2011/3/4
'            If textCU148 = "" And Not IsNull(.Fields("cu148")) Then
'               strExc(1) = strExc(1) & vbCrLf & vbTab & "商標D/N幣別：" & .Fields("cu148") & "（U：美金 N：台幣 R：人民幣）"
'            End If
'            '2011/3/4 End
            If textCU74 = "" And Not IsNull(.Fields("cu74")) Then
               'Modified by Lydia 2016/08/18
               'strExc(1) = strExc(1) & vbCrLf & vbTab & "FCP年費自動代繳：" & .Fields("cu74")
               strExc(1) = strExc(1) & vbCrLf & vbTab & "FCP年費自動代繳(Y)/寄證書後年費不續辦(N)：" & .Fields("cu74")
            End If
            If textCU122 = "" And Not IsNull(.Fields("cu122")) Then
               strExc(1) = strExc(1) & vbCrLf & vbTab & "FCP是否核對已准專利：" & .Fields("cu122")
            End If
            If textCU57 = "" And Not IsNull(.Fields("cu57")) Then
               strExc(1) = strExc(1) & vbCrLf & vbTab & "專利固定請款對象：" & .Fields("cu57")
            End If
            'Add By Sindy 2011/3/4
            If textCU147 = "" And Not IsNull(.Fields("cu147")) Then
               strExc(1) = strExc(1) & vbCrLf & vbTab & "商標固定請款對象：" & .Fields("cu147")
            End If
            '2011/3/4 End
            If textCU96 = "" And Not IsNull(.Fields("cu96")) Then
               strExc(1) = strExc(1) & vbCrLf & vbTab & "年費代理人：" & .Fields("cu96")
            End If
            If textCU97 = "" And Not IsNull(.Fields("cu97")) Then
               strExc(1) = strExc(1) & vbCrLf & vbTab & "年費請款對象：" & .Fields("cu97")
            End If
            If textCU105 = "" And Not IsNull(.Fields("cu105")) Then
               strExc(1) = strExc(1) & vbCrLf & vbTab & "專利D/N固定列印對象：" & .Fields("cu105")
            End If
            'Add By Sindy 2011/3/4
            If textCU151 = "" And Not IsNull(.Fields("cu151")) Then
               strExc(1) = strExc(1) & vbCrLf & vbTab & "商標D/N固定列印對象：" & .Fields("cu151")
            End If
            '2011/3/4 End
            If txtCU(130) = "" And Not IsNull(.Fields("cu130")) Then
               strExc(1) = strExc(1) & vbCrLf & vbTab & "專利領證折扣：" & .Fields("cu130") & "%"
            End If
            If txtCU(131) = "" And Not IsNull(.Fields("cu131")) Then
               strExc(1) = strExc(1) & vbCrLf & vbTab & "專利年費折扣：" & .Fields("cu131") & "%"
            End If
            If textCU36 = "" And Not IsNull(.Fields("cu36")) Then
               strExc(1) = strExc(1) & vbCrLf & vbTab & "專利全部折扣：" & .Fields("cu36") & "%"
            End If
            If textCU37 = "" And Not IsNull(.Fields("cu37")) Then
               strExc(1) = strExc(1) & vbCrLf & vbTab & "專利申請/翻譯折扣：" & .Fields("cu37") & "%"
            End If
            If textCU38 = "" And Not IsNull(.Fields("cu38")) Then
               strExc(1) = strExc(1) & vbCrLf & vbTab & "專利全部折扣起始日：" & TransDate(.Fields("cu38"), 1)
            End If
            If txtCU(133) = "" And Not IsNull(.Fields("cu133")) Then
               strExc(1) = strExc(1) & vbCrLf & vbTab & "專利定稿份數：" & .Fields("cu133")
            End If
            If txtCU(135) = "" And Not IsNull(.Fields("cu135")) Then
               strExc(1) = strExc(1) & vbCrLf & vbTab & "專利請款單份數：" & .Fields("cu135")
            End If
            If txtCU(124) = "" And Not IsNull(.Fields("cu124")) Then
               strExc(1) = strExc(1) & vbCrLf & vbTab & "專利以 EMail 通知：" & .Fields("cu124")
            End If
            If txtCU(137) = "" And Not IsNull(.Fields("cu137")) Then
               strExc(1) = strExc(1) & vbCrLf & vbTab & "專利 Email 同時寄紙本：" & .Fields("cu137")
            End If
   
            If strExc(1) <> "" Then
               If MsgBox("下列欄位母號有設定但本關係企業並未設定，請確認是否要繼續？" & vbCrLf & strExc(1), vbYesNo + vbDefaultButton2) = vbNo Then
                  Exit Function
               End If
            End If
         End With
      End If
   End If
   'end 2010/7/16
   
   '2010/10/20 add by sonia
   If textCU32 <> "N" And InStr(1, textCU79, "不寄雜誌日期", 1) > 0 Then
      MsgBox "注意：客戶備註有 不寄雜誌日期 字樣，若要寄台一雜誌請取消 不寄雜誌日期... 字樣！", vbInformation, "請注意！"
      Exit Function
   End If
   '2010/10/20 end
   
   'Add By Sindy 2012/5/24
   If Me.textCU100.Enabled = True Then
      Cancel = False
      textCU100_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If Me.textCU103.Enabled = True Then
      Cancel = False
      textCU103_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If Me.textCU104.Enabled = True Then
      Cancel = False
      textCU104_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If Me.textCU105.Enabled = True Then
      Cancel = False
      textCU105_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If Me.textCU151.Enabled = True Then
      Cancel = False
      textCU151_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If Me.textCU106.Enabled = True Then
      Cancel = False
      textCU106_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If Me.textCU152.Enabled = True Then
      Cancel = False
      textCU152_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   'Add By Sindy 2013/1/17
   For i = 0 To 1
      If Me.Combo2(i).Enabled = True Then
         Cancel = False
         Combo2_Validate i, Cancel
         If Cancel = True Then
            Exit Function
         End If
      End If
   Next i
   '2013/1/17 End

   If Me.textCU107.Enabled = True Then
      Cancel = False
      textCU107_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If Me.textCU108.Enabled = True Then
      Cancel = False
      textCU108_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If Me.textCU109.Enabled = True Then
      Cancel = False
      textCU109_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   'Add By Sindy 2025/3/10
   If Me.textCU203.Enabled = True Then
      Cancel = False
      textCU203_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If Me.textCU204.Enabled = True Then
      Cancel = False
      textCU204_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If Me.textCU205.Enabled = True Then
      Cancel = False
      textCU205_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   '2025/3/10 END
   
   If Me.textCU125.Enabled = True Then
      Cancel = False
      textCU125_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If Me.textCU112.Enabled = True Then
      Cancel = False
      textCU112_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If Me.textCU114.Enabled = True Then
      Cancel = False
      textCU114_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If Me.textCU12.Enabled = True Then
      Cancel = False
      textCU12_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If Me.textCU13.Enabled = True Then
      Cancel = False
      textCU13_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If Me.textCU132.Enabled = True Then
      Cancel = False
      textCU132_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   'Add By Sindy 2013/8/15
   If Me.textCU139.Enabled = True Then
      Cancel = False
      textCU139_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textCU145.Enabled = True Then
      Cancel = False
      textCU145_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If Me.textCU153.Enabled = True Then
      Cancel = False
      textCU153_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   'Modify By Sindy 2023/9/4 mark,秀玲說財務處自行維護
'   'Add By Sindy 2013/12/17
'   If Me.textCU144.Enabled = True Then
'      Cancel = False
'      textCU144_Validate Cancel
'      If Cancel = True Then
'         Exit Function
'      End If
'   End If
'   '2013/12/17 END

   If Me.textCU14.Enabled = True Then
      Cancel = False
      textCU14_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If Me.textCU20.Enabled = True Then
      Cancel = False
      textCU20_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If Me.textCU116.Enabled = True Then
      Cancel = False
      textCU116_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If Me.textCU117.Enabled = True Then
      Cancel = False
      textCU117_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If Me.textCU118.Enabled = True Then
      Cancel = False
      textCU118_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   'Added by Morgan 2018/11/14
   If Me.textCU176.Enabled = True Then
      Cancel = False
      textCU176_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   'end 2018/11/14
   
   'Added by Morgan 2021/10/7
   If Me.textCU185.Enabled = True Then
      Cancel = False
      textCU185_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   'Removed by Morgan 2025/2/27
   'If Me.textCU186.Enabled = True Then
   '   Cancel = False
   '   textCU186_Validate Cancel
   '   If Cancel = True Then
   '      Exit Function
   '   End If
   'End If
   'end 2025/2/27
   If Me.textCU187.Enabled = True Then
      Cancel = False
      textCU187_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If Me.textCU188.Enabled = True Then
      Cancel = False
      textCU188_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   'end 2021/10/7
   
   If Me.textCU32.Enabled = True Then
      Cancel = False
      textCU32_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If Me.textCU34.Enabled = True Then
      Cancel = False
      textCU34_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If Me.textCU36.Enabled = True Then
      Cancel = False
      textCU36_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If Me.textCU37.Enabled = True Then
      Cancel = False
      textCU37_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If Me.textCU45.Enabled = True Then
      Cancel = False
      textCU45_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If Me.textCU47.Enabled = True Then
      Cancel = False
      textCU47_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If Me.textCU48.Enabled = True Then
      Cancel = False
      textCU48_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If Me.textCU50.Enabled = True Then
      Cancel = False
      textCU50_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If Me.textCU51.Enabled = True Then
      Cancel = False
      textCU51_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If Me.textCU53.Enabled = True Then
      Cancel = False
      textCU53_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If Me.textCU57.Enabled = True Then
      Cancel = False
      textCU57_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If Me.textCU147.Enabled = True Then
      Cancel = False
      textCU147_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If Me.textCU60.Enabled = True Then
      Cancel = False
      textCU60_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If Me.textCU63.Enabled = True Then
      Cancel = False
      textCU63_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If Me.textCU77.Enabled = True Then
      Cancel = False
      textCU77_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If Me.textCU149.Enabled = True Then
      Cancel = False
      textCU149_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If Me.textCU79.Enabled = True Then
      Cancel = False
      textCU79_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If Me.textCU91.Enabled = True Then
      Cancel = False
      textCU91_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If Me.textCU93.Enabled = True Then
      Cancel = False
      textCU93_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If Me.textCU94.Enabled = True Then
      Cancel = False
      textCU94_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If Me.textCU96.Enabled = True Then
      Cancel = False
      textCU96_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If Me.textCU97.Enabled = True Then
      Cancel = False
      textCU97_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If Me.textCU98.Enabled = True Then
      Cancel = False
      textCU98_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If Me.textCU99.Enabled = True Then
      Cancel = False
      textCU99_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If cboStatus.Locked = False Then
      Cancel = False
      cboStatus_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   '2012/5/24 End
   
   'Added by Morgan 2022/9/13
   'E化及份數檢查
   If txtCU(124) <> "" And txtCU(137) = "" Then
      If txtCU(133) <> "" Then
         MsgBox "當設定【專利以 EMail 通知】，需同時設定【專利 Email 同時寄紙本】才可指定【專利定稿份數】！", vbExclamation
         tabCustomer.Tab = 4
         txtCU(137).SetFocus
         Exit Function
      ElseIf txtCU(135) <> "" Then
         MsgBox "當設定【專利以 EMail 通知】，需同時設定【專利 Email 同時寄紙本】才可指定【專利請款單份數】！", vbExclamation
         tabCustomer.Tab = 4
         txtCU(137).SetFocus
         Exit Function
      End If
   End If
   If txtCU(126) <> "" And txtCU(138) = "" Then
      If txtCU(134) <> "" Then
         MsgBox "當設定【商標以 EMail 通知】，需同時設定【商標 Email 同時寄紙本】才可指定【商標定稿份數】！", vbExclamation
         tabCustomer.Tab = 5
         txtCU(138).SetFocus
         Exit Function
      ElseIf txtCU(136) <> "" Then
         MsgBox "當設定【商標以 EMail 通知】，需同時設定【商標 Email 同時寄紙本】才可指定【商標請款單份數】！", vbExclamation
         tabCustomer.Tab = 5
         txtCU(138).SetFocus
         Exit Function
      End If
   End If
   'end 2022/9/13

   TxtValidate = True
End Function

'add by nickc 2006/10/24
' 設定欄位的內容
Private Sub SetFieldNewData(ByVal strName As String, Optional ByVal strData As String = "#==#")
Dim nIndex As Integer
Dim intNewIdx As Integer
   
   For nIndex = 0 To TF_CU - 1 'edit by nickc 2006/10/24  MAX_FIELD - 1
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

'add by nickc  2006/10/24
' 從記錄中更新欄位內容
Private Sub UpdateFieldOldData(ByRef rsTmp As ADODB.Recordset)
Dim nIndex As Integer
Dim strTmp As String

   For nIndex = 0 To TF_CU - 1 'edit by nickc 2006/10/24  MAX_FIELD - 1
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

Private Sub textCU01_GotFocus()
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textCU01.IMEMode = 2
   CloseIme
   TextInverse textCU01
End Sub

Private Sub textCU01_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 And m_EditMode = 3 Then OnAction vbKeyF10
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textCU01_LostFocus()
Dim strTmp1 As String
 
 If m_EditMode = 1 Then
    strTmp1 = Right(textCU01, 2)
    If textCU01 <> "" And Len(textCU01) > 6 And strTmp1 <> "00" Then
       textCU32 = "N"
    Else
       textCU32 = ""
    End If
End If
End Sub

Private Sub textCU01_Validate(Cancel As Boolean)
Dim strTmp As String, i As Integer
 
 If m_EditMode = 4 Then Exit Sub
    If Not IsEmptyText(textCU01) Then
       Dim strTmp1 As String
       Dim strTmp2 As String
       If Mid(textCU01, 1, 1) <> "X" Then
          Cancel = True
          MsgBox "客戶編號必須為X開頭", vbCritical + vbOKOnly, "檢核資料"
           Me.textCU01.Text = ""
          textCU01_GotFocus
          Exit Sub
       End If
       If Len(textCU01) < 6 Then
          Cancel = True
          MsgBox "客戶編號請至少輸入六碼", vbCritical + vbOKOnly, "檢核資料"
          textCU01_GotFocus
          Exit Sub
       End If
       If m_EditMode = 1 And IsEmptyText(textCU01) = False Then
          strTmp1 = textCU01 & String(8 - Len(textCU01), "0")
          strTmp2 = textCU02 & String(1 - Len(textCU02), "0")
          If IsRecordExist(strTmp1, strTmp2) = True Then
             Cancel = True
             MsgBox "該筆客戶已存在! ", vbCritical + vbOKOnly, "檢核資料"
             textCU01_GotFocus
             Exit Sub
          End If
          If IsOverAutoNumber("X", Empty, Mid(strTmp1, 2, 5)) = True Then
             Cancel = True
             MsgBox "客戶代碼超過自動編號! ", vbCritical + vbOKOnly, "檢核資料"
             textCU01_GotFocus
             Exit Sub
          End If
       End If
    End If
End Sub

Private Sub textCU02_GotFocus()
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textCU02.IMEMode = 2
   CloseIme
   TextInverse textCU02
End Sub

Private Sub textCU02_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 And m_EditMode = 3 Then OnAction vbKeyF10
End Sub

Private Sub textCU03_GotFocus()
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textCU03.IMEMode = 2
   CloseIme
   TextInverse textCU03
End Sub

Private Sub textCU03_Validate(Cancel As Boolean)
Dim strTmp As String, i As Integer
   
   If m_EditMode = 4 Then Exit Sub
   If Not IsEmptyText(textCU03) And (m_EditMode = 1 Or m_EditMode = 2) Then
      Dim strAgent As String
      Dim strAgentName As String
      If Mid(textCU03, 1, 1) <> "Y" Then
         Cancel = True
         MsgBox "代理人編號開頭必須為Y! ", vbCritical + vbOKOnly, "檢核資料"
         textCU03_GotFocus
         Exit Sub
      End If
      strAgent = textCU03 & String(9 - Len(textCU03), "0")
      strAgentName = GetFAgentName(strAgent)
      If IsEmptyText(strAgentName) Then
         Cancel = True
         MsgBox "無此代理人編號! ", vbCritical + vbOKOnly, "檢核資料"
         textCU03_GotFocus
         Exit Sub
      End If
   End If
End Sub

Private Sub textCU04_GotFocus()
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textCU04.IMEMode = 1
   OpenIme
   TextInverse textCU04
End Sub

Private Sub textCU04_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = ChangeZIP(KeyAscii, textCU04) 'Modify By Sindy 2021/12/13 +, textCU04
End Sub

Private Sub textCU04_Validate(Cancel As Boolean)
   'add by nickc 2008/01/17 若不是修改狀態，將會出不去
   If m_EditMode <> 1 And m_EditMode <> 2 Then Exit Sub
   If textCU04.Text = "" Then Exit Sub
   '中日文欄位尾碼加空白，最大可輸長度減一(因可能會有造字無法存入問題)
   'Modified by Lydia 2021/01/07 中、英、日文名稱改成判斷字串個數
   'If Not CheckLengthIsOK(textCU04, textCU04.MaxLength - 1) Then
   If Len(textCU04) > textCU04.MaxLength Then
      Cancel = True
   End If
   
End Sub

Private Sub textCU05_GotFocus()
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textCU05.IMEMode = 2
   CloseIme
   TextInverse textCU05
End Sub

Private Sub textCU06_GotFocus()
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textCU06.IMEMode = 1
   OpenIme
   TextInverse textCU06
End Sub

Private Sub textCU06_Validate(Cancel As Boolean)
   'add by nickc 2008/01/17 若不是修改狀態，將會出不去
   If m_EditMode <> 1 And m_EditMode <> 2 Then Exit Sub

   If textCU06.Text = "" Then Exit Sub
    '中日文欄位尾碼加空白，最大可輸長度減一(因可能會有造字無法存入問題)
    'Modified by Lydia 2021/01/07 額外判斷字串個數
   'If Not CheckLengthIsOK(textCU06, textCU06.MaxLength - 1) Then
   If Len(textCU06) > textCU06.MaxLength Then
      Cancel = True
   End If
End Sub

Private Sub textCU07_GotFocus()
   If textCU10 > "010" Then
      'edit by nickc 2007/06/06 切換輸入法改用API
      'textCU07.IMEMode = 0
      CloseIme
   Else
      'edit by nickc 2007/06/06 切換輸入法改用API
      'textCU07.IMEMode = 1
      OpenIme
   End If
   TextInverse textCU07
End Sub

Private Sub textCU07_Validate(Cancel As Boolean)
Dim strTmp As String, i As Integer
   
   If m_EditMode = 4 Then Exit Sub
   
   'Memo by Amy 2023/04/28 客戶為公司, 公司負責人欄不可空白 檢查搬至 TxtValidate
   If textCU07.Text = "" Then Exit Sub
   If Not CheckLengthIsOK(textCU07, textCU07.MaxLength) Then
      Cancel = True
      textCU07_GotFocus
   Else
      If textCU07 <> "" And optCustomer(0).Value = True Then
         ShowMsg "客戶為個人不可輸入公司負責人 !"
         Cancel = True
         textCU07_GotFocus
         Exit Sub
      End If
   End If
End Sub

Private Sub textCU09_GotFocus()
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textCU09.IMEMode = 2
   CloseIme
   TextInverse textCU09
End Sub

Private Sub textCU09_Validate(Cancel As Boolean)
   'If m_EditMode = 4 Then Exit Sub
   Label30(1).Caption = ""
   If textCU09.Text = "" Then Exit Sub
   Label30(1).Caption = ChgType(3, textCU09.Text)
   'Add by Amy 2016/07/06 解:案件來源輸不正確代號仍可存檔
    If Label30(1).Caption = "" Then
        Cancel = True
        textCU09.SetFocus
        textCU09_GotFocus
    End If
    'end 2016/07/06
End Sub

Private Sub textCU10_GotFocus()
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textCU10.IMEMode = 2
   CloseIme
   TextInverse textCU10
End Sub

'Modifiy by Amy 2016/11/24 改成Public
'Private Sub textCU10_Validate(Cancel As Boolean)
Public Sub textCU10_Validate(Cancel As Boolean)

   'If m_EditMode = 4 Then Exit Sub
   Label30(0).Caption = ""
   'Mark by Amy 2022/06/06 不使用
'   '2008/9/4 add by sonia 2010/3/17自UpdateFieldNewData移過來
'   If textCU10 <> "" Then
'      m_CU10 = textCU10
'   Else
'      m_CU10 = ""
'   End If
'   '2008/9/4 end
   
   If textCU10.Text = "" Then Exit Sub
  
   If textCU10.Text = 台灣國家代號 Then
      'add by nickc 2008/01/23 若不是修改狀態，將會出不去
      If m_EditMode <> 1 And m_EditMode <> 2 Then Exit Sub
      Cancel = True
      ShowMsg "客戶" & MsgText(9153)
      textCU10.SetFocus
      textCU10_GotFocus
   Else
      Label30(0).Caption = ChgType(2, textCU10.Text)
      If Label30(0).Caption = "" Then
          'add by nickc 2008/01/23 若不是修改狀態，將會出不去
          If m_EditMode <> 1 And m_EditMode <> 2 Then Exit Sub
         Cancel = True
         textCU10.SetFocus
         textCU10_GotFocus
      Else
         If m_EditMode = 1 And textCU87.Text = "" Then
            textCU87.Text = textCU10.Text
            Label30(4).Caption = Label30(0).Caption
         End If
         If textCU64 = "" Then
            If Val(textCU10.Text) < 9 Or textCU10.Text = "013" Or textCU10.Text = "020" Then
               textCU64.Text = 1
            '2012/4/13 ADD BY SONIA
            ElseIf Left(textCU10.Text, 3) = "011" Then
               textCU64.Text = 3
            '2012/4/13 END
            Else
               textCU64.Text = 2
            End If
         End If
      End If
   End If
End Sub

Private Sub textCU100_GotFocus()
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textCU11.IMEMode = 2
   CloseIme
   TextInverse textCU100
End Sub

Private Sub textCU100_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textCU100_Validate(Cancel As Boolean)
   'add by nickc 2008/01/17 若不是修改狀態，將會出不去
   If m_EditMode <> 1 And m_EditMode <> 2 Then Exit Sub
   
   If textCU100.Text = "" Then Exit Sub
   If textCU100.Text <> "Y" Then
      ShowMsg "輸入錯誤 !"
      Cancel = True
   End If
End Sub

Private Sub textCU102_GotFocus()
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textCU102.IMEMode = 2
   CloseIme
   TextInverse textCU102
End Sub

Private Sub textCU103_GotFocus()
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textCU103.IMEMode = 1
   OpenIme
   TextInverse textCU103
End Sub

Private Sub textCU103_Validate(Cancel As Boolean)
   'add by nickc 2008/01/17 若不是修改狀態，將會出不去
   If m_EditMode <> 1 And m_EditMode <> 2 Then Exit Sub
   
   If textCU103.Text = "" Then Exit Sub
   If Not CheckLengthIsOK(textCU103, textCU103.MaxLength) Then
      Cancel = True
   End If
End Sub

Private Sub textCU104_GotFocus()
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textCU104.IMEMode = 1
   OpenIme
   TextInverse textCU104
End Sub

Private Sub textCU104_Validate(Cancel As Boolean)
   'add by nickc 2008/01/17 若不是修改狀態，將會出不去
   If m_EditMode <> 1 And m_EditMode <> 2 Then Exit Sub
   
   If textCU104.Text = "" Then Exit Sub
   If Not CheckLengthIsOK(textCU104, textCU104.MaxLength) Then
      Cancel = True
   End If
End Sub

Private Sub textCU105_GotFocus()
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textCU105.IMEMode = 2
   CloseIme
   TextInverse textCU105
End Sub

Private Sub textCU105_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textCU105_Validate(Cancel As Boolean)
   'If m_EditMode = 4 Then Exit Sub
   Label30(12).Caption = ""
   If textCU105.Text = "" Then Exit Sub
   If textCU105 <> "" Then textCU105 = textCU105 & String(9 - Len(textCU105), "0")
   Label30(12).Caption = ChgType(4, textCU105.Text)
   'add by nickc 2008/01/23 若不是修改狀態，將會出不去
   If m_EditMode <> 1 And m_EditMode <> 2 Then Exit Sub
               
   If Label30(12).Caption = "" Then Cancel = True
End Sub

Private Sub textCU113_GotFocus()
   TextInverse textCU113
   CloseIme
End Sub

Private Sub textCU113_Validate(Cancel As Boolean)
   If GetTextLength(textCU113) > textCU113.MaxLength Then
      ShowMsg "欄位長度超過限制(" & textCU113.MaxLength & "個字元)!"
      Cancel = True
   End If
End Sub

Private Sub textCU115_KeyPress(KeyAscii As Integer)
   PUB_EMailFilter KeyAscii 'Added by Morgan 2011/11/30 Email輸入字元檢查
End Sub

Private Sub textCU116_KeyPress(KeyAscii As Integer)
   PUB_EMailFilter KeyAscii 'Added by Morgan 2011/11/30 Email輸入字元檢查
End Sub

Private Sub textCU117_KeyPress(KeyAscii As Integer)
   PUB_EMailFilter KeyAscii 'Added by Morgan 2011/11/30 Email輸入字元檢查
End Sub

Private Sub textCU118_KeyPress(KeyAscii As Integer)
   PUB_EMailFilter KeyAscii 'Added by Morgan 2011/11/30 Email輸入字元檢查
End Sub
'Added by Morgan 2018/11/14
Private Sub textCU176_KeyPress(KeyAscii As Integer)
   PUB_EMailFilter KeyAscii
End Sub
'Added by Morgan 2021/10/7
Private Sub textCU185_KeyPress(KeyAscii As Integer)
   PUB_EMailFilter KeyAscii
End Sub
'Added by Morgan 2018/11/14
Private Sub textCU187_KeyPress(KeyAscii As Integer)
   PUB_EMailFilter KeyAscii
End Sub
'Added by Morgan 2018/11/14
Private Sub textCU188_KeyPress(KeyAscii As Integer)
   PUB_EMailFilter KeyAscii
End Sub

'Add By Sindy 2011/3/4
Private Sub textCU139_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   Select Case textCU139
      Case "Y", "":
      Case Else:
         Select Case m_EditMode
            Case 1, 2:
               Cancel = True
               strTit = "檢核資料"
               strMsg = "不催延展只可輸入Y或空白"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textCU139_GotFocus
         End Select
   End Select
End Sub
'2011/3/4 End

'Add By Sindy 2011/3/4
Private Sub textCU151_GotFocus()
   CloseIme
   TextInverse textCU151
End Sub

'Add By Sindy 2011/3/4
Private Sub textCU151_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

'Add By Sindy 2011/3/4
Private Sub textCU151_Validate(Cancel As Boolean)
   Label30(15).Caption = ""
   If textCU151.Text = "" Then Exit Sub
   If textCU151 <> "" Then textCU151 = textCU151 & String(9 - Len(textCU151), "0")
   Label30(15).Caption = ChgType(4, textCU151.Text)
   If m_EditMode <> 1 And m_EditMode <> 2 Then Exit Sub
   
   If Label30(15).Caption = "" Then Cancel = True
End Sub

Private Sub textCU106_GotFocus()
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textCU106.IMEMode = 2
   CloseIme
   TextInverse textCU106
End Sub

Private Sub textCU106_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textCU106_Validate(Cancel As Boolean)
   'If m_EditMode = 4 Then Exit Sub
   Label30(13).Caption = ""
   If textCU106.Text = "" Then Exit Sub
   If textCU106 <> "" Then textCU106 = textCU106 & String(9 - Len(textCU106), "0")
   Label30(13).Caption = ChgType(4, textCU106.Text)
   'add by nickc 2008/01/23 若不是修改狀態，將會出不去
   If m_EditMode <> 1 And m_EditMode <> 2 Then Exit Sub
               
   If Label30(13).Caption = "" Then Cancel = True
End Sub

'Add By Sindy 2011/3/4
Private Sub textCU152_GotFocus()
   CloseIme
   TextInverse textCU152
End Sub

'Add By Sindy 2011/3/4
Private Sub textCU152_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

'Add By Sindy 2011/3/4
Private Sub textCU152_Validate(Cancel As Boolean)
   Label30(14).Caption = ""
   If textCU152.Text = "" Then Exit Sub
   If textCU152 <> "" Then textCU152 = textCU152 & String(9 - Len(textCU152), "0")
   Label30(14).Caption = ChgType(4, textCU152.Text)
   If m_EditMode <> 1 And m_EditMode <> 2 Then Exit Sub
   
   If Label30(14).Caption = "" Then Cancel = True
End Sub

Private Sub textCU107_GotFocus()
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textCU107.IMEMode = 2
   CloseIme
   TextInverse textCU107
End Sub

Private Sub textCU107_Validate(Cancel As Boolean)
   'add by nickc 2008/01/17 若不是修改狀態，將會出不去
   If m_EditMode <> 1 And m_EditMode <> 2 Then Exit Sub
   
   If textCU107.Text = "" Then Exit Sub
   If Val(textCU107.Text) > 101 Then
      ShowMsg "折扣不可大於 100 !"
      Cancel = True
   End If
End Sub

Private Sub textCU108_GotFocus()
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textCU108.IMEMode = 2
   CloseIme
   TextInverse textCU108
End Sub

Private Sub textCU108_Validate(Cancel As Boolean)
   'add by nickc 2008/01/17 若不是修改狀態，將會出不去
   If m_EditMode <> 1 And m_EditMode <> 2 Then Exit Sub
   
   If textCU108.Text = "" Then Exit Sub
   If Val(textCU108.Text) > 101 Then
      ShowMsg "折扣不可大於 100 !"
      Cancel = True
   End If
End Sub

Private Sub textCU109_GotFocus()
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textCU109.IMEMode = 2
   CloseIme
   TextInverse textCU109
End Sub

Private Sub textCU109_Validate(Cancel As Boolean)
   'add by nickc 2008/01/17 若不是修改狀態，將會出不去
   If m_EditMode <> 1 And m_EditMode <> 2 Then Exit Sub
   
   If textCU109.Text = "" Then Exit Sub
   If CheckIsTaiwanDate(textCU109.Text) = False Then
      Cancel = True
   End If
End Sub

'Add By Sindy 2025/3/10
Private Sub textCU203_GotFocus()
   CloseIme
   TextInverse textCU203
End Sub
Private Sub textCU203_Validate(Cancel As Boolean)
   '若不是修改狀態，將會出不去
   If m_EditMode <> 1 And m_EditMode <> 2 Then Exit Sub
   
   If textCU203.Text = "" Then Exit Sub
   If Val(textCU203.Text) > 101 Then
      ShowMsg "折扣不可大於 100 !"
      Cancel = True
   End If
End Sub
Private Sub textCU204_GotFocus()
   CloseIme
   TextInverse textCU204
End Sub
Private Sub textCU204_Validate(Cancel As Boolean)
   '若不是修改狀態，將會出不去
   If m_EditMode <> 1 And m_EditMode <> 2 Then Exit Sub
   
   If textCU204.Text = "" Then Exit Sub
   If Val(textCU204.Text) > 101 Then
      ShowMsg "折扣不可大於 100 !"
      Cancel = True
   End If
End Sub
Private Sub textCU205_GotFocus()
   CloseIme
   TextInverse textCU205
End Sub
Private Sub textCU205_Validate(Cancel As Boolean)
   '若不是修改狀態，將會出不去
   If m_EditMode <> 1 And m_EditMode <> 2 Then Exit Sub
   
   If textCU205.Text = "" Then Exit Sub
   If CheckIsTaiwanDate(textCU205.Text) = False Then
      Cancel = True
   End If
End Sub
'2025/3/10 END

'Add By Sindy 2009/10/26
Private Sub textCU125_GotFocus()
   OpenIme
   TextInverse textCU125
End Sub

'Add By Sindy 2009/10/26
Private Sub textCU125_Validate(Cancel As Boolean)
   If m_EditMode <> 1 And m_EditMode <> 2 Then Exit Sub
   If textCU125.Text = "" Then Exit Sub
   If Not CheckLengthIsOK(textCU125, textCU125.MaxLength) Then
      Cancel = True
   End If
End Sub

Private Sub textCU11_GotFocus()
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textCU11.IMEMode = 2
   CloseIme
   TextInverse textCU11
End Sub

Private Sub textCU11_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

'Mark by Amy 2024/03/28 改Form2.0後,欄位檢查不符合彈訊息後,造成無法跳出檢查,改至最後做
'Private Sub textCU11_Validate(Cancel As Boolean)
'Dim strTmp As String, i As Integer
'
'   'add by nickc 2008/01/17 若不是修改狀態，將會出不去
'   If m_EditMode <> 1 And m_EditMode <> 2 Then Exit Sub
'
'   If textCU11.Text = "" Then Exit Sub
'   i = -1 'Add by Amy 2023/12/20
'   '個人
'   If optCustomer(0).Value = True Then
'      'Modify by Amy 2020/08/04 +大陸公民身份證證規則
'      If textCU10 = "020" Then
'        If GetTextLength(textCU11.Text) <> 18 Then
'            Call textCU11_GotFocus
'            strTmp = "大陸身份證必須是18碼 ! 請確定 ?"
'            If MsgBox(strTmp, vbYesNo + vbCritical) = vbNo Then
'              Cancel = True
'              Exit Sub
'            End If
'        End If
'        i = 3
'      '台灣
'      ElseIf textCU10 < "010" Then
'        If GetTextLength(textCU11.Text) <> 10 Then
'           Call textCU11_GotFocus
'           strTmp = "身份證必須是10碼 ! 請確定 ?"
'           If MsgBox(strTmp, vbYesNo + vbCritical) = vbNo Then
'              Cancel = True
'              Exit Sub
'           End If
'        End If
'        i = 0
'      End If
'   '非個人且台灣
'   ElseIf textCU10 < "010" Then
'      If GetTextLength(textCU11.Text) <> 8 Then
'         Call textCU11_GotFocus
'         strTmp = "統編必須是8碼 ! 請確定 ?"
'         If MsgBox(strTmp, vbYesNo + vbCritical) = vbNo Then
'            Cancel = True
'            Exit Sub
'         End If
'      End If
'      i = 1
'   End If
'   'Modify by Amy 2023/12/20 ex:X04292040 解決大陸會一直彈統編8碼,無法存檔問題
'   If i <> -1 Then
'      If CheckID(i, textCU11.Text) = False Then
'         'Modify by Amy 2020/08/04 +i=3
'         If i = 0 Or i = 3 Then
'            strTmp = ""
'            If i = 3 Then strTmp = "大陸"
'            strTmp = strTmp & "身分證字號錯誤，是否確定 ?"
'         'end 2020/08/04
'         Else
'            strTmp = "統一編號錯誤，是否確定 ?"
'         End If
'         If MsgBox(strTmp, vbYesNo + vbCritical) = vbNo Then
'            Cancel = True
'         End If
'      End If
'   End If 'i <> -1
'End Sub

Private Sub textCU111_GotFocus()
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textCU111.IMEMode = 2
   CloseIme
   TextInverse textCU111
End Sub

Private Sub textCU111_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 8 And KeyAscii <> Asc("Y") Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub textCU112_GotFocus()
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textCU112.IMEMode = 2
   CloseIme
   TextInverse textCU112
End Sub

Private Sub textCU112_KeyPress(KeyAscii As Integer)
   KeyAscii = ChangeZIP(KeyAscii)
End Sub

Private Sub textCU112_Validate(Cancel As Boolean)
    
   'add by nickc 2008/01/17 若不是修改狀態，將會出不去
   If m_EditMode <> 1 And m_EditMode <> 2 Then Exit Sub
   
   If textCU112.Text = "" Then Exit Sub
   If Not CheckLengthIsOK(textCU112, textCU112.MaxLength) Then
      Cancel = True
   End If
End Sub

Private Sub textCU114_GotFocus()
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textCU114.IMEMode = 2
   CloseIme
   TextInverse textCU114
End Sub

Private Sub textCU114_Validate(Cancel As Boolean)
   'add by nickc 2008/01/17 若不是修改狀態，將會出不去
   If m_EditMode <> 1 And m_EditMode <> 2 Then Exit Sub

   If textCU114.Text = "" Then Exit Sub
    '中日文欄位尾碼加空白，最大可輸長度減一(因可能會有造字無法存入問題)
    'Modified by Lydia 2017/06/14
   'If Not CheckLengthIsOK(textCU114, textCU114.MaxLength - 1) Then
   If Not CheckLengthIsOK(textCU114, 60) Then
      Cancel = True
   End If
End Sub

Private Sub textCU12_GotFocus()
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textCU12.IMEMode = 2
   CloseIme
   TextInverse textCU12
End Sub

Private Sub textCU12_Validate(Cancel As Boolean)
   If textCU12.Text = "" Then Exit Sub
   '2008/9/4 modify by sonia
   'If m_EditMode = 1 Then
   If m_EditMode = 1 Or m_EditMode = 2 Then
   '2008/9/4 end
      '2010/4/19 MODIFY BY SONIA 改判斷非國外部
      'If UCase(Mid(textCU12.Text, 1, 1)) = "S" Then
      If UCase(Mid(textCU12.Text, 1, 1)) <> "F" Then
         'Modified by Morgan 2022/1/20
         'textCU64 = "1"
         If textCU64 = "" Then
            textCU64 = "1"
         ElseIf textCU64 <> "1" And textCU64.Tag <> textCU64 Then
            If MsgBox("此非國外部客戶，定稿語文是否確定不為中文？", vbYesNo + vbQuestion) = vbYes Then
               textCU64.Tag = textCU64 '控制只需回答一次
            Else
               Cancel = True
            End If
         End If
         'end 2022/1//20
      End If
   End If
   Label30(3).Caption = GetPrjSalesBlack(textCU12)
End Sub

Private Sub textCU122_GotFocus()
   CloseIme
   TextInverse textCU122
End Sub

Private Sub textCU122_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 8 And KeyAscii <> Asc("N") Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub textCU123_GotFocus()
   CloseIme
   TextInverse textCU123
End Sub

Private Sub textCU123_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 8 And KeyAscii <> Asc("N") And KeyAscii <> Asc("Y") Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub TextCu128_GotFocus()
   CloseIme
   TextInverse TextCu128
End Sub

Private Sub TextCu128_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 8 And KeyAscii <> Asc("Y") Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub textCU13_GotFocus()
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textCU13.IMEMode = 2
   CloseIme
   TextInverse textCU13
End Sub

'Add By Sindy 2010/11/25
Private Sub textCU13_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textCU13_Validate(Cancel As Boolean)
Dim strTmp As String, i As Integer

   'If m_EditMode = 4 Then Exit Sub
   Label30(2).Caption = ""
   'edit by nickc 2007/09/11
   'Label30(3).Caption = ""
   If textCU13.Text = "" Then Exit Sub
   
   strTmp = ChgType(5, textCU13.Text)
   If strTmp <> "" Then
      i = InStr(strTmp, ",")
      Label30(2).Caption = Left(strTmp, i - 1)
      'edit by nickc 2007/09/11 修正應該同步抓ST15
      'Label30(3).Caption = Mid(strTmp, i + 1)
      If textCU13.Tag <> textCU13.Text Then
           'Add by Amy 2019/12/10 分所人員修改,解除狀態改為空,改智權人員只能改同區人員
           If (m_EditMode = 1 Or m_EditMode = 2) And Left(textCU12, 1) = "S" _
              And textCU12.Enabled = False And cboStatus.Text = MsgText(601) And m_FieldList(79).fiOldData <> MsgText(601) Then
                If GetST15(textCU13.Text) <> textCU12 Then
                    MsgBox "智權人員只能為同區人員！"
                    textCU13.SetFocus
                    textCU13_GotFocus
                    Cancel = True
                    Exit Sub
                End If
           End If
           'end 2019/12/10
           textCU12.Text = GetST15(textCU13.Text)
           'add by nickc 2007/09/11 修正應該同步抓ST15
           Label30(3).Caption = GetPrjSalesBlack(textCU12)
       End If
       'Added by Lydia 2019/02/14 創新業務部人員收文控管
      If m_EditMode = 1 Or m_EditMode = 2 Then
        If PUB_ChkIsT10T20("2", textCU13.Text, m_Tuser, strTmp) = True Then
            textCU13.Text = m_Tuser
            Label30(2).Caption = strTmp
            textCU13.SetFocus
            Call textCU13_GotFocus
            Cancel = True
            Exit Sub
        End If
      End If
      'end 2019/02/14
   Else
      Me.textCU12.Text = ""
      'Add By Sindy 2012/7/25
      If textCU13.Text <> "" Then
         Cancel = True
      End If
      '2012/7/25 End
      'edit by nickc 2007/04/27 智權人員不存在也可以存檔
      'Cancel = True
   End If
End Sub

'2008/12/9 add by sonia
Private Sub textCU132_GotFocus()
   CloseIme
   TextInverse textCU132
End Sub

Private Sub textCU132_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 8 And KeyAscii <> Asc("N") Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub textCU132_Validate(Cancel As Boolean)
   If m_EditMode <> 1 And m_EditMode <> 2 Then Exit Sub
   
   If textCU132.Text = "" Then Exit Sub
   If textCU132.Text <> "N" Then
      ShowMsg "輸入錯誤 !"
      Cancel = True
   End If
End Sub
'2008/12/9 end

'Add By Sindy 2011/1/14
Private Sub textCU145_GotFocus()
   CloseIme
   TextInverse textCU145
End Sub
Private Sub textCU145_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   'Modified by Morgan 2012/1/2 改放 N
   If KeyAscii <> 8 And KeyAscii <> Asc("N") Then
      KeyAscii = 0
      Beep
   End If
End Sub
Private Sub textCU145_Validate(Cancel As Boolean)
   If m_EditMode <> 1 And m_EditMode <> 2 Then Exit Sub
   If textCU145.Text = "" Then Exit Sub
   'Modified by Morgan 2012/1/2 改放 N
   If textCU145.Text <> "N" Then
      ShowMsg "輸入錯誤 !"
      Cancel = True
   End If
End Sub
'2011/1/14 End

'Add By Sindy 2011/3/17
Private Sub textCU153_GotFocus()
   CloseIme
   TextInverse textCU153
End Sub
Private Sub textCU153_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 8 And KeyAscii <> Asc("Y") And KeyAscii <> Asc("N") Then
      KeyAscii = 0
      Beep
   End If
End Sub
Private Sub textCU153_Validate(Cancel As Boolean)
   If m_EditMode <> 1 And m_EditMode <> 2 Then Exit Sub
   If textCU153.Text = "" Then Exit Sub
   If textCU153.Text <> "Y" And textCU153.Text <> "N" Then
      ShowMsg "輸入錯誤 !"
      Cancel = True
   End If
End Sub
'2011/3/17 End

Private Sub textCU14_GotFocus()
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textCU14.IMEMode = 2
   CloseIme
   TextInverse textCU14
End Sub

Private Sub textCU14_Validate(Cancel As Boolean)
'add by nickc 2008/01/17 若不是修改狀態，將會出不去
   If m_EditMode <> 1 And m_EditMode <> 2 Then Exit Sub
   
   If textCU14.Text = "" Then Exit Sub
   If CheckIsTaiwanDate(textCU14.Text) = False Then
      Cancel = True
   End If
End Sub

Private Sub textCU15_GotFocus()
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textCU15.IMEMode = 2
   CloseIme
   TextInverse textCU15
End Sub

Private Sub textCU16_GotFocus()
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textCU16.IMEMode = 2
   CloseIme
   TextInverse textCU16
End Sub

Private Sub textCU17_GotFocus()
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textCU17.IMEMode = 2
   CloseIme
   TextInverse textCU17
End Sub

Private Sub textCU18_GotFocus()
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textCU18.IMEMode = 2
   CloseIme
   TextInverse textCU18
End Sub

'Add by Amy 2019/08/27
Private Sub textCU180_GotFocus()
    CloseIme
    TextInverse textCU180
End Sub

Private Sub textCU180_Validate(Cancel As Boolean)
    If textCU180.Text = "" Or m_EditMode = "0" Then Exit Sub
        
    If Not CheckLengthIsOK(textCU180, textCU180.MaxLength) Then
        Cancel = True
    End If
End Sub
'end 2019/08/27

Private Sub textCU19_GotFocus()
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textCU19.IMEMode = 2
   CloseIme
   TextInverse textCU19
End Sub

'Add by Amy 2024/08/26
Private Sub textCU191_Validate(Cancel As Boolean)
   If textCU191 = MsgText(601) Or m_EditMode = "0" Then Exit Sub
   If Len(textCU191) < 4 Then
      MsgBox Replace(Label41(39), "：", "") & "需輸入姓名+職稱", vbExclamation
      If Me.ActiveControl.Name <> "textCU191" Then Cancel = True
   End If
End Sub

Private Sub textCU20_GotFocus()
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textCU20.IMEMode = 2
   CloseIme
   TextInverse textCU20
End Sub

Private Sub textCU20_KeyPress(KeyAscii As Integer)
   PUB_EMailFilter KeyAscii 'Added by Morgan 2011/11/30 Email輸入字元檢查
End Sub

Private Sub textCU20_Validate(Cancel As Boolean)
   If textCU20.Text = "" Or m_EditMode = "0" Then Exit Sub
   Cancel = Not PUB_CheckMail(textCU20.Text)
End Sub

'Add by Morgan 2008/1/16
Private Sub textCU116_GotFocus()
   CloseIme
   TextInverse textCU116
End Sub

Private Sub textCU116_Validate(Cancel As Boolean)
   If textCU116.Text = "" Or m_EditMode = "0" Then Exit Sub
   Cancel = Not PUB_CheckMail(textCU116.Text)
End Sub

Private Sub textCU117_GotFocus()
   CloseIme
   TextInverse textCU117
End Sub

Private Sub textCU117_Validate(Cancel As Boolean)
   If textCU117.Text = "" Or m_EditMode = "0" Then Exit Sub
   Cancel = Not PUB_CheckMail(textCU117.Text)
End Sub

Private Sub textCU118_GotFocus()
   CloseIme
   TextInverse textCU118
End Sub

Private Sub textCU118_Validate(Cancel As Boolean)
   If textCU118.Text = "" Or m_EditMode = "0" Then Exit Sub
   Cancel = Not PUB_CheckMail(textCU118.Text)
End Sub
'Added by Morgan 2018/11/14
Private Sub textCU176_GotFocus()
   CloseIme
   TextInverse textCU176
End Sub
'Added by Morgan 2021/10/7
Private Sub textCU185_GotFocus()
   CloseIme
   TextInverse textCU185
End Sub

'Added by Morgan 2021/10/7
Private Sub textCU187_GotFocus()
   CloseIme
   TextInverse textCU187
End Sub
'Added by Morgan 2021/10/7
Private Sub textCU188_GotFocus()
   CloseIme
   TextInverse textCU188
End Sub

'Added by Morgan 2018/11/14
Private Sub textCU176_Validate(Cancel As Boolean)
   If textCU176.Text = "" Or m_EditMode = "0" Then Exit Sub
   Cancel = Not PUB_CheckMail(textCU176.Text)
End Sub

'Added by Morgan 2021/10/7
Private Sub textCU185_Validate(Cancel As Boolean)
   If textCU185.Text = "" Or m_EditMode = "0" Then Exit Sub
   Cancel = Not PUB_CheckMail(textCU185.Text)
End Sub

'Added by Morgan 2021/10/7
Private Sub textCU187_Validate(Cancel As Boolean)
   If textCU187.Text = "" Or m_EditMode = "0" Then Exit Sub
   Cancel = Not PUB_CheckMail(textCU187.Text)
End Sub

'Added by Morgan 2021/10/7
Private Sub textCU188_Validate(Cancel As Boolean)
   If textCU188.Text = "" Or m_EditMode = "0" Then Exit Sub
   Cancel = Not PUB_CheckMail(textCU188.Text)
End Sub

Private Sub textCU21_GotFocus()
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textCU21.IMEMode = 2
   CloseIme
   TextInverse textCU21
End Sub

Private Sub textCU22_GotFocus()
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textCU22.IMEMode = 2
   CloseIme
   TextInverse textCU22
End Sub

Private Sub textCU23_GotFocus()
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textCU23.IMEMode = 1
   OpenIme
   TextInverse textCU23
End Sub

Private Sub textCU23_KeyPress(KeyAscii As MSForms.ReturnInteger)
    Dim strAddr As String, strNewArea As String, strZipCode As String, strCountry As String, strROC As String, strIndArea As String
    Dim intArea As Integer, intFocus As Integer
    Dim bolMany As Boolean
    
    'Modify by Amy 2023/06/19 從下面搬上來
    KeyAscii = ChangeZIP(KeyAscii, textCU23) 'Modify By Sindy 2021/12/13 +, textCU23
    
    'Add by Amy 2016/12/20 +新增才做,因若加區會與案件之地址不一致,接洽單會一直彈訊息
    If m_EditMode <> 1 Then Exit Sub
    
    '臺灣地址判斷 Add by Amy 2016/05/26
    If LTrim(textCU23) <> MsgText(601) And textCU10 < "010" And InStr(LTrim(textCU23), "後補") = 0 Then
        strROC = ""
        strAddr = textCU23
        If Left(strAddr, 4) = "中華民國" Then strROC = strROC & Left(strAddr, 4): strAddr = Mid(strAddr, 5)
        If Left(strAddr, 3) = "臺灣省" Or Left(strAddr, 3) = "台灣省" Then strROC = strROC & Left(strAddr, 3): strAddr = Mid(strAddr, 4)
        If Left(strAddr, 2) = "臺灣" Or Left(strAddr, 2) = "台灣" Then strROC = strROC & Left(strAddr, 2): strAddr = Mid(strAddr, 3)
        '去除xx工業區查(台中工業區/台塑工業園區不取代,可能抓錯zip)
        strIndArea = "True"
        strAddr = ReplaceIndArea(strAddr, strIndArea)
        If strIndArea = "True" Then strIndArea = MsgText(601)
        If Left(strAddr, 4) = "新竹新竹" And (strIndArea = "科學工業園區" Or strIndArea = "科學園區") Then
            strIndArea = "新竹" & strIndArea
            strAddr = Mid(strAddr, 3)
        End If
        If Len(LTrim(strAddr)) > 4 And (Mid(strAddr, 3, 1) = "市" Or Mid(strAddr, 3, 1) = "縣") Then
            '輸到路/街/段
            If Asc("路") = KeyAscii Or Asc("街") = KeyAscii Or Asc("段") = KeyAscii Then
                intFocus = Val(textCU23.SelStart) - Len(strROC) - Len(strIndArea)
                strAddr = Mid(strAddr, 1, intFocus) & Chr(KeyAscii) & Mid(strAddr, intFocus + 1) 'KeyPress未完成時地址欄位尚未顯示目前字,故先加入當下的字查
                '有鄉/鎮/市/區
                'Modify by Amy 2018/12/19 +判斷第七個字 ex:嘉義縣阿里山鄉 X80024
                If Mid(strAddr, 7, 1) = "市" Or Mid(strAddr, 7, 1) = "區" Or Mid(strAddr, 7, 1) = "鄉" Or Mid(strAddr, 7, 1) = "鎮" _
                  Or Mid(strAddr, 6, 1) = "市" Or Mid(strAddr, 6, 1) = "區" Or Mid(strAddr, 6, 1) = "鄉" Or Mid(strAddr, 6, 1) = "鎮" _
                  Or Mid(strAddr, 5, 1) = "市" Or Mid(strAddr, 5, 1) = "區" Or Mid(strAddr, 5, 1) = "鄉" Or Mid(strAddr, 5, 1) = "鎮" Then
                    strZipCode = GetZipCode_Tai(1, strAddr, , bolMany, , strCountry)
                    If strZipCode <> MsgText(601) Then
                        If bolMany = False Then
                            Call ChkZipData(2, textCU23, strZipCode, , strCountry)
                            textCU23.SelStart = intFocus + Len(strROC) + Len(strIndArea)
                            textCU23.SelLength = 0
                            OpenIme
                        Else
                            '多筆以縣/市+鄉/鎮/市/區及路名查
                            bolMany = False
                            strZipCode = GetZipCode_Tai(3, Mid(strAddr, 1, intFocus + 1), intArea, bolMany, , strCountry)
                            If strZipCode <> MsgText(601) And bolMany = False Then
                                Call ChkZipData(2, textCU23, strZipCode, , strCountry)
                                textCU23.SelStart = intFocus + Len(strROC) + Len(strIndArea)
                                textCU23.SelLength = 0
                                OpenIme
                            End If
                        End If
                    End If
                '沒鄉/鎮/市/區
                Else
                    '取 段/路/街 查
                    strZipCode = GetZipCode_Tai(2, Mid(strAddr, 1, intFocus + 1), intArea, bolMany, strNewArea, strCountry)
                    If strZipCode <> MsgText(601) And bolMany = False Then
                        '補上查到的區,避免輸入兩個同樣的字(路/街/段)被取代,故不用Replace
                        textCU23 = strROC & Left(strAddr, 3) & strNewArea & strIndArea & Mid(strAddr, 4, intArea - 4) & Mid(strAddr, intFocus + 2)
                        Call ChkZipData(2, textCU23, strZipCode, , strCountry)
                        textCU23.SelStart = intFocus + Len(strROC) + Len(strNewArea) + Len(strIndArea)
                        textCU23.SelLength = 0
                        OpenIme
                    End If
                End If
            End If
        End If
        
    End If
    'end 2016/05/26
End Sub

'Add by Amy 2016/05/26 中文地址
Private Sub textCU23_LostFocus()
    Dim strZipCode As String, strAddr As String, strCountry As String, strCityN As String, strIndArea As String, strNewArea As String, strROC As String
    Dim bolMany As Boolean, intArea As Integer
    Dim intAns As Integer 'Add by Amy 2025/03/18

    If m_EditMode <> 1 And m_EditMode <> 2 Then Exit Sub
    'Modify by Amy 2019/09/02 原:textCU23 = MsgText(601) Or textCU10 >= "010"(客戶國籍) and InStr(LTrim(textCU23), "後補") > 0 ex:X58062
    'Modify by Amy 2022/06/06 中文地有輸,客戶國籍空白,彈是否為臺灣國籍,按否 備註會加 臺灣地址格式不檢查,再補上客戶國籍,備註不會拿掉
    'If textCU23 = MsgText(601) Or textCU10 >= "010" Then Exit Sub
    If textCU23 = MsgText(601) Or textCU10 = MsgText(601) Or textCU10 >= "010" Then Exit Sub
    'end 2022/06/06
    If InStr(LTrim(textCU23), "後補") > 0 Then Exit Sub
    'end 2019/09/02
    If InStr(textCU79, "臺灣地址格式不檢查") > 0 Then Exit Sub 'Add byAmy 2016/12/20
    If m_FieldList(22).fiOldData = textCU23 And m_FieldList(111).fiOldData = textCU112 Then Exit Sub 'Add by Amy 2019/10/03
    
    textCU23 = ReplaceAddrTW(textCU23, , textCU10) 'Add by Amy 2025/06/30
    If m_EditMode = 2 Then textCU23.Tag = textCU23
AgainCheck1:
    strROC = ""
    'Add by Amy 2016/12/20 +if
    If m_EditMode = 1 Then
        strAddr = textCU23
    Else
        '修改時,原沒區不帶區但仍需判斷 zip是否正確
        strAddr = textCU23.Tag
    End If
    If Left(strAddr, 4) = "中華民國" Then strROC = strROC & Left(strAddr, 4): strAddr = Mid(strAddr, 5)
    If Left(strAddr, 3) = "臺灣省" Or Left(strAddr, 3) = "台灣省" Then strROC = strROC & Left(strAddr, 3): strAddr = Mid(strAddr, 4)
    If Left(strAddr, 2) = "臺灣" Or Left(strAddr, 2) = "台灣" Then strROC = strROC & Left(strAddr, 2): strAddr = Mid(strAddr, 3)
    '去除xx工業區查(台中工業區/台塑工業園區不取代,可能抓錯zip)
    strIndArea = "True"
    strAddr = ReplaceIndArea(strAddr, strIndArea)
    If strIndArea = "True" Then strIndArea = MsgText(601)
    If Left(strAddr, 4) = "新竹新竹" And (strIndArea = "科學工業園區" Or strIndArea = "科學園區") Then
        strIndArea = "新竹" & strIndArea
        strAddr = Mid(strAddr, 3)
    End If
    '** 第3個字是 縣 / 市
    If Mid(strAddr, 3, 1) = "市" Or Mid(strAddr, 3, 1) = "縣" Or Mid(strAddr, 1, 3) = "釣魚臺" Or Mid(strAddr, 1, 3) = "海南島" Then
                
        'Modify by Amy 2018/12/19 +判斷第七個字 ex:嘉義縣阿里山鄉 X80024
        If Mid(strAddr, 7, 1) = "市" Or Mid(strAddr, 7, 1) = "區" Or Mid(strAddr, 7, 1) = "鄉" Or Mid(strAddr, 7, 1) = "鎮" _
          Or Mid(strAddr, 6, 1) = "市" Or Mid(strAddr, 6, 1) = "區" Or Mid(strAddr, 6, 1) = "鄉" Or Mid(strAddr, 6, 1) = "鎮" _
          Or Mid(strAddr, 5, 1) = "市" Or Mid(strAddr, 5, 1) = "區" Or Mid(strAddr, 5, 1) = "鄉" Or Mid(strAddr, 5, 1) = "鎮" Then
            '傳入地址前6個字抓到郵遞區號
            intArea = 6
            strZipCode = GetPostZip(Left(strAddr, 6), 6, , strCountry, bolMany)
            '傳入地址前5個字取郵遞區號
            If strZipCode = MsgText(601) Then strZipCode = GetPostZip(Left(strAddr, 5), 5, , strCountry, bolMany): intArea = 5
            '抓到郵遞區號
            If strZipCode <> MsgText(601) Then
                If bolMany = True Then
                    '多筆以縣/市+鄉/鎮/市/區及路名查
                    bolMany = False
                    strZipCode = GetZipCode_Tai(3, strAddr, , bolMany, , strCountry)
                    If strZipCode <> MsgText(601) Then
                        '限制縣/市+鄉/鎮/市/區及路名查:一筆-直接帶/多筆-進查詢畫面
                        If bolMany = False Then
                            Call ChkZipData(2, textCU23, strZipCode, , strCountry)
                        Else
                            Call ChkZipData(1, textCU23, strZipCode, intArea, strCountry)
                        End If
                    End If
                Else
                    '非多筆
                    Call ChkZipData(2, textCU23, strZipCode, intArea, strCountry)
                End If
            
            Else
                '判斷是否有此區/鄉/鎮
                strZipCode = GetPostZip(Mid(strAddr, 4, intArea - 3), intArea - 3, , strCountry, bolMany, "Pzd03")
                If strZipCode <> MsgText(601) Then
                    '區別錯,進入查詢畫面
                    Call ChkZipData(3, textCU23, strZipCode, intArea, strCountry)
                Else
                    '當作沒區只有路 ex:新竹縣or市園區二路
                    bolMany = False
                    strZipCode = GetZipCode_Tai(2, strAddr, intArea, bolMany, strNewArea, strCountry)
                    If strZipCode <> MsgText(601) Then
                        '以縣/市及路名查:一筆-直接帶/多筆-進查詢畫面
                        If bolMany = False Then
                            'Modify by Amy 2016/12/20 +if
                            If m_EditMode = 1 Then
                                textCU23 = strROC & Left(textCU23, 3) & strNewArea & strIndArea & _
                                    Replace(Replace(Replace(Replace(textCU23, strIndArea, ""), strNewArea, ""), Left(strAddr, 3), ""), strROC, "")
                            Else
                                '修改時,原沒區不帶區但仍需判斷 zip是否正確
                                textCU23.Tag = strROC & Left(strAddr, 3) & strNewArea & strIndArea & _
                                    Replace(Replace(Replace(Replace(textCU23.Tag, strIndArea, ""), strNewArea, ""), Left(strAddr, 3), ""), strROC, "")
                            End If
                            Call ChkZipData(2, textCU23, strZipCode, intArea, strCountry)
                            
                        Else
                            intArea = 0
                            Call ChkZipData(4, textCU23, strZipCode, intArea, strCountry)
                            Exit Sub
                        End If
                    End If
                End If
            End If
                    
        '無鄉/鎮/市/區
        Else
            '以路/街 抓是否有zip
            strZipCode = GetZipCode_Tai(2, strAddr, intArea, bolMany, strNewArea, strCountry)
            If strZipCode <> MsgText(601) Then
                If bolMany = True Then
                    '多筆
                    intArea = 0
                    Call ChkZipData(4, textCU23, strZipCode, intArea, strCountry)
                Else
                    '非多筆
                    'Modify by Amy 2016/12/20 +if
                    If m_EditMode = 1 Then
                        textCU23 = strROC & Left(textCU23, 3) & strNewArea & strIndArea & _
                            Replace(Replace(Replace(Replace(textCU23, strIndArea, ""), strNewArea, ""), Left(strAddr, 3), ""), strROC, "")
                    Else
                        '修改時,原沒區不帶區但仍需判斷 zip是否正確
                        textCU23.Tag = strROC & Left(strAddr, 3) & strNewArea & strIndArea & _
                            Replace(Replace(Replace(Replace(textCU23.Tag, strIndArea, ""), strNewArea, ""), Left(strAddr, 3), ""), strROC, "")
                    End If
                    Call ChkZipData(2, textCU23, strZipCode, intArea, strCountry)
                End If
            '都抓不到ZipCode
            ElseIf strZipCode = MsgText(601) Then
                If CheckTaiwanAddr_Tai(textCU23, textCU112, "000", "", strZipCode, , False, Me.Name) = False Then
                    'Add by Amy 2016/12/20 +詢問是否為臺灣地址(ex:國籍為台灣但地址為大陸 X26021020不判斷)
                    'Modify by Amy 2025/03/18 晉溢 修改 X72368 中文地址時,少輸臺北的臺字,導致出現下列訊息,因無「取消」鈕,故按「是」而寫入cu79中
                    'If MsgBox("國籍為「臺灣」但中文地址非臺灣正確格式，請問是臺灣地址嗎？", vbYesNo + vbCritical) = vbYes Then
                    intAns = MsgBox("國籍為「臺灣」但中文地址非臺灣正確格式，請問是臺灣地址嗎？" & vbCrLf & _
                                    "  是  :為臺灣地址,存檔" & vbCrLf & _
                                    "  否  :非臺灣地址,存檔" & vbCrLf & _
                                    "取消:不存檔,回前畫面", vbExclamation + vbYesNoCancel + vbDefaultButton1, "重要訊息！")
                    If intAns = vbCancel Then
                        Exit Sub
                    ElseIf intAns = vbYes Then
                        'Modify by Amy 2025/03/18 「臺灣地址格式不檢查」字樣前後是否有字,是否加;
                        If InStr(textCU79, "臺灣地址格式不檢查") > 0 Then
'                            strExc(0) = Mid(textCU79, InStr(textCU79, ";臺灣地址格式不檢查") + 10)
'                            'Modify by Amy 2019/10/2 strExc(0)為空會錯
'                            If strExc(0) <> MsgText(601) Then strExc(0) = Mid(strExc(0), InStr(strExc(0), ";"))
'                            textCU79 = Replace(textCU79, ";臺灣地址格式不檢查" & strExc(0), "")
                            strExc(0) = Mid(textCU79, InStr(textCU79, "臺灣地址格式不檢查"), 16)
                            textCU79 = Replace(Replace(textCU79, strExc(0), ""), ";;", ";")
                        End If
                        If strZipCode = "格式錯誤" Then
                            'Modify by Amy 2020/04/10 +if  因X5044001 地址:臺中大里工業區仁化路 彈二次強制表單 會錯
                            If PUB_CheckFormExist("frm100135") = False Then
                                frm100135.Show vbModal
                                Call ChkZipData(9, textCU23, strZipCode)
                            End If
                        Else
                            Call ChkZipData(3, textCU23, strZipCode)
                        End If
                    Else
                        'textCU79 = textCU79 & ";臺灣地址格式不檢查" & strSrvDate(2) & ";"
                        textCU79 = Replace(textCU79 & IIf(Len(Trim(textCU79)) > 0, ";", "") & "臺灣地址格式不檢查" & strSrvDate(2), ";;", ";")
                    End If
                    'end 2025/03/18
                    'end 2016/12/20
                    Exit Sub
                End If
            End If
        End If
    
    '** 第三3個字無 縣 / 市
    Else
        '傳入地址前2個字判斷是否有其縣/市
        strCityN = "Pzd02"
        strZipCode = GetPostZip(Left(strAddr, 2), 2, 1, strCountry, bolMany, "Pzd02", strCityN)
        If strZipCode <> MsgText(601) Then
            If bolMany = False Then
                '只有一筆
                'Modify by Amy 2016/12/20 +if
                If m_EditMode = 1 Then
                    textCU23 = strROC & strCityN & strIndArea & _
                        Replace(Replace(Replace(textCU23, strIndArea, ""), strCityN, ""), strROC, "")
                Else
                    '修改時,原沒區不帶區但仍需判斷 zip是否正確
                    textCU23.Tag = strROC & strCityN & strIndArea & _
                        Replace(Replace(Replace(textCU23.Tag, strIndArea, ""), strCityN, ""), strROC, "")
                End If
                GoTo AgainCheck1
            Else
                '新竹、嘉義會有2筆
                intArea = 0
                Call ChkZipData(5, textCU23, strZipCode, intArea)
            End If
        Else
            'Add by Amy 2016/12/20 +詢問是否為臺灣地址(ex:國籍為台灣但地址為大陸 X26021020不判斷)
            'Modify by Amy 2025/03/18 晉溢 修改 X72368 中文地址時,少輸臺北的臺字,導致出現下列訊息,因無「取消」鈕,故按「是」而寫入cu79中
            'If MsgBox("國籍為「臺灣」但中文地址非臺灣正確格式，請問是臺灣地址嗎？", vbYesNo + vbCritical) = vbYes Then
            intAns = MsgBox("國籍為「臺灣」但中文地址非臺灣正確格式，請問是臺灣地址嗎？" & vbCrLf & _
                                    "  是  :為臺灣地址,存檔" & vbCrLf & _
                                    "  否  :非臺灣地址,存檔" & vbCrLf & _
                                    "取消:不存檔,回前畫面", vbExclamation + vbYesNoCancel + vbDefaultButton1, "重要訊息！")
            If intAns = vbCancel Then
               Exit Sub
            ElseIf intAns = vbYes Then
               If InStr(textCU79, "臺灣地址格式不檢查") > 0 Then
                    'Modify by Amy 2025/03/18 「臺灣地址格式不檢查」民國年月日 字樣前後是否有文字加;
'                    strExc(0) = Mid(textCU79, InStr(textCU79, ";臺灣地址格式不檢查") + 10)
'                    'Modify by Amy 2019/10/2 strExc(0)為空會錯
'                    If strExc(0) <> MsgText(601) Then strExc(0) = Mid(strExc(0), 1, InStr(strExc(0), ";"))
'                    textCU79 = Replace(textCU79, ";臺灣地址格式不檢查" & strExc(0), "")
'                    textCU79 = textCU79 & IIf(Len(textCU79) > 0, ";", "")
                    strExc(0) = Mid(textCU79, InStr(textCU79, "臺灣地址格式不檢查"), 16)
                    textCU79 = Replace(Replace(textCU79, strExc(0), ""), ";;", ";")
                End If
                MsgBox "臺灣地址格式錯誤!請確認"
                Exit Sub
            ElseIf InStr(textCU79, "臺灣地址格式不檢查") = 0 Then
                textCU79 = Replace(textCU79 & IIf(Len(Trim(textCU79)) > 0, ";", "") & "臺灣地址格式不檢查" & strSrvDate(2), ";;", ";")
            End If
            'end 2025/03/18
        End If
    End If
End Sub

Public Sub textCU23_Validate(Cancel As Boolean)
   'Add by Amy 2015/09/10
   Dim strZipCode As String, strCountryCode As String, strAddress  As String
   Dim bolMany As Boolean, intArea As Integer
   
   'add by nickc 2008/01/17 若不是修改狀態，將會出不去
   If m_EditMode <> 1 And m_EditMode <> 2 Then Exit Sub
   
   If Trim(textCU23.Text) = "" Then Exit Sub
   If Not CheckLengthIsOK(textCU23, textCU23.MaxLength) Then
      Cancel = True
   End If
   'Add by Amy 2016/12/20 +?判斷-陳金蓮
   If InStr(textCU23, "?") > 0 Then
      MsgBox Left(Label41(20), Len(Label41(20)) - 1) & " 有「?」請確認！", vbExclamation
      Cancel = True
      textCU23.SetFocus
      textCU23_GotFocus
   End If
   'Add by Amy 2015/09/10 +確認臺灣地址格式
   If 案件預設收據公司別啟用日 <= Val(strSrvDate(1)) Then
        'Modify by Amy 2016/12/20 修改時地址無區不自動帶
        If m_EditMode = 1 Then
        textCU23 = ReplaceAddrTW(textCU23, , textCU10) 'Modify by Amy 2025/06/30 +textCU10
        strAddress = IIf(Left(textCU23, 4) = "中華民國", Mid(textCU23, 5), textCU23)
        'Modify by Amy 2020/09/10 傳入地址前7個字抓到郵遞區號
        intArea = 7
        strZipCode = GetPostZip(Left(strAddress, 7), 7, , strCountryCode, bolMany)
        '傳入地址前6個字抓到郵遞區號
        If strZipCode = MsgText(601) Then strZipCode = GetPostZip(Left(strAddress, 6), 6, , strCountryCode, bolMany): intArea = 6
        'end 2020/09/10
        '傳入地址前5個字取郵遞區號
        If strZipCode = MsgText(601) Then strZipCode = GetPostZip(Left(strAddress, 5), 5, , strCountryCode, bolMany): intArea = 5
        '抓到郵遞區號
        If strZipCode <> MsgText(601) Then
            If bolMany = True Then
                '若多筆(同區/鄉 ZipCode不同)且與畫面上欄位資料前3碼不同或空值,彈郵遞區號查詢畫面
                If InStr(strZipCode, Left(Trim(textCU112), 3)) = 0 Or Trim(textCU112) = MsgText(601) Then
                    If Trim(textCU112) <> MsgText(601) Then MsgBox "中文地址郵遞區號有誤,請選擇正確郵遞區號！"
                    Call frm100134.SetParent(Me)
                    Me.Hide
                    frm100134.BFormZip = "textCU112"
                    frm100134.BFormStatus = m_EditMode
                    frm100134.GetStreet textCU23, 1, intArea, strZipCode
                    frm100134.Show
                    Exit Sub
                End If
                'Modify by Amy 2024/06/17 +地址國籍為台灣或空白,與ZipCode 國籍不同才更正
                If (Trim(textCU10) < "010" Or Trim(textCU10) = MsgText(601)) And Trim(textCU10) <> strCountryCode Then
                    If textCU10 <> MsgText(601) Then MsgBox "客戶國籍有誤,系統將自動更正！", , MsgText(5)
                    textCU10 = strCountryCode
                    textCU10_Validate False
                End If
            Else
                '非多筆判斷抓到的郵遞區號是否與畫面上欄位資料前3碼相同
                If Left(textCU112, 3) <> strZipCode Then
                    If textCU112 <> MsgText(601) Then MsgBox "中文地址郵遞區號有誤,系統將自動更正！", , MsgText(5)
                    textCU112 = strZipCode
                End If
                'Modify by Amy 2024/06/17 +地址國籍為台灣或空白,與ZipCode 國籍不同才更正
                If (Trim(textCU10) < "010" Or Trim(textCU10) = MsgText(601)) And Trim(textCU10) <> strCountryCode Then
                    If textCU10 <> MsgText(601) Then MsgBox "客戶國籍有誤,系統將自動更正！", , MsgText(5)
                    textCU10 = strCountryCode
                    textCU10_Validate False
                End If
            End If
        End If
        '修改
        'Modify by Amy 2019/10/03
        ElseIf InStr(textCU79, "臺灣地址格式不檢查") = 0 And m_FieldList(22).fiOldData <> textCU23 Then
            textCU23_LostFocus
        End If
   End If
End Sub

Private Sub textCU24_GotFocus()
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textCU24.IMEMode = 2
   CloseIme
   TextInverse textCU24
End Sub

Private Sub textCU25_GotFocus()
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textCU25.IMEMode = 2
   CloseIme
   TextInverse textCU25
End Sub

Private Sub textCU26_GotFocus()
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textCU26.IMEMode = 2
   CloseIme
   TextInverse textCU26
End Sub

Private Sub textCU27_GotFocus()
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textCU27.IMEMode = 2
   CloseIme
   TextInverse textCU27
End Sub

Private Sub textCU28_GotFocus()
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textCU28.IMEMode = 2
   CloseIme
   TextInverse textCU28
End Sub

Private Sub textCU29_GotFocus()
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textCU29.IMEMode = 1
   OpenIme
   TextInverse textCU29
End Sub

Private Sub textCU29_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = ChangeZIP(KeyAscii, textCU29) 'Modify By Sindy 2021/12/13 +, textCU29
End Sub

Private Sub textCU29_Validate(Cancel As Boolean)
   'add by nickc 2008/01/17 若不是修改狀態，將會出不去
   If m_EditMode <> 1 And m_EditMode <> 2 Then Exit Sub
   
   If textCU29.Text = "" Then Exit Sub
    '中日文欄位尾碼加空白，最大可輸長度減一(因可能會有造字無法存入問題)
   If Not CheckLengthIsOK(textCU29, textCU29.MaxLength - 1) Then
      Cancel = True
   End If
End Sub

Private Sub textCU30_GotFocus()
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textCU30.IMEMode = 2
   CloseIme
   TextInverse textCU30
End Sub

Private Sub textCU30_KeyPress(KeyAscii As Integer)
   KeyAscii = ChangeZIP(KeyAscii, textCU30)
End Sub

Private Sub textCU30_Validate(Cancel As Boolean)
   
   'add by nickc 2008/01/17 若不是修改狀態，將會出不去
   If m_EditMode <> 1 And m_EditMode <> 2 Then Exit Sub
   
   If textCU30.Text = "" Then Exit Sub
   If Not CheckLengthIsOK(textCU30, textCU30.MaxLength) Then
      Cancel = True
   End If
End Sub

Private Sub textCU31_GotFocus()
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textCU31.IMEMode = 2
   OpenIme
   TextInverse textCU31
End Sub

Private Sub textCU31_KeyPress(KeyAscii As MSForms.ReturnInteger)
    Dim strAddr As String, strNewArea As String, strZipCode As String, strCountry As String, strROC As String, strIndArea As String
    Dim intArea As Integer, intFocus As Integer
    Dim bolMany As Boolean
    
    KeyAscii = ChangeZIP(KeyAscii, textCU31) 'Modify By Sindy 2021/12/13 +, textCU31
    '臺灣地址判斷 Add by Amy 2016/05/26
    If LTrim(textCU31) <> MsgText(601) And textCU87 < "010" And InStr(LTrim(textCU31), "後補") = 0 Then
        strROC = ""
        strAddr = textCU31
        If Left(strAddr, 4) = "中華民國" Then strROC = strROC & Left(strAddr, 4): strAddr = Mid(strAddr, 5)
        If Left(strAddr, 3) = "臺灣省" Or Left(strAddr, 3) = "台灣省" Then strROC = strROC & Left(strAddr, 3): strAddr = Mid(strAddr, 4)
        If Left(strAddr, 2) = "臺灣" Or Left(strAddr, 2) = "台灣" Then strROC = strROC & Left(strAddr, 2): strAddr = Mid(strAddr, 3)
        '去除xx工業區查(台中工業區/台塑工業園區不取代,可能抓錯zip)
        strIndArea = "True"
        strAddr = ReplaceIndArea(strAddr, strIndArea)
        If strIndArea = "True" Then strIndArea = MsgText(601)
        If Left(strAddr, 4) = "新竹新竹" And (strIndArea = "科學工業園區" Or strIndArea = "科學園區") Then
            strIndArea = "新竹" & strIndArea
            strAddr = Mid(strAddr, 3)
        End If
        If Len(LTrim(strAddr)) > 4 And (Mid(strAddr, 3, 1) = "市" Or Mid(strAddr, 3, 1) = "縣") Then
            '輸到路/街/段
            If Asc("路") = KeyAscii Or Asc("街") = KeyAscii Or Asc("段") = KeyAscii Then
                intFocus = Val(textCU31.SelStart) - Len(strROC) - Len(strIndArea)
                strAddr = Mid(strAddr, 1, intFocus) & Chr(KeyAscii) & Mid(strAddr, intFocus + 1) 'KeyPress未完成時地址欄位尚未顯示目前字,故先加入當下的字查
                '有鄉/鎮/市/區
                'Modify by Amy 2018/12/19 +判斷第七個字 ex:嘉義縣阿里山鄉 X80024
                If Mid(strAddr, 7, 1) = "市" Or Mid(strAddr, 7, 1) = "區" Or Mid(strAddr, 7, 1) = "鄉" Or Mid(strAddr, 7, 1) = "鎮" _
                  Or Mid(strAddr, 6, 1) = "市" Or Mid(strAddr, 6, 1) = "區" Or Mid(strAddr, 6, 1) = "鄉" Or Mid(strAddr, 6, 1) = "鎮" _
                  Or Mid(strAddr, 5, 1) = "市" Or Mid(strAddr, 5, 1) = "區" Or Mid(strAddr, 5, 1) = "鄉" Or Mid(strAddr, 5, 1) = "鎮" Then
                    strZipCode = GetZipCode_Tai(1, strAddr, , bolMany, , strCountry)
                    If strZipCode <> MsgText(601) Then
                        If bolMany = False Then
                            Call ChkZipData(2, textCU31, strZipCode, , strCountry)
                            textCU31.SelStart = intFocus + Len(strROC) + Len(strIndArea)
                            textCU31.SelLength = 0
                            OpenIme
                        Else
                            '多筆以縣/市+鄉/鎮/市/區及路名查
                            bolMany = False
                            strZipCode = GetZipCode_Tai(3, Mid(strAddr, 1, intFocus + 1), intArea, bolMany, , strCountry)
                            If strZipCode <> MsgText(601) And bolMany = False Then
                                Call ChkZipData(2, textCU31, strZipCode, , strCountry)
                                textCU31.SelStart = intFocus + Len(strROC) + Len(strIndArea)
                                textCU31.SelLength = 0
                                OpenIme
                            End If
                        End If
                    End If
                '沒鄉/鎮/市/區
                Else
                    '取 段/路/街 查
                    strZipCode = GetZipCode_Tai(2, Mid(strAddr, 1, intFocus + 1), intArea, bolMany, strNewArea, strCountry)
                    If strZipCode <> MsgText(601) And bolMany = False Then
                        '補上查到的區,避免輸入兩個同樣的字(路/街/段)被取代,故不用Replace
                        textCU31 = strROC & Left(strAddr, 3) & strNewArea & strIndArea & Mid(strAddr, 4, intArea - 4) & Mid(strAddr, intFocus + 2)
                        Call ChkZipData(2, textCU31, strZipCode, , strCountry)
                        textCU31.SelStart = intFocus + Len(strROC) + Len(strNewArea) + Len(strIndArea)
                        textCU31.SelLength = 0
                        OpenIme
                    End If
                End If
            End If
        End If
        
    End If
    'end 2016/05/26
End Sub

'Add by Amy 2016/05/26 聯絡地址
Private Sub textCU31_LostFocus()
    Dim strZipCode As String, strAddr As String, strCountry As String, strCityN As String, strIndArea As String, strNewArea As String, strROC As String
    Dim bolMany As Boolean, intArea As Integer

    If m_EditMode <> 1 And m_EditMode <> 2 Then Exit Sub
    'Modify by Amy 2016/12/20 拿掉textCU87 >= "010"(地址國籍) 先判斷若為臺灣地址則將地址國籍更正 ex:X6189501 原地址國籍為136
    'Modify by Amy 2025/06/30 加回textCU87 >= "010",X90619 地址國籍大陸,聯絡地址為中國浙江省台州市溫嶺市大溪區高速公路道口一級公路北側-->無法改成大溪鎮
    If textCU31 = MsgText(601) Or InStr(LTrim(textCU31), "後補") > 0 Or textCU87 >= "010" Then Exit Sub
    If m_FieldList(30).fiOldData = textCU31 And m_FieldList(29).fiOldData = textCU30 Then Exit Sub  'Modify by Amy 2020/06/12 bug-抓錯欄位 原:textCU30
    
    textCU31 = ReplaceAddrTW(textCU31, , textCU87) 'Add by Amy 2025/06/30
AgainCheck1:
    strROC = ""
    strAddr = textCU31
    If Left(strAddr, 4) = "中華民國" Then strROC = strROC & Left(strAddr, 4): strAddr = Mid(strAddr, 5)
    If Left(strAddr, 3) = "臺灣省" Or Left(strAddr, 3) = "台灣省" Then strROC = strROC & Left(strAddr, 3): strAddr = Mid(strAddr, 4)
    If Left(strAddr, 2) = "臺灣" Or Left(strAddr, 2) = "台灣" Then strROC = strROC & Left(strAddr, 2): strAddr = Mid(strAddr, 3)
    '去除xx工業區查(台中工業區/台塑工業園區不取代,可能抓錯zip)
    strIndArea = "True"
    strAddr = ReplaceIndArea(strAddr, strIndArea)
    If strIndArea = "True" Then strIndArea = MsgText(601)
    If Left(strAddr, 4) = "新竹新竹" And (strIndArea = "科學工業園區" Or strIndArea = "科學園區") Then
        strIndArea = "新竹" & strIndArea
        strAddr = Mid(strAddr, 3)
    End If
    'Add by Amy 2016/12/20 國籍非臺灣檢查臺灣地址格式是否正確,若是則判斷 ZipCode 是否正確 ex:X6189501地址為台灣 原地址國籍為136-預帶
    If textCU87 >= "010" Then
        '** 第3個字是 縣 / 市
        If Not (Mid(strAddr, 3, 1) = "市" Or Mid(strAddr, 3, 1) = "縣" Or Mid(strAddr, 1, 3) = "釣魚臺" Or Mid(strAddr, 1, 3) = "海南島") Then
            Exit Sub
        '有鄉鎮市區
        'Modify by Amy 2018/12/19 +判斷第七個字 ex:嘉義縣阿里山鄉 X80024
        ElseIf Not (Mid(strAddr, 7, 1) = "市" Or Mid(strAddr, 7, 1) = "區" Or Mid(strAddr, 7, 1) = "鄉" Or Mid(strAddr, 7, 1) = "鎮" _
              Or Mid(strAddr, 6, 1) = "市" Or Mid(strAddr, 6, 1) = "區" Or Mid(strAddr, 6, 1) = "鄉" Or Mid(strAddr, 6, 1) = "鎮" _
              Or Mid(strAddr, 5, 1) = "市" Or Mid(strAddr, 5, 1) = "區" Or Mid(strAddr, 5, 1) = "鄉" Or Mid(strAddr, 5, 1) = "鎮") Then
              Exit Sub
        End If
    End If
    'end 2016/12/20
  
    '** 第3個字是 縣 / 市
    If Mid(strAddr, 3, 1) = "市" Or Mid(strAddr, 3, 1) = "縣" Or Mid(strAddr, 1, 3) = "釣魚臺" Or Mid(strAddr, 1, 3) = "海南島" Then
                
        'Modify by Amy 2018/12/19 +判斷第七個字 ex:嘉義縣阿里山鄉 X80024
        If Mid(strAddr, 7, 1) = "市" Or Mid(strAddr, 7, 1) = "區" Or Mid(strAddr, 7, 1) = "鄉" Or Mid(strAddr, 7, 1) = "鎮" _
          Or Mid(strAddr, 6, 1) = "市" Or Mid(strAddr, 6, 1) = "區" Or Mid(strAddr, 6, 1) = "鄉" Or Mid(strAddr, 6, 1) = "鎮" _
          Or Mid(strAddr, 5, 1) = "市" Or Mid(strAddr, 5, 1) = "區" Or Mid(strAddr, 5, 1) = "鄉" Or Mid(strAddr, 5, 1) = "鎮" Then
            '傳入地址前6個字抓到郵遞區號
            intArea = 6
            strZipCode = GetPostZip(Left(strAddr, 6), 6, , strCountry, bolMany)
            '傳入地址前5個字取郵遞區號
            If strZipCode = MsgText(601) Then strZipCode = GetPostZip(Left(strAddr, 5), 5, , strCountry, bolMany): intArea = 5
            '抓到郵遞區號
            If strZipCode <> MsgText(601) Then
                If bolMany = True Then
                    '多筆以縣/市+鄉/鎮/市/區及路名查
                    bolMany = False
                    strZipCode = GetZipCode_Tai(3, strAddr, , bolMany, , strCountry)
                    If strZipCode <> MsgText(601) Then
                        '限制縣/市+鄉/鎮/市/區及路名查:一筆-直接帶/多筆-進查詢畫面
                        If bolMany = False Then
                            Call ChkZipData(2, textCU31, strZipCode, , strCountry)
                        Else
                            Call ChkZipData(1, textCU31, strZipCode, intArea, strCountry)
                        End If
                    End If
                Else
                    '非多筆
                    Call ChkZipData(2, textCU31, strZipCode, intArea, strCountry)
                End If
            
            Else
                '判斷是否有此區/鄉/鎮
                strZipCode = GetPostZip(Mid(strAddr, 4, intArea - 3), intArea - 3, , strCountry, bolMany, "Pzd03")
                If strZipCode <> MsgText(601) Then
                    '區別錯,進入查詢畫面
                    Call ChkZipData(3, textCU31, strZipCode, intArea, strCountry)
                Else
                    '當作沒區只有路 ex:新竹縣or市園區二路
                    bolMany = False
                    strZipCode = GetZipCode_Tai(2, strAddr, intArea, bolMany, strNewArea, strCountry)
                    If strZipCode <> MsgText(601) Then
                        '以縣/市及路名查:一筆-直接帶/多筆-進查詢畫面
                        If bolMany = False Then
                            textCU31 = strROC & Left(strAddr, 3) & strNewArea & strIndArea & _
                            Replace(Replace(Replace(Replace(textCU31, strIndArea, ""), strNewArea, ""), Left(strAddr, 3), ""), strROC, "")
                            Call ChkZipData(2, textCU31, strZipCode, intArea, strCountry)
                        Else
                            intArea = 0
                            Call ChkZipData(4, textCU31, strZipCode, intArea, strCountry)
                            Exit Sub
                        End If
                    End If
                End If
            End If
                    
        '無鄉/鎮/市/區
        Else
            '以路/街 抓是否有zip
            strZipCode = GetZipCode_Tai(2, strAddr, intArea, bolMany, strNewArea, strCountry)
            If strZipCode <> MsgText(601) Then
                If bolMany = True Then
                    '多筆
                    intArea = 0
                    Call ChkZipData(4, textCU31, strZipCode, intArea, strCountry)
                Else
                    '非多筆
                    textCU31 = strROC & Left(strAddr, 3) & strNewArea & strIndArea & _
                            Replace(Replace(Replace(Replace(textCU31, strIndArea, ""), strNewArea, ""), Left(strAddr, 3), ""), strROC, "")
                    Call ChkZipData(2, textCU31, strZipCode, intArea, strCountry)
                End If
            '都抓不到ZipCode
            ElseIf strZipCode = MsgText(601) Then
                If CheckTaiwanAddr_Tai(textCU31, textCU30, "000", "", strZipCode, , False, Me.Name) = False Then
                    If strZipCode = "格式錯誤" Then
                        'Modify by Amy 2020/04/10 +if  因X5044001 地址:臺中大里工業區仁化路 彈二次強制表單 會錯
                        If PUB_CheckFormExist("frm100135") = False Then
                            frm100135.Show vbModal
                            Call ChkZipData(9, textCU31, strZipCode)
                        End If
                    Else
                        Call ChkZipData(3, textCU31, strZipCode)
                    End If
                    Exit Sub
                End If
            End If
        End If
    
    '** 第三3個字無 縣 / 市
    Else
        '傳入地址前2個字判斷是否有其縣/市
        strCityN = "Pzd02"
        strZipCode = GetPostZip(Left(strAddr, 2), 2, 1, strCountry, bolMany, "Pzd02", strCityN)
        If strZipCode <> MsgText(601) Then
            If bolMany = False Then
                '只有一筆
                textCU31 = strROC & strCityN & strIndArea & _
                        Replace(Replace(Replace(textCU31, strIndArea, ""), strCityN, ""), strROC, "")
                GoTo AgainCheck1
            Else
                '新竹、嘉義會有2筆
                intArea = 0
                Call ChkZipData(5, textCU31, strZipCode, intArea)
            End If
        End If
    End If
End Sub

Private Sub textCU31_Validate(Cancel As Boolean)
Dim strTmp As String   '2008/9/4 add by sonia

   'add by nickc 2008/01/17 若不是修改狀態，將會出不去
   If m_EditMode <> 1 And m_EditMode <> 2 Then Exit Sub
   
   If Trim(textCU31.Text) = "" Then Exit Sub
   If Not CheckLengthIsOK(textCU31, textCU31.MaxLength) Then
      Cancel = True
   End If
    'Add by Amy 2016/12/20 +?判斷-陳金蓮
   If InStr(textCU31, "?") > 0 Then
      MsgBox Left(Label41(18), Len(Label41(18)) - 1) & " 有「?」請確認！", vbExclamation
      Cancel = True
      textCU31.SetFocus
      textCU31_GotFocus
   End If
   'Mark by Amy 2015/09/09 秀玲:取消此判斷
'   '2008/9/4 add by sonia
'   If m_EditMode = 2 Then    '2009/7/30 ADD BY SONIA
'      If textCU31.Text <> m_CU31 And m_CU31 <> "" Then
'         strTmp = "修改聯絡地址，請確認地址國籍是否正確 ? 地址國籍為" & textCU87
'         If MsgBox(strTmp, vbYesNo + vbCritical) = vbNo Then
'            m_CU31 = textCU31
'            Cancel = True
'            textCU87.SetFocus
'         Else
'            m_CU31 = textCU31
'         End If
'      End If
'   End If
'   '2008/9/4 end
End Sub

Private Sub textCU32_GotFocus()
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textCU32.IMEMode = 2
   CloseIme
   TextInverse textCU32
End Sub

Private Sub textCU32_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 8 And KeyAscii <> Asc("N") Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub textCU32_Validate(Cancel As Boolean)
   'add by nickc 2008/01/17 若不是修改狀態，將會出不去
   If m_EditMode <> 1 And m_EditMode <> 2 Then Exit Sub
   
   If textCU32.Text = "" Then Exit Sub
   If textCU32.Text <> "N" Then
      ShowMsg "輸入錯誤 !"
      Cancel = True
   End If
End Sub

Private Sub textCU33_GotFocus()
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textCU33.IMEMode = 2
   CloseIme
   TextInverse textCU33
End Sub

'Add By Sindy 2011/3/4
Private Sub textCU146_GotFocus()
   CloseIme
   TextInverse textCU146
End Sub

Private Sub textCU34_GotFocus()
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textCU34.IMEMode = 1
   OpenIme
   TextInverse textCU34
End Sub

Private Sub textCU34_Validate(Cancel As Boolean)
   'add by nickc 2008/01/17 若不是修改狀態，將會出不去
   If m_EditMode <> 1 And m_EditMode <> 2 Then Exit Sub
   
   If textCU34.Text = "" Then Exit Sub
   If Not CheckLengthIsOK(textCU34, textCU34.MaxLength) Then
      Cancel = True
   End If
End Sub

Private Sub textCU35_GotFocus()
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textCU35.IMEMode = 2
   CloseIme
   TextInverse textCU35
End Sub

Private Sub textCU36_GotFocus()
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textCU36.IMEMode = 2
   CloseIme
   TextInverse textCU36
End Sub

Private Sub textCU36_Validate(Cancel As Boolean)
   'add by nickc 2008/01/17 若不是修改狀態，將會出不去
   If m_EditMode <> 1 And m_EditMode <> 2 Then Exit Sub
   
   If textCU36.Text = "" Then Exit Sub
   If Val(textCU36.Text) > 101 Then
      ShowMsg "折扣不可大於 100 !"
      Cancel = True
   End If
End Sub

Private Sub textCU37_GotFocus()
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textCU37.IMEMode = 2
   CloseIme
   TextInverse textCU37
End Sub

Private Sub textCU37_Validate(Cancel As Boolean)
   'add by nickc 2008/01/17 若不是修改狀態，將會出不去
   If m_EditMode <> 1 And m_EditMode <> 2 Then Exit Sub
   
   If textCU37.Text = "" Then Exit Sub
   If Val(textCU37.Text) > 101 Then
       ShowMsg "折扣不可大於 100 !"
       Cancel = True
   End If
End Sub

Private Sub textCU38_GotFocus()
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textCU38.IMEMode = 2
   CloseIme
   TextInverse textCU38
End Sub

Private Sub textCU38_Validate(Cancel As Boolean)
   'add by nickc 2008/01/17 若不是修改狀態，將會出不去
   If m_EditMode <> 1 And m_EditMode <> 2 Then Exit Sub
   
   If textCU38.Text = "" Then Exit Sub
   If CheckIsTaiwanDate(textCU38.Text) = False Then
      Cancel = True
   End If
End Sub

Private Sub textCU39_GotFocus()
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textCU39.IMEMode = 1
   OpenIme
   TextInverse textCU39
End Sub

Private Sub textCU39_Validate(Cancel As Boolean)
   'add by nickc 2008/01/17 若不是修改狀態，將會出不去
   If m_EditMode <> 1 And m_EditMode <> 2 Then Exit Sub
   
   If textCU39.Text = "" Then Exit Sub
   If Not CheckLengthIsOK(textCU39, textCU39.MaxLength) Then
      Cancel = True
   End If
End Sub

Private Sub textCU40_GotFocus()
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textCU40.IMEMode = 2
   CloseIme
   TextInverse textCU40
End Sub

Private Sub textCU41_GotFocus()
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textCU41.IMEMode = 1
   OpenIme
   TextInverse textCU41
End Sub

Private Sub textCU41_Validate(Cancel As Boolean)
   'add by nickc 2008/01/17 若不是修改狀態，將會出不去
   If m_EditMode <> 1 And m_EditMode <> 2 Then Exit Sub
   
   If textCU41.Text = "" Then Exit Sub
   If Not CheckLengthIsOK(textCU41, textCU41.MaxLength) Then
      Cancel = True
   End If
End Sub

Private Sub textCU42_GotFocus()
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textCU42.IMEMode = 1
   OpenIme
   TextInverse textCU42
End Sub

Private Sub textCU42_Validate(Cancel As Boolean)
   'add by nickc 2008/01/17 若不是修改狀態，將會出不去
   If m_EditMode <> 1 And m_EditMode <> 2 Then Exit Sub
   
   If textCU42.Text = "" Then Exit Sub
   If Not CheckLengthIsOK(textCU42, textCU42.MaxLength) Then
      Cancel = True
   End If
End Sub

Private Sub textCU43_GotFocus()
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textCU43.IMEMode = 2
   CloseIme
   TextInverse textCU43
End Sub

Private Sub textCU44_GotFocus()
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textCU44.IMEMode = 1
   OpenIme
   TextInverse textCU44
End Sub

Private Sub textCU44_Validate(Cancel As Boolean)
   'add by nickc 2008/01/17 若不是修改狀態，將會出不去
   If m_EditMode <> 1 And m_EditMode <> 2 Then Exit Sub
   
   If textCU44.Text = "" Then Exit Sub
   If Not CheckLengthIsOK(textCU44, textCU44.MaxLength) Then
      Cancel = True
   End If
End Sub

Private Sub textCU45_GotFocus()
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textCU45.IMEMode = 1
   OpenIme
   TextInverse textCU45
End Sub

Private Sub textCU45_Validate(Cancel As Boolean)
   'add by nickc 2008/01/17 若不是修改狀態，將會出不去
   If m_EditMode <> 1 And m_EditMode <> 2 Then Exit Sub
   
   If textCU45.Text = "" Then Exit Sub
   If Not CheckLengthIsOK(textCU45, textCU45.MaxLength) Then
      Cancel = True
   End If
End Sub

Private Sub textCU46_GotFocus()
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textCU46.IMEMode = 2
   CloseIme
   TextInverse textCU46
End Sub

Private Sub textCU47_GotFocus()
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textCU47.IMEMode = 1
   OpenIme
   TextInverse textCU47
End Sub

Private Sub textCU47_Validate(Cancel As Boolean)
   'add by nickc 2008/01/17 若不是修改狀態，將會出不去
   If m_EditMode <> 1 And m_EditMode <> 2 Then Exit Sub
   
   If textCU47.Text = "" Then Exit Sub
   If Not CheckLengthIsOK(textCU47, textCU47.MaxLength) Then
      Cancel = True
   End If
End Sub

Private Sub textCU48_GotFocus()
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textCU48.IMEMode = 1
   OpenIme
   TextInverse textCU48
End Sub

Private Sub textCU48_Validate(Cancel As Boolean)
      'add by nickc 2008/01/17 若不是修改狀態，將會出不去
      If m_EditMode <> 1 And m_EditMode <> 2 Then Exit Sub
      
      If textCU48.Text = "" Then Exit Sub
      If Not CheckLengthIsOK(textCU48, textCU48.MaxLength) Then
         Cancel = True
      End If
End Sub

Private Sub textCU49_GotFocus()
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textCU49.IMEMode = 2
   CloseIme
   TextInverse textCU49
End Sub

Private Sub textCU50_GotFocus()
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textCU50.IMEMode = 1
   OpenIme
   TextInverse textCU50
End Sub

Private Sub textCU50_Validate(Cancel As Boolean)
   'add by nickc 2008/01/17 若不是修改狀態，將會出不去
   If m_EditMode <> 1 And m_EditMode <> 2 Then Exit Sub
   
   If textCU50.Text = "" Then Exit Sub
   If Not CheckLengthIsOK(textCU50, textCU50.MaxLength) Then
      Cancel = True
   End If
End Sub

Private Sub textCU51_GotFocus()
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textCU51.IMEMode = 1
   OpenIme
   TextInverse textCU51
End Sub

Private Sub textCU51_Validate(Cancel As Boolean)
   'add by nickc 2008/01/17 若不是修改狀態，將會出不去
   If m_EditMode <> 1 And m_EditMode <> 2 Then Exit Sub
   
   If textCU51.Text = "" Then Exit Sub
   If Not CheckLengthIsOK(textCU51, textCU51.MaxLength) Then
      Cancel = True
   End If
End Sub

Private Sub textCU52_GotFocus()
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textCU52.IMEMode = 2
   CloseIme
   TextInverse textCU52
End Sub

Private Sub textCU53_GotFocus()
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textCU53.IMEMode = 1
   OpenIme
   TextInverse textCU53
End Sub

Private Sub textCU53_Validate(Cancel As Boolean)
   'add by nickc 2008/01/17 若不是修改狀態，將會出不去
   If m_EditMode <> 1 And m_EditMode <> 2 Then Exit Sub
   
   If textCU53.Text = "" Then Exit Sub
   If Not CheckLengthIsOK(textCU53, textCU53.MaxLength) Then
      Cancel = True
   End If
End Sub

Private Sub textCU54_GotFocus()
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textCU54.IMEMode = 1
   OpenIme
   TextInverse textCU54
End Sub

Private Sub textCU54_Validate(Cancel As Boolean)
   'add by nickc 2008/01/17 若不是修改狀態，將會出不去
   If m_EditMode <> 1 And m_EditMode <> 2 Then Exit Sub
   
   If textCU54.Text = "" Then Exit Sub
   If Not CheckLengthIsOK(textCU54, textCU54.MaxLength) Then
      Cancel = True
   End If
End Sub

Private Sub textCU55_GotFocus()
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textCU55.IMEMode = 2
   CloseIme
   TextInverse textCU55
End Sub

Private Sub textCU56_GotFocus()
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textCU56.IMEMode = 1
   OpenIme
   TextInverse textCU56
End Sub

Private Sub textCU56_Validate(Cancel As Boolean)
   'add by nickc 2008/01/17 若不是修改狀態，將會出不去
   If m_EditMode <> 1 And m_EditMode <> 2 Then Exit Sub
   
   If textCU56.Text = "" Then Exit Sub
   If Not CheckLengthIsOK(textCU56, textCU56.MaxLength) Then
      Cancel = True
   End If
End Sub

Private Sub textCU57_GotFocus()
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textCU55.IMEMode = 2
   CloseIme
   TextInverse textCU57
End Sub

Private Sub textCU57_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textCU57_Validate(Cancel As Boolean)
   'If m_EditMode = 4 Then Exit Sub
   Label30(6).Caption = ""
   If textCU57.Text = "" Then Exit Sub
   If textCU57 <> "" Then textCU57 = textCU57 & String(9 - Len(textCU57), "0")
   Label30(6).Caption = ChgType(4, textCU57.Text)
   'add by nickc 2008/01/23 若不是修改狀態，將會出不去
   If m_EditMode <> 1 And m_EditMode <> 2 Then Exit Sub
               
   If Label30(6).Caption = "" Then Cancel = True
End Sub

'Add By Sindy 2011/3/4
   Private Sub textCU147_GotFocus()
   CloseIme
   TextInverse textCU147
End Sub

'Add By Sindy 2011/3/4
Private Sub textCU147_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

'Add By Sindy 2011/3/4
Private Sub textCU147_Validate(Cancel As Boolean)
   Label30(16).Caption = ""
   If textCU147.Text = "" Then Exit Sub
   If textCU147 <> "" Then textCU147 = textCU147 & String(9 - Len(textCU147), "0")
   Label30(16).Caption = ChgType(4, textCU147.Text)
   If m_EditMode <> 1 And m_EditMode <> 2 Then Exit Sub
   
   If Label30(16).Caption = "" Then Cancel = True
End Sub

Private Sub textCU58_GotFocus()
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textCU58.IMEMode = 1
   OpenIme
   TextInverse textCU58
End Sub

Private Sub textCU58_Validate(Cancel As Boolean)
   'add by nickc 2008/01/17 若不是修改狀態，將會出不去
   If m_EditMode <> 1 And m_EditMode <> 2 Then Exit Sub
   
   If textCU58.Text = "" Then Exit Sub
   'Modified by Lydia 2017/06/14 聯絡人(中)改為30字
   'If Not CheckLengthIsOK(textCU58, textCU58.MaxLength) Then
   If Not CheckLengthIsOK(textCU58, 30) Then
      Cancel = True
   End If
End Sub

Private Sub textCU59_GotFocus()
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textCU59.IMEMode = 2
   CloseIme
   TextInverse textCU59
End Sub

Private Sub textCU60_GotFocus()
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textCU60.IMEMode = 1
   OpenIme
   TextInverse textCU60
End Sub

Private Sub textCU60_Validate(Cancel As Boolean)
   'add by nickc 2008/01/17 若不是修改狀態，將會出不去
   If m_EditMode <> 1 And m_EditMode <> 2 Then Exit Sub
   
   If textCU60.Text = "" Then Exit Sub
    '中日文欄位尾碼加空白，最大可輸長度減一(因可能會有造字無法存入問題)
    'Modified by Lydia 2017/06/14
   'If Not CheckLengthIsOK(textCU60, textCU60.MaxLength - 1) Then
   If Not CheckLengthIsOK(textCU60, 60) Then
      Cancel = True
   End If
End Sub
Private Sub textCU61_GotFocus()
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textCU61.IMEMode = 1
   OpenIme
   TextInverse textCU61
End Sub

Private Sub textCU61_Validate(Cancel As Boolean)
   'add by nickc 2008/01/17 若不是修改狀態，將會出不去
   If m_EditMode <> 1 And m_EditMode <> 2 Then Exit Sub
   
   If textCU61.Text = "" Then Exit Sub
   'Modified by Lydia 2017/06/14 聯絡人(中)改為30字
   'If Not CheckLengthIsOK(textCU61, textCU61.MaxLength) Then
   If Not CheckLengthIsOK(textCU61, 30) Then
      Cancel = True
   End If
End Sub

Private Sub textCU62_GotFocus()
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textCU62.IMEMode = 2
   CloseIme
   TextInverse textCU62
End Sub

Private Sub textCU63_GotFocus()
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textCU63.IMEMode = 1
   OpenIme
   TextInverse textCU63
End Sub

Private Sub textCU63_Validate(Cancel As Boolean)
   'add by nickc 2008/01/17 若不是修改狀態，將會出不去
   If m_EditMode <> 1 And m_EditMode <> 2 Then Exit Sub

   If textCU63.Text = "" Then Exit Sub
    '中日文欄位尾碼加空白，最大可輸長度減一(因可能會有造字無法存入問題)
   'Modified by Lydia 2022/03/16
   'If Not CheckLengthIsOK(textCU63, textCU63.MaxLength - 1) Then
   If Not CheckLengthIsOK(textCU63, 60) Then
      Cancel = True
   End If
End Sub

Private Sub textCU64_GotFocus()
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textCU64.IMEMode = 2
   CloseIme
   TextInverse textCU64
End Sub

Private Sub textCU64_KeyPress(KeyAscii As Integer)
   If (KeyAscii > 51 Or KeyAscii < 49) And KeyAscii <> 8 Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub textCU65_GotFocus()
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textCU65.IMEMode = 2
   CloseIme
   TextInverse textCU65
End Sub

Private Sub textCU66_GotFocus()
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textCU66.IMEMode = 2
   CloseIme
   TextInverse textCU66
End Sub

Private Sub textCU67_GotFocus()
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textCU67.IMEMode = 2
   CloseIme
   TextInverse textCU67
End Sub

Private Sub textCU68_GotFocus()
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textCU68.IMEMode = 2
   CloseIme
   TextInverse textCU68
End Sub

Private Sub textCU69_GotFocus()
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textCU69.IMEMode = 2
   CloseIme
   TextInverse textCU69
End Sub

'Add by Amy 2016/12/20 +?判斷-陳金蓮
Private Sub textCU65_Validate(Cancel As Boolean)
    If m_EditMode <> 1 And m_EditMode <> 2 Then Exit Sub
    
    If InStr(textCU65, "?") > 0 Then
        MsgBox Left(Label41(23), Len(Label41(23)) - 1) & " 有「?」請確認！", vbExclamation
        Cancel = True
        textCU65.SetFocus
        textCU65_GotFocus
   End If
End Sub

Private Sub textCU66_Validate(Cancel As Boolean)
    If m_EditMode <> 1 And m_EditMode <> 2 Then Exit Sub
    
    If InStr(textCU66, "?") > 0 Then
        MsgBox Left(Label41(24), Len(Label41(24)) - 1) & " 有「?」請確認！", vbExclamation
        Cancel = True
        textCU66.SetFocus
        textCU66_GotFocus
   End If
End Sub

Private Sub textCU67_Validate(Cancel As Boolean)
    If m_EditMode <> 1 And m_EditMode <> 2 Then Exit Sub
    
    If InStr(textCU67, "?") > 0 Then
        MsgBox Left(Label41(25), Len(Label41(25)) - 1) & " 有「?」請確認！", vbExclamation
        Cancel = True
        textCU67.SetFocus
        textCU67_GotFocus
   End If
End Sub

Private Sub textCU68_Validate(Cancel As Boolean)
    If m_EditMode <> 1 And m_EditMode <> 2 Then Exit Sub
    
    If InStr(textCU68, "?") > 0 Then
        MsgBox Left(Label41(26), Len(Label41(26)) - 1) & " 有「?」請確認！", vbExclamation
        Cancel = True
        textCU68.SetFocus
        textCU68_GotFocus
   End If
End Sub

Private Sub textCU69_Validate(Cancel As Boolean)
    If m_EditMode <> 1 And m_EditMode <> 2 Then Exit Sub
    
    If InStr(textCU69, "?") > 0 Then
        MsgBox Left(Label41(27), Len(Label41(27)) - 1) & " 有「?」請確認！", vbExclamation
        Cancel = True
        textCU69.SetFocus
        textCU69_GotFocus
   End If
End Sub
'end 2016/12/20

Private Sub textCU70_GotFocus()
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textCU70.IMEMode = 2
   CloseIme
   TextInverse textCU70
End Sub

Private Sub textCU71_GotFocus()
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textCU71.IMEMode = 2
   CloseIme
   TextInverse textCU71
End Sub

Private Sub textCU71_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textCU71_Validate(Cancel As Boolean)
   'If m_EditMode = 4 Then Exit Sub
   Label30(5).Caption = ""
   If textCU71.Text = "" Then Exit Sub
   If textCU71 <> "" Then textCU71 = textCU71 & String(9 - Len(textCU71), "0")
   Label30(5).Caption = ChgType(4, textCU71.Text)
   'add by nickc 2008/01/23 若不是修改狀態，將會出不去
   If m_EditMode <> 1 And m_EditMode <> 2 Then Exit Sub
   
   If Label30(5).Caption = "" Then Cancel = True
End Sub

Private Sub textCU72_GotFocus()
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textCU72.IMEMode = 2
   CloseIme
   TextInverse textCU72
End Sub

Private Sub textCU72_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textCU72_Validate(Cancel As Boolean)
   'add by nickc 2008/01/17 若不是修改狀態，將會出不去
   If m_EditMode <> 1 And m_EditMode <> 2 Then Exit Sub
   
   If textCU72.Text = "" Then Exit Sub
   If textCU72.Text <> "Y" Then
      ShowMsg "輸入錯誤 !"
      Cancel = True
   End If
End Sub

Private Sub textCU73_GotFocus()
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textCU73.IMEMode = 2
   CloseIme
   TextInverse textCU73
End Sub

Private Sub textCU73_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textCU73_Validate(Cancel As Boolean)
   'add by nickc 2008/01/17 若不是修改狀態，將會出不去
   If m_EditMode <> 1 And m_EditMode <> 2 Then Exit Sub
   
   If textCU73.Text = "" Then Exit Sub
   If textCU73.Text <> "Y" Then
      ShowMsg "輸入錯誤 !"
      Cancel = True
   End If
End Sub

Private Sub textCU74_GotFocus()
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textCU74.IMEMode = 2
   CloseIme
   TextInverse textCU74
End Sub

Private Sub textCU74_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textCU74_Validate(Cancel As Boolean)
   'add by nickc 2008/01/17 若不是修改狀態，將會出不去
   If m_EditMode <> 1 And m_EditMode <> 2 Then Exit Sub
   
   If textCU74.Text = "" Then Exit Sub
   'Modified by Lydia 2016/08/18 +N:寄證書後年費不續辦
   If textCU74.Text <> "Y" And textCU74.Text <> "N" Then
      ShowMsg "輸入錯誤 !"
      Cancel = True
   'Added by Morgan 2019/12/24
   '若客戶設定年費自動代繳時證書定稿需增加相關敘述
   ElseIf textCU74.Text = "Y" Then
      ShowMsg "客戶目前暫不開放設定年費動代繳 !" & vbCrLf & "若有實務需求請先與承辦組主管確認 !"
      Cancel = True
   
   End If
End Sub

Private Sub textCU75_GotFocus()
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textCU75.IMEMode = 2
   CloseIme
   TextInverse textCU75
End Sub

Private Sub textCU75_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textCU75_Validate(Cancel As Boolean)
   'add by nickc 2008/01/17 若不是修改狀態，將會出不去
   If m_EditMode <> 1 And m_EditMode <> 2 Then Exit Sub
   
   If textCU75.Text = "" Then Exit Sub
   If textCU75.Text <> "Y" Then
      ShowMsg "輸入錯誤 !"
      Cancel = True
   End If
End Sub

'Private Sub textCU76_GotFocus()
'   'edit by nickc 2007/06/06 切換輸入法改用API
'   'textCU76.IMEMode = 2
'   CloseIme
'   TextInverse textCU76
'End Sub
'
'Private Sub textCU76_KeyPress(KeyAscii As Integer)
'   KeyAscii = UpperCase(KeyAscii)
'End Sub
'
'Private Sub textCU76_Validate(Cancel As Boolean)
'   'add by nickc 2008/01/17 若不是修改狀態，將會出不去
'   If m_EditMode <> 1 And m_EditMode <> 2 Then Exit Sub
'
'   If textCU76.Text = "" Then Exit Sub
'   Select Case textCU76.Text
'      Case "U", "N", "R"
'
'      Case Else
'         ShowMsg "輸入錯誤 !"
'         Cancel = True
'   End Select
'End Sub
'
''Add By Sindy 2011/3/4
'Private Sub textCU148_GotFocus()
'   CloseIme
'   TextInverse textCU148
'End Sub
'
''Add By Sindy 2011/3/4
'Private Sub textCU148_KeyPress(KeyAscii As Integer)
'   KeyAscii = UpperCase(KeyAscii)
'End Sub
'
''Add By Sindy 2011/3/4
'Private Sub textCU148_Validate(Cancel As Boolean)
'   If m_EditMode <> 1 And m_EditMode <> 2 Then Exit Sub
'
'   If textCU148.Text = "" Then Exit Sub
'   Select Case textCU148.Text
'      Case "U", "N", "R"
'
'      Case Else
'         ShowMsg "輸入錯誤 !"
'         Cancel = True
'   End Select
'End Sub

Private Sub textCU77_GotFocus()
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textCU77.IMEMode = 2
   CloseIme
   TextInverse textCU77
End Sub

Private Sub textCU77_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textCU77_Validate(Cancel As Boolean)
   'add by nickc 2008/01/17 若不是修改狀態，將會出不去
   If m_EditMode <> 1 And m_EditMode <> 2 Then Exit Sub
   
   If textCU77.Text = "" Then Exit Sub
   If textCU77.Text <> "Y" Then
      ShowMsg "輸入錯誤 !"
      Cancel = True
   End If
End Sub

'Add By Sindy 2011/3/4
Private Sub textCU149_GotFocus()
   CloseIme
   TextInverse textCU149
End Sub

'Add By Sindy 2011/3/4
Private Sub textCU149_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

'Add By Sindy 2011/3/4
Private Sub textCU149_Validate(Cancel As Boolean)
   If m_EditMode <> 1 And m_EditMode <> 2 Then Exit Sub
   
   If textCU149.Text = "" Then Exit Sub
   If textCU149.Text <> "Y" Then
      ShowMsg "輸入錯誤 !"
      Cancel = True
   End If
End Sub

'Add By Sindy 2013/8/15
Private Sub textCU139_GotFocus()
   CloseIme
   TextInverse textCU139
End Sub

'Add By Sindy 2013/8/15
Private Sub textCU139_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

'Mark by Amy 2015/09/10 改為label顯示
''Add By Sindy 2013/11/19
'Private Sub textCU143_GotFocus()
'   CloseIme
'   TextInverse textCU143
'End Sub
'
''Add By Sindy 2013/11/19
'Private Sub textCU143_KeyPress(KeyAscii As Integer)
'   KeyAscii = Pub_NumAscii(KeyAscii)
'End Sub
'end 2015/09/10

Private Sub textCU78_GotFocus()
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textCU78.IMEMode = 2
   CloseIme
   TextInverse textCU78
End Sub

'Add By Sindy 2011/3/4
Private Sub textCU150_GotFocus()
   CloseIme
   TextInverse textCU150
End Sub

Private Sub textCU79_GotFocus()
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textCU79.IMEMode = 1
   OpenIme
   TextInverse textCU79
End Sub

Private Sub textCU79_Validate(Cancel As Boolean)
   'add by nickc 2008/01/17 若不是修改狀態，將會出不去
   If m_EditMode <> 1 And m_EditMode <> 2 Then Exit Sub
   
   If textCU79.Text = "" Then Exit Sub
   If Not CheckLengthIsOK(textCU79, textCU79.MaxLength) Then
      Cancel = True
   End If
End Sub

Private Sub cboStatus_LostFocus()
   'modify by sonia 2022/3/31
   'If cboStatus <> "" Then
   If cboStatus <> "" And cboStatus <> "其他" And cboStatus <> "業務自行處理" And cboStatus <> "解除對造" Then
      textCU32 = "N"
      textCU132 = "N" '2008/12/9 add by sonia
   End If
End Sub

'2008/6/26 add by sonia X30504仍可自行輸入故加入此控制
Private Sub cboStatus_Validate(Cancel As Boolean)
   If m_EditMode <> 1 And m_EditMode <> 2 Then Exit Sub
   Select Case cboStatus
      '2012/10/11 modify by sonia 加 不得代理
      'Modify by Amy 2019/08/27 Cancel 刪址、倒閉,往生改死亡
      'Modify by Amy 2021/11/29 +國內同業
      Case "", "遷移不明", "解散", "廢止", "撤銷", "停業", "死亡", "其他", "業務自行處理", "不再使用", "不得代理", "不得代理專利", "不得代理商標", "宣告破產", "國內同業"
      Case Else
         'Add by Amy 2015/08/24 +if 電腦中心可自行輸入
         'Modify by Amy 2022/06/20 +客戶狀態開放才檢查(客戶狀態,並非操作者權限的下拉選項內容會鎖住)
         If Pub_StrUserSt03 <> "M51" And cboStatus.Locked = False Then
            ShowMsg "客戶狀態錯誤, 請以下拉方式點選 !"
            Cancel = True
         End If
   End Select
End Sub

Private Sub cboStatus_GotFocus()
   OpenIme
End Sub

'2008/6/26 END
Private Sub textCU87_GotFocus()
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textCU87.IMEMode = 2
   CloseIme
   TextInverse textCU87
End Sub

Public Sub textCU87_Validate(Cancel As Boolean)
   
   'If m_EditMode = 4 Then Exit Sub
   Label30(4).Caption = ""
   If textCU87.Text = "" Then Exit Sub
   
   'Add by Amy 2015/08/24 國籍不可輸入000
   If textCU87.Text = 台灣國家代號 Then
      If m_EditMode <> 1 And m_EditMode <> 2 Then Exit Sub
      Cancel = True
      ShowMsg "地址" & MsgText(9153)
      textCU87.SetFocus
      textCU87_GotFocus
      Exit Sub
   End If
   'end 2015/08/24
    Label30(4).Caption = ChgType(2, textCU87.Text)
    
   'add by nickc 2008/01/23 若不是修改狀態，將會出不去
   If m_EditMode <> 1 And m_EditMode <> 2 Then Exit Sub
      
   If Label30(4).Caption = "" Then Cancel = True
  
End Sub

Private Sub textCU88_GotFocus()
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textCU88.IMEMode = 2
   CloseIme
   TextInverse textCU88
End Sub

Private Sub textCU89_GotFocus()
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textCU89.IMEMode = 2
   CloseIme
   TextInverse textCU89
End Sub

Private Sub textCU91_GotFocus()
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textCU91.IMEMode = 1
   OpenIme
   TextInverse textCU91
End Sub

Private Sub textCU91_Validate(Cancel As Boolean)
   'add by nickc 2008/01/17 若不是修改狀態，將會出不去
   If m_EditMode <> 1 And m_EditMode <> 2 Then Exit Sub
   
   If textCU91.Text = "" Then Exit Sub
   'Modified by Lydia 2021/01/07 中、英、日文名稱改成判斷字串個數
   'If Not CheckLengthIsOK(textCU91, textCU91.MaxLength) Then
   If Len(textCU91) > textCU91.MaxLength Then
      Cancel = True
   End If
End Sub

Private Sub textCU92_GotFocus()
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textCU92.IMEMode = 2
   CloseIme
   TextInverse textCU92
End Sub

Private Sub textCU93_GotFocus()
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textCU93.IMEMode = 1
   OpenIme
   TextInverse textCU93
End Sub

Private Sub textCU93_Validate(Cancel As Boolean)
   'add by nickc 2008/01/17 若不是修改狀態，將會出不去
   If m_EditMode <> 1 And m_EditMode <> 2 Then Exit Sub
   
   If textCU93.Text = "" Then Exit Sub
    '中日文欄位尾碼加空白，最大可輸長度減一(因可能會有造字無法存入問題)
   If Not CheckLengthIsOK(textCU93, textCU93.MaxLength - 1) Then
      Cancel = True
   End If
End Sub

Private Sub textCU94_GotFocus()
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textCU94.IMEMode = 2
   CloseIme
   TextInverse textCU94
End Sub

Private Sub textCU94_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textCU94_Validate(Cancel As Boolean)
   'If m_EditMode = 4 Then Exit Sub
   Label30(7).Caption = ""
   If textCU94.Text = "" Then Exit Sub
   If textCU94 <> "" Then textCU94 = textCU94 & String(9 - Len(textCU94), "0")
   Label30(7).Caption = ChgType(4, textCU94.Text)
   'add by nickc 2008/01/23 若不是修改狀態，將會出不去
   If m_EditMode <> 1 And m_EditMode <> 2 Then Exit Sub
               
   If Label30(7).Caption = "" Then Cancel = True
End Sub

Private Sub textCU95_GotFocus()
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textCU95.IMEMode = 2
   CloseIme
   TextInverse textCU95
End Sub

Private Sub textCU96_GotFocus()
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textCU96.IMEMode = 2
   CloseIme
   TextInverse textCU96
End Sub

Private Sub textCU96_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textCU96_Validate(Cancel As Boolean)
   'If m_EditMode = 4 Then Exit Sub
   Label30(8).Caption = ""
   If textCU96.Text = "" Then Exit Sub
   If textCU96 <> "" Then textCU96 = textCU96 & String(9 - Len(textCU96), "0")
   Label30(8).Caption = ChgType(4, textCU96.Text)
   'add by nickc 2008/01/23 若不是修改狀態，將會出不去
   If m_EditMode <> 1 And m_EditMode <> 2 Then Exit Sub
               
   If Label30(8).Caption = "" Then Cancel = True
End Sub

Private Sub textCU97_GotFocus()
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textCU97.IMEMode = 2
   CloseIme
   TextInverse textCU97
End Sub

Private Sub textCU97_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textCU97_Validate(Cancel As Boolean)
   'If m_EditMode = 4 Then Exit Sub
   Label30(9).Caption = ""
   If textCU97.Text = "" Then Exit Sub
   If textCU97 <> "" Then textCU97 = textCU97 & String(9 - Len(textCU97), "0")
   Label30(9).Caption = ChgType(4, textCU97.Text)
   'add by nickc 2008/01/23 若不是修改狀態，將會出不去
   If m_EditMode <> 1 And m_EditMode <> 2 Then Exit Sub
               
   If Label30(9).Caption = "" Then Cancel = True
End Sub

Private Sub textCU98_GotFocus()
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textCU98.IMEMode = 2
   CloseIme
   TextInverse textCU98
End Sub

Private Sub textCU98_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textCU98_Validate(Cancel As Boolean)
   'If m_EditMode = 4 Then Exit Sub
   Label30(10).Caption = ""
   If textCU98.Text = "" Then Exit Sub
   If textCU98 <> "" Then textCU98 = textCU98 & String(9 - Len(textCU98), "0")
   Label30(10).Caption = ChgType(4, textCU98.Text)
   'add by nickc 2008/01/23 若不是修改狀態，將會出不去
   If m_EditMode <> 1 And m_EditMode <> 2 Then Exit Sub
               
   If Label30(10).Caption = "" Then Cancel = True
End Sub

Private Sub textCU99_GotFocus()
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textCU99.IMEMode = 2
   CloseIme
   TextInverse textCU99
End Sub

Private Sub textCU99_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textCU99_Validate(Cancel As Boolean)
   'If m_EditMode = 4 Then Exit Sub
   Label30(11).Caption = ""
   If textCU99.Text = "" Then Exit Sub
   If textCU99 <> "" Then textCU99 = textCU99 & String(9 - Len(textCU99), "0")
   Label30(11).Caption = ChgType(4, textCU99.Text)
   'add by nickc 2008/01/23 若不是修改狀態，將會出不去
   If m_EditMode <> 1 And m_EditMode <> 2 Then Exit Sub
               
   If Label30(11).Caption = "" Then Cancel = True
End Sub

Private Sub textCUID_GotFocus()
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textCUID.IMEMode = 2
   CloseIme
   TextInverse textCUID
End Sub

' 新增記錄
Private Function AddRecord() As Boolean
Dim strSql As String
Dim strTmp As String
Dim strTit As String
Dim strMsg As String
Dim nResponse
Dim nIndex As Integer
Dim bDifference As Boolean
Dim bFirst As Boolean
Dim strCU01 As String
Dim strCU02 As String
Dim strTo As String 'Add by Amy 2024/05/15
   
   AddRecord = False
   
   strCU01 = textCU01 & String(8 - Len(textCU01), "0")
   strCU02 = textCU02 & String(1 - Len(textCU02), "0")
   
   ' 檢查記錄是否已存在
   If IsRecordExist(strCU01, strCU02) = True Then
      strTit = "新增資料"
      strMsg = "該筆記錄已存在"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      UpdateCtrlData
      Exit Function
   End If
   
   '2007/12/12 ADD BY SONIA 新增時未設定m_CU01,造成更新FAGENT之FA03錯誤
   m_CU01 = strCU01
   '2007/12/12 END
   
   bFirst = True
   bDifference = False
   strSql = "INSERT INTO customer ("
   For nIndex = 0 To TF_CU - 1 'edit by nickc 2006/10/24  MAX_FIELD - 1
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
   For nIndex = 0 To TF_CU - 1 'edit by nickc 2006/10/24  MAX_FIELD - 1
      strTmp = Empty
      If m_FieldList(nIndex).fiOldData <> m_FieldList(nIndex).fiNewData Then
         If m_FieldList(nIndex).fiType = 0 Then
            ' 90.12.18 modify by louis 字串中有單引號的處理
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
   
On Error GoTo ErrHand
   cnnConnection.BeginTrans
   'add by nickc 2006/12/20
   Pub_SeekTbLog strSql
   
   cnnConnection.Execute strSql
   
    'add by nickc 2006/03/09 新增或刪除時，有輸入代理人時，若代理人那邊的客戶是空白的，自動更新進去
   If (m_EditMode = 1 Or m_EditMode = 2) And textCU03 <> "" Then
      '2008/7/17 modify by sonia 改為詢問是否更新
      strTit = "select * from fagent where fa01='" & Left(textCU03 & "00000000", 8) & "' and fa02='0' "
      CheckOC3
      AdoRecordSet3.CursorLocation = adUseClient
      AdoRecordSet3.Open strTit, cnnConnection, adOpenStatic, adLockReadOnly
      If AdoRecordSet3.RecordCount <> 0 Then
         If Mid(CheckStr(AdoRecordSet3.Fields("fa03")), 1, 8) <> textCU01 And CheckStr(AdoRecordSet3.Fields("fa03")) <> "" Then
            strTmp = "客戶編號與代理人檔之客戶編號(" & Mid(CheckStr(AdoRecordSet3.Fields("fa03")), 1, 8) & ")不同 ! 是否更新代理人檔之客戶編號 ?"
            If MsgBox(strTmp, vbYesNo + vbCritical) = vbYes Then
               StrSQLa = "update fagent set fa03='" & m_CU01 & "',fa76='B' where fa01='" & Left(textCU03 & "00000000", 8) & "' and fa02='0'  "
               'Modify by Amy 2025/09/15 +strCU01 & strCU02 ,log記錄此客戶編號
               'Modify by Amy 2025/09/18 改只記錄8碼-秀玲
               Pub_SeekTbLog StrSQLa, , , , , strCU01
               cnnConnection.Execute StrSQLa
            End If
         ElseIf CheckStr(AdoRecordSet3.Fields("fa03")) = "" Then
            StrSQLa = "update fagent set fa03='" & m_CU01 & "',fa76='B' where fa01='" & Left(textCU03 & "00000000", 8) & "' and fa02='0'  "
            'Modify by Amy 2025/09/15 +strCU01 & strCU02 ,log記錄此客戶編號
            'Modify by Amy 2025/09/18 改只記錄8碼-秀玲
            Pub_SeekTbLog StrSQLa, , , , , strCU01
            cnnConnection.Execute StrSQLa
         End If
      End If
      '2008/7/17 ADD BY SONIA
      If textCU03.Tag <> "" And textCU03.Tag <> textCU03 Then
         strTmp = "原代理人編號(" & textCU03.Tag & ")之代理人檔的客戶編號是否清除 ?"
         If MsgBox(strTmp, vbYesNo + vbCritical) = vbYes Then
            StrSQLa = "UPDATE FAGENT SET FA03=NULL WHERE FA01='" & textCU03.Tag & "' and FA02='0' "
            'Modify by Amy 2025/09/15 +strCU01 & strCU02 ,log記錄此客戶編號
            'Modify by Amy 2025/09/18 改只記錄8碼-秀玲
            Pub_SeekTbLog StrSQLa, , , , , strCU01
            cnnConnection.Execute StrSQLa
         End If
      End If
      '2008/7/17 END
   
   'add by nickc 2006/03/16 若是清空，也要順便清空
   ElseIf (m_EditMode = 1 Or m_EditMode = 2) And textCU03 = "" And textCU03.Tag <> "" Then
        'edit by nickc 2008/03/12
        'StrSQLa = "update fagent set fa03=null where fa01='" & m_CU03 & "' and fa02='0'  "
        StrSQLa = "update fagent set fa03=null,fa76='" & m_fa76 & "' where fa01='" & Left(textCU03.Tag & "00000000", 8) & "' and fa02='0'  "
        'add by nickc 2006/12/20
        'Modify by Amy 2025/09/15 +strCU01 & strCU02 ,log記錄此客戶編號
        'Modify by Amy 2025/09/18 改只記錄8碼-秀玲
        Pub_SeekTbLog StrSQLa, , , , , strCU01
        cnnConnection.Execute StrSQLa
   End If
   
   If ((strCU01 & strCU02) < (m_FirstKEY(0) & m_FirstKEY(1))) Or ((strCU01 & strCU02) > (m_LastKEY(0) & m_LastKEY(1))) Then
      RefreshRange
   End If
   
   'Add by Morgan 2008/7/30
   '新增聯絡人資料
   If Trim(cboContact.Text) <> "" Then
      'Add by Amy 2025/05/19 名稱前後有空白要Trim
      If Left(cboContact.Text, 1) = " " Or Left(cboContact.Text, 1) = "　" Then
         cboContact.Text = Mid(cboContact.Text, 2)
      End If
      If Right(cboContact.Text, 1) = " " Or Right(cboContact.Text, 1) = "　" Then
         cboContact.Text = Mid(cboContact.Text, 1, Len(cboContact.Text) - 1)
      End If
      'end 2025/05/19
      StrSQLa = "insert into potcustcont (pcc01,pcc02,pcc05) VALUES ('" & m_CU01 & "','01','" & ChgSQL(cboContact.Text) & "')"
      'Memo by Amy 2025/09/18 log記錄此客戶編號(8碼)
      Pub_SeekTbLog StrSQLa
      cnnConnection.Execute StrSQLa
      StrSQLa = "update customer set cu127='01' where cu01='" & m_CU01 & "' and cu02='0'"
      Pub_SeekTbLog StrSQLa
      cnnConnection.Execute StrSQLa
   End If
   
   'Add By Sindy 2016/12/2 繳款書寄件處,預設值為1客戶
   StrSQLa = "update customer set cu169='1' where cu01='" & strCU01 & "' and cu02='" & strCU02 & "'"
   cnnConnection.Execute StrSQLa
   '2016/12/2 END
   
   'Add by Amy 2022/10/06 回寫接洽記錄單申請人編號
   If Left(m_PrevNo, 3) = "Add" Then
        StrSQLa = Replace(Replace(m_PrevNo, "Add ", ""), "-" & strCra02, "")
        StrSQLa = "Update ConsultRecApp Set cra05='" & strCU01 & "',cra06='" & strCU02 & "' Where cra01='" & StrSQLa & "' And cra02='" & strCra02 & "' "
        cnnConnection.Execute StrSQLa
   End If
   'Added by Lydia 2024/09/18 客戶檔新增待活化客戶的關係企業時：若所有關係企業之智權人員與新建編號的智權人員相同時，直接將此新建編號寫入OldCustomer，以利後續收文時更新關係企業待活化資料。
   If textCU10 <= "010" Then
      'Modified by Lydia 2025/01/14 處於待活化+and ocu03 is null
      strExc(0) = "select cu12,cu13,sum(decode(ocu01,'" & strCU01 & "',0,1)) cnt1,sum(decode(ocu01,'" & strCU01 & "',1,0)) cnt2 " & _
              " from oldcustomer,customer where ocu01 like '" & Left(strCU01, 6) & "%' and ocu03 is null " & _
              " and substr(ocu01,1,8)=cu01(+) and '0'=cu02(+) group by cu12,cu13 "
      intI = 1 'Add by Amy 2025/05/19  沒資料會彈訊息
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         If RsTemp.RecordCount = 1 And "" & RsTemp.Fields("cu13") = Trim(textCU13) And Val("" & RsTemp.Fields("cnt2")) = 0 Then
            strSql = "Insert Into OldCustomer (OCU01,OCU02) VALUES ('" & strCU01 & "', to_char(sysdate,'yyyymmdd')) "
            cnnConnection.Execute strSql
         'Added by Lydia 2024/09/25 debug:Email通知在新建客戶，不是在收文; ex.X25862030在9/25收文FCP案
         ElseIf RsTemp.RecordCount > 0 Then
            'Memo by Lydia 2024/09/25 從Service1搬來
            '若關係企業智權人員只要有一筆與新建編號的智權人員不同時，即發EMAIL給系統特殊設定人員：全所智權部主管(智權部)或總經理員工編號(非智權部)、副本都要給程式管理人員
            'Modified by Lydia 2025/01/14 處於待活化+and ocu03 is null
            strExc(0) = "select cu01||cu02 as custno, nvl(cu04,nvl(cu05,cu06)) custname,cu12,cu13,st02 from customer,staff where cu01='" & strCU01 & "' and cu02='" & strCU02 & "' " & _
                      " and cu13=st01(+) and substr(cu01,1,6) in (select substr(ocu01,1,6) pno from oldcustomer,customer where ocu01 like '" & Mid(strCU01, 1, 6) & "%' and ocu01<>'" & Mid(strCU01, 1, 8) & "' " & _
                      " and ocu01=cu01(+) and '0'=cu02(+) and cu13<>'" & Trim(textCU13) & "' and ocu03 is null )"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               If Left("" & RsTemp.Fields("cu12"), 1) = "S" Then
                  strExc(1) = Pub_GetSpecMan("全所智權部主管")
               Else
                  strExc(1) = Pub_GetSpecMan("總經理員工編號")
               End If
               If strExc(1) <> "" Then
                   strExc(2) = Pub_GetSpecMan("程式管理人員")
                   'Modified by Lydia 2024/11/26 +P.S.第2點
                   strExc(3) = "新關係企業：" & RsTemp.Fields("custno") & "　" & RsTemp.Fields("custname") & vbCrLf & _
                             "智權人員：" & RsTemp.Fields("cu13") & "　" & RsTemp.Fields("st02") & vbCrLf & vbCrLf & vbCrLf & _
                             "PS：1.若需調整智權人員，請通知檔案室處理。" & vbCrLf & _
                             "    2.請通知電腦中心將待活化關係企業編號補上活化記錄。"
                   strSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08,mc09)" & _
                           " values( '" & strUserNum & "','" & strExc(1) & "',to_char(sysdate,'yyyymmdd'),to_char(sysdate,'hh24miss')" & _
                           ",'待活化客戶建立了新關係企業且智權人員不同，請考慮此組客戶之智權人員是否統一？','" & ChgSQL(strExc(3)) & "'," & CNULL(strExc(2)) & ")"
                   cnnConnection.Execute strSql
               End If
            End If
         'end 2024/09/25
         End If

      End If
      
      
   End If
   'end 2024/09/18

   cnnConnection.CommitTrans
   'Add by Amy 2023/02/14 櫃台新客戶建檔,彈身分證字號錯誤(輸E222649287),是否確定按否,雖停於畫面上,但客戶編號會顯示
   '                                         導致客戶編號會帶回新客戶建檔,但卻未新增完成,若退智權會mail給秀玲又會有查不到此客戶編號的情況
   bolAddFinish = True
   
   'Add By Sindy 2014/3/25 以客戶中文抓ACC420的A4201,若存在,則發E-mail給特殊人員財務處總帳人員
   'Modify By Sindy 2015/3/23 客戶名稱有造字不可以下Trim函數,不然sql會錯誤”ORA-01756: 引號字串未以恰當方式終止”
   'strTit = "select * from acc420 where a4201='" & Trim(textCU04) & "'"
   'strTit = "select * from acc420 where a4201='" & textCU04 & "'"
   strTit = "select * from acc420 where ltrim(rtrim(a4201))=ltrim(rtrim('" & textCU04 & "'))"
   '2015/3/23 END
   CheckOC3
   AdoRecordSet3.CursorLocation = adUseClient
   AdoRecordSet3.Open strTit, cnnConnection, adOpenStatic, adLockReadOnly
   If AdoRecordSet3.RecordCount <> 0 Then
      strMsg = "客戶編號：" & strCU01 & strCU02 & vbCrLf & _
               "客戶名稱：" & Trim(textCU04) & vbCrLf & _
               "客戶檔統一編號：" & textCU11 & vbCrLf & _
               "客戶檔中文地址：" & textCU23 & vbCrLf & _
               "客戶檔聯絡地址：" & textCU31
      'Modify by Amy 2024/05/15 財務2個特殊設定拆成3個
      If Val(strSrvDate(1)) >= Val(財務拆總帳出納國內應收啟用日) Then
          strTo = Pub_GetSpecMan("財務處應收處理人員")
      Else
         strTo = Pub_GetSpecMan("財務處總帳人員")
      End If
      PUB_SendMail strUserNum, strTo, "", "收據抬頭已新建客戶資料，請確認若相同則刪除收據抬頭資料！", strMsg
      'end 2024/05/15
   End If
   '2014/3/25 END
   
   Call ChkCustNameAndPotCust 'Add By Sindy 2012/7/17 比對國內外潛在客戶名稱相同者寄Mail通知電腦中心
   Call ChkCustName 'Add By Sindy 2018/1/5 新增非個人之國內客戶時,若已有相同名稱的資料,系統自動發信給財務處
   Call ChkCustName2 'Add by Amy 2015/08/10 比對名稱是對造者寄Mail通知電腦中心
   
   ShowCurrRecord strCU01, strCU02
   AddRecord = True
   Exit Function
ErrHand:
    cnnConnection.RollbackTrans
    MsgBox " 新增失敗！" & vbCrLf & Err.Description
    'Resume Next 'Mark by Amy 2017/01/03
End Function

' 修改記錄
Private Function ModRecord() As Boolean
Dim strSql As String
Dim strTmp As String
Dim strTit As String
Dim strMsg As String
Dim nResponse
Dim nIndex As Integer
Dim bDifference As Boolean
Dim bFirst As Boolean
Dim strCU01 As String
Dim strCU02 As String
   
   ModRecord = False
   
   strCU01 = m_CurrKEY(0)
   strCU02 = m_CurrKEY(1)
   '910910  nick tigger
   '***** start
   'strSQL = "UPDATE FAGENT SET "
   strSql = "begin user_data.user_enabled:=1; UPDATE customer SET "
   '***** end
   bFirst = True
   bDifference = False
   For nIndex = 0 To TF_CU - 1 'edit by nickc 2006/10/24  MAX_FIELD - 1
      strTmp = Empty
'      '92.05.22 nick 跳過 create & update
      If nIndex < 80 Or nIndex > 85 Then
            If m_FieldList(nIndex).fiOldData <> m_FieldList(nIndex).fiNewData Then
               If m_FieldList(nIndex).fiType = 0 Then
                  If m_FieldList(nIndex).fiNewData = Empty Then
                     strTmp = m_FieldList(nIndex).fiName & " = NULL "
                  Else
                     ' 90.12.18 modify by louis 字串中有單引號的處理
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

   strSql = strSql & " " & _
                  "WHERE CU01 = '" & strCU01 & "' AND " & _
                        "CU02 = '" & strCU02 & "'; end; "
    '***** end
'910910 nick tigger
'***** start
On Error GoTo ErrHand
'***** end

   If bDifference = True Then
      '910910 nick tigger
      '**** start
      cnnConnection.BeginTrans
      '***** end
      'Add by Amy 2022/08/23 修改智權人員
      If m_EditMode = 2 And Me.textCU13.Text <> Me.textCU13.Tag Then
            '非F部門時,客戶未發文進度之智權人員同步修改(改更新要於更新客戶檔之前先做,避免智權人員已更新為新人員而抓不到資料)
            Call Pub_ChangeSaleUpdCP13(Me.textCU01.Text, Me.textCU13.Tag, Me.textCU13.Text)
            'Add by Amy 2024/05/21 更新智權文件寄送確認未確認資料 (條件有改需確認frm12040129 智權人員客戶轉移作業是否也要改)
'            strExc(2) = "UPDATE LETTERPROGRESS SET LP06 = '" & Me.textCU13.Text & "' WHERE LP06 = '" & Me.textCU13.Tag & "' AND NVL(LP07,0)=0 AND LP15<>'Y' "
'            cnnConnection.Execute strExc(2)
      End If
      
      'add by nickc 2006/12/20
      Pub_SeekTbLog strSql, , , True 'Modified by Morgan 2019/10/4 +第4參數
      cnnConnection.Execute strSql
      
      'Added by Lydia 2023/12/29 改成共用模組PUB_ChangeSaleUpdNP10
      If m_EditMode = 2 And Me.textCU13.Text <> Me.textCU13.Tag Then
         Call PUB_ChangeSaleUpdNP10(Left(textCU01 & "00000000", 8) & IIf(textCU02 = "", "0", textCU02), Me.textCU13.Tag, Me.textCU13.Text, False)
         'Added by Lydia 2024/05/09 待活化客戶：判斷智權人員為ＸＸ無效，更新OCU03取消待活化
         'Mark by Lydia 2025/01/15 因為智權部或管理部以外的部門，沒有人員為ＸＸ無效，改成全用狀態判斷
         'If InStr(Label30(2), "無效") > 0 Then
         '   strSql = "Update OldCustomer set ocu03=to_char(sysdate,'yyyymmdd') where ocu01='" & Left(textCU01 & "00000000", 8) & "' and ocu03 is null "
         '   cnnConnection.Execute strSql
         'Else
         '   '判斷智權人員從ＸＸ無效修改為非無效，檢查所有關係企業若有待活化時，則將目前修改編號CU01的OCU03取消
         '   If InStr(GetStaffName(textCU13.Tag, True), "無效") > 0 Then
         '       'Modified by Lydia 2025/01/14 + 排除本身and ocu01 <>'" & Left(textCU01 & "00000000", 8) & "'
         '       StrSQLa = "Select * From OldCustomer Where substr(ocu01,1,6)='" & Left(textCU01 & "00000000", 6) & "' and ocu03 is null and ocu01 <>'" & Left(textCU01 & "00000000", 8) & "' "
         '       If rsA.State = adStateOpen Then rsA.Close
         '       rsA.CursorLocation = adUseClient
         '       rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
         '       If rsA.RecordCount > 0 Then
         '          strSql = "Update OldCustomer set ocu03=null where ocu01='" & Left(textCU01 & "00000000", 8) & "' "
         '          cnnConnection.Execute strSql
         '       End If
         '   End If
         'End If
         ''end 2024/05/09
         'end 2025/01/15
      End If
      'end 2023/12/29
      
      'Added by Lydia 2025/01/15 變更客戶狀態，取消待活化客戶
      If m_EditMode = 2 And m_FieldList(79).fiOldData <> cboStatus.Text Then
         strExc(0) = Pub_GetSpecMan("待活化客戶-無效狀態設定")
         'Modified by Lyida 2025/02/04 排除空白+And Trim(cboStatus.Text) <> ""
         If InStr(strExc(0), Trim(cboStatus.Text)) > 0 And Trim(cboStatus.Text) <> "" Then
            strSql = "Update OldCustomer set ocu03=to_char(sysdate,'yyyymmdd') where ocu01='" & Left(textCU01 & "00000000", 8) & "' and ocu03 is null "
            cnnConnection.Execute strSql
         Else
            '從ＸＸ無效修改為非無效，檢查所有關係企業若有待活化時，則將目前修改編號CU01的OCU03取消
            If InStr(strExc(0), m_FieldList(79).fiOldData) > 0 And InStr(strExc(0), Trim(cboStatus.Text)) = 0 Then
                StrSQLa = "Select * From OldCustomer Where substr(ocu01,1,6)='" & Left(textCU01 & "00000000", 6) & "' and ocu03 is null and ocu01 <>'" & Left(textCU01 & "00000000", 8) & "' "
                If rsA.State = adStateOpen Then rsA.Close
                rsA.CursorLocation = adUseClient
                rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
                If rsA.RecordCount > 0 Then
                   'Modified by Lydia 2025/03/07 +OCU02
                   'Modify by Amy 2025/03/19 bug-原:rsA.Fields("ouc02")
                   strSql = "Update OldCustomer set ocu03=null where ocu01='" & Left(textCU01 & "00000000", 8) & "' and ocu02=" & rsA.Fields("ocu02")
                   cnnConnection.Execute strSql
                End If
            End If
         End If
      End If
      'end 2025/01/15
      
      'Add by Amy 2022/07/01 修改母號狀態,更名前資料 cu12、cu13、cu80、cu132、cu180 一併改
      'Modify by Amy 2024/05/22 修改母號 不提供ID欄,更名前資料 一併改
      If m_EditMode = 2 And textCU02 = "0" _
         And (m_FieldList(79).fiOldData <> m_FieldList(79).fiNewData Or m_FieldList(181).fiOldData <> m_FieldList(181).fiNewData) Then
            StrSQLa = "Select  cu02,cu12,cu13,cu80,cu132,cu180,cu182 From Customer Where cu01='" & Left(textCU01 & "00000000", 8) & "' and cu02<>'0' "
            If rsA.State = adStateOpen Then rsA.Close
            rsA.CursorLocation = adUseClient
            rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
            If rsA.RecordCount > 0 Then
                Do While Not rsA.EOF
                    strSql = ""
                    'Modify by Amy 2024/05/22 修改母號 不提供ID欄,更名前資料 一併改
                    If m_FieldList(79).fiOldData <> m_FieldList(79).fiNewData Then
                        If textCU12 <> "" & rsA.Fields("cu12") Then strSql = strSql & ",cu12='" & textCU12.Text & "'"
                        If textCU13 <> "" & rsA.Fields("cu13") Then strSql = strSql & ",cu13='" & textCU13.Text & "'"
                        If cboStatus <> "" & rsA.Fields("cu80") Then strSql = strSql & ",cu80='" & cboStatus & "'"
                        If textCU132 <> "" & rsA.Fields("cu132") Then strSql = strSql & ",cu132='" & textCU132.Text & "'"
                        'Modify by Amy 2024/09/23 cu180(客戶狀態備註)無資料不加; -秀玲 ex:X0844807 與母號備註欄不同,每日檢查語法會檢查出不一致
                        If IsNull(rsA.Fields("cu180")) Then
                           strSql = strSql & ",cu180='" & textCU180 & "'"
                        Else
                           strSql = strSql & ",cu180='" & textCU180 & ";" & rsA.Fields("cu180") & "'"
                        End If
                        'end 2024/09/23
                    End If
                    strExc(2) = IIf(ChkID.Value = 1, "Y", IIf(m_FieldList(181).fiOldData = "W", "W", ""))
                    If strExc(2) <> "" & rsA.Fields("cu182") Then strSql = strSql & ",cu182=" & CNULL(strExc(2))
                    'end 2024/05/22
                    If strSql <> MsgText(601) Then
                        strSql = "Update Customer Set " & Mid(strSql, 2) & " Where cu01='" & Left(textCU01 & "00000000", 8) & "' And cu02='" & rsA.Fields("cu02") & "' "
                        Pub_SeekTbLog strSql, , , True
                        cnnConnection.Execute strSql
                    End If
                    rsA.MoveNext
                Loop
            End If
      End If
      'end 2022/07/01
      
      'Add By Sindy 2013/7/2
      '修改公司負責人CU07時 (改名稱)
      '若業務區非 Fxx 部門時, 若cu39或cu40或cu41有值時,
      '存檔時自動清除 cu39,cu40,cu41 為 null
      'Remove by Lydia 2019/04/17 取消自動清空，改在存檔彈提醒
      'If Trim(m_CU07) <> Trim(textCU07) And Left(Trim(textCU12), 1) <> "F" Then
      '   strSql = "update customer set cu39=null,cu40=null,cu41=null " & _
      '            "WHERE CU01 = '" & strCU01 & "' AND " & _
      '                  "CU02 = '" & strCU02 & "'"
      '   Pub_SeekTbLog strSql
      '   cnnConnection.Execute strSql
      'End If
      '2013/7/2 END
      
'edit by nickc 2007/03/05 根本沒用到
'    strExc(0) = "Select CU01, CU02, CU03 From Customer Where CU01='" & m_CU01 & "' And CU02='" & m_CU02 & "' "
'    If RcMain.State = 1 Then RcMain.Close
'    RcMain.CursorLocation = adUseClient
'    RcMain.Open strExc(0), cnnConnection, adOpenDynamic, adLockReadOnly
    '若修改智權人員
    '2006/8/2 MODIFY BY SONIA 國外部的不更新 X47749
    'If ActionEdit = 1 And Me.Text1(15).Text <> Me.Text1(15).Tag Then
    '2012/8/7 modify by sonia 國內智權人員改國外智權人員也要做X44833之77050->80030
    'If m_EditMode = 2 And Me.textCU13.Text <> Me.textCU13.Tag And Mid(Me.textCU12, 1, 1) <> "F" Then
    'Mark by Lydia 2023/12/29 改成共用模組PUB_ChangeSaleUpdNP10
'    If m_EditMode = 2 And Me.textCU13.Text <> Me.textCU13.Tag And (Mid(Me.textCU12, 1, 1) <> "F" Or Mid(Me.textCU12.Tag, 1, 1) <> "F") Then
'        'Modify By Cheng 2003/09/29
''        strSQLA = "Select NP01, NP07, NP22 From NextProgress, Patent Where NP02=PA01 And NP03=PA02 And NP04=PA03 And NP05=PA04 And NP10='" & Me.Text1(15).Tag & "' And NP06 Is Null And NP08>=" & strSrvDate(1) & " And  PA26='" & Me.textcu01.Text & IIf(Me.textcu02.Text = "", "0", Me.textcu02.Text) & "' "
''        strSQLA = strSQLA & " Union Select NP01, NP07, NP22 From NextProgress, Trademark Where NP02=TM01 And NP03=TM02 And NP04=TM03 And NP05=TM04 And NP10='" & Me.Text1(15).Tag & "' And NP06 Is Null And NP08>=" & strSrvDate(1) & " And  TM23='" & Me.textcu01.Text & IIf(Me.textcu02.Text = "", "0", Me.textcu02.Text) & "' "
''        strSQLA = strSQLA & " Union Select NP01, NP07, NP22 From NextProgress, Lawcase Where NP02=LC01 And NP03=LC02 And NP04=LC03 And NP05=LC04 And NP10='" & Me.Text1(15).Tag & "' And NP06 Is Null And NP08>=" & strSrvDate(1) & " And  LC11='" & Me.textcu01.Text & IIf(Me.textcu02.Text = "", "0", Me.textcu02.Text) & "' "
''        strSQLA = strSQLA & " Union Select NP01, NP07, NP22 From NextProgress, Hirecase Where NP02=HC01 And NP03=HC02 And NP04=HC03 And NP05=HC04 And NP10='" & Me.Text1(15).Tag & "' And NP06 Is Null And NP08>=" & strSrvDate(1) & " And  HC05='" & Me.textcu01.Text & IIf(Me.textcu02.Text = "", "0", Me.textcu02.Text) & "' "
''        strSQLA = strSQLA & " Union Select NP01, NP07, NP22 From NextProgress, ServicePractice Where NP02=SP01 And NP03=SP02 And NP04=SP03 And NP05=SP04 And NP10='" & Me.Text1(15).Tag & "' And NP06 Is Null And NP08>=" & strSrvDate(1) & " And  SP08='" & Me.textcu01.Text & IIf(Me.textcu02.Text = "", "0", Me.textcu02.Text) & "' "
'        '93.5.11 MODIFY BY SONIA "FCT","FCP","FCL","FG"不更新
'        'strSQLA = "Select NP01, NP07, NP22 From NextProgress, Patent Where NP02=PA01 And NP03=PA02 And NP04=PA03 And NP05=PA04 And NP06 Is Null And NP08>=" & strSrvDate(1) & " And  PA26='" & Me.textcu01.Text & IIf(Me.textcu02.Text = "", "0", Me.textcu02.Text) & "' "
'        'strSQLA = strSQLA & " Union Select NP01, NP07, NP22 From NextProgress, Trademark Where NP02=TM01 And NP03=TM02 And NP04=TM03 And NP05=TM04 And NP06 Is Null And NP08>=" & strSrvDate(1) & " And  TM23='" & Me.textcu01.Text & IIf(Me.textcu02.Text = "", "0", Me.textcu02.Text) & "' "
'        'strSQLA = strSQLA & " Union Select NP01, NP07, NP22 From NextProgress, Lawcase Where NP02=LC01 And NP03=LC02 And NP04=LC03 And NP05=LC04 And NP06 Is Null And NP08>=" & strSrvDate(1) & " And  LC11='" & Me.textcu01.Text & IIf(Me.textcu02.Text = "", "0", Me.textcu02.Text) & "' "
'        'strSQLA = strSQLA & " Union Select NP01, NP07, NP22 From NextProgress, Hirecase Where NP02=HC01 And NP03=HC02 And NP04=HC03 And NP05=HC04 And NP06 Is Null And NP08>=" & strSrvDate(1) & " And  HC05='" & Me.textcu01.Text & IIf(Me.textcu02.Text = "", "0", Me.textcu02.Text) & "' "
'        'strSQLA = strSQLA & " Union Select NP01, NP07, NP22 From NextProgress, ServicePractice Where NP02=SP01 And NP03=SP02 And NP04=SP03 And NP05=SP04 And NP06 Is Null And NP08>=" & strSrvDate(1) & " And  SP08='" & Me.textcu01.Text & IIf(Me.textcu02.Text = "", "0", Me.textcu02.Text) & "' "
'        'Modify By Sindy 2009/07/24 增加LIN系統類別
'        '2011/9/28 MODIFY BY SONIA NP08改為NP09
'        'Modify by Amy 2022/0/31 +NP02~05 本所案號
'        StrSQLa = "Select NP01, NP07, NP22,NP02,NP03,NP04,NP05 From NextProgress, Patent Where NP02<>'FCP' AND NP02=PA01 And NP03=PA02 And NP04=PA03 And NP05=PA04 And NP06 Is Null And NP09>=" & strSrvDate(1) & " And  PA26='" & Me.textCU01.Text & IIf(textCU02.Text = "", "0", textCU02.Text) & "' "
'        StrSQLa = StrSQLa & " Union Select NP01, NP07, NP22,NP02,NP03,NP04,NP05 From NextProgress, Trademark Where NP02<>'FCT' AND NP02=TM01 And NP03=TM02 And NP04=TM03 And NP05=TM04 And NP06 Is Null And NP09>=" & strSrvDate(1) & " And  TM23='" & textCU01.Text & IIf(textCU02.Text = "", "0", textCU02.Text) & "' "
'        StrSQLa = StrSQLa & " Union Select NP01, NP07, NP22,NP02,NP03,NP04,NP05 From NextProgress, Lawcase Where (NP02<>'FCL' and NP02<>'LIN') AND NP02=LC01 And NP03=LC02 And NP04=LC03 And NP05=LC04 And NP06 Is Null And NP09>=" & strSrvDate(1) & " And  LC11='" & textCU01.Text & IIf(textCU02.Text = "", "0", textCU02.Text) & "' "
'        StrSQLa = StrSQLa & " Union Select NP01, NP07, NP22,NP02,NP03,NP04,NP05 From NextProgress, Hirecase Where NP02=HC01 And NP03=HC02 And NP04=HC03 And NP05=HC04 And NP06 Is Null And NP09>=" & strSrvDate(1) & " And  HC05='" & textCU01.Text & IIf(textCU02.Text = "", "0", textCU02.Text) & "' "
'        StrSQLa = StrSQLa & " Union Select NP01, NP07, NP22,NP02,NP03,NP04,NP05 From NextProgress, ServicePractice Where NP02<>'FG' AND NP02=SP01 And NP03=SP02 And NP04=SP03 And NP05=SP04 And NP06 Is Null And NP09>=" & strSrvDate(1) & " And  SP08='" & textCU01.Text & IIf(textCU02.Text = "", "0", textCU02.Text) & "' "
'        '93.5.11 END
'        If rsA.State = 1 Then rsA.Close
'        rsA.CursorLocation = adUseClient
'        rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
'        If rsA.RecordCount > 0 Then
'            While Not rsA.EOF
'                'Modify by Amy 2022/08/31 +if MCT案件的智權人員是依FC代理人的管控智權人員FA120,而不是客戶檔的智權人員
'                If Left(Me.textCU13.Text, 3) = "MCT" Then
'                    'Memo 原若已是MCT則不會更新,因可能同一案子不同代理人,但原不是MCT改MCT要更新(但要先設定FA120)
'                    StrSQLa = "Update Nextprogress Set NP10='" & PUB_GetAKindSalesNo("" & rsA.Fields("NP02"), "" & rsA.Fields("NP03"), "" & rsA.Fields("NP04"), "" & rsA.Fields("NP05")) & "' " & _
'                                    "Where NP01='" & rsA.Fields(0).Value & "' And NP07='" & rsA.Fields(1).Value & "' And NP22=" & rsA.Fields(2).Value & strNpSqlOfNoSalesDuty
'                Else
'                    '2010/3/22 MODIFY BY SONIA 剔除下一程序非智權人員掌控之案件性質改以strNpSqlOfNoSalesDuty控制
'                    StrSQLa = "Update Nextprogress Set NP10='" & Me.textCU13.Text & "' Where NP01='" & rsA.Fields(0).Value & "' And NP07='" & rsA.Fields(1).Value & "' And NP22=" & rsA.Fields(2).Value & strNpSqlOfNoSalesDuty
'                End If
'                Pub_SeekTbLog StrSQLa
'                cnnConnection.Execute "begin user_data.user_notrigger:=1; end;" 'Add by Morgan 2012/8/22 +控制來函期限通知的 Trigger 不被觸發
'                cnnConnection.Execute StrSQLa
'                cnnConnection.Execute "begin user_data.user_notrigger:=0; end;" 'Add by Morgan 2012/8/22 +控制來函期限通知的 Trigger 不被觸發
'                rsA.MoveNext
'            Wend
'        End If
'        If rsA.State <> adStateClosed Then rsA.Close
'        Set rsA = Nothing
'    End If
    
    'add by nickc 2006/06/28 修改智權人員或業務區時時，一併修改舊名稱的資料
    'Mark by Lydia 2023/12/29 改成共用模組PUB_ChangeSaleUpdNP10
    'If m_EditMode = 2 And Me.textCU12.Text <> Me.textCU12.Tag And m_CU02 = "0" Then
    '    StrSQLa = "update customer set cu12='" & Me.textCU12.Text & "' where cu01='" & m_CU01 & "' and cu02<>'0' "
    '    'add by nickc 2006/12/20
    '    Pub_SeekTbLog StrSQLa
    '    cnnConnection.Execute StrSQLa
    'End If
    'Mark by Lydia 2023/12/29 改成共用模組PUB_ChangeSaleUpdNP10
    'If m_EditMode = 2 And Me.textCU13.Text <> Me.textCU13.Tag And m_CU02 = "0" Then
    '    StrSQLa = "update customer set cu13='" & Me.textCU13.Text & "' where cu01='" & m_CU01 & "' and cu02<>'0' "
    '    'add by nickc 2006/12/20
    '    Pub_SeekTbLog StrSQLa
    '    cnnConnection.Execute StrSQLa
    '    'add by nickc 2007/11/26 改舊名稱時，一併修改
    '    'Modify By Sindy 2009/07/24 增加LIN系統類別
    '    '2011/9/28 MODIFY BY SONIA NP08改為NP09
    '    'Modify by Amy 2022/0/31 +NP02~05 本所案號
     '   StrSQLa = "Select NP01, NP07, NP22,NP02,NP03,NP04,NP05 From NextProgress, Patent Where NP02<>'FCP' AND NP02=PA01 And NP03=PA02 And NP04=PA03 And NP05=PA04 And NP06 Is Null And NP09>=" & strSrvDate(1) & " And  PA26>='" & textCU01.Text & "1'  and pa26<='" & textCU01.Text & "9' "
     '   StrSQLa = StrSQLa & " Union Select NP01, NP07, NP22,NP02,NP03,NP04,NP05 From NextProgress, Trademark Where NP02<>'FCT' AND NP02=TM01 And NP03=TM02 And NP04=TM03 And NP05=TM04 And NP06 Is Null And NP09>=" & strSrvDate(1) & " And  TM23>='" & textCU01.Text & "1'  and tm23<='" & textCU01.Text & "9' "
     '   StrSQLa = StrSQLa & " Union Select NP01, NP07, NP22,NP02,NP03,NP04,NP05 From NextProgress, Lawcase Where (NP02<>'FCL' and NP02<>'LIN') AND NP02=LC01 And NP03=LC02 And NP04=LC03 And NP05=LC04 And NP06 Is Null And NP09>=" & strSrvDate(1) & " And  LC11>='" & textCU01.Text & "1' and lc11<='" & textCU01.Text & "9'  "
     '   StrSQLa = StrSQLa & " Union Select NP01, NP07, NP22,NP02,NP03,NP04,NP05 From NextProgress, Hirecase Where NP02=HC01 And NP03=HC02 And NP04=HC03 And NP05=HC04 And NP06 Is Null And NP09>=" & strSrvDate(1) & " And  HC05>='" & textCU01.Text & "1'  and hc05<='" & textCU01.Text & "9' "
     '   StrSQLa = StrSQLa & " Union Select NP01, NP07, NP22,NP02,NP03,NP04,NP05 From NextProgress, ServicePractice Where NP02<>'FG' AND NP02=SP01 And NP03=SP02 And NP04=SP03 And NP05=SP04 And NP06 Is Null And NP09>=" & strSrvDate(1) & " And  SP08>='" & textCU01.Text & "1' and sp08<='" & textCU01.Text & "9' "
     '   If rsA.State = 1 Then rsA.Close
     '   rsA.CursorLocation = adUseClient
     '   rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
     '   If rsA.RecordCount > 0 Then
     '       While Not rsA.EOF
     '            'Modify by Amy 2022/08/31 +if MCT案件的智權人員是依FC代理人的管控智權人員FA120,而不是客戶檔的智權人員
     '           If Left(Me.textCU13.Text, 3) = "MCT" Then
     '               'Memo 原若已是MCT則不會更新,因可能同一案子不同代理人,但原不是MCT改MCT要更新(但要先設定FA120)
      '              StrSQLa = "Update Nextprogress Set NP10='" & PUB_GetAKindSalesNo("" & rsA.Fields("NP02"), "" & rsA.Fields("NP03"), "" & rsA.Fields("NP04"), "" & rsA.Fields("NP05")) & "' " & _
      '                              "Where NP01='" & rsA.Fields(0).Value & "' And NP07='" & rsA.Fields(1).Value & "' And NP22=" & rsA.Fields(2).Value & strNpSqlOfNoSalesDuty
      '          Else
      '              '2010/3/22 MODIFY BY SONIA 剔除下一程序非智權人員掌控之案件性質改以strNpSqlOfNoSalesDuty控制
      '              StrSQLa = "Update Nextprogress Set NP10='" & Me.textCU13.Text & "' Where NP01='" & rsA.Fields(0).Value & "' And NP07='" & rsA.Fields(1).Value & "' And NP22=" & rsA.Fields(2).Value & strNpSqlOfNoSalesDuty
      '          End If
      '          Pub_SeekTbLog StrSQLa
      '          cnnConnection.Execute "begin user_data.user_notrigger:=1; end;" 'Add by Morgan 2012/8/22 +控制來函期限通知的 Trigger 不被觸發
      '          cnnConnection.Execute StrSQLa
      '          cnnConnection.Execute "begin user_data.user_notrigger:=0; end;" 'Add by Morgan 2012/8/22 +控制來函期限通知的 Trigger 不被觸發
      '          rsA.MoveNext
       '     Wend
       ' End If
       ' If rsA.State <> adStateClosed Then rsA.Close
       ' Set rsA = Nothing
    'End If
'end 2023/12/29 ---'Mark by Lydia 2023/12/29 改成共用模組 PUB_ChangeSaleUpdNP10

    'add by nickc 2006/03/09 新增或刪除時，有輸入代理人時，若代理人那邊的客戶是空白的，自動更新進去
   If (m_EditMode = 1 Or m_EditMode = 2) And textCU03 <> "" Then
      '2008/7/17 modify by sonia 改為詢問是否更新
      strTit = "select * from fagent where fa01='" & Left(textCU03 & "00000000", 8) & "' and fa02='0' "
      CheckOC3
      AdoRecordSet3.CursorLocation = adUseClient
      AdoRecordSet3.Open strTit, cnnConnection, adOpenStatic, adLockReadOnly
      If AdoRecordSet3.RecordCount <> 0 Then
         If Mid(CheckStr(AdoRecordSet3.Fields("fa03")), 1, 8) <> textCU01 And CheckStr(AdoRecordSet3.Fields("fa03")) <> "" Then
            strTmp = "客戶編號與代理人檔之客戶編號(" & Mid(CheckStr(AdoRecordSet3.Fields("fa03")), 1, 8) & ")不同 ! 是否更新代理人檔之客戶編號 ?"
            If MsgBox(strTmp, vbYesNo + vbCritical) = vbYes Then
               StrSQLa = "update fagent set fa03='" & m_CU01 & "',fa76='B' where fa01='" & Left(textCU03 & "00000000", 8) & "' and fa02='0'  "
               'Modify by Amy 2025/09/15 +strCU01 & strCU02 ,log記錄此客戶編號
               'Modify by Amy 2025/09/18 改只記錄8碼-秀玲
               Pub_SeekTbLog StrSQLa, , , , , strCU01
               cnnConnection.Execute StrSQLa
            End If
         ElseIf CheckStr(AdoRecordSet3.Fields("fa03")) = "" Then
            StrSQLa = "update fagent set fa03='" & m_CU01 & "',fa76='B' where fa01='" & Left(textCU03 & "00000000", 8) & "' and fa02='0'  "
            'Modify by Amy 2025/09/15 +strCU01 & strCU02 ,log記錄此客戶編號
            'Modify by Amy 2025/09/18 改只記錄8碼-秀玲
            Pub_SeekTbLog StrSQLa, , , , , strCU01
            cnnConnection.Execute StrSQLa
         End If
      End If
      '2008/7/17 ADD BY SONIA
      If textCU03.Tag <> "" And textCU03.Tag <> textCU03 Then
         strTmp = "原代理人編號(" & textCU03.Tag & ")之代理人檔的客戶編號是否清除 ?"
         If MsgBox(strTmp, vbYesNo + vbCritical) = vbYes Then
            StrSQLa = "UPDATE FAGENT SET FA03=NULL WHERE FA01='" & textCU03.Tag & "' and FA02='0' "
            'Modify by Amy 2025/09/15 +strCU01 & strCU02 ,log記錄此客戶編號
            'Modify by Amy 2025/09/18 改只記錄8碼-秀玲
            Pub_SeekTbLog StrSQLa, , , , , strCU01
            cnnConnection.Execute StrSQLa
         End If
      End If
      '2008/7/17 END
   'add by nickc 2006/03/16 若是清空，也要順便清空
   ElseIf m_EditMode = 2 And textCU03 = "" And textCU03.Tag <> "" Then
        'edit by nickc 2008/03/12
        'StrSQLa = "update fagent set fa03=null where fa01='" & m_CU03 & "' and fa02='0'  "
        StrSQLa = "update fagent set fa03=null,fa76='" & m_fa76 & "' where fa01='" & Left(textCU03.Tag & "00000000", 8) & "' and fa02='0'  "
        'add by nickc 2006/12/20
        'Modify by Amy 2025/09/15 +strCU01 & strCU02 ,log記錄此客戶編號
        'Modify by Amy 2025/09/18 改只記錄8碼-秀玲
        Pub_SeekTbLog StrSQLa, , , , , strCU01
        cnnConnection.Execute StrSQLa
   End If

      'Added by Lydia 2019/11/27 年費不續辦CU74=N => 目前案件的年費期限自動上不續辦
      strMsg = "": strTmp = ""  'Added by Lydia 2020/03/17
      If textCU74.Text = "N" And textCU74.Tag <> textCU74.Text Then
          'Modified by Lydia 2020/03/17 回傳FMP案範圍，發清單通知程序
          'Call Pub_AutoUpdFCP605(textCU01 & textCU02)
          If Pub_AutoUpdFCP605(textCU01 & textCU02, strTmp, strMsg) = False Then
               GoTo ErrHand
          End If
          'end 2020/03/17
      End If
      'end 2019/11/27

   'Added by Morgan 2021/11/29
   '身分變更時檢查若有設定減免身分時EMail通知最後修改人員
   If m_CU15 <> textCU15 Then
      'Modified by Morgan 2022/12/15 若為QPGMR時改發智權人員(舊資料)
      strExc(0) = "select nvl(ad07,ad04) Usr,na03 Cty,cu13" & _
         " from APPLICANTDISCOUNT,nation,customer where ad01='" & textCU01 & "' and na01(+)=ad02 and cu01(+)=ad01 and cu02(+)='0'"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         With RsTemp
         .MoveFirst
         Do While Not .EOF
            If .Fields("Usr") <> "QPGMR" Then
               strExc(0) = .Fields("Usr")
            Else
               strExc(0) = .Fields("cu13")
            End If
            If strExc(0) <> "" Then
               strExc(1) = "【" & textCU01 & " " & textCU04 & "】的身分已由【" & optCustomer(Val(m_CU15)).Caption & "】變更為【" & optCustomer(Val(textCU15)).Caption & "】，請檢查該編號【" & .Fields("Cty") & "】的減免身分設定是否正確！"
               strSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08)" & _
                  " values( '" & strUserNum & "','" & strExc(0) & "',to_char(sysdate,'yyyymmdd'),to_char(sysdate,'hh24miss')" & _
                  ",'" & ChgSQL(strExc(1)) & "','如旨')"
               cnnConnection.Execute strSql, intI
            End If
            .MoveNext
         Loop
         End With
      End If
   End If
   'end 2021/11/29

      '910910 nick tigger
      '***** start
      cnnConnection.CommitTrans
      '***** end
        
     'Added by Lydia 2020/03/17 FMP案件不自動上年費不續辦，改發清單給程序，由各區程序逐筆產生定稿通知大陸代理人
     If strTmp <> "" Then
        If PUB_GetP605Email("1", strTmp, strMsg) = False Then
           If strMsg <> "" Then
               MsgBox strMsg, vbCritical
           End If
        End If
     End If
     'end 2020/03/17
     
      'Modify By Sindy 2022/9/19 Mark
'      'Add By Sindy 2016/5/20
'      '原智權人員離職時,調整待會稿區正在送會中及會圖中的收受者
'      If textCU13.Tag <> textCU13.Text Then
'         Call PUB_SalseLeaveUpEEP05(textCU13.Tag)
'      End If
      
      'Add by Morgan 2006/1/10 若代理人有修改且有相對的匯款銀行資料時需發Mail通知婧瑄
      '2011/12/26 MODIFY BY SONIA 加入中日文名稱欄位
      'If (textCU05.Tag & textCU88.Tag & textCU89.Tag & textCU90.Tag <> "") And (textCU05.Tag & textCU88.Tag & textCU89.Tag & textCU90.Tag <> textCU05 & textCU88 & textCU89 & textCU90) Then
      '   PUB_AccDataCheck textCU01 & textCU02, "英：" & textCU05.Tag & " " & textCU88.Tag & " " & textCU89.Tag & " " & textCU90.Tag & " --> " & textCU05 & " " & textCU88 & " " & textCU89 & " " & textCU90
      'End If
      If (textCU04.Tag & textCU06.Tag & textCU05.Tag & textCU88.Tag & textCU89.Tag & textCU90.Tag <> "") And (textCU04.Tag & textCU06.Tag & textCU05.Tag & textCU88.Tag & textCU89.Tag & textCU90.Tag <> textCU04 & textCU06 & textCU05 & textCU88 & textCU89 & textCU90) Then
         PUB_AccDataCheck textCU01 & textCU02, _
         "中：" & textCU04.Tag & " --> " & textCU04 & Chr(13) & _
         "英：" & textCU05.Tag & " " & textCU88.Tag & " " & textCU89.Tag & " " & textCU90.Tag & " --> " & textCU05 & " " & textCU88 & " " & textCU89 & " " & textCU90 & Chr(13) & _
         "日：" & textCU06.Tag & " --> " & textCU06
      End If
      '2011/12/26 END
      '2006/1/10 end
      
      '2010/5/4 add by sonia 修改時若有更名前資料,提醒使用者詢問智權人員 X61578020
      If m_EditMode = 2 And textCU02 = "0" Then
         strTit = "select * from customer where cu01='" & Left(textCU01 & "00000000", 8) & "' and cu02<>'0' "
         CheckOC3
         AdoRecordSet3.CursorLocation = adUseClient
         AdoRecordSet3.Open strTit, cnnConnection, adOpenStatic, adLockReadOnly
         If AdoRecordSet3.RecordCount <> 0 Then
            ShowMsg "此客戶曾經變更名稱，請與智權人員確認是否修改更名前資料(除智權人員欄外) ! 若需要修改請查更名前編號再人工修改 !"
         End If
      End If
      '2010/5/4 end
      
      'Add By Sindy 2013/12/12 原為學校改為非學校者,寄mail通知秀玲
      'Modify by Amy 2023/11/16 原寄給83002
      If m_CU15 = "2" And optCustomer(2).Value = False Then
         PUB_SendMail strUserNum, Pub_GetSpecMan("程式管理人員"), "", strCU01 & strCU02 & "原為學校" & ChangeWStringToTDateString(strSrvDate(1)) & strUserName & "改為非學校，確認是否不可開發票", "同主旨"
      End If
      '2013/12/12 END
      
      Call ChkCustNameAndPotCust 'Add By Sindy 2012/7/17 比對國內外潛在客戶名稱相同者寄Mail通知電腦中心
           
      ShowCurrRecord strCU01, strCU02
   End If
'910910 nick tigger
'***** start
    ModRecord = True
   Exit Function
ErrHand:
    cnnConnection.RollbackTrans
    'Modified by Lydia 2020/03/17
    'MsgBox (Err.Description)
    MsgBox (Err.Description) & vbCrLf & strMsg
'******* end
    Resume 'Mark by Amy 2017/01/03
End Function

' 刪除記錄
Private Function DelRecord() As Boolean
Dim strSql As String
Dim strCU01 As String
Dim strCU02 As String
Dim lngDel As Long 'Add by Amy 2025/09/15

   DelRecord = False
   
On Error GoTo ErrHand
   
   cnnConnection.BeginTrans
   
   strCU01 = m_CurrKEY(0)
   strCU02 = m_CurrKEY(1)
   
   'Add by Amy 2022/08/23
   If strCU02 = "0" Then
        'Modify by Amy 2025/09/15 +lngDEL ,有刪除資料才寫log,並傳入筆數
        'Add by Amy 2022/08/23 刪除發明人
        strSql = "Delete From INVENTOR Where IN01 in (Select cu01 From Customer Where cu01='" & strCU01 & "') "
        cnnConnection.Execute strSql, lngDel
        'Modify by Amy 2025/09/18 改只記錄8碼-秀玲,原:strCU01 & strCU02
        If lngDel > 0 Then Pub_SeekTbLog strSql, , , , , strCU01 & ";" & lngDel
        
        'Add by Amy 2022/08/26 刪除客戶減免身分檔
        strSql = "Delete From APPLICANTDISCOUNT Where AD01 in (Select cu01 From Customer Where cu01='" & strCU01 & "') "
        cnnConnection.Execute strSql, lngDel
        'Modify by Amy 2025/09/18 改只記錄8碼-秀玲,原:strCU01 & strCU02
        If lngDel > 0 Then Pub_SeekTbLog strSql, , , , , strCU01 & ";" & lngDel
        'end 2025/09/15
   End If
   
   strSql = "DELETE FROM customer " & _
            "WHERE CU01 = '" & strCU01 & "' AND " & _
                  "CU02 = '" & strCU02 & "' "
   
   'add by nickc 2006/12/20
   Pub_SeekTbLog strSql
   cnnConnection.Execute strSql
  
   'Modify by Amy 2025/09/19 原2025/09/15 判斷有刪除資料才寫log,改至Pub_SeekTbLog判斷,故程式改回(先寫log再執行刪除)並傳入客戶編號8碼
   If strCU02 = "0" Then 'Added by Lydia 2021/10/27 增加判斷更名後的號碼不可刪除;  ex.刪除代理人Y34013002一併刪除各項指示
        '2012/11/9 ADD BY SONIA 同時刪除接洽人
        strSql = "DELETE FROM POTCUSTCONT  " & _
                 "WHERE PCC01 = '" & strCU01 & "' "
        Pub_SeekTbLog strSql, , , , , strCU01
        cnnConnection.Execute strSql
        '2012/11/9 END
        
        'Added by Lydia 2016/10/28 一併刪除申請人指定國外代理人檔
        strSql = "delete from CustAssignAgent where caa02=" & CNULL(strCU01)
        Pub_SeekTbLog strSql, , , , , strCU01
        cnnConnection.Execute strSql
        'end 2016/10/28
   End If 'end 2021/10/27
   
   'Add By Sindy 2016/11/1 同時刪除會計師資料檔
   strSql = "DELETE FROM ACC490  " & _
            "WHERE A4901 = '" & strCU01 & strCU02 & "' "
   Pub_SeekTbLog strSql, , , , , strCU01
   cnnConnection.Execute strSql
   '2016/11/1 END
   
   If strCU02 = "0" Then 'Added by Lydia 2021/10/27 增加判斷更名後的號碼不可刪除;  ex.刪除代理人Y34013002一併刪除各項指示
        'Added by Lydia 2016/11/22 一併刪除國外固定寄催款單代理人檔
        strSql = "delete from Acc225 where a2251=" & CNULL(strCU01)
        Pub_SeekTbLog strSql, , , , , strCU01
        cnnConnection.Execute strSql
        'end 2016/11/22
        
        'Added by Lydia 2016/11/24 一併刪除各項指示
         strSql = "DELETE FROM INSTRUCTIONS WHERE ITS01=" & CNULL(Pub_GetITS01Type(strCU01)) & " AND ITS02=" & CNULL(strCU01)
         Pub_SeekTbLog strSql, , , , , strCU01
         cnnConnection.Execute strSql
         'end 2016/11/24
         
        'Added by Lydia 2016/11/30 一併刪除國外部關聯企業資料
        strSql = "delete from frelation where fr01=" & CNULL(strCU01) & " or fr02=" & CNULL(strCU01)
        Pub_SeekTbLog strSql, , , , , strCU01
        cnnConnection.Execute strSql
        'end 2016/11/30
        
       'Added by Lydia 2023/01/03 刪除外專特殊設定備註; ex.Y53912直接刪除代理人編號
       '下一程序固定備註(NpMemo)
       If ChkExistSpec("NPMEMO", strCU01, 8) = True Then
          strSql = "Delete From NPMEMO WHERE NM05='" & Left(strCU01, 8) & "' AND NM04 IS NULL "
          Pub_SeekTbLog strSql, , , , , strCU01
          cnnConnection.Execute strSql
       End If
       If ChkExistSpec("NPMEMO", strCU01, 6) = True Then
          strSql = "Delete From NPMEMO WHERE NM05='" & Left(strCU01, 6) & "' AND NM04 IS NULL "
          Pub_SeekTbLog strSql, , , , , strCU01
          cnnConnection.Execute strSql
       End If
      '核准函輸入備註(ApprovalMemo2)
      If ChkExistSpec("APPROVALMEMO2", strCU01, 8) = True Then
          strSql = "Delete From ApprovalMemo2 WHERE AM05='" & Left(strCU01, 8) & "' AND AM04 IS NULL "
          Pub_SeekTbLog strSql, , , , , strCU01
          cnnConnection.Execute strSql
       End If
       If ChkExistSpec("APPROVALMEMO2", strCU01, 6) = True Then
          strSql = "Delete From ApprovalMemo2 WHERE AM05='" & Left(strCU01, 6) & "' AND AM04 IS NULL "
          Pub_SeekTbLog strSql, , , , , strCU01
          cnnConnection.Execute strSql
       End If
       '核駁及審查意見通知函備註(IncomMemo)
       If ChkExistSpec("INCOMMEMO", strCU01, 8) = True Then
           strSql = "Delete From IncomMemo WHERE IM05='" & Left(strCU01, 8) & "' AND IM04 IS NULL "
           Pub_SeekTbLog strSql, , , , , strCU01
           cnnConnection.Execute strSql
       End If
       If ChkExistSpec("INCOMMEMO", strCU01, 6) = True Then
          strSql = "Delete From IncomMemo WHERE IM05='" & Left(strCU01, 6) & "' AND IM04 IS NULL "
          Pub_SeekTbLog strSql, , , , , strCU01
          cnnConnection.Execute strSql
       End If
       '請款函預設備註維護檔(DebitNotePS)
       If ChkExistSpec("DEBITNOTEPS", strCU01, 8) = True Then
           strSql = "Delete From DEBITNOTEPS WHERE DNPS05='" & Left(strCU01, 8) & "' AND DNPS04 IS NULL "
           Pub_SeekTbLog strSql, , , , , strCU01
           cnnConnection.Execute strSql
       End If
       If ChkExistSpec("DEBITNOTEPS", strCU01, 6) = True Then
           strSql = "Delete From DEBITNOTEPS WHERE DNPS05='" & Left(strCU01, 6) & "' AND DNPS04 IS NULL "
           Pub_SeekTbLog strSql, , , , , strCU01
           cnnConnection.Execute strSql
       End If
       'FCP承辦單設定維護(FcpEMPbill)
       If ChkExistSpec("FCPEMPBILL", strCU01, 8) = True Then
           strSql = "Delete From FcpEMPbill WHERE FEB05='" & Left(strCU01, 8) & "' AND FEB04 IS NULL "
           Pub_SeekTbLog strSql, , , , , strCU01
           cnnConnection.Execute strSql
       End If
       If ChkExistSpec("FCPEMPBILL", strCU01, 6) = True Then
           strSql = "Delete From FcpEMPbill WHERE FEB05='" & Left(strCU01, 6) & "' AND FEB04 IS NULL "
           Pub_SeekTbLog strSql, , , , , strCU01
           cnnConnection.Execute strSql
       End If
       '通知告准加註(ApprovalPS)
       If ChkExistSpec("APPROVALPS", strCU01, 8) = True Then
           strSql = "Delete From APPROVALPS WHERE APS05='" & Left(strCU01, 8) & "' AND APS04 IS NULL "
           Pub_SeekTbLog strSql, , , , , strCU01
           cnnConnection.Execute strSql
       End If
       If ChkExistSpec("APPROVALPS", strCU01, 6) = True Then
           strSql = "Delete From APPROVALPS WHERE APS05='" & Left(strCU01, 6) & "' AND APS04 IS NULL "
           Pub_SeekTbLog strSql, , , , , strCU01
           cnnConnection.Execute strSql
       End If
       'end 2023/01/03
   End If 'Added by Lydia 2021/10/27
    
   'Added by Lydia 2022/03/28 一併刪除DHL輸入資料
   strSql = "delete from dhl_input_data where did01=" & CNULL(strCU01) & " and did02=" & CNULL(strCU02)
   Pub_SeekTbLog strSql, , , , , strCU01  'add by sonia 2025/7/24
   cnnConnection.Execute strSql
   'end 2022/03/28
   
   'Added by Lydia 2023/11/03 刪除活化客戶
   If textCU02 = "0" Then
      strSql = "delete from oldcustomer where ocu01=" & CNULL(strCU01)
      Pub_SeekTbLog strSql, , , , , strCU01  'add by sonia 2025/7/24
      cnnConnection.Execute strSql
   End If
   'end 2023/11/03
   'end 2025/09/19
    
   '93.10.7 ADD BY SONIA
   If textCU03 <> "" Then
      'edit by nickc 2008/03/12
      'strSQL = "UPDATE fagent SET fa03=NULL WHERE fa01='" & textCU03 & "'"
      'Memo by Amy 2025/09/15 此句不管有沒有執行都寫log-秀玲
      strSql = "UPDATE fagent SET fa03=NULL,fa76='" & m_fa76 & "' WHERE fa01='" & textCU03 & "'"
      'Modify by Amy 2025/09/15 +strCU01 & strCU02 ,log記錄此客戶編號
      'Modify by Amy 2025/09/18 改只記錄8碼-秀玲
      Pub_SeekTbLog strSql, , , , , strCU01   'add by nickc 2006/12/20
      cnnConnection.Execute strSql
   End If
   '93.10.7 END
    
   ' 只有刪除的是最後一筆才須重新取的第一筆及最後一筆的本所案號
   If (strCU01 = m_LastKEY(0) And strCU02 = m_LastKEY(1)) Or (strCU01 = m_FirstKEY(0) And strCU02 = m_FirstKEY(1)) Then
      RefreshRange
   End If
   ShowCurrRecord strCU01, strCU02
   DelRecord = True
   cnnConnection.CommitTrans
   
   Exit Function
ErrHand:
    cnnConnection.RollbackTrans
    MsgBox "刪除失敗！" & vbCrLf & Err.Description
End Function

' 查詢記錄
Private Function QueryRecord() As Boolean
Dim strCU01 As String
Dim strCU02 As String

   QueryRecord = False
   strCU01 = textCU01 & String(8 - Len(textCU01), "0")
   strCU02 = textCU02 & String(1 - Len(textCU02), "0")
   'add by nickc 2006/03/17
   textCUID = ""
   If IsRecordExist(strCU01, strCU02) = True Then
      m_CurrKEY(0) = strCU01
      m_CurrKEY(1) = strCU02
      QueryRecord = True
      UpdateCtrlData
   Else
      QueryRecord = False
   End If

   UpdateToolbarState
End Function

' 使用者按下確定的按紐
Private Function OnWork() As Boolean
Dim strMsg As String, strTit As String
Dim nResponse
Dim bolChk2nd As Boolean 'Added by Lydia 2019/05/27 是否修改FCP是否核對已准專利=N
Dim stTO As String 'Add by Amy 2023/09/01
   
   OnWork = False
   Select Case m_EditMode
      Case 1: '新增
         If CheckDataValid() = True Then
            'Add By Cheng 2002/05/22
            '重新檢查欄位有效性
            If TxtValidate = False Then Exit Function
'            'Mark by Amy 2025/05/19 避免有訊息未回完(ex:統一編號重覆),就已產生編號,故自動取號往下搬
'            'Add by Amy 2023/07/06 自動取號及檢查由CheckDataValid拆出,避免電子收文已給號後,因錯誤跳離開,但自動取號已跳號
'            If GetChkAutoNo = False Then Exit Function
            
            'Add By Sindy 2012/2/8
            '新增客戶檔時,客戶國籍<="010",若公司負責人欄有值且代表人１（中）為空白時,同時將公司負責人欄存入代表人１（中）欄
            If Trim(textCU10) <= "010" Then
               If Trim(textCU07) > "" And Trim(textCU39) = "" Then
                  textCU39 = Trim(textCU07)
               End If
            End If
            
            If ChkCuPerson() = False Then Exit Function 'Added by Lydia 2019/04/16
            'Modify by Amy 2023/05/03 智權人員與客戶可收文所別不同,且跨所同意主管未輸不可存檔
            'Call ChkPZD07 'Add by Amy 2020/08/04 判斷智權人員與客戶可收文所別不同時彈訊息
            'Modify by Amy bug-原:ChkPZD07(True),導致判斷到修改的程式
            If ChkPZD07(False) = False Then
                Exit Function
            End If
            'Modify by Amy 2023/06/06 原sub 只彈訊息,改為彈訊息選是否
            If ChkShowMsg = False Then 'Add by Amy 2023/05/03
               Exit Function
            End If
            
            'Add by Amy 2025/05/19 程式從上面搬下來,避免有訊息未回完就已產生編號
            If GetChkAutoNo = False Then Exit Function
            
            ' 將所有欄位的內容更新到欄位串列中的欄位內容項目
            UpdateFieldNewData
            If AddRecord = True Then
                RefreshRange
                'Add by Amy 2023/09/01 身份證/統編 重覆發信
                If strIDRepeat <> MsgText(601) Then
                    stTO = Pub_GetSpecMan("程式管理人員")
                    PUB_SendMail strUserNum, stTO, "", Me.textCU01.Text & Me.textCU02.Text & " 新客戶身份證字號/統一編號重覆通知！", strIDRepeat
                End If
            Else
                Exit Function
            End If
         Else
            GoTo EXITSUB
         End If
      Case 2: '修改
         If CheckDataValid() = True Then
            'Add By Cheng 2002/05/22
            '重新檢查欄位有效性
            If TxtValidate = False Then Exit Function
            If ChkCuPerson() = False Then Exit Function 'Added by Lydia 2019/04/16
            'Modify by Amy 2023/06/06 改檢查[地址]有修改才彈訊息,舊資料因「中文地址郵遞區號」為空，且是跨所又沒跨所同意主管資料,會無法存(目前修改中文地址郵遞區號為空需補資料)
            'Add by Amy 2023/05/03 郵遞區號有修改,若改為跨所需彈訊息
            'If Left(textCU30.Text, 3) <> Left(m_FieldList(29).fiOldData, 3) Or Left(textCU112.Text, 3) <> Left(m_FieldList(111).fiOldData, 3) Then
            If textCU31.Text <> m_FieldList(30).fiOldData Or textCU23.Text <> m_FieldList(22).fiOldData Then
                If ChkPZD07(True) = False Then
                    Exit Function
                End If
            End If
            'Modify by Amy 2023/06/06 原sub 只彈訊息,改為彈訊息選是否
            If ChkShowMsg = False Then 'Add by Amy 2023/05/03
               Exit Function
            End If
            
            ' 將所有欄位的內容更新到欄位串列中的欄位內容項目
            UpdateFieldNewData
            'Add by Amy 2015/08/24 +檢查客戶狀態未改過且其他資料有改過彈訊息
            If ChkDataNotSave = True Then Exit Function
            'Add by Amy 2023/09/01 身份證/統編 重覆發信
            If strIDRepeat <> MsgText(601) Then
               stTO = Pub_GetSpecMan("程式管理人員")
               PUB_SendMail strUserNum, stTO, "", Me.textCU01.Text & Me.textCU02.Text & " 修改客戶身份證字號/統一編號重覆通知！", strIDRepeat
            End If
            'edit by nickc 2007/04/17 往下搬
            'If ModRecord = False Then Exit Function
            Dim strVer As String
            If (textCU04.Tag <> "") And (textCU04.Tag <> CheckStr(textCU04)) Then
               strVer = strVer & "中：" & textCU04.Tag & " --> " & textCU04 & vbCrLf
            End If
            If (textCU05.Tag & textCU88.Tag & textCU89.Tag & textCU90.Tag <> "") And (textCU05.Tag & textCU88.Tag & textCU89.Tag & textCU90.Tag <> textCU05 & textCU88 & textCU89 & textCU90) Then
               strVer = strVer & "英：" & textCU05.Tag & " " & textCU88.Tag & " " & textCU89.Tag & " " & textCU90.Tag & " --> " & textCU05 & " " & textCU88 & " " & textCU89 & " " & textCU90 & vbCrLf
               'Add by Morgan 2006/1/10
               '2011/12/26 cancel by sonia 因為ModRecord已有寫
               'PUB_AccDataCheck Me.textCU01.Text & Me.textCU02.Text, "英：" & textCU05.Tag & " " & textCU88.Tag & " " & textCU89.Tag & " " & textCU90.Tag & " --> " & textCU05 & " " & textCU88 & " " & textCU89 & " " & textCU90
            End If
            If (textCU06.Tag <> "") And (textCU06.Tag <> CheckStr(textCU06)) Then
               strVer = strVer & "日：" & textCU06.Tag & " --> " & textCU06 & vbCrLf
            End If
            'Added by Lydia 2019/05/27
            bolChk2nd = False
            If textCU122.Tag <> textCU122.Text And textCU122.Text = "N" Then
                bolChk2nd = True
            End If
            
            'add by nickc 2007/04/17 上面搬下來
            If ModRecord = False Then Exit Function
            If strVer <> "" Then
               'Modify by Amy 2023/09/01 原:83002
               stTO = Pub_GetSpecMan("程式管理人員")
               PUB_SendMail strUserNum, stTO, "", Me.textCU01.Text & Me.textCU02.Text, " 客戶名稱修改！", strVer
            End If
            'Added by Lydia 2019/05/27 設定"FCP是否核對已准專利"上" N"，則出"核對已准專利"未發文之清單
            If bolChk2nd = True Then
                If Pub_GetFA85CU122List(textCU01 & textCU02) = True Then
                End If
            End If
         Else
            GoTo EXITSUB
         End If
      Case 3: '刪除
         ' 將所有欄位的內容更新到欄位串列中的欄位內容項目
         UpdateFieldNewData
         If DelRecord = True Then
            RefreshRange
         Else
            Exit Function
         End If
      Case 4: '查詢
         ' 將所有欄位的內容更新到欄位串列中的欄位內容項目
         UpdateFieldNewData
         'If CheckDataValid() = True Then
         If textCU01 <> "" Then
            If QueryRecord = False Then
               strMsg = "無此資料"
               strTit = "查詢資料"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               UpdateCtrlData
            End If
         Else
            GoTo EXITSUB
         End If
   End Select
   
   If m_EditMode <> 4 Then PUB_SendMailCache 'Added by Morgan 2021/11/29
   
   m_EditMode = 0
   SetCtrlReadOnly True
   OnWork = True
EXITSUB:
End Function

' 開始輸入資料
Private Sub SetInputEntry()
   Select Case m_EditMode
      Case 1: textCU01.SetFocus
      Case 2:
         'Modify by Amy 2024/01/22 +if 國外潛在客戶維護轉號存檔切至此畫面欄位鎖住會錯
         If textCU03.Enabled = False Then
            textCU03.SetFocus
         End If
         'end 2024/01/22
      Case 4: If Me.Visible = True Then textCU01.SetFocus
   End Select
End Sub

' 檢查記錄是否已經存在
Private Function IsRecordExist(ByVal strKEY01 As String, ByVal strKEY02 As String) As Boolean
Dim rsTmp As New ADODB.Recordset
Dim strSql As String
   
   IsRecordExist = False
   strSql = "SELECT * FROM customer " & _
            "WHERE cu01 = '" & strKEY01 & "' AND " & _
                  "cu02 = '" & strKEY02 & "' "
                  
   ' 讀取資料庫
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   ' 檢查讀取的資料筆數
   If rsTmp.RecordCount > 0 Then
      IsRecordExist = True
   Else
      IsRecordExist = False
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Function

' 顯示資料
Private Sub ShowCurrRecord(ByVal strKEY01 As String, ByVal strKEY02 As String)
Dim strSql As String
Dim rsTmp As New ADODB.Recordset

   If IsRecordExist(strKEY01, strKEY02) = True Then
      m_CurrKEY(0) = strKEY01
      m_CurrKEY(1) = strKEY02
   Else
      strSql = "SELECT CU01,CU02 FROM customer " & _
               "WHERE CU01 = '" & m_CurrKEY(0) & "' AND " & _
                     "CU02 = (SELECT MIN(CU02) FROM customer " & _
                             "WHERE CU01 = '" & m_CurrKEY(0) & "' AND " & _
                                   "CU02 > '" & m_CurrKEY(1) & "' )"
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         If IsNull(rsTmp.Fields("CU01")) = False Then: m_CurrKEY(0) = rsTmp.Fields("CU01")
         If IsNull(rsTmp.Fields("CU02")) = False Then: m_CurrKEY(1) = rsTmp.Fields("CU02")
         rsTmp.Close
         UpdateCtrlData
         GoTo EXITSUB
      End If
      rsTmp.Close
      
      strSql = "SELECT CU01,CU02 FROM customer " & _
               "WHERE CU01 = (SELECT MIN(CU01) FROM customer " & _
                              "WHERE CU01 > '" & m_CurrKEY(0) & "') AND " & _
                     "CU02 = (SELECT MIN(CU02) FROM customer " & _
                              "WHERE CU01 = (SELECT MIN(CU01) FROM customer " & _
                                             "WHERE CU01 > '" & m_CurrKEY(0) & "')) "
   
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         If IsNull(rsTmp.Fields("CU01")) = False Then: m_CurrKEY(0) = rsTmp.Fields("CU01")
         If IsNull(rsTmp.Fields("CU02")) = False Then: m_CurrKEY(1) = rsTmp.Fields("CU02")
      Else
         ShowLastRecord
         GoTo EXITSUB
      End If
      rsTmp.Close
   End If
   UpdateCtrlData
EXITSUB:
End Sub

' 顯示第一筆資料
Private Sub ShowFirstRecord()
   m_CurrKEY(0) = m_FirstKEY(0)
   m_CurrKEY(1) = m_FirstKEY(1)
   
   UpdateCtrlData
End Sub

' 顯示上一筆資料
Private Sub ShowPrevRecord()
Dim strSql As String
Dim rsTmp As New ADODB.Recordset

   If m_CurrKEY(0) = m_FirstKEY(0) And m_CurrKEY(1) = m_FirstKEY(1) Then
      ShowMsg MsgText(9008)
      GoTo EXITSUB
   End If
   
   strSql = "SELECT CU01,CU02 FROM customer " & _
            "WHERE CU01 = '" & m_CurrKEY(0) & "' AND " & _
                  "CU02 = (SELECT MAX(CU02) FROM customer " & _
                          "WHERE CU01 = '" & m_CurrKEY(0) & "' AND " & _
                                "CU02 < '" & m_CurrKEY(1) & "' )"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("CU01")) = False Then: m_CurrKEY(0) = rsTmp.Fields("CU01")
      If IsNull(rsTmp.Fields("CU02")) = False Then: m_CurrKEY(1) = rsTmp.Fields("CU02")
      rsTmp.Close
      UpdateCtrlData
      GoTo EXITSUB
   End If
   rsTmp.Close
   
   strSql = "SELECT CU01,CU02 FROM customer " & _
            "WHERE CU01 = (SELECT MAX(CU01) FROM customer " & _
                           "WHERE CU01 < '" & m_CurrKEY(0) & "') AND " & _
                  "CU02 = (SELECT MAX(CU02) FROM customer " & _
                           "WHERE CU01 = (SELECT MAX(CU01) FROM customer " & _
                                          "WHERE CU01 < '" & m_CurrKEY(0) & "')) "

   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("CU01")) = False Then: m_CurrKEY(0) = rsTmp.Fields("CU01")
      If IsNull(rsTmp.Fields("CU02")) = False Then: m_CurrKEY(1) = rsTmp.Fields("CU02")
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
   
   If m_CurrKEY(0) = m_LastKEY(0) And m_CurrKEY(1) = m_LastKEY(1) Then
      ShowMsg MsgText(9009)
      GoTo EXITSUB
   End If
   
   strSql = "SELECT CU01,CU02 FROM customer " & _
            "WHERE CU01 = '" & m_CurrKEY(0) & "' AND " & _
                  "CU02 = (SELECT MIN(CU02) FROM customer " & _
                          "WHERE CU01 = '" & m_CurrKEY(0) & "' AND " & _
                                "CU02 > '" & m_CurrKEY(1) & "' )"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("CU01")) = False Then: m_CurrKEY(0) = rsTmp.Fields("CU01")
      If IsNull(rsTmp.Fields("CU02")) = False Then: m_CurrKEY(1) = rsTmp.Fields("CU02")
      rsTmp.Close
      UpdateCtrlData
      GoTo EXITSUB
   End If
   rsTmp.Close
   
   strSql = "SELECT CU01,CU02 FROM customer " & _
            "WHERE CU01 = (SELECT MIN(CU01) FROM customer " & _
                           "WHERE Cu01 > '" & m_CurrKEY(0) & "') AND " & _
                  "CU02 = (SELECT MIN(CU02) FROM customer " & _
                           "WHERE CU01 = (SELECT MIN(CU01) FROM customer " & _
                                          "WHERE CU01 > '" & m_CurrKEY(0) & "')) "

   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("CU01")) = False Then: m_CurrKEY(0) = rsTmp.Fields("CU01")
      If IsNull(rsTmp.Fields("CU02")) = False Then: m_CurrKEY(1) = rsTmp.Fields("CU02")
   End If
   rsTmp.Close
   
   UpdateCtrlData
   
EXITSUB:
   Set rsTmp = Nothing
End Sub

' 顯示最後一筆資料
Private Sub ShowLastRecord()
   m_CurrKEY(0) = m_LastKEY(0)
   m_CurrKEY(1) = m_LastKEY(1)
   
   UpdateCtrlData
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
         Me.Caption = Replace(Me.Caption, "(修改)", "") & "(新增)" 'Add by Amy 2022/09/16
         m_EditMode = 1
         ClearField
         SetCtrlReadOnly False
         UpdateToolbarState
         textCU111.Locked = False 'Add by Amy 2013/11/04
         textCU111.Enabled = True
         If Left(m_PrevNo, 3) <> "Add" Then SetInputEntry 'Add by Amy 2022/09/14
      ' 修改
      Case vbKeyF3:
         Me.Caption = Replace(Me.Caption, "(新增)", "") & "(修改)" 'Add by Amy 2022/09/16
         
         'add by nickc 2007/06/05
         UpdateCtrlData
         
         m_EditMode = 2
         SetCtrlReadOnly False
         SetKeyReadOnly True
         UpdateToolbarState
        '92.8.28 ADD BY SONIA 分所人員不可修改智權人員及業務區
        StrSQLa = "SELECT ST06 FROM STAFF WHERE ST01 = '" & Trim(strUserNum) & "'"
        Set rsA = New ADODB.Recordset
        rsA.CursorLocation = adUseClient
        rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
        If rsA.RecordCount > 0 Then
           If rsA.Fields("ST06").Value <> "1" Then
              textCU12.Enabled = False
              textCU13.Enabled = False
           Else
              textCU12.Enabled = True
              textCU13.Enabled = True
           End If
        Else
           textCU12.Enabled = False
           textCU13.Enabled = False
        End If
        '92.8.28 END
        'Add by Amy 2013/10/29 CU142為B.宣告破產 則不可修改呆帳記錄
        strExc(0) = "Select CU142 From Customer Where CU01 = '" & textCU01 & "' AND " & _
                         "CU02 = '" & textCU02 & "' And CU142='B' "
        intI = 1
        Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
        If intI = 1 Then
            textCU111.Locked = True
            textCU111.Enabled = False
        Else
            textCU111.Locked = False
            textCU111.Enabled = True
        End If
        'end 2013/10/29
         SetInputEntry
      ' 刪除
      Case vbKeyF5:
         strTit = "詢問"
         strMsg = "是否要刪除此筆資料?"
         nResponse = MsgBox(strMsg, vbYesNo + vbCritical + vbDefaultButton2, strTit)
         If nResponse = vbYes Then
            'Add by Amy 2018/07/18 有往來記錄不可刪除
            strExc(0) = "Select CR03 From ContactRecord Where CR03 = '" & textCU01 & textCU02 & "'  " & _
                    "Union Select COR03 From ContactRecord1 Where COR03 = '" & textCU01 & textCU02 & "'   "
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
                MsgBox "有往來記錄不可刪除！"
                Exit Sub
            End If
            'end 2018/07/18
            'Add by Amy 2022/12/05 若存在XYS02介紹來源編號,則不可刪
            'Modify by Amy 2024/11/29 考慮多筆,改訊息至共用
            If textCU02 = "0" And Pub_GetXYSource(2, textCU01, , , , Me.Name, strMsg) = True Then
                MsgBox strMsg, vbOKOnly, "注意"
                Exit Sub
            End If
            'Add by Amy 2022/09/29 商申承辦人責任業務區分配人員確認
            If textCU02 = "0" Then
               If ChkDutyZoneAssign(Me.Name, textCU01, True, True) = True Then
                  Exit Sub
               End If
            End If
         
            m_fa76 = ""
            '2008/7/17 MODIFY BY SONIA
            'Do While Not m_fa76 = "A" And Not m_fa76 = "C" And textCU03 <> ""
            '    m_fa76 = UCase(InputBox("取消代理人編號與客戶編號的關連，代理人的性質應為??(A、C)"))
            old_fa76 = PUB_GetFAgentFA76(textCU03 & String(9 - Len(textCU03), "0"))
            Do While Not m_fa76 = "A" And Not m_fa76 = "B" And Not m_fa76 = "C" And textCU03 <> ""
                m_fa76 = UCase(InputBox("取消代理人編號與客戶編號的關連，代理人的性質應為??(A、B、C)，原代理人性質為" & old_fa76))
            Loop
            m_EditMode = 3
            If OnWork = True Then
                UpdateToolbarState
            Else
                Exit Sub
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
         'Modify By Sindy 2014/8/29 Mark
         'PUB_FilterFormText Me 'Add by Morgan 2008/6/20 修正畫面所有含跳行符號的文字框
         '2014/8/29 END
         'Modify By Sindy 2014/12/27 +不過濾的文字框.name
         'Modified by Lydia 2017/03/31 +,textCU78,textCU150
         PUB_FilterFormText Me, "textCU79,textCU78,textCU150"
         '2014/12/27 END
         If OnWork = True Then
            UpdateToolbarState
            Me.Caption = "客戶基本資料維護" 'Add by Amy 2022/09/16
         Else
            Exit Sub
         End If
      ' 取消
      Case vbKeyF10:
         Select Case m_EditMode
            Case 1, 2:
               strTit = "詢問"
               strMsg = "你並未存檔, 確定離開嗎?"
               nResponse = MsgBox(strMsg, vbYesNo + vbCritical + vbDefaultButton2, strTit)
               If nResponse = vbYes Then
                  Me.Caption = "客戶基本資料維護" 'Add by Amy 2022/09/16
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
   If KeyCode <> vbKeyEscape And KeyCode <> vbKeyF3 Then
      'tabCustomer.Tab = 0
   End If
End Sub

Private Sub RefreshRange()
Dim strSql As String
Dim rsTmp As New ADODB.Recordset

   strSql = "SELECT cu01,cu02 FROM customer " & _
            "WHERE cu01 = (SELECT MIN(cu01) FROM customer) AND " & _
                  "cu02 = (SELECT MIN(cu02) FROM customer " & _
                           "WHERE cu01 = (SELECT MIN(cu01) FROM customer)) "
   
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("cu01")) = False Then: m_FirstKEY(0) = rsTmp.Fields("cu01")
      If IsNull(rsTmp.Fields("cu02")) = False Then: m_FirstKEY(1) = rsTmp.Fields("cu02")
   End If
   rsTmp.Close

   strSql = "SELECT cu01,cu02 FROM customer " & _
            "WHERE cu01 = (SELECT MAX(cu01) FROM customer) AND " & _
                  "cu02 = (SELECT MAX(cu02) FROM customer " & _
                           "WHERE cu01 = (SELECT MAX(cu01) FROM customer)) "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("cu01")) = False Then: m_LastKEY(0) = rsTmp.Fields("cu01")
      If IsNull(rsTmp.Fields("cu02")) = False Then: m_LastKEY(1) = rsTmp.Fields("cu02")
   End If
   rsTmp.Close
   
   Set rsTmp = Nothing
End Sub

' 將資料庫中的資料更新到所有欄位中
Private Sub UpdateCtrlData()
Dim rsTmp As New ADODB.Recordset
Dim strSql As String
Dim nIndex As Integer
Dim m_lbl As Object 'Add by Amy 2015/09/09
Dim arrID 'Add By Sindy 2025/1/6
   
'add by nickc 2007/09/21
m_CU01 = m_CurrKEY(0)
m_CU02 = m_CurrKEY(1)
bolMsg = False 'Added by Lydia 2019/04/17

   strSql = "SELECT * FROM customer " & _
            "WHERE cu01 = '" & m_CurrKEY(0) & "' AND " & _
                  "cu02 = '" & m_CurrKEY(1) & "' "
                  
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      ClearField
      If IsNull(rsTmp.Fields("CU01")) = False Then: textCU01 = rsTmp.Fields("CU01")
      If IsNull(rsTmp.Fields("CU02")) = False Then: textCU02 = rsTmp.Fields("CU02")
      If IsNull(rsTmp.Fields("CU03")) = False Then: textCU03 = rsTmp.Fields("CU03")
      If IsNull(rsTmp.Fields("CU04")) = False Then: textCU04 = rsTmp.Fields("CU04")
      If IsNull(rsTmp.Fields("CU05")) = False Then: textCU05 = rsTmp.Fields("CU05")
      If IsNull(rsTmp.Fields("CU06")) = False Then: textCU06 = rsTmp.Fields("CU06")
      m_CU07 = "" 'Add By Sindy 2013/7/2
      If IsNull(rsTmp.Fields("CU07")) = False Then: textCU07 = rsTmp.Fields("CU07"): m_CU07 = rsTmp.Fields("CU07")
      If IsNull(rsTmp.Fields("CU09")) = False Then: textCU09 = rsTmp.Fields("CU09")
      If IsNull(rsTmp.Fields("CU10")) = False Then: textCU10 = rsTmp.Fields("CU10")
      If IsNull(rsTmp.Fields("CU11")) = False Then: textCU11 = rsTmp.Fields("CU11")
      If IsNull(rsTmp.Fields("CU12")) = False Then: textCU12 = rsTmp.Fields("CU12")
      If IsNull(rsTmp.Fields("CU13")) = False Then: textCU13 = rsTmp.Fields("CU13")
      'add by nickc 2007/09/11 紀錄原先智權人員
      textCU13.Tag = textCU13.Text
      textCU12.Tag = textCU12.Text   '2012/8/7 ADD BY SONIA 記錄原業務區
      ' 開發日期
      If IsNull(rsTmp.Fields("CU14")) = False Then
         If rsTmp.Fields("CU14") <> "0" Then
            textCU14 = TAIWANDATE(rsTmp.Fields("CU14"))
         Else
            textCU14 = ""
         End If
      Else
         textCU14 = ""
      End If
      m_CU15 = "" 'Add By Sindy 2013/12/12
      If IsNull(rsTmp.Fields("CU15")) = False Then
         textCU15 = rsTmp.Fields("CU15")
         m_CU15 = rsTmp.Fields("CU15") 'Add By Sindy 2013/12/12
      End If
      'Add By Sindy 2012/7/12 清欄位值
      For nIndex = 0 To 3
         optCustomer(nIndex).Value = False
         If nIndex = 0 Then optCustomer(nIndex).Tag = False 'Add by Amy 2015/10/20
      Next nIndex
      '2012/7/12 End
      If textCU15.Text = "0" Then
         optCustomer(0).Value = True
         optCustomer(0).Tag = True 'Add by Amy 2015/10/20
      'Modify By Sindy 2012/5/24
      ElseIf textCU15.Text = "1" Then
         optCustomer(1).Value = True
      ElseIf textCU15.Text = "2" Then
         optCustomer(2).Value = True
      ElseIf textCU15.Text = "3" Then 'Modify By Sindy 2012/7/12
         optCustomer(3).Value = True
      End If
      '2012/5/24 End
      If IsNull(rsTmp.Fields("CU16")) = False Then: textCU16 = rsTmp.Fields("CU16")
      If IsNull(rsTmp.Fields("CU17")) = False Then: textCU17 = rsTmp.Fields("CU17")
      If IsNull(rsTmp.Fields("CU18")) = False Then: textCU18 = rsTmp.Fields("CU18")
      If IsNull(rsTmp.Fields("CU19")) = False Then: textCU19 = rsTmp.Fields("CU19")
      If IsNull(rsTmp.Fields("CU20")) = False Then: textCU20 = rsTmp.Fields("CU20")
      If IsNull(rsTmp.Fields("CU21")) = False Then: textCU21 = rsTmp.Fields("CU21")
      If IsNull(rsTmp.Fields("CU22")) = False Then: textCU22 = rsTmp.Fields("CU22")
      If IsNull(rsTmp.Fields("CU70")) = False Then: textCU70 = rsTmp.Fields("CU70")
      'Modify by Amy 2025/06/30 +textCU10,X90619 客戶國籍大陸,中文地址為中國浙江省台州市溫嶺市大溪區高速公路道口一級公路北側-->無法改成大溪鎮 (ReplaceAddr函數會取代)
      If IsNull(rsTmp.Fields("CU23")) = False Then: textCU23 = ReplaceAddrTW(rsTmp.Fields("CU23"), , textCU10) 'Modify by Amy 2015/08/24 +取代臺灣地址
      If IsNull(rsTmp.Fields("CU24")) = False Then: textCU24 = rsTmp.Fields("CU24")
      If IsNull(rsTmp.Fields("CU25")) = False Then: textCU25 = rsTmp.Fields("CU25")
      If IsNull(rsTmp.Fields("CU26")) = False Then: textCU26 = rsTmp.Fields("CU26")
      If IsNull(rsTmp.Fields("CU27")) = False Then: textCU27 = rsTmp.Fields("CU27")
      If IsNull(rsTmp.Fields("CU28")) = False Then: textCU28 = rsTmp.Fields("CU28")
      If IsNull(rsTmp.Fields("CU29")) = False Then: textCU29 = rsTmp.Fields("CU29")
      If IsNull(rsTmp.Fields("CU87")) = False Then: textCU87 = rsTmp.Fields("CU87") 'Modify by Amy 2025/06/30 地址國籍,從下面搬上來
      If IsNull(rsTmp.Fields("CU30")) = False Then: textCU30 = rsTmp.Fields("CU30")
      'Modify by Amy 2025/06/30 +textCU87,X90619 地址國籍大陸,聯絡地址為中國浙江省台州市溫嶺市大溪區高速公路道口一級公路北側-->無法改成大溪鎮
      If IsNull(rsTmp.Fields("CU31")) = False Then: textCU31 = ReplaceAddrTW(rsTmp.Fields("CU31"), , textCU87) 'Modify by Amy 2015/08/24 +取代臺灣地址
      If IsNull(rsTmp.Fields("CU32")) = False Then: textCU32 = rsTmp.Fields("CU32")
      '2008/9/4 add by sonia
      If textCU31 <> "" Then
         m_CU31 = textCU31
      Else
         m_CU31 = ""
      End If
      '2008/9/4 end
      SeekOldCu32 = textCU32
      If IsNull(rsTmp.Fields("CU33")) = False Then: textCU33 = rsTmp.Fields("CU33")
      If IsNull(rsTmp.Fields("CU34")) = False Then: textCU34 = rsTmp.Fields("CU34")
      If IsNull(rsTmp.Fields("CU35")) = False Then: textCU35 = rsTmp.Fields("CU35")
      If IsNull(rsTmp.Fields("CU36")) = False Then: textCU36 = rsTmp.Fields("CU36")
      If IsNull(rsTmp.Fields("CU37")) = False Then: textCU37 = rsTmp.Fields("CU37")
      ' 全部折扣起始日
      If IsNull(rsTmp.Fields("CU38")) = False Then
         If rsTmp.Fields("CU38") <> "0" Then
            textCU38 = TAIWANDATE(rsTmp.Fields("CU38"))
         Else
            textCU38 = ""
         End If
      Else
         textCU38 = ""
      End If
      If IsNull(rsTmp.Fields("CU39")) = False Then: textCU39 = rsTmp.Fields("CU39")
      If IsNull(rsTmp.Fields("CU40")) = False Then: textCU40 = rsTmp.Fields("CU40")
      If IsNull(rsTmp.Fields("CU41")) = False Then: textCU41 = rsTmp.Fields("CU41")
      If IsNull(rsTmp.Fields("CU42")) = False Then: textCU42 = rsTmp.Fields("CU42")
      If IsNull(rsTmp.Fields("CU43")) = False Then: textCU43 = rsTmp.Fields("CU43")
      If IsNull(rsTmp.Fields("CU44")) = False Then: textCU44 = rsTmp.Fields("CU44")
      If IsNull(rsTmp.Fields("CU45")) = False Then: textCU45 = rsTmp.Fields("CU45")
      If IsNull(rsTmp.Fields("CU46")) = False Then: textCU46 = rsTmp.Fields("CU46")
      If IsNull(rsTmp.Fields("CU47")) = False Then: textCU47 = rsTmp.Fields("CU47")
      If IsNull(rsTmp.Fields("CU48")) = False Then: textCU48 = rsTmp.Fields("CU48")
      If IsNull(rsTmp.Fields("CU49")) = False Then: textCU49 = rsTmp.Fields("CU49")
      If IsNull(rsTmp.Fields("CU50")) = False Then: textCU50 = rsTmp.Fields("CU50")
      If IsNull(rsTmp.Fields("CU51")) = False Then: textCU51 = rsTmp.Fields("CU51")
      If IsNull(rsTmp.Fields("CU52")) = False Then: textCU52 = rsTmp.Fields("CU52")
      If IsNull(rsTmp.Fields("CU53")) = False Then: textCU53 = rsTmp.Fields("CU53")
      If IsNull(rsTmp.Fields("CU54")) = False Then: textCU54 = rsTmp.Fields("CU54")
      If IsNull(rsTmp.Fields("CU55")) = False Then: textCU55 = rsTmp.Fields("CU55")
      If IsNull(rsTmp.Fields("CU56")) = False Then: textCU56 = rsTmp.Fields("CU56")
      If IsNull(rsTmp.Fields("CU57")) = False Then: textCU57 = rsTmp.Fields("CU57")
      If IsNull(rsTmp.Fields("CU58")) = False Then: textCU58 = rsTmp.Fields("CU58")
      If IsNull(rsTmp.Fields("CU59")) = False Then: textCU59 = rsTmp.Fields("CU59")
      If IsNull(rsTmp.Fields("CU60")) = False Then: textCU60 = rsTmp.Fields("CU60")
      If IsNull(rsTmp.Fields("CU61")) = False Then: textCU61 = rsTmp.Fields("CU61")
      If IsNull(rsTmp.Fields("CU62")) = False Then: textCU62 = rsTmp.Fields("CU62")
      If IsNull(rsTmp.Fields("CU63")) = False Then: textCU63 = rsTmp.Fields("CU63")
      If IsNull(rsTmp.Fields("CU64")) = False Then: textCU64 = rsTmp.Fields("CU64")
      textCU64.Tag = textCU64.Text 'Added by Morgan 2022/1/20
      If IsNull(rsTmp.Fields("CU65")) = False Then: textCU65 = rsTmp.Fields("CU65")
      If IsNull(rsTmp.Fields("CU66")) = False Then: textCU66 = rsTmp.Fields("CU66")
      If IsNull(rsTmp.Fields("CU67")) = False Then: textCU67 = rsTmp.Fields("CU67")
      If IsNull(rsTmp.Fields("CU68")) = False Then: textCU68 = rsTmp.Fields("CU68")
      If IsNull(rsTmp.Fields("CU69")) = False Then: textCU69 = rsTmp.Fields("CU69")
      If IsNull(rsTmp.Fields("CU71")) = False Then: textCU71 = rsTmp.Fields("CU71")
      If IsNull(rsTmp.Fields("CU72")) = False Then: textCU72 = rsTmp.Fields("CU72")
      If IsNull(rsTmp.Fields("CU73")) = False Then: textCU73 = rsTmp.Fields("CU73")
      If IsNull(rsTmp.Fields("CU74")) = False Then: textCU74 = rsTmp.Fields("CU74")
      textCU74.Tag = textCU74.Text 'Added by Lydia 2019/11/27
      If IsNull(rsTmp.Fields("CU75")) = False Then: textCU75 = rsTmp.Fields("CU75")
      'Modify By Sindy 2013/1/17
'      If IsNull(rsTmp.Fields("CU76")) = False Then: textCU76 = rsTmp.Fields("CU76")
      If IsNull(rsTmp.Fields("CU76")) = False Then
         For i = 0 To Combo2(0).ListCount - 1
            Combo2(0).ListIndex = i
            If InStr(Combo2(0).Text, rsTmp.Fields("CU76")) > 0 Then
               Exit For
            End If
         Next
      Else
         Combo2(0).ListIndex = 0
      End If
      '2013/1/17 End
      If IsNull(rsTmp.Fields("CU77")) = False Then: textCU77 = rsTmp.Fields("CU77")
      If IsNull(rsTmp.Fields("CU78")) = False Then: textCU78 = rsTmp.Fields("CU78")
      If IsNull(rsTmp.Fields("CU79")) = False Then: textCU79 = rsTmp.Fields("CU79")
      If IsNull(rsTmp.Fields("CU80")) = False Then: cboStatus = rsTmp.Fields("CU80")
      If IsNull(rsTmp.Fields("CU88")) = False Then: textCU88 = rsTmp.Fields("CU88")
      If IsNull(rsTmp.Fields("CU89")) = False Then: textCU89 = rsTmp.Fields("CU89")
      If IsNull(rsTmp.Fields("CU90")) = False Then: textCU90 = rsTmp.Fields("CU90")
      If IsNull(rsTmp.Fields("CU91")) = False Then: textCU91 = rsTmp.Fields("CU91")
      If IsNull(rsTmp.Fields("CU92")) = False Then: textCU92 = rsTmp.Fields("CU92")
      If IsNull(rsTmp.Fields("CU93")) = False Then: textCU93 = rsTmp.Fields("CU93")
      If IsNull(rsTmp.Fields("CU94")) = False Then: textCU94 = rsTmp.Fields("CU94")
      If IsNull(rsTmp.Fields("CU95")) = False Then: textCU95 = rsTmp.Fields("CU95")
      If IsNull(rsTmp.Fields("CU96")) = False Then: textCU96 = rsTmp.Fields("CU96")
      If IsNull(rsTmp.Fields("CU97")) = False Then: textCU97 = rsTmp.Fields("CU97")
      If IsNull(rsTmp.Fields("CU98")) = False Then: textCU98 = rsTmp.Fields("CU98")
      If IsNull(rsTmp.Fields("CU99")) = False Then: textCU99 = rsTmp.Fields("CU99")
      If IsNull(rsTmp.Fields("CU100")) = False Then: textCU100 = rsTmp.Fields("CU100")
      If IsNull(rsTmp.Fields("CU102")) = False Then: textCU102 = rsTmp.Fields("CU102")
      If IsNull(rsTmp.Fields("CU103")) = False Then: textCU103 = rsTmp.Fields("CU103")
      If IsNull(rsTmp.Fields("CU104")) = False Then: textCU104 = rsTmp.Fields("CU104")
      If IsNull(rsTmp.Fields("CU105")) = False Then: textCU105 = rsTmp.Fields("CU105")
      If IsNull(rsTmp.Fields("CU106")) = False Then: textCU106 = rsTmp.Fields("CU106")
      If IsNull(rsTmp.Fields("CU107")) = False Then: textCU107 = rsTmp.Fields("CU107")
      If IsNull(rsTmp.Fields("CU108")) = False Then: textCU108 = rsTmp.Fields("CU108")
      ' 全部折扣起始日
      If IsNull(rsTmp.Fields("CU109")) = False Then
         If rsTmp.Fields("CU109") <> "0" Then
            textCU109 = TAIWANDATE(rsTmp.Fields("CU109"))
         Else
            textCU109 = ""
         End If
      Else
         textCU109 = ""
      End If
      
      'Add By Sindy 2025/3/10
      If IsNull(rsTmp.Fields("CU203")) = False Then: textCU203 = rsTmp.Fields("CU203")
      If IsNull(rsTmp.Fields("CU204")) = False Then: textCU204 = rsTmp.Fields("CU204")
      ' 全部折扣起始日
      If IsNull(rsTmp.Fields("CU205")) = False Then
         If rsTmp.Fields("CU205") <> "0" Then
            textCU205 = TAIWANDATE(rsTmp.Fields("CU205"))
         Else
            textCU205 = ""
         End If
      Else
         textCU205 = ""
      End If
      '2025/3/10 END
      
      If IsNull(rsTmp.Fields("CU111")) = False Then: textCU111 = rsTmp.Fields("CU111")
      If textCU111 = "Y" Then textCU01.ForeColor = &HFF&: textCU02.ForeColor = &HFF& Else textCU01.ForeColor = &H80000008: textCU02.ForeColor = &H80000008
      If IsNull(rsTmp.Fields("CU112")) = False Then: textCU112 = rsTmp.Fields("CU112")
      If IsNull(rsTmp.Fields("CU113")) = False Then: textCU113 = rsTmp.Fields("CU113") 'Added by Morgan 2012/8/28
      If IsNull(rsTmp.Fields("CU114")) = False Then: textCU114 = rsTmp.Fields("CU114")
      If IsNull(rsTmp.Fields("CU115")) = False Then: textCU115 = rsTmp.Fields("CU115") 'Add by Morgan 2008/1/16
      If IsNull(rsTmp.Fields("CU116")) = False Then: textCU116 = rsTmp.Fields("CU116") 'Add by Morgan 2008/1/16
      If IsNull(rsTmp.Fields("CU117")) = False Then: textCU117 = rsTmp.Fields("CU117") 'Add by Morgan 2008/1/16
      If IsNull(rsTmp.Fields("CU118")) = False Then: textCU118 = rsTmp.Fields("CU118") 'Add by Morgan 2008/1/16
      If IsNull(rsTmp.Fields("CU176")) = False Then: textCU176 = rsTmp.Fields("CU176") 'Add by Morgan 2018/11/14
      If IsNull(rsTmp.Fields("CU185")) = False Then: textCU185 = rsTmp.Fields("CU185") 'Add by Morgan 2021/10/7
      'If IsNull(rsTmp.Fields("CU186")) = False Then: textCU186 = rsTmp.Fields("CU186") 'Add by Morgan 2021/10/7
      If IsNull(rsTmp.Fields("CU187")) = False Then: textCU187 = rsTmp.Fields("CU187") 'Add by Morgan 2021/10/7
      If IsNull(rsTmp.Fields("CU188")) = False Then: textCU188 = rsTmp.Fields("CU188") 'Add by Morgan 2021/10/7
      
      If IsNull(rsTmp.Fields("CU122")) = False Then: textCU122 = rsTmp.Fields("CU122") 'Add by Morgan 2007/10/26
      textCU122.Tag = textCU122.Text  'Added by Lydia 2019/05/27
      If IsNull(rsTmp.Fields("CU123")) = False Then: textCU123 = rsTmp.Fields("CU123") 'Add by Morgan 2008/1/7
      If IsNull(rsTmp.Fields("CU125")) = False Then: textCU125 = rsTmp.Fields("CU125") 'Add By Sindy 2009/10/26
      If IsNull(rsTmp.Fields("CU128")) = False Then: TextCu128 = rsTmp.Fields("CU128") 'Add by Toni 2008/10/21
      If IsNull(rsTmp.Fields("CU132")) = False Then: textCU132 = rsTmp.Fields("CU132") '2008/12/9 add by sonia
      If IsNull(rsTmp.Fields("CU145")) = False Then: textCU145 = rsTmp.Fields("CU145") 'Add By Sindy 2011/1/14
      'Add By Sindy 2011/3/4
      If IsNull(rsTmp.Fields("CU146")) = False Then: textCU146 = rsTmp.Fields("CU146")
      If IsNull(rsTmp.Fields("CU147")) = False Then: textCU147 = rsTmp.Fields("CU147")
      'Modify By Sindy 2013/1/17
'      If IsNull(rsTmp.Fields("CU148")) = False Then: textCU148 = rsTmp.Fields("CU148")
      If IsNull(rsTmp.Fields("CU148")) = False Then
         For i = 0 To Combo2(1).ListCount - 1
            Combo2(1).ListIndex = i
            If InStr(Combo2(1).Text, rsTmp.Fields("CU148")) > 0 Then
               Exit For
            End If
         Next
      Else
         Combo2(1).ListIndex = 0
      End If
      '2013/1/17 End
      If IsNull(rsTmp.Fields("CU139")) = False Then: textCU139 = rsTmp.Fields("CU139") 'Add By Sindy 2013/8/15
      'Modify by Amy 2015/09/10 改為label顯示
      'If IsNull(rsTmp.Fields("CU143")) = False Then: textCU143 = rsTmp.Fields("CU143") 'Add By Sindy 2013/11/19
      If IsNull(rsTmp.Fields("CU143")) = False Then: lblCU143 = rsTmp.Fields("CU143")
      'end 2015/09/10
      
      If IsNull(rsTmp.Fields("CU144")) = False Then: textCU144 = rsTmp.Fields("CU144") 'Add By Sindy 2013/12/17
      LblCU144 = ShowLblCU144(textCU144) 'Add By Sindy 2023/9/4
      
      If IsNull(rsTmp.Fields("CU149")) = False Then: textCU149 = rsTmp.Fields("CU149")
      If IsNull(rsTmp.Fields("CU150")) = False Then: textCU150 = rsTmp.Fields("CU150")
      If IsNull(rsTmp.Fields("CU151")) = False Then: textCU151 = rsTmp.Fields("CU151")
      If IsNull(rsTmp.Fields("CU152")) = False Then: textCU152 = rsTmp.Fields("CU152")
      '2011/3/4 End
      If IsNull(rsTmp.Fields("CU153")) = False Then: textCU153 = rsTmp.Fields("CU153") 'Add By Sindy 2011/3/17
      'Modified by Lydia 2022/12/20 改成「FCP提申急件預設組別」
      'If IsNull(rsTmp.Fields("CU154")) = False Then: textCU154 = rsTmp.Fields("CU154") 'Added by Morgan 2012/8/20
      If "" & rsTmp.Fields("CU154") = "" Then
          Combo4.Text = ""
      Else
          Combo4 = rsTmp.Fields("CU154") + "." + PUB_GetFCPGrpName(rsTmp.Fields("CU154"))
      End If
      'end 2022/12/20
      
      'Add by Amy 2015/09/09 +cu160~cu165 使用label顯示
      If 案件預設收據公司別啟用日 <= Val(strSrvDate(1)) Then
        For Each m_lbl In lblCU16X
           m_lbl = "" 'Add by Amy 2016/07/06
           If IsNull(rsTmp.Fields("CU" & m_lbl.Index + 160)) = False Then m_lbl = rsTmp.Fields("CU" & m_lbl.Index + 160)
        Next
      End If
      'end 2015/09/09
      
      'Add By Sindy 2013/1/17
      If IsNull(rsTmp.Fields("CU156")) = False Then
         Combo3(0).ListIndex = rsTmp.Fields("CU156")
      Else
         Combo3(0).ListIndex = 0
      End If
      If IsNull(rsTmp.Fields("CU157")) = False Then
         Combo3(1).ListIndex = rsTmp.Fields("CU157")
      Else
         Combo3(1).ListIndex = 0
      End If
      '2013/1/17 End
      If IsNull(rsTmp.Fields("CU180")) = False Then: textCU180 = rsTmp.Fields("CU180") 'Add by Amy 2019/08/27 客戶狀態備註
      'Add by Amy 2024/05/22 +不提供ID
      ChkID.Value = 0
      If "" & rsTmp.Fields("CU182") = "Y" Then ChkID.Value = 1
      'end 2024/05/22
      If IsNull(rsTmp.Fields("CU191")) = False Then: textCU191 = rsTmp.Fields("CU191") 'Add by Amy 2023/05/03 跨所同意主管(中文字)
      Label30(17).Caption = "" & rsTmp.Fields("CU199") 'Added by Lydia 2024/01/15 顧問專用信箱；只有智權在修改客戶資料可以維護，只可以輸入"無信箱或Email"。
      
      'Add by Morgan 2008/11/13 改用陣列以免控制項超過且以後再新增欄位也不必改
      For Each m_Txt In txtCU
         If IsNull(rsTmp.Fields("CU" & m_Txt.Index)) = False Then: m_Txt = rsTmp.Fields("CU" & m_Txt.Index)
      Next
      'Modified by Lydia 2021/12/14 改為Form 2.0元件
      PUB_SetUserList lstDeveloper, "" & rsTmp.Fields("CU129"), True
      'end 2008/11/13
      
      'Add By Sindy 2025/1/6
      If IsNull(rsTmp.Fields("CU201")) = False Then
         arrID = Split(rsTmp.Fields("CU201"), ",")
         For intI = UBound(arrID) To LBound(arrID) Step -1
            Chk1K(Val(arrID(intI)) - 1).Value = 1
         Next intI
      End If
      '2025/1/6 END
      
      'Added by Morgan 2025/2/27
      If IsNull(rsTmp.Fields("CU186")) = False Then
         arrID = Split(rsTmp.Fields("CU186"), ",")
         For intI = LBound(arrID) To UBound(arrID)
            ChkCU186(Val(arrID(intI))).Value = 1
         Next intI
      End If
      'end 2025/2/27
      
      'Add by Morgan 2008/7/30
      PUB_AddContact m_CurrKEY(0), cboContact, "" & rsTmp.Fields("CU127"), , True
      ' 更新CUID
      UpdateCUID rsTmp
      ' 更新暫存區的資料
      UpdateFieldOldData rsTmp
      
      textCU10_Validate False
      textCU09_Validate False
      textCU13_Validate False
      textCU12_Validate False
      textCU87_Validate False
      textCU57_Validate False
      textCU71_Validate False
      textCU94_Validate False
      textCU96_Validate False
      textCU97_Validate False
      textCU98_Validate False
      textCU99_Validate False
      textCU105_Validate False
      textCU106_Validate False
      'Add By Sindy 2011/3/4
      textCU147_Validate False
      textCU151_Validate False
      textCU152_Validate False
      '2011/3/4 End
   End If
   rsTmp.Close
   textCU12.Tag = textCU12.Text
   textCU13.Tag = textCU13.Text
   textCU04.Tag = textCU04.Text
   textCU05.Tag = textCU05.Text
   textCU88.Tag = textCU88.Text
   textCU89.Tag = textCU89.Text
   textCU90.Tag = textCU90.Text
   textCU06.Tag = textCU06.Text
   '2006/1/10 end
   'add by nickc 2006/03/16
   textCU03.Tag = textCU03.Text
   
EXITSUB:
   Set rsTmp = Nothing
End Sub

'Modify By Sindy 2021/12/7 因改Form2.0無法用 p_Listbox.ItemData(0) = PUB_Id2Num(.Fields(0)) '員工編號
'先 Public 改 Private
'Mark by Lydia 2021/12/14 改basQuery的共用模組
'Private Sub PUB_SetUserList(p_Listbox As Object, p_stNums As String)
'   Dim arrID, stSQL As String, intR As Integer, rstTmp As ADODB.Recordset
'   p_Listbox.Clear
'   If p_stNums <> "" Then
'      stSQL = "select st01,st02 from staff where instr('" & p_stNums & "',st01)>0"
'      intR = 1
'      Set rstTmp = ClsLawReadRstMsg(intR, stSQL)
'      If intR = 1 Then
'         arrID = Split(p_stNums, ",")
'         With rstTmp
'         '照原順序排
'         For intI = UBound(arrID) To LBound(arrID) Step -1
'            .MoveFirst
'            Do While Not .EOF
'               If .Fields("st01") = arrID(intI) Then
'                  p_Listbox.AddItem "" & .Fields(1), 0
'                  '2012/2/14 MODIFY BY SONIA 員工編號已可非數字需做轉換
'                  'Modify By Sindy 2021/12/7 Mark
'                  'p_Listbox.ItemData(0) = PUB_Id2Num(.Fields(0)) '員工編號
'                  '2021/12/7 END
'                  .MoveLast
'               End If
'               .MoveNext
'            Loop
'         Next
'         End With
'      End If
'   End If
'   Set rstTmp = Nothing
'End Sub
'end 2021/12/14

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
         'Modify by Amy 2022/09/14 收文新建客戶進入者不使用
         If m_bInsert And Left(m_PrevNo, 3) <> "Add" Then
            TBar1.Buttons(1).Enabled = True
         Else
            TBar1.Buttons(1).Enabled = False
         End If
         
         If m_bUpdate Then
            'Modify by Amy 2022/09/16 收文新建客戶進入新增未存檔不可用修改鈕
            If Left(m_PrevNo, 3) = "Add" And ExistCheck("Customer", "CU01||CU02", textCU01 & "0", "", False) = False Then
                TBar1.Buttons(2).Enabled = False
            Else
                TBar1.Buttons(2).Enabled = True
            End If
         Else
            TBar1.Buttons(2).Enabled = False
         End If
         If m_bDelete Then
            TBar1.Buttons(3).Enabled = True
         Else
            TBar1.Buttons(3).Enabled = False
         End If
         'Modify by Amy 2022/09/14 收文新建客戶進入者不使用
         If m_bQuery And Left(m_PrevNo, 3) <> "Add" Then
            TBar1.Buttons(4).Enabled = True
         Else
            TBar1.Buttons(4).Enabled = False
         End If
         'Modify by Amy 2022/09/14 收文新建客戶進入者不使用上/下/第一/最後 筆
         If m_bQuery And Left(m_PrevNo, 3) <> "Add" Then
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
         '2007/8/30 add by sonia 分所修改鎖住智權人員,再新增時要放開
         textCU12.Enabled = True
         textCU13.Enabled = True
         '2007/8/30 end
         ' 新增
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
End Sub

Private Function CheckDataValid() As Boolean
Dim strTit As String, strMsg As String, strTmp As String
Dim nResponse
Dim strTmp1 As String, strZipCode As String, strCountryCode As String 'Add by Amy 2015/08/24
Dim iRtn As Integer 'Add by Amy 2021/11/26
Dim bCancel As Boolean 'Add by Amy 2024/08/26
   
   CheckDataValid = False
  
   'add by nickc 2008/03/12 加控制，有輸入代理人編號時，必須第一碼為Y
   If textCU03 <> "" Then
        If Mid(textCU03, 1, 1) <> "Y" Then
            ShowMsg "代理人編號應該為 Y 開頭 !"
            textCU03.SetFocus
            tabCustomer.Tab = 0
            Exit Function
        End If
        
   End If
   If textCU04 = "" And textCU05 = "" And textCU88 = "" And textCU89 = "" And textCU90 = "" And textCU06 = "" Then
      ShowMsg "中、英、日文客戶名稱不可同時為空白 !"
      textCU04.SetFocus
      'Add By Cheng 2002/03/12
      tabCustomer.Tab = 0
      Exit Function
   End If
    '客戶名稱(中)
    'Modified by Lydia 2021/01/07 中、英、日文名稱改成判斷字串個數
    'If Not CheckLengthIsOK(textCU04, textCU04.MaxLength) Then
    If Len(textCU04) > textCU04.MaxLength Then
        textCU04.SetFocus
        textCU04_GotFocus
        tabCustomer.Tab = 0
        Exit Function
    End If
    '客戶名稱(日)
     'Modified by Lydia 2021/01/07 中、英、日文名稱改成判斷字串個數
     'If Not CheckLengthIsOK(textCU06, textCU06.MaxLength) Then
     If Len(textCU06) > textCU06.MaxLength Then
        textCU06.SetFocus
        textCU06_GotFocus
        tabCustomer.Tab = 0
        Exit Function
    End If
    '中文地址
     If Not CheckLengthIsOK(textCU23, textCU23.MaxLength) Then
        textCU23.SetFocus
        textCU23_GotFocus
        tabCustomer.Tab = 2
        Exit Function
    End If
    '公司負責人
    If Not CheckLengthIsOK(textCU07, textCU07.MaxLength) Then
        textCU07.SetFocus
        textCU07_GotFocus
        tabCustomer.Tab = 0
        Exit Function
    End If
    'Add by Amy 2025/03/05 非個人客戶,客戶狀態欄不能有遷移不明(避免國籍非台灣也彈,加國籍條件)
    If textCU10 < "010" And optCustomer(0).Value = 0 And cboStatus = "遷移不明" Then
      MsgBox "非個人客戶不能有遷移不明的客戶狀態" & vbCrLf & _
                     "請查詢商工登記公示資料" & vbCrLf & _
                     "依網站內資料更正地址並刪除遷移不明的客戶狀態"
      tabCustomer.Tab = 0
      If cboStatus.Locked = False Then cboStatus.SetFocus
      Exit Function
    End If
    'end 2025/03/05
    
    'Add by Amy 2019/08/27 客戶狀態為空
    'modify by sonia 2019/9/9 限智權部X38695(W1001)
    If cboStatus.Text = MsgText(601) And Left(textCU12, 1) = "S" Then
        '解除狀態時，智權人員不可為「該區」ex:10011(北一區)-2022/05/18 已改為區無效編號 ex:10019
        If m_FieldList(79).fiOldData <> MsgText(601) Then
            strTmp = GetAreaEmpNo(textCU12)
            If strTmp = textCU13 Then
                MsgBox "客戶狀態改為空,則智權人員不可為「" & Label30(2) & "」！"
                If textCU13.Enabled = False Then textCU13.Enabled = True 'Add by Amy 2019/12/10 X23353 84019 不可改會錯
                textCU13.SetFocus
                textCU13_GotFocus
                tabCustomer.Tab = 0
                Exit Function
            End If
        End If
    End If
    'add by nickc 2006/07/24 有輸入客戶狀態時，不印雜誌，不管新增或修改
    If cboStatus.Text <> "" Then
        'Modify  by Amy 2019/09/06
        If m_FieldList(79).fiOldData <> cboStatus Then
            '客戶狀態為其他或業務自行處理時需輸狀態備註
            If cboStatus = "其他" Or cboStatus = "業務自行處理" Then
                'Mark by Amy 2022/05/23 改至外層
'                If textCU180 = MsgText(601) Then
'                    MsgBox "客戶狀態為" & cboStatus & ",狀態備註不可為空！"
'                    textCU180.SetFocus
'                    textCU180_GotFocus
'                    tabCustomer.Tab = 0
'                    Exit Function
'                End If
                'Modify by Amy 2020/09/26 +if 判斷業務區別為S才判斷 ex:X79885 為林總 客戶不用判斷
                If Left(textCU12, 1) = "S" Then
                    strTmp = GetAreaEmpNo(textCU12)
                    If strTmp = textCU13 Then
                        MsgBox "客戶狀態為" & cboStatus & ",則智權人員不可為「" & Label30(2) & "」！"
                        textCU13.SetFocus
                        textCU13_GotFocus
                        tabCustomer.Tab = 0
                        Exit Function
                    End If
                    'Add by Amy 2022/05/23 客戶狀態為其他或業務自行處理時需輸狀態備註
                    'Modify by Amy 2023/02/01 +智權部客戶 字樣,因非智權部可不輸
                    If textCU180 = MsgText(601) Then
                        MsgBox "智權部客戶之客戶狀態為" & cboStatus & ",狀態備註不可為空！"
                        textCU180.SetFocus
                        textCU180_GotFocus
                        tabCustomer.Tab = 0
                        Exit Function
                    End If
                End If
            End If
        End If
        
        '客戶狀態為下列項目,智權人員自動改為「該區」
'*** Memo [此處狀態有修改],需確認 stNotModStatus 是否也需改 ***
        'Modify by Amy 2019/09/06 +限於智權部
        'Moidfy by Amy 2019/10/21 原有修改狀態才改,若只改智權人員也需改回區 ex:X64137
        'Modify by Amy 2022/05/23 遷移不明及停業 保留於個人
        'If Left(textCU12, 1) = "S" And (cboStatus = "遷移不明" Or cboStatus = "解散" Or cboStatus = "廢止" Or cboStatus = "撤銷" Or cboStatus = "停業" Or cboStatus = "死亡") Then
        'Modify by Amy 2023/06/06 修正訊息,以畫面上區的無效編號為主-秀玲
        'If Left(textCU12, 1) = "S" And (cboStatus = "解散" Or cboStatus = "廢止" Or cboStatus = "撤銷" Or cboStatus = "死亡") Then
        'Modify by Amy 2023/11/16 +M開頭部門(管理部)的無效客戶也要轉給10009
        If (Left(textCU12, 1) = "S" Or Left(textCU12, 1) = "M") And InStr(stNotModStatus, cboStatus) > 0 Then
            strTmp1 = GetAreaEmpNo(textCU12)
            '避免已更正又再彈,故需判斷目前畫面欄位是否已更正
            If m_FieldList(12).fiOldData = textCU13.Tag And strTmp1 <> textCU13 Then
               strMsg = "[智權部]"
               If Left(textCU12, 1) = "M" Then strMsg = "[管理部]"
               MsgBox strMsg & "客戶狀態為「" & cboStatus & "」" & vbCrLf & _
                               "智權人員只能為區無效編號！" & vbCrLf & _
                               "系統將自動更正"
               Label30(2) = GetPrjSalesNM(strTmp1)
            End If
            textCU13 = strTmp1
        End If
        'end 2023/11/16
        'end 2023/06/06
        'end 2019/09/06
        'modify by sonia 2022/3/31
        'textCU32 = "N"
        'textCU132 = "N" '2008/12/9 add by sonia
        If cboStatus <> "其他" And cboStatus <> "業務自行處理" And cboStatus <> "解除對造" Then
            textCU32 = "N"
            textCU132 = "N"
            textCU145 = "N" 'Add by Amy 2024/06/21 專利雙週報
        End If
        'end 2022/3/31
    'Add by Amy 2022/05/23 客戶狀態為空,狀態備註不可有值
    ElseIf textCU180 <> MsgText(601) Then
        MsgBox "客戶狀態為空白,狀態備註不可輸入！"
        textCU180.SetFocus
        textCU180_GotFocus
        tabCustomer.Tab = 0
        Exit Function
    End If
    
    'Add by Amy 2025/02/10
    'TEL2[不是]空且TEL1[是]空彈訊息
    If Trim(textCU17) <> MsgText(601) And Trim(textCU16) = MsgText(601) Then
       MsgBox "TEL2有值,TEL1不可為空！"
       textCU17.SetFocus
       textCU17_GotFocus
       tabCustomer.Tab = 1
       Exit Function
    End If
    'FAX2[不是]空且FAX1[是]空彈訊息
    If Trim(textCU19) <> MsgText(601) And Trim(textCU18) = MsgText(601) Then
       MsgBox "FAX2有值,FAX1不可為空！"
       textCU19.SetFocus
       textCU19_GotFocus
       tabCustomer.Tab = 1
       Exit Function
    End If
    'end 2025/02/10
    
    '聯絡人1(中)
    'Modified by Lydia 2017/06/14
    'If Not CheckLengthIsOK(textCU58, textCU58.MaxLength) Then
    If Not CheckLengthIsOK(textCU58, 30) Then
       textCU58.SetFocus
       textCU58_GotFocus
       tabCustomer.Tab = 1
       Exit Function
    End If
    '聯絡人2(中)
    'Modified by Lydia 2017/06/14
    'If Not CheckLengthIsOK(textCU61, textCU61.MaxLength) Then
    If Not CheckLengthIsOK(textCU61, 30) Then
       textCU61.SetFocus
       textCU61_GotFocus
       tabCustomer.Tab = 1
       Exit Function
    End If
    '實體聯絡人(中)
    If Not CheckLengthIsOK(textCU91, textCU91.MaxLength) Then
       textCU91.SetFocus
       textCU91_GotFocus
       tabCustomer.Tab = 1
       Exit Function
    End If
    '行業別
    If Not CheckLengthIsOK(textCU34, textCU34.MaxLength) Then
       textCU34.SetFocus
       textCU34_GotFocus
       tabCustomer.Tab = 0
       Exit Function
    End If
    '聯絡人1(日)
    'Modified by Lydia 2017/06/14
    'If Not CheckLengthIsOK(textCU60, textCU60.MaxLength) Then
    If Not CheckLengthIsOK(textCU60, 60) Then
       textCU60.SetFocus
       textCU60_GotFocus
       tabCustomer.Tab = 1
       Exit Function
    End If
    '聯絡人2(日)
    'Modified by Lydia 2017/06/14
    'If Not CheckLengthIsOK(textCU63, textCU63.MaxLength) Then
    If Not CheckLengthIsOK(textCU63, 60) Then
       textCU63.SetFocus
       textCU63_GotFocus
       tabCustomer.Tab = 1
       Exit Function
    End If
    '實體聯絡人(日)
    If Not CheckLengthIsOK(textCU93, textCU93.MaxLength) Then
       textCU93.SetFocus
       textCU93_GotFocus
       tabCustomer.Tab = 1
       Exit Function
    End If
    '客戶備註
    If Not CheckLengthIsOK(textCU79, textCU79.MaxLength) Then
       textCU79.SetFocus
       textCU79_GotFocus
       tabCustomer.Tab = 0
       Exit Function
    End If
    '聯絡地址
    If Not CheckLengthIsOK(textCU31, textCU31.MaxLength) Then
       textCU31.SetFocus
       textCU31_GotFocus
       tabCustomer.Tab = 2
       Exit Function
    End If
    '日文地址
    If Not CheckLengthIsOK(textCU29, textCU29.MaxLength) Then
       textCU29.SetFocus
       textCU29_GotFocus
       tabCustomer.Tab = 2
       Exit Function
    End If
    '代表人1(中), 代表人1(日), 代表人2(中), 代表人2(日), ...
    If Not CheckLengthIsOK(textCU39, textCU39.MaxLength) Then
       textCU39.SetFocus
       textCU39_GotFocus
       tabCustomer.Tab = 3
       Exit Function
    End If
    If Not CheckLengthIsOK(textCU41, textCU41.MaxLength) Then
       textCU41.SetFocus
       textCU41_GotFocus
       tabCustomer.Tab = 3
       Exit Function
    End If
    If Not CheckLengthIsOK(textCU42, textCU42.MaxLength) Then
       textCU42.SetFocus
       textCU42_GotFocus
       tabCustomer.Tab = 3
       Exit Function
    End If
    If Not CheckLengthIsOK(textCU44, textCU44.MaxLength) Then
       textCU44.SetFocus
       textCU44_GotFocus
       tabCustomer.Tab = 3
       Exit Function
    End If
    If Not CheckLengthIsOK(textCU45, textCU45.MaxLength) Then
       textCU45.SetFocus
       textCU45_GotFocus
       tabCustomer.Tab = 3
       Exit Function
    End If
    If Not CheckLengthIsOK(textCU47, textCU47.MaxLength) Then
       textCU47.SetFocus
       textCU47_GotFocus
       tabCustomer.Tab = 3
       Exit Function
    End If
    If Not CheckLengthIsOK(textCU48, textCU48.MaxLength) Then
       textCU48.SetFocus
       textCU48_GotFocus
       tabCustomer.Tab = 3
       Exit Function
    End If
    If Not CheckLengthIsOK(textCU50, textCU50.MaxLength) Then
       textCU50.SetFocus
       textCU50_GotFocus
       tabCustomer.Tab = 3
       Exit Function
    End If
    If Not CheckLengthIsOK(textCU51, textCU51.MaxLength) Then
       textCU51.SetFocus
       textCU51_GotFocus
       tabCustomer.Tab = 3
       Exit Function
    End If
    If Not CheckLengthIsOK(textCU53, textCU53.MaxLength) Then
       textCU53.SetFocus
       textCU53_GotFocus
       tabCustomer.Tab = 3
       Exit Function
    End If
    If Not CheckLengthIsOK(textCU54, textCU54.MaxLength) Then
       textCU54.SetFocus
       textCU54_GotFocus
       tabCustomer.Tab = 3
       Exit Function
    End If
    If Not CheckLengthIsOK(textCU56, textCU56.MaxLength) Then
       textCU56.SetFocus
       textCU56_GotFocus
       tabCustomer.Tab = 3
       Exit Function
    End If
    'Add by Amy 2015/10/24 +代表人2~6其中一個有值,則代表人1不可為空
    If Trim(textCU39) = MsgText(601) And Trim(textCU40) = MsgText(601) And Trim(textCU41) = MsgText(601) Then
        If Trim(textCU42) <> MsgText(601) Or Trim(textCU43) <> MsgText(601) Or Trim(textCU44) <> MsgText(601) Then
            ShowMsg "代表人2有值代表人1不可為空!"
            textCU42.SetFocus
            textCU42_GotFocus
            tabCustomer.Tab = 3
            Exit Function
        End If
        If Trim(textCU45) <> MsgText(601) Or Trim(textCU46) <> MsgText(601) Or Trim(textCU47) <> MsgText(601) Then
            ShowMsg "代表人3有值代表人1不可為空!"
            textCU45.SetFocus
            textCU45_GotFocus
            tabCustomer.Tab = 3
            Exit Function
        End If
        If Trim(textCU48) <> MsgText(601) Or Trim(textCU49) <> MsgText(601) Or Trim(textCU50) <> MsgText(601) Then
            ShowMsg "代表人4有值代表人1不可為空!"
            textCU48.SetFocus
            textCU48_GotFocus
            tabCustomer.Tab = 3
            Exit Function
        End If
        If Trim(textCU51) <> MsgText(601) Or Trim(textCU52) <> MsgText(601) Or Trim(textCU53) <> MsgText(601) Then
            ShowMsg "代表人5有值代表人1不可為空!"
            textCU51.SetFocus
            textCU51_GotFocus
            tabCustomer.Tab = 3
            Exit Function
        End If
        If Trim(textCU54) <> MsgText(601) Or Trim(textCU55) <> MsgText(601) Or Trim(textCU56) <> MsgText(601) Then
            ShowMsg "代表人6有值代表人1不可為空!"
            textCU54.SetFocus
            textCU54_GotFocus
            tabCustomer.Tab = 3
            Exit Function
        End If
    End If
    'end 2015/10/24
   If textCU10.Text = "" Then
      ShowMsg "客戶國籍不可為空白 !"
      textCU10.SetFocus
      tabCustomer.Tab = 0
      Exit Function
   'Add by Amy 2015/10/07 直接key000 按Enter未檢查到不可輸台灣國家代號
   Else
        'Modify by Amy 2019/09/05 改與聯絡地址相同判斷,ex:X3658701 中文地址郵遞區號為空
        If textCU10.Text = 台灣國家代號 Or Len(textCU10) <= 2 Then
            ShowMsg "客戶" & MsgText(9153)
            textCU10.SetFocus
            textCU10_GotFocus
            tabCustomer.Tab = 0
            Exit Function
         End If
         '客戶國籍為台灣且中文地址不為空,中文地址郵遞區號不可為空
         'Modify by Amy 2023/08/10 +無效客戶不檢查
         'Modify by Amy 2024/08/30 bug-cboStatus為"",InStr(stNotModStatus, cboStatus) = 0會回傳False,加地址不是後補 or 備註[沒]臺灣地址格式不檢查 字樣
         If textCU10 < "010" And Trim(textCU23) <> MsgText(601) And Trim(textCU23) <> "後補" And Trim(textCU112) = MsgText(601) _
           And ((cboStatus <> MsgText(601) And InStr(stNotModStatus, cboStatus) = 0) Or cboStatus = MsgText(601)) _
           And InStr(textCU79, "臺灣地址格式不檢查") = 0 Then
            ShowMsg "客戶中文地址有資料" & vbCrLf & _
                                 "中文地址郵遞區號不可為空！"
            textCU112.SetFocus
            textCU112_GotFocus
            tabCustomer.Tab = 2
            Exit Function
         End If
         'Add by Amy 2023/05/03 地址中不可有刪址字樣
         If InStr(textCU23, "刪址") > 0 Then
            ShowMsg "中文地址不可為刪址！"
            textCU23.SetFocus
            textCU23_GotFocus
            tabCustomer.Tab = 2
            Exit Function
         End If
         If InStr(textCU31, "刪址") > 0 Then
            ShowMsg "聯絡地址不可為刪址！"
            textCU31.SetFocus
            textCU31_GotFocus
            tabCustomer.Tab = 2
            Exit Function
         End If
         'Add by Amy 2021/11/26 客戶國籍為台灣且中文名稱有「事務所」字樣且狀態 非「國內同業」,彈 國內同業控制
         If textCU10 < "010" And InStr(textCU04, "事務所") > 0 And cboStatus <> "國內同業" Then
             iRtn = MsgBox("國籍在台灣之事務所，請與智權人員確認是否為國內同業？" & vbCrLf & _
                                        "是:為國內同業　否:非國內同業", vbYesNoCancel + vbDefaultButton3)
            '取消
            If iRtn = 2 Then
                tabCustomer.Tab = 0
               Exit Function
            '是
            ElseIf iRtn = 6 Then
                cboStatus = "國內同業"
            End If 'iRtn
         End If
         'end 2021/11/26
   End If
   'Add by Amy 2021/11/26
   If cboStatus = "國內同業" Then
        '非財務Email不可輸
        'Modify by Amy 2022/09/23 訊息一次顯示 原:ShowMsg
        strMsg = ""
        If textCU20 <> MsgText(601) Then
            strMsg = strMsg & "此客戶為國內同業,不可輸入E-Mail(代表)以免誤發電子郵件, 如有需要請加註於備註欄 ！" & vbCrLf
'            textCU20.SetFocus
'            textCU20_GotFocus
'            tabCustomer.Tab = 1
'            Exit Function
        End If
        If textCU116 <> MsgText(601) Then
            strMsg = strMsg & "此客戶為國內同業,不可輸入E-Mail(其他1)以免誤發電子郵件, 如有需要請加註於備註欄 ！" & vbCrLf
'            textCU116.SetFocus
'            textCU116_GotFocus
'            tabCustomer.Tab = 1
'            Exit Function
        End If
        If textCU117 <> MsgText(601) Then
            strMsg = strMsg & "此客戶為國內同業,不可輸入E-Mail(其他2)以免誤發電子郵件, 如有需要請加註於備註欄 ！" & vbCrLf
'            textCU117.SetFocus
'            textCU117_GotFocus
'            tabCustomer.Tab = 1
'            Exit Function
        End If
        If textCU118 <> MsgText(601) Then
            strMsg = strMsg & "此客戶為國內同業,不可輸入E-Mail(其他3)以免誤發電子郵件, 如有需要請加註於備註欄 ！" & vbCrLf
'            textCU118.SetFocus
'            textCU118_GotFocus
'            tabCustomer.Tab = 1
'            Exit Function
        End If
        '電子報要設定不寄
        If textCU32 <> "N" Then
            strMsg = strMsg & "此客戶為國內同業, 不可寄台一雜誌 ！" & vbCrLf
'            textCU32.SetFocus
'            textCU32_GotFocus
'            tabCustomer.Tab = 0
'            Exit Function
        End If
        If textCU132 <> "N" Then
            strMsg = strMsg & "此客戶為國內同業, 不可寄電子報 ！" & vbCrLf
'            textCU132.SetFocus
'            textCU132_GotFocus
'            tabCustomer.Tab = 0
'            Exit Function
        End If
        If textCU145 <> "N" Then
            strMsg = strMsg & "此客戶為國內同業, 不可專利雙週報 ！" & vbCrLf
'            textCU145.SetFocus
'            textCU145_GotFocus
'            tabCustomer.Tab = 0
'            Exit Function
        End If
        If textCU153 <> "N" Then
            strMsg = strMsg & "此客戶為國內同業, 不可寄顧問電子報！" & vbCrLf
'            textCU153.SetFocus
'            textCU153_GotFocus
'            tabCustomer.Tab = 0
'            Exit Function
        End If
        If strMsg <> MsgText(601) Then
            MsgBox strMsg, vbCritical + vbOKOnly, MsgText(9001)
            Exit Function
        End If
        'end 2022/09/23
   End If
   'end 2021/11/26

   If textCU64.Text = "" Then
      ShowMsg "定稿語文不可為空白 !"
      textCU64.SetFocus
      'Add By Cheng 2002/03/12
      tabCustomer.Tab = 0
      Exit Function
   'Add by Amy 2021/11/26 以國籍判斷定稿語文彈提醒-外商阿蓮
   'Modify by Amy 2021/12/09 +Left(textCU10, 3)
   Else
        If Left(textCU10, 3) = "011" And textCU64 <> "3" Then
            If MsgBox("國籍為「日本」，定稿語文確定「不是」日文？", vbYesNo + vbDefaultButton2 + vbQuestion) = vbNo Then
                textCU64.SetFocus
                tabCustomer.Tab = 0
                Exit Function
            End If
        ElseIf Left(textCU10, 3) <> "011" And textCU64 = "3" Then
            If MsgBox("國籍為「不是」日本，定稿語文確定是「日文」？", vbYesNo + vbDefaultButton2 + vbQuestion) = vbNo Then
                textCU64.SetFocus
                tabCustomer.Tab = 0
                Exit Function
            End If
        End If
   End If
   'Modify by Amy 2015/08/24 原判斷客戶國籍textCU10.Text
   '因先key客戶國籍會設地址國籍=客戶國籍,若聯絡地址與客戶國籍不同時,會忘記改,導致資料不一致
   If textCU87.Text <> "" Then
      If Val(textCU87.Text) < 9 Then
         If textCU87 = "000" Or Len(textCU87) <= 2 Then
            ShowMsg "地址國籍有誤!"
            textCU87_GotFocus
             tabCustomer.Tab = 2
            Exit Function
         End If
         If textCU31 = "" Then
            ShowMsg "國籍為台灣，聯絡地址不可為空白 !"
            textCU31.SetFocus
            'Add By Cheng 2002/03/12
            tabCustomer.Tab = 2
            Exit Function
         End If
         If textCU30 = "" Then
            ShowMsg "國籍為台灣，聯絡地址郵遞區號不可為空白 !"
            textCU30.SetFocus
            'Add By Cheng 2002/03/12
            tabCustomer.Tab = 2
            Exit Function
         End If
         '2006/5/2 ADD BY SONIA
'edit by nickc 2007/04/27 改新增時，皆不可空白
'         If textCU13 = "" Then
'            ShowMsg "國籍為台灣，智權人員不可為空白 !"
'            textCU13.SetFocus
'            tabCustomer.Tab = 0
'            Exit Function
'         End If
         '2006/5/2 END
      'ElseIf Val(Text1(12).Text) > 9 Then
         'If Text1(61).Text = "" And Text1(62).Text = "" And Text1(63).Text = "" Then
         '   ShowMsg "國籍非台灣時，聯絡人中、英、日不可同時為空白 !"
         '   Text1(61).Text = ""
         '   Exit Function
         'End If
      End If
   End If
   'Modify by Amy 2015/10/23 +修改時
   'add by nickc 2007/04/27 新增時，智權人員不可空白
   If m_EditMode = 1 Or m_EditMode = 2 Then
      If textCU13 = "" Then
         ShowMsg "智權人員不可為空白 !"
         textCU13.SetFocus
         tabCustomer.Tab = 0
         Exit Function
      'Add by Amy 2015/09/09 +離職人員彈訊息
      Else
           If PUB_GetStaffNameDept(textCU13, strTmp, strTmp1, True, True) = False Then
                textCU13.SetFocus
                tabCustomer.Tab = 0
                Exit Function
           End If
      'end 2015/09/09
      End If
   End If
'   If Text1(13).Text = "" Then
'      If optCustomer(0).Value Then
'         ShowMsg "身份證字號不可為空白 !"
'      Else
'         ShowMsg "統一編號不可為空白 !"
'      End If
'      Text1(13).SetFocus
'      Exit Function
'   End If
   'Modify by Amy 2024/11/29 +新增時不可同時空白 or 修改時不允許,有資料改為都是空白
   If (m_EditMode = 1 Or (m_EditMode = 2 _
     And (m_FieldList(22).fiOldData <> "" Or m_FieldList(23).fiOldData <> "" Or m_FieldList(24).fiOldData <> "" Or m_FieldList(25).fiOldData <> "" Or m_FieldList(26).fiOldData <> "" Or m_FieldList(27).fiOldData <> "" Or m_FieldList(28).fiOldData <> "" Or m_FieldList(101).fiOldData <> ""))) _
     And textCU23 = "" And textCU24 = "" And textCU25 = "" And textCU26 = "" _
      And textCU27 = "" And textCU28 = "" And textCU29 = "" And textCU102 = "" Then
      ShowMsg "中、英、日文地址不可同時為空白 !"
      textCU23.SetFocus
      'Add By Cheng 2002/03/12
      tabCustomer.Tab = 2
      Exit Function
   End If
   'Mark by Amy 2015/09/09 秀玲:取消此判斷
'   If textCU24 <> "" Or textCU25 <> "" Or textCU26 <> "" _
'      Or textCU27 <> "" Or textCU28 <> "" Or textCU102 <> "" Then
'         If textCU87 = "" Then
'            ShowMsg "英文地址不為空白時，地址國籍不可為空白 !"
'            textCU87.SetFocus
'            'Add By Cheng 2002/03/12
'            tabCustomer.Tab = 2
'            Exit Function
'         End If
'   End If
   '2008/9/4 add by sonia
   If textCU87.Text = "" Then
      ShowMsg "地址國籍不可為空白 !"
      textCU87.SetFocus
      tabCustomer.Tab = 2
      Exit Function
   End If
   '2008/9/4 end
   If textCU71 <> "" Then
      If textCU70 = "" Then
         ShowMsg "副本收受人不為空白時，副本聯絡人不可空白 !"
         textCU70.SetFocus
         'Add By Cheng 2002/03/12
         tabCustomer.Tab = 4
         Exit Function
      End If
   End If
   If textCU94 <> "" Then
      If textCU95 = "" Then
         ShowMsg "實體副本收受人不為空白時，實體副本聯絡人不可空白 !"
         textCU95.SetFocus
         'Add By Cheng 2002/03/12
         tabCustomer.Tab = 4
         Exit Function
      End If
   End If
   
   'Add By Sindy 2013/1/17
   If Trim(Me.Combo2(0).Text) <> "" Then
      '若輸入幣別就一定要選格式
      If Trim(Me.Combo3(0).Text) = "" Then
         ShowMsg "專利請款單列印幣別格式不可空白 !"
         Me.Combo3(0).SetFocus
         GoTo EXITSUB
      End If
      '請款幣別<>NTD時不可輸入1
      If Trim(Me.Combo2(0).Text) <> "NTD" And Me.Combo3(0).ListIndex = 1 Then
         ShowMsg "專利請款幣別<>NTD時，專利請款單列印幣別格式不可選純台幣 !"
         Me.Combo3(0).SetFocus
         GoTo EXITSUB
      End If
   End If
   If Trim(Me.Combo2(1).Text) <> "" Then
      '若輸入幣別就一定要選格式
      If Trim(Me.Combo3(1).Text) = "" Then
         ShowMsg "商標請款單列印幣別格式不可空白 !"
         Me.Combo3(1).SetFocus
         GoTo EXITSUB
      End If
      '請款幣別<>NTD時不可輸入1
      If Trim(Me.Combo2(1).Text) <> "NTD" And Me.Combo3(1).ListIndex = 1 Then
         ShowMsg "商標請款幣別<>NTD時，商標請款單列印幣別格式不可選純台幣 !"
         Me.Combo3(1).SetFocus
         GoTo EXITSUB
      End If
   End If
   '2013/1/17 End
   'Add by Amy 2024/08/26 跨所同意主管
    Call textCU191_Validate(bCancel)
    If bCancel = True Then
      textCU191.SetFocus
      GoTo EXITSUB
    End If
   
   'Added by Lydia 2019/04/17　修改公司負責人CU07時 (改名稱),若業務區非 Fxx 部門時(例外:F3X投資法律), 彈提醒檢查代表人資料
   If m_EditMode = 2 And Trim(m_CU07) <> Trim(textCU07) And Left(Trim(textCU12), 1) <> "F" And bolMsg = False Then
        MsgBox "修改公司負責人，請一併檢查代表人資料！", vbInformation, "資料檢核"
        Me.tabCustomer.Tab = 3
        bolMsg = True
        Exit Function
   End If
   'end 2019/04/17
   
   'add by nickc 2006/03/09 檢查用客戶和代理人是否與代理人檔相同
'2008/7/17 CANCEL BY SONIA 改至存檔時詢問
'   If textCU03 <> "" Then
'      strTmp = "select * from fagent where fa01='" & Left(textCU03 & "00000000", 8) & "' and fa02='0' "
'      CheckOC3
'      AdoRecordSet3.CursorLocation = adUseClient
'      AdoRecordSet3.Open strTmp, cnnConnection, adOpenStatic, adLockReadOnly
'      If AdoRecordSet3.RecordCount <> 0 Then
'         If CheckStr(AdoRecordSet3.Fields("fa03")) <> textCU01 And CheckStr(AdoRecordSet3.Fields("fa03")) <> "" Then
'            ShowMsg "客戶編號與代理人檔客戶編號不同 !"
'            textCU03.SetFocus
'            tabCustomer.Tab = 0
'            Exit Function
'         End If
'      End If
'   End If
'2008/7/17 END
   If textCU01.Text <> "" Then
      strTmp = textCU01.Text
      'Add by Amy 2020/04/30 不允許輸空白 ex:X775710空白
      If textCU01.Locked = False Then
        If InStr(strTmp, " ") > 0 Then
            ShowMsg "客戶編號不可輸空白,請重新輸入 !"
            Exit Function
        End If
      End If
      ' 90.12.18 modify by louis
      strTmp = strTmp & String(8 - Len(strTmp), "0")
      If textCU02.Text <> "" Then
         strTmp = strTmp & textCU02.Text
      Else
         strTmp = strTmp & "0"
      End If
      strExc(0) = "SELECT COUNT(*) FROM CUSTOMER WHERE " & ChgCustomer(strTmp)
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         If RsTemp.Fields(0) > 0 Then
            If m_EditMode = 1 Then
               ShowMsg "客戶編號重覆，請重新輸入 !"
               Exit Function
            End If
         End If
      End If
   Else
      If m_EditMode = 1 Then
         'edit by nickc 2007/02/02 不用 dll 了
         'If objPublicData.GetAutoNumber("X", strTmp, True, False) Then
         'Move by Lydia 2017/05/09 自動編號移到檢查的後面
      Else
         ShowMsg "客戶編號不得為空值，請重新輸入 !"
         Exit Function
      End If
   End If
'edit by nickc 2008/05/08 改共用 function
'   'add by nickc 2006/06/15 加入檢查英文名稱第一碼
'   If Mid(textCU10, 1, 3) = "101" Then
'       '2008/1/4 MODIFY BY SONIA 原為A~I為101,J~Z為1011,2008年改為分四段
'       If Mid(UCase(LTrim(textCU05)), 1, 1) >= "A" And Mid(UCase(LTrim(textCU05)), 1, 1) <= "E" Then
'            If Trim(textCU10) <> "101" Then
'                ShowMsg "客戶英文名稱第一碼介於 A~E 之間，客戶國籍應該為 101 !"
'                Exit Function
'            End If
'       ElseIf Mid(UCase(LTrim(textCU05)), 1, 1) >= "F" And Mid(UCase(LTrim(textCU05)), 1, 1) <= "I" Then
'            If Trim(textCU10) <> "1011" Then
'                ShowMsg "客戶英文名稱第一碼介於 F~I 之間，客戶國籍應該為 1011 !"
'                Exit Function
'            End If
'       ElseIf Mid(UCase(LTrim(textCU05)), 1, 1) >= "J" And Mid(UCase(LTrim(textCU05)), 1, 1) <= "N" Then
'            If Trim(textCU10) <> "1012" Then
'                ShowMsg "客戶英文名稱第一碼介於 J~N 之間，客戶國籍應該為 1012 !"
'                Exit Function
'            End If
'       ElseIf Mid(UCase(LTrim(textCU05)), 1, 1) >= "O" And Mid(UCase(LTrim(textCU05)), 1, 1) <= "Z" Then
'            If Trim(textCU10) <> "1013" Then
'                ShowMsg "客戶英文名稱第一碼介於 O~Z 之間，客戶國籍應該為 1013 !"
'                Exit Function
'            End If
'       '2008/1/9 add by sonia
'       Else
'            If Trim(textCU10) <> "1013" Then
'                ShowMsg "客戶英文名稱第一碼非英文字母或無英文名稱，客戶國籍應該為 1013 !"
'                Exit Function
'            End If
'       '2008/1/9 end
'       End If
'   ElseIf Mid(textCU10, 1, 3) = "011" Then
'       '2008/4/21 MODIFY BY SONIA 原為A~L為011,M~Z為0111,2008/4/22改為分三段(將M~Z再細分成二段)
'       If Mid(UCase(LTrim(textCU05)), 1, 1) >= "A" And Mid(UCase(LTrim(textCU05)), 1, 1) <= "L" Then
'            If Trim(textCU10) <> "011" Then
'                ShowMsg "客戶英文名稱第一碼介於 A~L 之間，客戶國籍應該為 011 !"
'                Exit Function
'            End If
'       ElseIf Mid(UCase(LTrim(textCU05)), 1, 1) >= "M" And Mid(UCase(LTrim(textCU05)), 1, 1) <= "O" Then
'            If Trim(textCU10) <> "0111" Then
'                ShowMsg "客戶英文名稱第一碼介於 M~O 之間，客戶國籍應該為 0111 !"
'                Exit Function
'            End If
'       ElseIf Mid(UCase(LTrim(textCU05)), 1, 1) >= "P" And Mid(UCase(LTrim(textCU05)), 1, 1) <= "Z" Then
'            If Trim(textCU10) <> "0112" Then
'                ShowMsg "客戶英文名稱第一碼介於 P~Z 之間，客戶國籍應該為 0112 !"
'                Exit Function
'            End If
'       '2008/1/9 modify by sonia
'       'ElseIf Trim(textCU05) = "" Then
'       Else
'            If Trim(textCU10) <> "0112" Then
'                ShowMsg "客戶英文名稱第一碼非英文字母或無英文名稱，客戶國籍應該為 0112 !"
'                Exit Function
'            End If
'       End If
'   End If
    'Add by Amy 2022/09/14 收文新建客戶進入需檢查輸入之英文名稱地址是否一致
    If Left(m_PrevNo, 3) = "Add" Then
        'Add by Amy 2022/10/03 前畫面有關係企業,編號不可為空
        If m_Cra04 <> MsgText(601) And textCU01 = MsgText(601) Then
            ShowMsg "此為關係企業,請輸入關係企業編號！"
            textCU01_GotFocus
            Exit Function
        End If
        'end 2022/10/03
        
        'Modify by Amy 2023/06/27 改抓ReplaceSign DB函數
        '客戶英文名稱
        'Modify by Amy 2023/01/07 取代符號,改抓共用函數
        strTmp = Pub_GetField("Dual", "1=1", "ReplaceSign(TO_MULTI_BYTE(Upper('" & ChgSQL(strCra08) & "')))")
        strTmp1 = Pub_GetField("Dual", "1=1", "ReplaceSign(TO_MULTI_BYTE(Upper('" & ChgSQL(textCU05 & textCU88 & textCU89 & textCU90) & "')))")
        'If Pub_ReplaceSign(False, strCra08) <> Pub_ReplaceSign(False, textCU05 & textCU88 & textCU89 & textCU90) Then
        If strTmp <> strTmp1 Then
            ShowMsg "客戶英文名稱與接洽單資料不一致！"
            textCU05_GotFocus
            Exit Function
        End If
        '客戶英文地址
        strTmp = Pub_GetField("Dual", "1=1", "ReplaceSign(TO_MULTI_BYTE(Upper('" & ChgSQL(strCra21) & "')))")
        strTmp1 = Pub_GetField("Dual", "1=1", "ReplaceSign(TO_MULTI_BYTE(Upper('" & ChgSQL(textCU24 & textCU25 & textCU26 & textCU27 & textCU28 & textCU102) & "')))")
        'If Pub_ReplaceSign(False, strCra21) <> Pub_ReplaceSign(False, textCU24 & textCU25 & textCU26 & textCU27 & textCU28 & textCU102) Then
        If strTmp <> strTmp1 Then
        'end 2023/01/07
        'end 2023/06/27
            ShowMsg "客戶英文地址與接洽單資料不一致！"
            textCU05_GotFocus
            Exit Function
        End If
    End If
    'end 2022/09/14
    If Trim(textCU10) <> pub_NationByName(textCU05 & textCU88 & textCU89 & textCU90, Trim(textCU10), True, "客戶") Then
        'Added by Lydia 2016/08/10
        If Me.ActiveControl = textCU10 Then
           textCU10_GotFocus
        Else
           textCU10.SetFocus
        End If
        tabCustomer.Tab = 0
        'end 2016/08/10
        GoTo EXITSUB
    End If

   '2011/10/17 add by sonia
   'Modify by Amy 2016/06/15 依國籍確認臺灣地址格式(地址會有刪址字,故多加客戶狀態判斷)
   If Trim(textCU31.Text) <> "" And InStr(LTrim(textCU31), "後補") = 0 And textCU87 < "010" And (cboStatus = "" Or cboStatus = "其他" Or cboStatus = "業務自行處理") Then
      'Add by Amy 2019/10/03 +if 有修改才判斷
      If m_FieldList(30).fiOldData <> textCU31 Then
        If CheckAddrData(textCU31, textCU30, textCU87) = False Then
          GoTo EXITSUB
        End If
        If CheckTaiwanAddr_Tai(textCU31, textCU30, textCU87, "聯絡地址", strZipCode, strCountryCode) = False Then
          Call ChkZipData(9, textCU31)
          GoTo EXITSUB
        End If
      End If
        
'      If CheckTaiwanAddr(textCU31.Text, textCU87.Text, "聯絡地址") = False Then
'         Me.tabCustomer.Tab = 2
'         textCU31.SetFocus
'         textCU31_GotFocus
'         GoTo EXITSUB
'      End If
   End If
   '中文地址
    'Modify by Amy 2024/06/28 +"臺灣地址格式不檢查",每天秀玲會檢查資料所以排除,避免無法存檔 ex:國籍為台灣 中文地址[非]台灣,會一直檢查台灣地址是否符合
    '註:客戶國籍[非]台灣,中文地址[非]台灣 ->台灣地址,因國籍不會改,不會判斷地址是否正確,也不會判斷上述條件-秀玲說每天會用語法檢查
    '     客戶國籍 為 台灣,中文地址 為 台灣 ->台灣地址,但國籍不同,國籍會更正；中文地址 為 台灣 ->[非]台灣地址,會於備註顯示[臺灣地址格式不檢查]
    If Trim(textCU23.Text) <> "" And InStr(LTrim(textCU23), "後補") = 0 And InStr(textCU79, "臺灣地址格式不檢查") = 0 _
      And textCU10 < "010" And (cboStatus = "" Or cboStatus = "其他" Or cboStatus = "業務自行處理") Then
      'Add by Amy 2019/10/03 +if 有修改才判斷
      'Modify by Amy 2024/06/17 +國籍有修改
      'Modify by Amy 2025/03/06 +接洽單判斷textCU10.Tag ,避免一些不需彈的訊息被觸發
      '              ex:接洽單 1130021562,客戶國籍為香港,存檔時因m_FieldList(9).fiOldData =空,彈「修改客戶國籍,地址國籍是否同時修改...」,而修改成客戶國籍為台灣
      'If m_FieldList(22).fiOldData <> textCU23 Or m_FieldList(9).fiOldData <> textCU10 Then
      If m_FieldList(22).fiOldData <> textCU23 _
        Or ((Left(m_PrevNo, 3) <> "Add" And textCU10.Text <> m_FieldList(9).fiOldData) Or (Left(m_PrevNo, 3) = "Add" And textCU10 <> textCU10.Tag)) Then
        If CheckAddrData(textCU23, textCU112, textCU10) = False Then
          GoTo EXITSUB
        End If
        If CheckTaiwanAddr_Tai(textCU23, textCU112, textCU10, "中文地址", strZipCode, strCountryCode) = False Then
          Call ChkZipData(9, textCU23)
          GoTo EXITSUB
        End If
      End If
      'end 2019/10/03
'      If CheckTaiwanAddr(textCU23.Text, textCU10.Text, "中文地址") = False Then
'         Me.tabCustomer.Tab = 2
'         textCU23.SetFocus
'         textCU23_GotFocus
'         GoTo EXITSUB
'      End If
   End If
   'end 2016/06/15
   '2011/10/17 end
   
   'add by nickc 2007/03/05 將業務區同步
   textCU12.Text = GetST15(textCU13.Text)
   
   'Memo by Amy 2023/07/06 原:自動編號移到全部檢查完再run,避免電子收文已給號,TxtValidate 檢查有誤跳離開,但自動編號卻已跳號
   
   If m_EditMode = 2 Then
      '2012/8/7 modify by sonia 國內智權人員改國外智權人員也要做X44833之77050->80030
      'If Me.textCU13.Text <> Me.textCU13.Tag And Mid(Me.textCU12, 1, 1) <> "F" Then
      If Me.textCU13.Text <> Me.textCU13.Tag And (Mid(Me.textCU12, 1, 1) <> "F" Or Mid(Me.textCU12.Tag, 1, 1) <> "F") Then
          If MsgBox("是否確定修改智權人員？注意！期限資料也會修改！！！", vbExclamation + vbOKCancel + vbDefaultButton2) = vbCancel Then
               Exit Function
          End If
      End If
      'Add by Amy 2021/09/01 狀態修改為遷移不明、解散、廢止、撤銷、停業、死亡、其他、業務自行處理 彈訊息
      If m_FieldList(79).fiOldData <> cboStatus Then
            If cboStatus = "遷移不明" Or cboStatus = "解散" Or cboStatus = "廢止" Or cboStatus = "撤銷" Or cboStatus = "停業" Or cboStatus = "死亡" Then
                'Modify by Amy 2023/06/06 +客戶狀態為「" & cboStatus & "」
                MsgBox "請注意！！" & vbCrLf & "客戶狀態為「" & cboStatus & "」" & vbCrLf & "須有區主管簽准！"
            End If
            If cboStatus = "其他" Or cboStatus = "業務自行處理" Then
                'Modify by Amy 2023/06/06 +客戶狀態為「" & cboStatus & "」
                MsgBox "請注意！！" & vbCrLf & "客戶狀態為「" & cboStatus & "」" & vbCrLf & "須有區主管及權責主管(參閱異動表之註2)簽准！"
            End If
      End If
      'end 2021/09/01
      'Added by Lydia 2019/12/04 年費不續辦衝突管制：
                            '若設申請人設定年費自動代繳/年費不續辦，與代理人有衝突，發email通知承辦CC程序管制
      If textCU74.Text <> "" And textCU74.Tag <> textCU74.Text Then
          '保留
          'strExc(0) = "SELECT PA01||'-'||PA02||DECODE(PA03||PA04,'000','','-'||PA03||'-'||PA04) AS CASENO,PA75,NVL(FA05,NVL(FA04,FA06)) FNAME" & _
                           " FROM PATENT,FAGENT" & _
                           " WHERE INSTR(PA26||','||PA27||','||PA28||','||PA29||','||PA30,'" & Left(ChangeCustomerL(textCU01), 8) & "') >0" & _
                           " AND SUBSTR(PA75,1,8)=FA01(+) AND SUBSTR(PA75,9,1)=FA02(+) AND FA41=" & CNULL(IIf(textCU74.Text = "Y", "N", "Y"))
          strExc(0) = "SELECT COUNT(*) CNT " & _
                           " FROM PATENT,FAGENT" & _
                           " WHERE INSTR(PA26||','||PA27||','||PA28||','||PA29||','||PA30,'" & Left(ChangeCustomerL(textCU01), 8) & "') >0" & _
                           " AND SUBSTR(PA75,1,8)=FA01(+) AND SUBSTR(PA75,9,1)=FA02(+) AND FA41=" & CNULL(IIf(textCU74.Text = "Y", "N", "Y"))
          intI = 1
          Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
          If intI = 1 Then
              '保留
              'strExc(1) = ""
              'RsTemp.MoveFirst
              'Do While Not RsTemp.EOF
              '     strExc(1) = strExc(1) & vbCrLf & convForm("" & RsTemp.Fields("caseno"), 15) '& "  " & RsTemp.Fields("pa75") & "  " & RsTemp.Fields("fname")
              '     RsTemp.MoveNext
              'Loop
              'If strExc(1) <> "" Then strExc(1) = convForm("本所案號", 15) & vbCrLf & strExc(1) '& "  代理人編號" & " 代理人名稱" & vbCrLf & strExc(1)
              If Val("" & RsTemp.Fields("CNT")) > 0 Then
                    strExc(2) = "目前案件有代理人設為" & IIf(textCU74.Text = "N", "年費自動代繳=Y", "年費不續辦=N") & _
                                 "，此申請人不可設定" & IIf(textCU74.Text = "Y", "年費自動代繳=Y", "年費不續辦=N") & "，請改在個案設定！"
                    MsgBox strExc(2), vbExclamation
                    
                    If textCU10 <> "" Then
                        strExc(0) = "select NA16,NA51 from nation where na01='" & textCU10 & "' "
                        intI = 1
                        Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                        If intI = 1 Then
                            strExc(3) = "" & RsTemp.Fields("NA16") 'FCP管制
                            strExc(4) = "" & RsTemp.Fields("NA51") 'FCP承辦
                            If strExc(4) = "" Then
                                strExc(4) = strExc(3): strExc(3) = ""
                            End If
                            '保留
                            'PUB_SendMail strUserNum, strExc(4), "", textCU01 & textCU02 & "，此申請人不可設定" & IIf(textCU74.Text = "Y", "年費自動代繳=Y", "年費不續辦=N") & "，請改在個案設定！", strExc(2) & vbCrLf & vbCrLf & strExc(1), , , , , , strExc(3)
                            PUB_SendMail strUserNum, strExc(4), "", textCU01 & textCU02 & "申請人設定與代理人設定有衝突，請確認", "同主旨", , , , , , strExc(3)
                        End If
                    End If
                    textCU74.Text = ""
              End If
          End If
      End If
      'end 2019/12/04
   End If
   CheckDataValid = True
EXITSUB:
End Function

' 更新Key的狀態
Private Sub SetKeyReadOnly(ByVal bEnable As Boolean)
   textCU01.Locked = bEnable
   textCU02.Locked = bEnable
End Sub

' 更新各控制項的狀態
Private Sub SetCtrlReadOnly(ByVal bEnable As Boolean)
   textCU01.Locked = bEnable
   textCU02.Locked = bEnable
   textCU03.Locked = bEnable
   textCU04.Locked = bEnable
   textCU05.Locked = bEnable
   textCU06.Locked = bEnable
   textCU07.Locked = bEnable
   textCU09.Locked = bEnable
   textCU10.Locked = bEnable
   textCU11.Locked = bEnable
   textCU12.Locked = bEnable
   textCU13.Locked = bEnable
   textCU14.Locked = bEnable
   textCU15.Locked = bEnable
   textCU16.Locked = bEnable
   textCU17.Locked = bEnable
   textCU18.Locked = bEnable
   textCU19.Locked = bEnable
   textCU20.Locked = bEnable
   textCU21.Locked = bEnable
   textCU22.Locked = bEnable
   textCU23.Locked = bEnable
   textCU24.Locked = bEnable
   textCU25.Locked = bEnable
   textCU26.Locked = bEnable
   textCU27.Locked = bEnable
   textCU28.Locked = bEnable
   textCU29.Locked = bEnable
   textCU30.Locked = bEnable
   textCU31.Locked = bEnable
   textCU32.Locked = bEnable
   textCU33.Locked = bEnable
   textCU34.Locked = bEnable
   textCU35.Locked = bEnable
   textCU36.Locked = bEnable
   textCU37.Locked = bEnable
   textCU38.Locked = bEnable
   textCU39.Locked = bEnable
   textCU40.Locked = bEnable
   textCU41.Locked = bEnable
   textCU42.Locked = bEnable
   textCU43.Locked = bEnable
   textCU44.Locked = bEnable
   textCU45.Locked = bEnable
   textCU46.Locked = bEnable
   textCU47.Locked = bEnable
   textCU48.Locked = bEnable
   textCU49.Locked = bEnable
   textCU50.Locked = bEnable
   textCU51.Locked = bEnable
   textCU52.Locked = bEnable
   textCU53.Locked = bEnable
   textCU54.Locked = bEnable
   textCU55.Locked = bEnable
   textCU56.Locked = bEnable
   textCU57.Locked = bEnable
   textCU58.Locked = bEnable
   textCU59.Locked = bEnable
   textCU60.Locked = bEnable
   textCU61.Locked = bEnable
   textCU62.Locked = bEnable
   textCU63.Locked = bEnable
   textCU64.Locked = bEnable
   textCU65.Locked = bEnable
   textCU66.Locked = bEnable
   textCU67.Locked = bEnable
   textCU68.Locked = bEnable
   textCU69.Locked = bEnable
   textCU70.Locked = bEnable
   textCU71.Locked = bEnable
   textCU72.Locked = bEnable
   textCU73.Locked = bEnable
   textCU74.Locked = bEnable
   textCU75.Locked = bEnable
'   textCU76.Locked = bEnable
   textCU77.Locked = bEnable
   textCU78.Locked = bEnable
   textCU79.Locked = bEnable
   cboStatus.Locked = bEnable
   textCU87.Locked = bEnable
   textCU88.Locked = bEnable
   textCU89.Locked = bEnable
   textCU90.Locked = bEnable
   textCU91.Locked = bEnable
   textCU92.Locked = bEnable
   textCU93.Locked = bEnable
   textCU94.Locked = bEnable
   textCU95.Locked = bEnable
   textCU96.Locked = bEnable
   textCU97.Locked = bEnable
   textCU98.Locked = bEnable
   textCU99.Locked = bEnable
   textCU100.Locked = bEnable
   textCU102.Locked = bEnable
   textCU103.Locked = bEnable
   textCU104.Locked = bEnable
   textCU105.Locked = bEnable
   textCU106.Locked = bEnable
   textCU107.Locked = bEnable
   textCU108.Locked = bEnable
   textCU109.Locked = bEnable
   
   'Add By Sindy 2025/3/10
   textCU203.Locked = bEnable
   textCU204.Locked = bEnable
   textCU205.Locked = bEnable
   '2025/3/10 END
   
   textCU111.Locked = bEnable
   textCU112.Locked = bEnable
   textCU113.Locked = bEnable 'Added by Morgan 2012/8/28
   textCU114.Locked = bEnable
   textCU115.Locked = True 'Add by Morgan 2008/1/16 財務信箱不能從這裡改
   textCU116.Locked = bEnable 'Add by Morgan 2008/1/16
   textCU117.Locked = bEnable 'Add by Morgan 2008/1/16
   textCU118.Locked = bEnable 'Add by Morgan 2008/1/16
   
   'Add by Morgan 2018/11/14 全E化客戶先不開放其他單位修改
   If Pub_StrUserSt03 = "M51" Then
      'Modified by Morgan 2025/2/27
      'textCU176.Locked = bEnable
      ''Added by Morgan 2021/10/7
      'textCU185.Locked = bEnable
      'textCU186.Locked = bEnable
      'textCU187.Locked = bEnable
      'textCU188.Locked = bEnable
      ''end 2021/10/7
      Frame2.Enabled = Not bEnable
      'end 2025/2/27
   Else
      'Modified by Morgan 2025/2/27
      'textCU176.Locked = True
      ''Added by Morgan 2021/10/7
      'textCU185.Locked = True
      'textCU186.Locked = True
      'textCU187.Locked = True
      'textCU188.Locked = True
      ''end 2021/10/7
      Frame2.Enabled = False
      'end 2025/2/27
   End If
   'end 2018/11/14
   
   textCU122.Locked = bEnable 'Add by Morgan 2007/10/26
   textCU123.Locked = bEnable 'Add by Morgan 2008/1/7
'   textCU125.Locked = bEnable 'Add By Sindy 2009/10/26
   
   TextCu128.Locked = bEnable 'Add by Toni 2008/10/21
   textCU132.Locked = bEnable '2008/12/9 add by sonia
   textCU145.Locked = bEnable 'Add By Sindy 2011/1/14
   'Add By Sindy 2011/3/4
   textCU146.Locked = bEnable
   textCU147.Locked = bEnable
'   textCU148.Locked = bEnable
   textCU149.Locked = bEnable
   textCU150.Locked = bEnable
   textCU151.Locked = bEnable
   textCU152.Locked = bEnable
   '2011/3/4 End
   textCU153.Locked = bEnable 'Add By Sindy 2011/3/17
   'Modified by Lydia 2022/12/20 改成「FCP提申急件預設組別」
   'textCU154.Locked = bEnable 'Added by Morgan 2012/8/20
   Combo4.Locked = bEnable
   textCU139.Locked = bEnable 'Add By Sindy 2013/8/15
   textCU180.Locked = bEnable '客戶狀態備註
   ChkID.Enabled = Not bEnable 'Add by Amy 2024/05/22 不提供ID
   textCU191.Locked = bEnable 'Add by Amy 2023/05/03 跨所同意主管(中文字)
   'Mark by Amy 2015/09/10 改為label顯示
'   'Add By Sindy 2013/11/19
'   If bEnable = False Then
'      If Pub_StrUserSt03 = "M51" Then textCU143.Locked = bEnable
'   Else
'      textCU143.Locked = bEnable
'   End If
'   '2013/11/19 END
   textCU144.Locked = True 'Add By Sindy 2013/12/17
   'Add by Amy 2022/06/20 客戶狀態,並非操作者權限的下拉選項內容時,鎖住客戶狀態及狀態備註欄,不可修改
   If Pub_StrUserSt03 <> "M51" And (cboStatus = "設為對造" Or cboStatus = "解除對造" Or cboStatus = "不再使用" Or cboStatus = "不得代理" _
        Or cboStatus = "不得代理專利" Or cboStatus = "不得代理商標" Or cboStatus = "宣告破產") Then
      cboStatus.Locked = True
      textCU180.Locked = True '客戶狀態備註
   End If
   
   'Add By Sindy 2013/1/17
   For i = 0 To 1
      Combo2(i).Locked = bEnable
      Combo3(i).Locked = bEnable
   Next i
   '2013/1/17 End
   
   'Add By Sindy 2012/5/24 個人或公司
   Frame1.Enabled = Not bEnable
   '2012/5/24 End
   
   'Add by Morgan 2008/11/13 改用陣列以免控制項超過且以後再新增欄位也不必改
   For Each m_Txt In txtCU
      m_Txt.Locked = bEnable
   Next
   'end 2008/11/13
   
   'Add by Morgan 2008/7/30 只有新增時候可修改聯絡人資料
   If m_EditMode = 1 Then
      cboContact.Enabled = True 'Locked = False
   Else
      cboContact.Enabled = False 'Locked = True
   End If
   
   Frame1K.Enabled = Not bEnable 'Add By Sindy 2025/1/6
End Sub

Private Sub ClearField()
Dim nIndex As Integer
   
   textCU01 = Empty
   textCU02 = Empty
   textCU03 = Empty
   'Add by Amy 2025/09/15 bug-修改有cu03資料存檔後新增Tag未清,會寫一筆Update fa03=null
   textCU03.Tag = Empty
   textCU04 = Empty
   textCU05 = Empty
   textCU06 = Empty
   textCU07 = Empty
   textCU09 = Empty
   textCU10 = Empty
   textCU10.Tag = Empty 'Add by Amy 2025/03/07
   m_CU10 = Empty                                  '2008/11/21 add by sonia
   textCU11 = Empty
   textCU12 = Empty
   Label30(3).Caption = Empty                      '2008/9/4 add by sonia
   textCU13 = Empty
   If m_EditMode = 1 Then
      textCU12.Tag = Empty     '2012/8/7 add by sonia
      textCU13.Tag = Empty     '2008/9/4 add by sonia
   End If
   textCU14 = Empty
   If m_EditMode = 1 Then textCU14 = strSrvDate(2)
   textCU15 = Empty
   textCU16 = Empty
   textCU17 = Empty
   textCU18 = Empty
   textCU19 = Empty
   textCU20 = Empty
   textCU21 = Empty
   textCU22 = Empty
   textCU23 = Empty
   textCU23.Tag = Empty 'Add by Amy 2016/12/20
   textCU24 = Empty
   textCU25 = Empty
   textCU26 = Empty
   textCU27 = Empty
   textCU28 = Empty
   textCU29 = Empty
   textCU30 = Empty
   textCU31 = Empty
   textCU32 = Empty
   textCU33 = Empty
   textCU34 = Empty
   textCU35 = Empty
   textCU36 = Empty
   textCU37 = Empty
   textCU38 = Empty
   textCU39 = Empty
   textCU40 = Empty
   textCU41 = Empty
   textCU42 = Empty
   textCU43 = Empty
   textCU44 = Empty
   textCU45 = Empty
   textCU46 = Empty
   textCU47 = Empty
   textCU48 = Empty
   textCU49 = Empty
   textCU50 = Empty
   textCU51 = Empty
   textCU52 = Empty
   textCU53 = Empty
   textCU54 = Empty
   textCU55 = Empty
   textCU56 = Empty
   textCU57 = Empty
   textCU58 = Empty
   textCU59 = Empty
   textCU60 = Empty
   textCU61 = Empty
   textCU62 = Empty
   textCU63 = Empty
   textCU64 = Empty
   textCU65 = Empty
   textCU66 = Empty
   textCU67 = Empty
   textCU68 = Empty
   textCU69 = Empty
   textCU70 = Empty
   textCU71 = Empty
   textCU72 = Empty
   textCU73 = Empty
   textCU74 = Empty
   textCU74.Tag = Empty 'Added by Lydia 2019/11/27
   textCU75 = Empty
'   textCU76 = Empty
   textCU77 = Empty
   textCU78 = Empty
   textCU79 = Empty
   cboStatus = Empty
   textCU87 = Empty
   textCU88 = Empty
   textCU89 = Empty
   textCU90 = Empty
   textCU91 = Empty
   textCU92 = Empty
   textCU93 = Empty
   textCU94 = Empty
   textCU95 = Empty
   textCU96 = Empty
   textCU97 = Empty
   textCU98 = Empty
   textCU99 = Empty
   textCU100 = Empty
   textCU102 = Empty
   textCU103 = Empty
   textCU104 = Empty
   textCU105 = Empty
   textCU106 = Empty
   textCU107 = Empty
   textCU108 = Empty
   textCU109 = Empty
   
   'Add By Sindy 2025/3/10
   textCU203 = Empty
   textCU204 = Empty
   textCU205 = Empty
   '2025/3/10 END
   
   textCU111 = Empty
   textCU112 = Empty
   textCU113 = Empty 'Added by Morgan 2012/8/28
   textCU114 = Empty
   textCU115 = Empty 'Add by Morgan 2008/1/16
   textCU116 = Empty 'Add by Morgan 2008/1/16
   textCU117 = Empty 'Add by Morgan 2008/1/16
   textCU118 = Empty 'Add by Morgan 2008/1/16
   textCU176 = Empty 'Add by Morgan 2018/11/14
   textCU185 = Empty 'Add by Morgan 2021/10/7
   'textCU186 = Empty 'Add by Morgan 2021/10/7 'Removed by Morgan 2025/2/27
   textCU187 = Empty 'Add by Morgan 2021/10/7
   textCU188 = Empty 'Add by Morgan 2021/10/7
   
   textCU122 = Empty 'Add by Morgan 2007/10/26
   textCU123 = Empty 'Add by Morgan 2008/1/7
   textCU125 = Empty 'Add By Sindy 2009/10/26
   
   TextCu128 = Empty 'add by Toni 2008/10/21
   textCU132 = Empty '2008/12/9 add by sonia
   textCU145 = Empty 'Add By Sindy 2011/1/14
   'Add By Sindy 2011/3/4
   textCU146 = Empty
   textCU147 = Empty
'   textCU148 = Empty
   textCU149 = Empty
   textCU150 = Empty
   textCU151 = Empty
   textCU152 = Empty
   '2011/3/4 End
   textCU153 = Empty 'Add By Sindy 2011/3/17
   'Modified by Lydia 2022/12/20 改成「FCP提申急件預設組別」
   'textCU154 = Empty 'Added by Morgan 2012/8/20
   Combo4.Text = ""
   textCU139 = Empty 'Add By Sindy 2013/8/15
   'Modify by Amy 2015/09/10 改為Label顯示
   'textCU143 = Empty 'Add By Sindy 2013/11/19
   lblCU143.Caption = Empty
   'end 2015/09/10
   textCU144 = Empty 'Add By Sindy 2013/12/17
   LblCU144 = "(N:不開發票)" 'Add By Sindy 2023/9/4
   textCU180 = Empty 'Add by Amy 2019/08/27 客戶狀態備註
   ChkID.Value = 0  'Add by Amy 2024/05/22 不提供ID
   textCU191 = Empty 'Add by Amy 2023/05/03 跨所同意主管
   
   'Add by Morgan 2008/11/13 改用陣列以免控制項超過且以後再新增欄位也不必改
   For Each m_Txt In txtCU
      m_Txt = Empty
   Next
   lstDeveloper.Clear
   'end 2008/11/13
   
   textCU10_Validate False
   textCU09_Validate False
   textCU13_Validate False
   textCU12_Validate False
   textCU87_Validate False
   textCU57_Validate False
   textCU71_Validate False
   textCU94_Validate False
   textCU96_Validate False
   textCU97_Validate False
   textCU98_Validate False
   textCU99_Validate False
   textCU105_Validate False
   textCU106_Validate False
   'Add By Sindy 2011/3/4
   textCU147_Validate False
   textCU151_Validate False
   textCU152_Validate False
   '2011/3/4 End
   
   'Add By Sindy 2013/1/17
   For i = 0 To 1
      Me.Combo2(i).ListIndex = 0
      Me.Combo3(i).ListIndex = 0
   Next i
   '2013/1/17 End
   
   For nIndex = 0 To TF_CU - 1 'edit by nickc 2006/10/24  MAX_FIELD - 1
      m_FieldList(nIndex).fiOldData = Empty
      m_FieldList(nIndex).fiNewData = Empty
   Next nIndex
   'add by nickc 2006/03/17
   textCUID = ""
   
   cboContact.Clear 'Add by Morgan 2008/7/30
   
   'Add By Sindy 2012/6/12
   For nIndex = 0 To 3
      optCustomer(nIndex).Value = False
      If nIndex = 0 Then optCustomer(nIndex).Tag = "" 'Add by Amy 2015/10/20
   Next nIndex
   optCustomer(1).Value = True '預設為公司 Add By Sindy 2012/7/12
   '2012/6/12 End
   
   m_CU15 = Empty 'Add By Sindy 2013/12/12
   
   'Add By Sindy 2025/1/6
   For Each m_Txt In Chk1K
      m_Txt = Empty
   Next
   '2025/1/6 END
   
   'Added by Morgan 2025/2/27
   For Each m_Txt In ChkCU186
      m_Txt = Empty
   Next
   'end 2025/2/27
End Sub

Private Sub UpdateFieldNewData()
   '若新增資料
   If m_EditMode = 1 Then
        '若未輸入代理人編號
'      If IsEmptyText(textCU01) = True Then
'         textCU01 = GetNewAgentNo
'      End If
      If IsEmptyText(textCU02) = True Then
         textCU02 = "0"
      End If
   End If
   
   If IsEmptyText(textCU01) = False Then
      SetFieldNewData "CU01", textCU01 & String(8 - Len(textCU01), "0")
   Else
      SetFieldNewData "CU01", textCU01
   End If
   SetFieldNewData "CU02", textCU02
   ' 客戶編號
   If IsEmptyText(textCU03) = False Then
      SetFieldNewData "CU03", textCU03 & String(8 - Len(textCU03), "0")
   Else
      SetFieldNewData "CU03", textCU03
   End If
   SetFieldNewData "CU04", textCU04
   SetFieldNewData "CU05", textCU05
   SetFieldNewData "CU06", textCU06
   SetFieldNewData "CU07", textCU07
   SetFieldNewData "CU09", textCU09
   SetFieldNewData "CU10", textCU10
   SetFieldNewData "CU11", Trim(textCU11) 'Modify by Amy 2022/05/10 取代空白 (可能貼上,換行不用改,檢查時會取代)
   SetFieldNewData "CU12", textCU12
   SetFieldNewData "CU13", textCU13
   SetFieldNewData "CU14", DBDATE(textCU14)
   If optCustomer(0).Value = True Then textCU15 = "0"
   If optCustomer(1).Value = True Then textCU15 = "1"
   'Add By Sindy 2012/5/24
   If optCustomer(2).Value = True Then textCU15 = "2"
   If optCustomer(3).Value = True Then textCU15 = "3"
   '2012/5/24 End
   SetFieldNewData "CU15", textCU15
   SetFieldNewData "CU16", textCU16
   SetFieldNewData "CU17", textCU17
   SetFieldNewData "CU18", textCU18
   SetFieldNewData "CU19", textCU19
   SetFieldNewData "CU20", textCU20
   SetFieldNewData "CU21", textCU21
   SetFieldNewData "CU22", textCU22
   SetFieldNewData "CU23", textCU23
   SetFieldNewData "CU24", textCU24
   SetFieldNewData "CU25", textCU25
   SetFieldNewData "CU26", textCU26
   SetFieldNewData "CU27", textCU27
   SetFieldNewData "CU28", textCU28
   SetFieldNewData "CU29", textCU29
   SetFieldNewData "CU30", textCU30
   SetFieldNewData "CU31", textCU31
   SetFieldNewData "CU32", textCU32
   SeekNewCu32 = textCU32
   SetFieldNewData "CU33", textCU33
   SetFieldNewData "CU34", textCU34
   SetFieldNewData "CU35", textCU35
   SetFieldNewData "CU36", textCU36
   SetFieldNewData "CU37", textCU37
   SetFieldNewData "CU38", DBDATE(textCU38)
   SetFieldNewData "CU39", textCU39
   SetFieldNewData "CU40", textCU40
   SetFieldNewData "CU41", textCU41
   SetFieldNewData "CU42", textCU42
   SetFieldNewData "CU43", textCU43
   SetFieldNewData "CU44", textCU44
   SetFieldNewData "CU45", textCU45
   SetFieldNewData "CU46", textCU46
   SetFieldNewData "CU47", textCU47
   SetFieldNewData "CU48", textCU48
   SetFieldNewData "CU49", textCU49
   SetFieldNewData "CU50", textCU50
   SetFieldNewData "CU51", textCU51
   SetFieldNewData "CU52", textCU52
   SetFieldNewData "CU53", textCU53
   SetFieldNewData "CU54", textCU54
   SetFieldNewData "CU55", textCU55
   SetFieldNewData "CU56", textCU56
   If IsEmptyText(textCU57) = False Then
      SetFieldNewData "CU57", textCU57 & String(9 - Len(textCU57), "0")
   Else
      SetFieldNewData "CU57", textCU57
   End If
   SetFieldNewData "CU58", textCU58
   SetFieldNewData "CU59", textCU59
   SetFieldNewData "CU60", textCU60
   SetFieldNewData "CU61", textCU61
   SetFieldNewData "CU62", textCU62
   SetFieldNewData "CU63", textCU63
   SetFieldNewData "CU64", textCU64
   SetFieldNewData "CU65", textCU65
   SetFieldNewData "CU66", textCU66
   SetFieldNewData "CU67", textCU67
   SetFieldNewData "CU68", textCU68
   SetFieldNewData "CU69", textCU69
    'Add By Cheng 2003/09/23
    'Begin
   SetFieldNewData "CU70", textCU70
   If IsEmptyText(textCU71) = False Then
      'edit by nickc 2007/03/01 改 9 碼
      'SetFieldNewData "CU71", textCU71 & String(8 - Len(textCU71), "0")
      SetFieldNewData "CU71", textCU71 & String(9 - Len(textCU71), "0")
   Else
      SetFieldNewData "CU71", textCU71
   End If
   SetFieldNewData "CU72", textCU72
   SetFieldNewData "CU73", textCU73
   SetFieldNewData "CU74", textCU74
   SetFieldNewData "CU75", textCU75
   'Modify By Sindy 2013/1/17
'   SetFieldNewData "CU76", textCU76
   SetFieldNewData "CU76", Combo2(0).Text
   '2013/1/17 End
   SetFieldNewData "CU77", textCU77
   SetFieldNewData "CU78", textCU78
  If m_EditMode = 1 Then
      If Trim(SeekNewCu32) <> "" Then
         textCU79 = textCU79 & ";不寄雜誌日期：" & strSrvDate(2) & ";"
      End If
   ElseIf m_EditMode = 2 Then
      If Trim(SeekOldCu32) = "" And Trim(SeekNewCu32) <> "" Then
         textCU79 = textCU79 & ";不寄雜誌日期：" & strSrvDate(2) & ";"
      End If
   End If
   SetFieldNewData "CU79", textCU79
   SetFieldNewData "CU80", cboStatus
   SetFieldNewData "CU87", textCU87
   SetFieldNewData "CU88", textCU88
   SetFieldNewData "CU89", textCU89
   SetFieldNewData "CU90", textCU90
   SetFieldNewData "CU91", textCU91
   SetFieldNewData "CU92", textCU92
   SetFieldNewData "CU93", textCU93
   If IsEmptyText(textCU94) = False Then
      'edit by nickc 2007/03/01 改 9 碼
      'SetFieldNewData "CU94", textCU94 & String(8 - Len(textCU94), "0")
      SetFieldNewData "CU94", textCU94 & String(9 - Len(textCU94), "0")
   Else
      SetFieldNewData "CU94", textCU94
   End If
   SetFieldNewData "CU95", textCU95
   If IsEmptyText(textCU96) = False Then
      'edit by nickc 2007/03/01 改 9 碼
      'SetFieldNewData "CU96", textCU96 & String(8 - Len(textCU96), "0")
      SetFieldNewData "CU96", textCU96 & String(9 - Len(textCU96), "0")
   Else
      SetFieldNewData "CU96", textCU96
   End If
   If IsEmptyText(textCU97) = False Then
      'edit by nickc 2007/03/01 改 9 碼
      'SetFieldNewData "CU97", textCU97 & String(8 - Len(textCU97), "0")
      SetFieldNewData "CU97", textCU97 & String(9 - Len(textCU97), "0")
   Else
      SetFieldNewData "CU97", textCU97
   End If
   If IsEmptyText(textCU98) = False Then
      'edit by nickc 2007/03/01 改 9 碼
      'SetFieldNewData "CU98", textCU98 & String(8 - Len(textCU98), "0")
      SetFieldNewData "CU98", textCU98 & String(9 - Len(textCU98), "0")
   Else
      SetFieldNewData "CU98", textCU98
   End If
   If IsEmptyText(textCU99) = False Then
      'edit by nickc 2007/03/01 改 9 碼
      'SetFieldNewData "CU99", textCU99 & String(8 - Len(textCU99), "0")
      SetFieldNewData "CU99", textCU99 & String(9 - Len(textCU99), "0")
   Else
      SetFieldNewData "CU99", textCU99
   End If
   SetFieldNewData "CU100", textCU100
   SetFieldNewData "CU102", textCU102
   SetFieldNewData "CU103", textCU103
   SetFieldNewData "CU104", textCU104
   If IsEmptyText(textCU105) = False Then
      SetFieldNewData "CU105", textCU105 & String(9 - Len(textCU105), "0")
   Else
      SetFieldNewData "CU105", textCU105
   End If
   If IsEmptyText(textCU106) = False Then
      SetFieldNewData "CU106", textCU106 & String(9 - Len(textCU106), "0")
   Else
      SetFieldNewData "CU106", textCU106
   End If
   SetFieldNewData "CU107", textCU107
   SetFieldNewData "CU108", textCU108
   SetFieldNewData "CU109", DBDATE(textCU109)
   
   'Add By Sindy 2025/3/10
   SetFieldNewData "CU203", textCU203
   SetFieldNewData "CU204", textCU204
   SetFieldNewData "CU205", DBDATE(textCU205)
   '2025/3/10 END
   
   SetFieldNewData "CU111", textCU111
   SetFieldNewData "CU112", textCU112
   SetFieldNewData "CU113", textCU113 'Added by Morgan 2012/8/28
   SetFieldNewData "CU114", textCU114
   'SetFieldNewData "CU115", textCU115 'Add by Morgan 2008/1/16 Modify By Sindy 2018/3/16 Mark.不可在此作業異動該欄位值
   SetFieldNewData "CU116", textCU116 'Add by Morgan 2008/1/16
   SetFieldNewData "CU117", textCU117 'Add by Morgan 2008/1/16
   SetFieldNewData "CU118", textCU118 'Add by Morgan 2008/1/16
   SetFieldNewData "CU176", textCU176 'Add by Morgan 2018/11/14
   SetFieldNewData "CU185", textCU185 'Add by Morgan 2021/10/7
   'SetFieldNewData "CU186", textCU186 'Add by Morgan 2021/10/7 'Removed by Morgan 2025/2/27
   SetFieldNewData "CU187", textCU187 'Add by Morgan 2021/10/7
   SetFieldNewData "CU188", textCU188 'Add by Morgan 2021/10/7
   
   SetFieldNewData "CU122", textCU122 'Add by Morgan 2007/10/26
   SetFieldNewData "CU123", textCU123 'Add by Morgan 2008/1/7
'   SetFieldNewData "CU125", textCU125 'Add By Sindy 2009/10/26
   
   SetFieldNewData "CU128", TextCu128 'add by Toni 2008/10/21
   SetFieldNewData "CU132", textCU132 '2008/12/9 add by sonia
   SetFieldNewData "CU145", textCU145 'Add By Sindy 2011/1/14
   'Add By Sindy 2011/3/4
   SetFieldNewData "CU146", textCU146
   If IsEmptyText(textCU147) = False Then
      SetFieldNewData "CU147", textCU147 & String(9 - Len(textCU147), "0")
   Else
      SetFieldNewData "CU147", textCU147
   End If
   'Modify By Sindy 2013/1/17
'   SetFieldNewData "CU148", textCU148
   SetFieldNewData "CU148", Combo2(1).Text
   '2013/1/17 End
   SetFieldNewData "CU149", textCU149
   SetFieldNewData "CU150", textCU150
   If IsEmptyText(textCU151) = False Then
      SetFieldNewData "CU151", textCU151 & String(9 - Len(textCU151), "0")
   Else
      SetFieldNewData "CU151", textCU151
   End If
   If IsEmptyText(textCU152) = False Then
      SetFieldNewData "CU152", textCU152 & String(9 - Len(textCU152), "0")
   Else
      SetFieldNewData "CU152", textCU152
   End If
   '2011/3/4 End
   SetFieldNewData "CU153", textCU153 'Add By Sindy 2011/3/17
   'Modified by Lydia 2022/12/20 改成「FCP提申急件預設組別」
   'SetFieldNewData "CU154", textCU154 'Added by Morgan 2012/8/20
   SetFieldNewData "CU154", Trim(Left(Combo4.Text, 1))
   SetFieldNewData "CU139", textCU139 'Add By Sindy 2013/8/15
   'Mark by Amy 2015/09/10 改為label顯示
   'SetFieldNewData "CU143", textCU143 'Add By Sindy 2013/11/19
   'Add By Sindy 2013/12/12
   If optCustomer(2).Value = True Then '學校時,不可開立發票必須為N
      SetFieldNewData "CU144", "N"
   Else
      SetFieldNewData "CU144", textCU144 'Add By Sindy 2013/12/17
   End If
   '2013/12/12 END
   
   SetFieldNewData "CU156", IIf(Combo3(0).Text <> "", Combo3(0).ListIndex, "") 'Add By Sindy 2013/1/17 專利
   SetFieldNewData "CU157", IIf(Combo3(1).Text <> "", Combo3(1).ListIndex, "") 'Add By Sindy 2013/1/17 商標
   'Add by Amy 2023/05/16 新客戶建檔進入者(申請人1),收據公司為J公司需寫入對應欄位中
   If m_Crl49JCmp <> MsgText(601) Then
        strExc(1) = Replace(m_Crl49JCmp, ",'", "")
        strExc(1) = Mid(strExc(1), 1, InStr(strExc(1), "'") - 1)
        SetFieldNewData UCase(strExc(1)), lblCU16X(Val(Right(strExc(1), 1))) 'Right(strExc(1))=1 or 3 or 5
   End If
   SetFieldNewData "CU180", textCU180 'Add By Amy 2019/08/27 客戶狀態備註
   'Add by Amy 2024/05/22  不提供ID
   If ChkID.Value = 1 Then
      SetFieldNewData "CU182", "Y"
   ElseIf m_FieldList(181).fiOldData <> "W" And ChkID.Value = 0 Then
      SetFieldNewData "CU182", ""
   End If
   SetFieldNewData "CU191", textCU191 'Add By Amy 2023/05/03 跨所同意主管(中文字)
   
   'Add by Morgan 2008/11/13 改用陣列以免控制項超過且以後再新增欄位也不必改
   For Each m_Txt In txtCU
      SetFieldNewData "CU" & m_Txt.Index, m_Txt
   Next
   'end 2008/11/13
   
   'Add By Sindy 2025/1/6
   strExc(10) = ""
   For Each m_Txt In Chk1K
      If m_Txt.Value = 1 Then
         strExc(10) = strExc(10) & "," & m_Txt.Index + 1
      End If
   Next
   If strExc(10) <> "" Then strExc(10) = Mid(strExc(10), 2)
   SetFieldNewData "CU201", strExc(10)
   '2025/1/6 END
   
   'Added by Morgan 2025/2/27 全E化客戶特殊設定
   strExc(10) = ""
   For Each m_Txt In ChkCU186
      If m_Txt.Value = 1 Then
         strExc(10) = strExc(10) & "," & m_Txt.Index
      End If
   Next
   If strExc(10) <> "" Then strExc(10) = Mid(strExc(10), 2)
   SetFieldNewData "CU186", strExc(10)
   'end 2025/2/27
End Sub

' 初始化欄位陣列
Private Sub InitialField()
Dim nIndex As Integer
Dim strTmp As String
   
   ' 初始化欄位陣列
   For nIndex = 1 To TF_CU   'edit by nickc 2006/10/24  MAX_FIELD
      strTmp = Format(nIndex, "00")
      m_FieldList(nIndex - 1).fiName = "CU" & strTmp
      m_FieldList(nIndex - 1).fiOldData = Empty
      m_FieldList(nIndex - 1).fiNewData = Empty
      m_FieldList(nIndex - 1).fiType = 0 '文字型態
      Select Case nIndex
         Case 14, 35, 36, 37, 38, 82, 83, 85, 86, 107, 108, 109, 110:
            m_FieldList(nIndex - 1).fiType = 1 '數值型態
      End Select
   Next nIndex
End Sub

'Mark by Sindy 2017/03/09
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 取得國外代理人的性質
' Input : strAgent ==> 代理人的代碼
' Output : 傳回國外代理人的性質
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Public Function GetFAgentFA76(ByVal strAgent As String) As String
'Dim rsTmp As New ADODB.Recordset
'Dim strSql As String
'Dim strFA01 As String
'Dim strFA02 As String
'
'   ' 設定 KEY 值
'   strFA01 = Mid(strAgent, 1, 8)
'   If Len(strFA01) < 8 Then: strFA01 = strFA01 & String(8 - Len(strFA01), "0")
'   strFA02 = Mid(strAgent, 9, 1)
'   If strFA02 = Empty Then: strFA02 = "0"
'
'   ' 設定初始值
'   GetFAgentFA76 = Empty
'
'   ' 組成SQL語法
'   strSql = "SELECT * FROM FAGENT " & _
'            "WHERE FA01 = '" & strFA01 & "' AND " & _
'                  "FA02 = '" & strFA02 & "' "
'   rsTmp.CursorLocation = adUseClient
'   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
'
'   ' 判斷是否有資料
'   If rsTmp.RecordCount > 0 Then
'      If IsNull(rsTmp.Fields("FA76")) = False Then
'         GetFAgentFA76 = rsTmp.Fields("FA76")
'      End If
'   End If
'   rsTmp.Close
'   Set rsTmp = Nothing
'End Function

'Add by Morgan 2008/11/13
Private Sub txtCU_GotFocus(Index As Integer)
   If txtCU(Index).Enabled = True And txtCU(Index).Locked = False Then
      InverseTextBox txtCU(Index)
      CloseIme
   End If
End Sub

'Add by Morgan 2008/11/13
Private Sub txtCU_KeyPress(Index As Integer, KeyAscii As Integer)
   If txtCU(Index).Enabled = True And txtCU(Index).Locked = False Then
      Select Case Index
         'Modify by Morgan 2009/9/15 +133, 134, 135, 136
         Case 130, 131, 133, 134, 135, 136
            If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then
               KeyAscii = 0
               Beep
            End If
         'Add by Morgan 2009/9/15
         'Modify by Morgan 2011/9/23 +141
         'Modified by Morgan 2014/6/3
         'Case 124, 126, 137, 138, 141
         'Modified by Lydia 2017/11/30 +FCP是否電子送件(CU174)
         'Modified by Morgan 2019/1/25 +177
         'Modified by Morgan 2020/1/16 +182
         'Modified by Morgan 2024/1/30 -182
         'Modified by Morgan 2025/2/10 +202
         Case 137, 138, 141, 174, 177, 202
            KeyAscii = UpperCase(KeyAscii)
            If KeyAscii <> 89 And KeyAscii <> 8 Then
               KeyAscii = 0
               Beep
            End If
            
         'Added by Morgan 2014/6/3
         Case 124, 126
            KeyAscii = UpperCase(KeyAscii)
            If KeyAscii <> 89 And KeyAscii <> 68 And KeyAscii <> 8 Then
               KeyAscii = 0
               Beep
            End If
         'Added by Morgan 2022/12/1
         Case 189, 190
            If KeyAscii <> 49 And KeyAscii <> 50 And KeyAscii <> 8 Then
               KeyAscii = 0
               Beep
            End If
      End Select
   End If
End Sub

Private Sub txtCU_Validate(Index As Integer, Cancel As Boolean)
   Select Case Index
      Case 124, 137
         If (txtCU(124) = "" And txtCU(137) = "Y") Then
            MsgBox "【專利 EMail 同時寄紙本】為 Y 時，【專利以 EMail 通知】欄位也必須為 Y！"
            Cancel = True
            Exit Sub
         End If
      Case 126, 138
         If (txtCU(126) = "" And txtCU(138) = "Y") Then
            MsgBox "【商標 EMail 同時寄紙本】為 Y 時，【商標以 EMail 通知】欄位也必須為 Y！"
            Cancel = True
            Exit Sub
         End If
   End Select
End Sub

'Add By Sindy 2012/7/17 比對國內外潛在客戶名稱相同者寄Mail通知電腦中心
Private Sub ChkCustNameAndPotCust()
Dim rsTmp As New ADODB.Recordset
Dim strCustID As String, strContext As String
Dim bolModify As Boolean
Dim strST06 As String
   
   'Modify By Sindy 2016/12/6 sql裡加rtrim(),因有的客戶名稱最後面會有空格
   'Add By Sindy 2012/7/27
   bolModify = False
   If (textCU04.Tag <> "") And (textCU04.Tag <> CheckStr(textCU04)) Then
      bolModify = True
   End If
   If (textCU05.Tag & textCU88.Tag & textCU89.Tag & textCU90.Tag <> "") And (textCU05.Tag & textCU88.Tag & textCU89.Tag & textCU90.Tag <> textCU05 & textCU88 & textCU89 & textCU90) Then
      bolModify = True
   End If
   If (textCU06.Tag <> "") And (textCU06.Tag <> CheckStr(textCU06)) Then
      bolModify = True
   End If
   If bolModify = False And m_EditMode <> 1 Then Exit Sub
   '2012/7/27 End
   
   strCustID = "": strContext = ""
   '比對名稱
   If textCU04 <> "" Then
      strSql = "SELECT poc01||poc02" & _
                " FROM PotCustomer1" & _
               " WHERE rtrim(poc03)=rtrim('" & ChgSQL(textCU04) & "')"
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         rsTmp.MoveFirst
         Do While Not rsTmp.EOF
            If IsNull(rsTmp.Fields(0)) = False Then
               strCustID = strCustID & ",'" & rsTmp.Fields(0) & "'"
            End If
            rsTmp.MoveNext
         Loop
      End If
      rsTmp.Close
      strSql = "SELECT pcu01||pcu02" & _
                " FROM PotCustomer" & _
               " WHERE rtrim(pcu08)=rtrim('" & ChgSQL(textCU04) & "')"
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         rsTmp.MoveFirst
         Do While Not rsTmp.EOF
            If IsNull(rsTmp.Fields(0)) = False Then
               strCustID = strCustID & ",'" & rsTmp.Fields(0) & "'"
            End If
            rsTmp.MoveNext
         Loop
      End If
      rsTmp.Close
   End If
   If textCU05 <> "" Then
      strSql = "SELECT poc01||poc02" & _
                " FROM PotCustomer1" & _
               " WHERE rtrim(upper(poc23||' '||poc24||' '||poc25||' '||poc26))=rtrim('" & ChgSQL(UCase(Trim(Trim(textCU05) & " " & Trim(textCU88) & " " & Trim(textCU89) & " " & Trim(textCU90)))) & "')"
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         rsTmp.MoveFirst
         Do While Not rsTmp.EOF
            If IsNull(rsTmp.Fields(0)) = False Then
               strCustID = strCustID & ",'" & rsTmp.Fields(0) & "'"
            End If
            rsTmp.MoveNext
         Loop
      End If
      rsTmp.Close
      strSql = "SELECT pcu01||pcu02" & _
                " FROM PotCustomer" & _
               " WHERE rtrim(upper(pcu03||' '||pcu04||' '||pcu05||' '||pcu06))=rtrim('" & ChgSQL(UCase(Trim(Trim(textCU05) & " " & Trim(textCU88) & " " & Trim(textCU89) & " " & Trim(textCU90)))) & "')"
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         rsTmp.MoveFirst
         Do While Not rsTmp.EOF
            If IsNull(rsTmp.Fields(0)) = False Then
               strCustID = strCustID & ",'" & rsTmp.Fields(0) & "'"
            End If
            rsTmp.MoveNext
         Loop
      End If
      rsTmp.Close
   End If
   If textCU06 <> "" Then
      strSql = "SELECT poc01||poc02" & _
                " FROM PotCustomer1" & _
               " WHERE rtrim(poc27)=rtrim('" & ChgSQL(textCU06) & "')"
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         rsTmp.MoveFirst
         Do While Not rsTmp.EOF
            If IsNull(rsTmp.Fields(0)) = False Then
               strCustID = strCustID & ",'" & rsTmp.Fields(0) & "'"
            End If
            rsTmp.MoveNext
         Loop
      End If
      rsTmp.Close
      strSql = "SELECT pcu01||pcu02" & _
                " FROM PotCustomer" & _
               " WHERE rtrim(pcu07)=rtrim('" & ChgSQL(textCU06) & "')"
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         rsTmp.MoveFirst
         Do While Not rsTmp.EOF
            If IsNull(rsTmp.Fields(0)) = False Then
               strCustID = strCustID & ",'" & rsTmp.Fields(0) & "'"
            End If
            rsTmp.MoveNext
         Loop
      End If
      rsTmp.Close
   End If
   '有相同者寄Mail通知電腦中心
   If strCustID <> "" Then
      strCustID = Mid(strCustID, 2, Len(strCustID))
      'Modify By Sindy 2012/8/23
      'strContext = textCU01 & textCU02 & " 國籍：" & Label30(0) & vbCrLf
      strContext = Left(textCU01 & "00000000", 8) & Left(textCU02 & "0", 1) & " 國籍：" & Label30(0) & vbCrLf
      '2012/8/23 End
      strContext = strContext & "          中文名稱：" & textCU04 & vbCrLf
      strContext = strContext & "          英文名稱：" & Trim(Trim(textCU05) & " " & Trim(textCU88) & " " & Trim(textCU89) & " " & Trim(textCU90)) & vbCrLf
      strContext = strContext & "          日文名稱：" & textCU06 & vbCrLf
      strST06 = PUB_GetST06(textCU13)
      If strST06 = "1" Then
         strST06 = "北所"
      ElseIf strST06 = "2" Then
         strST06 = "中所"
      ElseIf strST06 = "3" Then
         strST06 = "南所"
      ElseIf strST06 = "4" Then
         strST06 = "高所"
      ElseIf strST06 = "5" Then
         strST06 = "其他"
      Else
         strST06 = ""
      End If
      strContext = strContext & "          智權人員：" & textCU13 & " " & Label30(2) & "　" & strST06 & IIf(ChkStaffST04(textCU13, False) = True, "　(已離職)", "") & vbCrLf & vbCrLf 'Add By Sindy 2013/12/30
      
      strSql = "SELECT poc01||poc02,na03,poc03,rtrim(poc23||' '||poc24||' '||poc25||' '||poc26),poc27,poc13||' '||st02,st04" & _
               " FROM PotCustomer1,nation,staff" & _
               " WHERE poc01||poc02 in(" & strCustID & ")" & _
                 " and poc04=na01(+) and poc13=st01(+)" & _
               " union" & _
               " SELECT pcu01||pcu02,na03,pcu08,rtrim(pcu03||' '||pcu04||' '||pcu05||' '||pcu06),pcu07,pcu38||' '||st02,st04" & _
               " FROM PotCustomer,nation,staff" & _
               " WHERE pcu01||pcu02 in(" & strCustID & ")" & _
                 " and pcu09=na01(+) and substr(pcu38,1,5)=st01(+)"
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         rsTmp.MoveFirst
         Do While Not rsTmp.EOF
            strContext = strContext & rsTmp.Fields(0) & " 國籍：" & "" & rsTmp.Fields(1) & vbCrLf
            strContext = strContext & "          中文名稱：" & "" & rsTmp.Fields(2) & vbCrLf
            strContext = strContext & "          英文名稱：" & "" & rsTmp.Fields(3) & vbCrLf
            strContext = strContext & "          日文名稱：" & "" & rsTmp.Fields(4) & vbCrLf
            strST06 = PUB_GetST06(Left(rsTmp.Fields(5), 5))
            If strST06 = "1" Then
               strST06 = "北所"
            ElseIf strST06 = "2" Then
               strST06 = "中所"
            ElseIf strST06 = "3" Then
               strST06 = "南所"
            ElseIf strST06 = "4" Then
               strST06 = "高所"
            ElseIf strST06 = "5" Then
               strST06 = "其他"
            Else
               strST06 = ""
            End If
            strContext = strContext & "          開發人員：" & "" & rsTmp.Fields(5) & "　" & strST06 & IIf(rsTmp.Fields(6) = "2", "　(已離職)", "") & vbCrLf & vbCrLf 'Add By Sindy 2013/12/30
            rsTmp.MoveNext
         Loop
      End If
      rsTmp.Close
      strContext = strContext & "請向智權人員確認是否相同, 並做後續處理." & vbCrLf & vbCrLf & vbCrLf & _
                                "注意：１若潛在客戶的智權人員為（副所長）時，必須將此Mail轉寄給副所長。" & vbCrLf & _
                                "　　　２潛在客戶的開發人員為（業務助理）時：" & vbCrLf & _
                                "　　　　向各所收文人員確認接洽單上的同仁介紹欄是否有註明是業務助理，若有則可改資料；" & vbCrLf & _
                                "　　　　若不是註明業務助理時，則請影印接洽單，並把客戶名稱圈起來，並標註為業務助理之潛在客戶，請秘書交總經理或主秘批示。"
      
      strCustID = Replace(strCustID, "'", "")
      PUB_SendMail strUserNum, "97038", "", Left(textCU01 & "00000000", 8) & Left(textCU02 & "0", 1) & " 與 " & strCustID & "名稱相同 通知 !", strContext
   End If
   Set rsTmp = Nothing
End Sub

'Add By Sindy 2018/1/5 新增非個人之國內客戶時,若已有相同名稱的資料,
'                      系統自動發信給特殊設定:(財務處總帳人員)檢查建檔
Private Sub ChkCustName()
Dim rsTmp As New ADODB.Recordset
Dim strCustID As String, strContext As String
Dim bolModify As Boolean
Dim strST06 As String
Dim strTo As String 'Add by Amy 2024/05/15
   
   If optCustomer(0).Value = True Then Exit Sub '個人不須檢查
   If textCU10 > "010" Then Exit Sub 'Add by Sindy 2018/01/31 非本國客戶不須檢查
   
   'sql裡加rtrim(),因有的客戶名稱最後面會有空格
   bolModify = False
   If (textCU04.Tag <> "") And (textCU04.Tag <> CheckStr(textCU04)) Then
      bolModify = True
   End If
   If (textCU05.Tag & textCU88.Tag & textCU89.Tag & textCU90.Tag <> "") And (textCU05.Tag & textCU88.Tag & textCU89.Tag & textCU90.Tag <> textCU05 & textCU88 & textCU89 & textCU90) Then
      bolModify = True
   End If
   If (textCU06.Tag <> "") And (textCU06.Tag <> CheckStr(textCU06)) Then
      bolModify = True
   End If
   If bolModify = False And m_EditMode <> 1 Then Exit Sub
   
   strCustID = "": strContext = ""
   '比對名稱
   If textCU04 <> "" Then
      strSql = "SELECT cu01||cu02" & _
                " FROM Customer" & _
               " WHERE rtrim(cu04)=rtrim('" & ChgSQL(textCU04) & "')" & _
               " and CU01||CU02<>'" & Left(textCU01 & "0000000", 8) & Left(textCU02 & "0", 1) & "'"
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         rsTmp.MoveFirst
         Do While Not rsTmp.EOF
            If IsNull(rsTmp.Fields(0)) = False Then
               strCustID = strCustID & ",'" & rsTmp.Fields(0) & "'"
            End If
            rsTmp.MoveNext
         Loop
      End If
      rsTmp.Close
   End If
   If textCU05 <> "" Then
      strSql = "SELECT cu01||cu02" & _
                " FROM Customer" & _
               " WHERE rtrim(upper(cu05||' '||cu88||' '||cu89||' '||cu90))=rtrim('" & ChgSQL(UCase(Trim(Trim(textCU05) & " " & Trim(textCU88) & " " & Trim(textCU89) & " " & Trim(textCU90)))) & "')" & _
               " and CU01||CU02<>'" & Left(textCU01 & "0000000", 8) & Left(textCU02 & "0", 1) & "'"
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         rsTmp.MoveFirst
         Do While Not rsTmp.EOF
            If IsNull(rsTmp.Fields(0)) = False Then
               strCustID = strCustID & ",'" & rsTmp.Fields(0) & "'"
            End If
            rsTmp.MoveNext
         Loop
      End If
      rsTmp.Close
   End If
   If textCU06 <> "" Then
      strSql = "SELECT cu01||cu02" & _
                " FROM Customer" & _
               " WHERE rtrim(cu06)=rtrim('" & ChgSQL(textCU06) & "')" & _
               " and CU01||CU02<>'" & Left(textCU01 & "0000000", 8) & Left(textCU02 & "0", 1) & "'"
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         rsTmp.MoveFirst
         Do While Not rsTmp.EOF
            If IsNull(rsTmp.Fields(0)) = False Then
               strCustID = strCustID & ",'" & rsTmp.Fields(0) & "'"
            End If
            rsTmp.MoveNext
         Loop
      End If
      rsTmp.Close
   End If
   '有相同者寄Mail通知財務處
   If strCustID <> "" Then
      strCustID = Mid(strCustID, 2, Len(strCustID))
      strContext = Left(textCU01 & "00000000", 8) & Left(textCU02 & "0", 1) & " 國籍：" & Label30(0) & vbCrLf
      strContext = strContext & "          中文名稱：" & textCU04 & vbCrLf
      strContext = strContext & "          英文名稱：" & Trim(Trim(textCU05) & " " & Trim(textCU88) & " " & Trim(textCU89) & " " & Trim(textCU90)) & vbCrLf
      strContext = strContext & "          日文名稱：" & textCU06 & vbCrLf
      strST06 = PUB_GetST06(textCU13)
      If strST06 = "1" Then
         strST06 = "北所"
      ElseIf strST06 = "2" Then
         strST06 = "中所"
      ElseIf strST06 = "3" Then
         strST06 = "南所"
      ElseIf strST06 = "4" Then
         strST06 = "高所"
      ElseIf strST06 = "5" Then
         strST06 = "其他"
      Else
         strST06 = ""
      End If
      strContext = strContext & "          智權人員：" & textCU13 & " " & Label30(2) & "　" & strST06 & IIf(ChkStaffST04(textCU13, False) = True, "　(已離職)", "") & vbCrLf & vbCrLf
      
      strSql = "SELECT CU01||CU02,na03,CU04,rtrim(CU05||' '||CU88||' '||CU89||' '||CU90),CU06,CU13||' '||st02,st04" & _
               " FROM Customer,nation,staff" & _
               " WHERE CU01||CU02 in(" & strCustID & ")" & _
                 " and CU10=na01(+) and CU13=st01(+)"
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         rsTmp.MoveFirst
         Do While Not rsTmp.EOF
            strContext = strContext & rsTmp.Fields(0) & " 國籍：" & "" & rsTmp.Fields(1) & vbCrLf
            strContext = strContext & "          中文名稱：" & "" & rsTmp.Fields(2) & vbCrLf
            strContext = strContext & "          英文名稱：" & "" & rsTmp.Fields(3) & vbCrLf
            strContext = strContext & "          日文名稱：" & "" & rsTmp.Fields(4) & vbCrLf
            strST06 = PUB_GetST06(Left(rsTmp.Fields(5), 5))
            If strST06 = "1" Then
               strST06 = "北所"
            ElseIf strST06 = "2" Then
               strST06 = "中所"
            ElseIf strST06 = "3" Then
               strST06 = "南所"
            ElseIf strST06 = "4" Then
               strST06 = "高所"
            ElseIf strST06 = "5" Then
               strST06 = "其他"
            Else
               strST06 = ""
            End If
            strContext = strContext & "          開發人員：" & "" & rsTmp.Fields(5) & "　" & strST06 & IIf(rsTmp.Fields(6) = "2", "　(已離職)", "") & vbCrLf & vbCrLf 'Add By Sindy 2013/12/30
            rsTmp.MoveNext
         Loop
      End If
      rsTmp.Close
      strCustID = Replace(strCustID, "'", "")
      'Modify by Amy 2024/05/15 財務2個特殊設定拆成3個
      If Val(strSrvDate(1)) >= Val(財務拆總帳出納國內應收啟用日) Then
          strTo = Pub_GetSpecMan("財務處應收處理人員")
      Else
         strTo = Pub_GetSpecMan("財務處總帳人員")
      End If
      PUB_SendMail strUserNum, strTo, "", "客戶名稱相同，請確認是否有資料需要調整！", strContext
      'end 2024/05/15
   End If
   Set rsTmp = Nothing
End Sub

'Modify By Sindy 2023/9/4 mark,秀玲說財務處自行維護
''Add By Sindy 2013/12/17
'Private Sub textCU144_GotFocus()
'   CloseIme
'   TextInverse textCU144
'End Sub
'Private Sub textCU144_KeyPress(KeyAscii As Integer)
'   KeyAscii = UpperCase(KeyAscii)
'   If KeyAscii <> 8 And KeyAscii <> Asc("N") Then
'      KeyAscii = 0
'      Beep
'   End If
'End Sub
'Private Sub textCU144_Validate(Cancel As Boolean)
'   If m_EditMode <> 1 And m_EditMode <> 2 Then Exit Sub
'   If textCU144.Text = "" Then Exit Sub
'   If textCU144.Text <> "N" Then
'      ShowMsg "輸入錯誤 !"
'      Cancel = True
'   End If
'End Sub
''2013/12/17 End

Private Function CheckAddrData(ByRef objTxt As Object, ByRef objZip As Object, ByRef objCountry As Object) As Boolean
    Dim strZipCode As String, strAddr As String, strCountry As String, strIndArea As String, strROC As String
    Dim bolMany As Boolean, intArea As Integer, strMsg As String
    
    CheckAddrData = False

    Select Case UCase(objTxt.Name)
        Case "TEXTCU23"
            strMsg = "中文地址"
          
        Case "TEXTCU31"
            strMsg = "聯絡地址"
    End Select
    
    'Modify by Amy 2025/06/30 +objCountry.Text
    objTxt.Text = ReplaceAddrTW(objTxt.Text, , objCountry.Text)
  
    strROC = ""
    strAddr = objTxt.Text
    If Left(strAddr, 4) = "中華民國" Then strROC = strROC & Left(strAddr, 4): strAddr = Mid(strAddr, 5)
    If Left(strAddr, 3) = "臺灣省" Or Left(strAddr, 3) = "台灣省" Then strROC = strROC & Left(strAddr, 3): strAddr = Mid(strAddr, 4)
    If Left(strAddr, 2) = "臺灣" Or Left(strAddr, 2) = "台灣" Then strROC = strROC & Left(strAddr, 2): strAddr = Mid(strAddr, 3)
    '去除xx工業區查(台中工業區/台塑工業園區不取代,可能抓錯zip)
    strIndArea = "True"
    strAddr = ReplaceIndArea(strAddr, strIndArea)
    If strIndArea = "True" Then strIndArea = MsgText(601)
    If Left(strAddr, 4) = "新竹新竹" And (strIndArea = "科學工業園區" Or strIndArea = "科學園區") Then
        strIndArea = "新竹" & strIndArea
        strAddr = Mid(strAddr, 3)
    End If
    'Modify by Amy 2020/09/10 傳7個字檢查 ex:高雄市那瑪夏區
    intArea = 7
    strZipCode = GetPostZip(Left(strAddr, 7), 7, , strCountry, bolMany)
    '傳入地址前6個字取郵遞區號
    If strZipCode = MsgText(601) Then strZipCode = GetPostZip(Left(strAddr, 6), 6, , strCountry, bolMany): intArea = 6
    'end 2020/09/10
    '傳入地址前5個字取郵遞區號
    If strZipCode = MsgText(601) Then strZipCode = GetPostZip(Left(strAddr, 5), 5, , strCountry, bolMany): intArea = 5
    If InStr(strZipCode, Left(objZip, 3)) = 0 Then MsgBox "地址對應之郵遞區號有誤請確認！": objTxt.SetFocus: Exit Function
    'Modify by Amy 2024/06/17 +地址國籍為台灣或空白,與ZipCode 國籍不同才更正
    If (UCase(objTxt.Name) = "TEXTCU23" And (Trim(textCU10) < "010" Or Trim(textCU10) = MsgText(601))) Or UCase(objTxt.Name) = "TEXTCU31" Then
      'Modify by Amy 2025/03/07 +strMsg避免不知中文or聯絡地址
      If strCountry <> objCountry Then MsgBox strMsg & "對應之國籍有誤請確認！":  objTxt.SetFocus: Exit Function
    End If
    
    If CheckTaiwanAddr(objTxt, "000", strMsg) = False Then
        Call ChkZipData(9, objTxt, strZipCode, , strCountry)
    End If
        
    CheckAddrData = True
End Function

'Modify by Amy 2020/08/04 改抓Function
'比對名稱是對造者寄Mail通知電腦中心-秀玲
Private Sub ChkCustName2()
    Dim RsQ As New ADODB.Recordset
    Dim strQ As String, strFind As String
    Dim strSQL1 As String, strSQL2 As String, StrSQL3 As String, StrSQL4 As String, strSQL5 As String
    Dim ii As Integer, jj As Integer
    Dim bolModify As Boolean
    Dim SelField(1) As String, strTp(2) As String, strContext As String, strCheckWay As String
    Dim strRCLSql As String, strTxt As String 'Add by Amy 2024/01/31
    Dim bolCP As Boolean, bolRCL As Boolean, strSubject As String 'Add by Amy 2024/02/01
    Dim strTPContext As String, strOldTxt As String 'Add by Amy 2025/04/25
   
   bolModify = False
   If (textCU04.Tag <> "") And (textCU04.Tag <> CheckStr(textCU04)) Then
      bolModify = True
   End If
   If (textCU05.Tag & textCU88.Tag & textCU89.Tag & textCU90.Tag <> "") And (textCU05.Tag & textCU88.Tag & textCU89.Tag & textCU90.Tag <> textCU05 & textCU88 & textCU89 & textCU90) Then
      bolModify = True
   End If
   If (textCU06.Tag <> "") And (textCU06.Tag <> CheckStr(textCU06)) Then
      bolModify = True
   End If
   If bolModify = False And m_EditMode <> 1 Then Exit Sub
   
    strCheckWay = ">0"
    strSQL1 = " And CP01 IN (" & SQLGrpStr(GetGroupKindByTwo, 2) & ") "
    strSQL2 = " And CP01 IN (" & SQLGrpStr("", 1) & ") "
    StrSQL3 = " And CP01 IN (" & SQLGrpStr("", 3) & ") "
    StrSQL4 = " And CP01 IN (" & SQLGrpStr("", 4) & ") "
    strSQL5 = " And CP01 IN (" & SQLGrpStr("", 5) & ") "
    
    For ii = 0 To 2
        Select Case ii
            Case 0
                 strFind = ChgSQL(UCase(Trim(textCU04)))
                 SelField(0) = "CP40"
                 SelField(1) = "CP50"
            Case 1
                 strFind = ChgSQL(UCase(Trim(textCU05) & IIf(Trim(textCU88) <> MsgText(601), " " & Trim(textCU88), "") & _
                            IIf(Trim(textCU89) <> MsgText(601), " " & Trim(textCU89), "") & IIf(Trim(textCU90) <> MsgText(601), " " & Trim(textCU90), "")))
                 SelField(0) = "CP41"
                 SelField(1) = "CP51"
            Case 2
                 strFind = ChgSQL(UCase(Trim(textCU06)))
                 SelField(0) = "CP42"
                 SelField(1) = "CP52"
        End Select
        
        If Trim(strFind) <> MsgText(601) Then
            Call Pub_ProcR100102_1(strUserNum & "@" & Me.Name, strSQL1, strSQL2, StrSQL3, StrSQL4, strSQL5, strFind, strCheckWay)
            'Modify by Amy 2024/01/31 +風險檢查名單
            Call ChkRiskData(2, Me.Name, , , strFind, strRCLSql)
            '若狀態為2.其他相關人,但抓的資料為 cp40-42 仍不發mail,故只抓狀態1
            'strQ = "Select Distinct R021002 From R100102_1 Where ID='" & strUserNum & "@" & Me.Name & "' And R021004='1' "
            strQ = "Select Distinct R021002,'' as No,1 as Sort From R100102_1 Where ID='" & strUserNum & "@" & Me.Name & "' And R021004='1' "
            strQ = strQ & " Union Select INTTxt||'@@'||RCLField,RCL01 as No,2 as Sort From(" & strRCLSql & ") Order by Sort"
            
            If RsQ.State = adStateOpen Then RsQ.Close
            RsQ.CursorLocation = adUseClient
            RsQ.Open strQ, cnnConnection, adOpenStatic, adLockReadOnly
            If RsQ.RecordCount > 0 Then
               RsQ.MoveFirst
               strTp(0) = "": strTp(1) = "": strTp(2) = ""
               strTPContext = "" 'Add by Amy 2025/04/25
               Do While RsQ.EOF = False
                  'Modify by Amy 2025/04/25 換行有誤,修改顯示內文
                  Select Case ii
                     Case 0
                        strTp(0) = "中文名稱對應系統資料如下：" & vbCrLf
                     Case 1
                        strTp(0) = "英文名稱對應系統資料如下：" & vbCrLf
                     Case 2
                        strTp(0) = "日文名稱對應系統資料如下：" & vbCrLf
                  End Select
                  strTp(2) = "" & RsQ.Fields("R021002") '名稱+@@+CP4x
                  strTp(1) = Mid(strTp(2), 1, Val(InStr(strTp(2), "@@") - 1)) '取名稱
                  
                  '對造
                  If RsQ.Fields("Sort") = 1 Then
                     strTp(2) = Right(Replace(strTp(2), strTp(1), ""), 1) '取CP4x的x
                  '風險檢查
                  Else
                     strTp(2) = Replace(strTp(2), strTp(1) & "@@", "") '風險檢查欄位 中/英/日字樣
                  End If
                  '對造
                  'Modify by Amy 2025/04/25 換行有誤,原:strTxt = "　　　　　　　(與對造中文名稱相同) "
                  If RsQ.Fields("Sort") = 1 Then
                     bolCP = True 'Add by Amy 2024/02/01
                    Select Case strTp(2)
                        Case "0"
                            strTxt = "(與對造中文名稱相同) "
                        Case "1"
                            strTxt = "(與對造英文名稱相同) "
                        Case "2"
                            strTxt = "(與對造日文名稱相同) "
                    End Select
                  '風險檢查
                  Else
                     bolRCL = True 'Add by Amy 2024/02/01
                     Select Case strTp(2)
                        Case "中"
                            strTxt = "(與風險檢查資料中文名稱相同-編號" & RsQ.Fields("No") & ") "
                        Case "英"
                            strTxt = "(與風險檢查資料英文名稱相同-編號" & RsQ.Fields("No") & ") "
                        Case "日"
                            strTxt = "(與風險檢查資料日文名稱相同-編號" & RsQ.Fields("No") & ") "
                     End Select
                  End If
                  'end 2025/04/25
                  '目前客戶 中/英/日 字樣
                  'Modify by Amy 2025/04/25 換行有誤,最後再加換行,原:strContext
                  If InStr(strTPContext, strTp(0)) = 0 Then
                     strTPContext = strTPContext & strTp(0)
                  End If
                  '對應到的 名稱
                  '        ex:輸[陳淑]對應到多個時,最後二筆會顯示 (與對造中文名稱相同) 陳淑美 換行(與對造中文名稱相同) 被誤解無對應資料
                  '        ex:輸[程裕智]對應到對造及風險檢查
                  '對應到的 名稱
                  If strOldTxt = strTp(1) Then
                     strTPContext = strTPContext & strTxt
                  ElseIf InStr(strTPContext, strTp(1)) = 0 Then
                     If strTPContext <> MsgText(601) And strTPContext <> strTp(0) Then
                        strTPContext = strTPContext & vbCrLf
                     End If
                     strTPContext = strTPContext & strTp(1) & strTxt
                  End If
'                  strContext = strContext & vbCrLf & strTxt
                  strOldTxt = strTp(1)
                  'end 2025/04/25
                  RsQ.MoveNext
               Loop
               '中/英/日不止1個對應到以換行區隔
               If strTPContext <> MsgText(601) Then
                  strContext = strContext & strTPContext & vbCrLf & vbCrLf
               End If
            End If
            'end 2024/01/31
        End If
    Next ii
    If strContext <> "" Then
        strContext = "客戶編號：" & Left(textCU01 & "00000000", 8) & Left(textCU02 & "0", 1) & vbCrLf & _
                            "客戶名稱：中：" & textCU04 & vbCrLf & _
                            "　　　　　英：" & ChgSQL(UCase(Trim(textCU05) & IIf(Trim(textCU88) <> MsgText(601), " " & Trim(textCU88), "") & _
                            IIf(Trim(textCU89) <> MsgText(601), " " & Trim(textCU89), "") & IIf(Trim(textCU90) <> MsgText(601), " " & Trim(textCU90), ""))) & vbCrLf & _
                            "　　　　　日：" & textCU06 & vbCrLf & vbCrLf & _
                            strContext & vbCrLf & _
                            "PS：要副總等級以上主管同意 ：" & vbCrLf & _
                            "同意：請於客戶備註欄加註何人同意可收文。" & vbCrLf & _
                            "不同意：請通知專業部及智權人員，刪除客戶及收文案件。"
        'Modify by Amy 2024/02/01 依對應到之資料顯示主旨,原:新客戶為對造資料，請協助確認！
        If bolCP = True Then strSubject = strSubject & "對造"
        If bolRCL = True Then
            If strSubject <> MsgText(601) Then strSubject = strSubject & "及"
            strSubject = strSubject & "風險檢查"
        End If
        strSubject = "新客戶為" & strSubject & "資料，請協助確認！"
        'Modify by Amy 2023/11/16 原寄給83002
        PUB_SendMail strUserNum, Pub_GetSpecMan("程式管理人員"), "", strSubject, strContext
        'PUB_SendMail strUserNum, "A2004", "", strSubject, strContext '測式用
        'end 2024/02/01
    End If
End Sub

'Mark by Amy 2020/08/04 改抓Function 避免條件不一致
'比對名稱是對造者寄Mail通知電腦中心-秀玲
Private Sub ChkCustName2_Old()
'    Dim RsQ As New ADODB.Recordset
'    Dim strQ As String, strFind As String
'    Dim strSQL1 As String, strSQL2 As String, StrSQL3 As String, StrSQL4 As String, strSQL5 As String
'    Dim strSubSQL1 As String, strSubSQL2 As String, strSwhSQL1 As String, strSwhSQL2 As String
'    Dim SelField(1) As String '查詢欄位
'    Dim ii As Integer, jj As Integer
'    Dim bolModify As Boolean
'    Dim strTp(1) As String, strContext As String
'
'   bolModify = False
'   If (textCU04.Tag <> "") And (textCU04.Tag <> CheckStr(textCU04)) Then
'      bolModify = True
'   End If
'   If (textCU05.Tag & textCU88.Tag & textCU89.Tag & textCU90.Tag <> "") And (textCU05.Tag & textCU88.Tag & textCU89.Tag & textCU90.Tag <> textCU05 & textCU88 & textCU89 & textCU90) Then
'      bolModify = True
'   End If
'   If (textCU06.Tag <> "") And (textCU06.Tag <> CheckStr(textCU06)) Then
'      bolModify = True
'   End If
'   If bolModify = False And m_EditMode <> 1 Then Exit Sub
'
'    strSQL1 = " And CP01 IN (" & SQLGrpStr(GetGroupKindByTwo, 2) & ") "
'    strSQL2 = " And CP01 IN (" & SQLGrpStr("", 1) & ") "
'    StrSQL3 = " And CP01 IN (" & SQLGrpStr("", 3) & ") "
'    StrSQL4 = " And CP01 IN (" & SQLGrpStr("", 4) & ") "
'    strSQL5 = " And CP01 IN (" & SQLGrpStr("", 5) & ") "
'
'    For ii = 0 To 2
'        Select Case ii
'            Case 0
'                 strFind = ChgSQL(UCase(Trim(textCU04)))
'                 SelField(0) = "CP40"
'                 SelField(1) = "CP50"
'            Case 1
'                 strFind = ChgSQL(UCase(Trim(textCU05) & IIf(Trim(textCU88) <> MsgText(601), " " & Trim(textCU88), "") & _
'                            IIf(Trim(textCU89) <> MsgText(601), " " & Trim(textCU89), "") & IIf(Trim(textCU90) <> MsgText(601), " " & Trim(textCU90), "")))
'                 SelField(0) = "CP41"
'                 SelField(1) = "CP51"
'            Case 2
'                 strFind = ChgSQL(UCase(Trim(textCU06)))
'                 SelField(0) = "CP42"
'                 SelField(1) = "CP52"
'        End Select
'
'        If Trim(strFind) <> MsgText(601) Then
'            For jj = 0 To 2
'                Select Case jj
'                    Case 0
'                        '對造(中)
'                        strSubSQL1 = " And InStr(Upper(CP40),'" & strFind & "') >0 "
'                        strSubSQL2 = " And InStr(Upper(CP50),'" & strFind & "') >0 "
'                        strSwhSQL1 = " CP40>' ' "
'                        strSwhSQL2 = " CP50>' ' "
'                    Case 1
'                        '對造(英)
'                        strSubSQL1 = " And InStr(Upper(CP41),'" & strFind & "') >0 "
'                        strSubSQL2 = " And InStr(Upper(CP51),'" & strFind & "') >0 "
'                        strSwhSQL1 = " CP41>' ' "
'                        strSwhSQL2 = " CP51>' ' "
'                    Case 2
'                        '對造(日)
'                        strSubSQL1 = " And InStr(Upper(CP42),'" & strFind & "') >0 "
'                        strSubSQL2 = " And InStr(Upper(CP52),'" & strFind & "') >0 "
'                        strSwhSQL1 = " CP42>' ' "
'                        strSwhSQL2 = " CP52>' ' "
'                End Select
'
'                 '商標
'                 If ii > 0 Then strQ = strQ & " Union "
'                 strQ = "Select CP01||'-'||CP02||'-'||CP03||'-'||CP04 as 編號," & SelField(0) & " as 名稱,CP09 as 總收文號,NVL(Decode(TM10,'000',CPM03,CPM04),CP10) AS 案件性質 " & _
'                             "From (Select * From CaseProgress Where " & strSwhSQL1 & "),TradeMark,CasePropertyMap " & _
'                             "Where CP01=TM01(+) And CP02=TM02(+) And CP03=TM03(+) And CP04=TM04(+) And CP01=CPM01(+) And CP10=CPM02(+) " & strSQL1 & strSubSQL1
'                 strQ = strQ & " Union " & _
'                             "Select CP01||'-'||CP02||'-'||CP03||'-'||CP04 as 編號, " & SelField(1) & " as 名稱,CP09 as 總收文號,NVL(Decode(TM10,'000',CPM03,CPM04),CP10) AS 案件性質 " & _
'                             "From (Select * From CaseProgress Where " & strSwhSQL2 & "),TradeMark,CasePropertyMap " & _
'                             "Where CP01=TM01(+) And CP02=TM02(+) And CP03=TM03(+) And CP04=TM04(+) And CP01=CPM01(+) And CP10=CPM02(+) " & strSQL1 & strSubSQL2
'                 '專利
'                 strQ = strQ & " Union " & _
'                             "Select CP01||'-'||CP02||'-'||CP03||'-'||CP04 as 編號, " & SelField(0) & " as 名稱,CP09 as 總收文號,NVL(Decode(PA09,'000',CPM03,CPM04),CP10) AS 案件性質 " & _
'                             "From (Select * From CaseProgress Where " & strSwhSQL1 & "),Patent,CasePropertyMap " & _
'                             "Where CP01=PA01(+) And CP02=PA02(+) And CP03=PA03(+) And CP04=PA04(+) And CP01=CPM01(+) And CP10=CPM02(+) " & strSQL2 & strSubSQL1
'                 strQ = strQ & " Union " & _
'                             "Select CP01||'-'||CP02||'-'||CP03||'-'||CP04 as 編號, " & SelField(1) & " as 名稱,CP09 as 總收文號,NVL(Decode(PA09,'000',CPM03,CPM04),CP10) AS 案件性質 " & _
'                             "From (Select * From CaseProgress Where " & strSwhSQL2 & "),Patent,CasePropertyMap " & _
'                             "Where CP01=PA01(+) And CP02=PA02(+) And CP03=PA03(+) And CP04=PA04(+) And CP01=CPM01(+) And CP10=CPM02(+) " & strSQL2 & strSubSQL2
'                '法務
'                 strQ = strQ & " Union " & _
'                             "Select CP01||'-'||CP02||'-'||CP03||'-'||CP04 as 編號, " & SelField(0) & " as 名稱,CP09 as 總收文號,NVL(Decode(LC15,'000',CPM03,CPM04),CP10) AS 案件性質 " & _
'                             "From (Select * From CaseProgress Where " & strSwhSQL1 & "),LawCase,CasePropertyMap " & _
'                             "Where CP01=LC01(+) And CP02=LC02(+) And CP03=LC03(+) And CP04=LC04(+) And CP01=CPM01(+) And CP10=CPM02(+) " & StrSQL3 & strSubSQL1
'                 strQ = strQ & " Union " & _
'                             "Select CP01||'-'||CP02||'-'||CP03||'-'||CP04 as 編號, " & SelField(1) & " as 名稱,CP09 as 總收文號,NVL(Decode(LC15,'000',CPM03,CPM04),CP10) AS 案件性質 " & _
'                             "From (Select * From CaseProgress Where " & strSwhSQL2 & "),LawCase,CasePropertyMap " & _
'                             "Where CP01=LC01(+) And CP02=LC02(+) And CP03=LC03(+) And CP04=LC04(+) And CP01=CPM01(+) And CP10=CPM02(+) " & StrSQL3 & strSubSQL2
'                 '顧問
'                 strQ = strQ & " Union " & _
'                             "Select CP01||'-'||CP02||'-'||CP03||'-'||CP04 as 編號, " & SelField(0) & " as 名稱,CP09 as 總收文號,NVL(Decode(CPM03,null,CPM04,CPM03),CP10) AS 案件性質 " & _
'                             "From (Select * From CaseProgress Where " & strSwhSQL1 & "),HireCase,CasePropertyMap " & _
'                             "Where CP01=HC01(+) And CP02=HC02(+) And CP03=HC03(+) And CP04=HC04(+) And CP01=CPM01(+) And CP10=CPM02(+) " & StrSQL4 & strSubSQL1
'                 strQ = strQ & " Union " & _
'                             "Select CP01||'-'||CP02||'-'||CP03||'-'||CP04 as 編號, " & SelField(1) & " as 名稱,CP09 as 總收文號,NVL(Decode(CPM03,null,CPM04,CPM03),CP10) AS 案件性質 " & _
'                             "From (Select * From CaseProgress Where " & strSwhSQL2 & "),HireCase,CasePropertyMap " & _
'                             "Where CP01=HC01(+) And CP02=HC02(+) And CP03=HC03(+) And CP04=HC04(+) And CP01=CPM01(+) And CP10=CPM02(+) " & StrSQL4 & strSubSQL2
'                 '服務
'                 strQ = strQ & " Union " & _
'                             "Select CP01||'-'||CP02||'-'||CP03||'-'||CP04 as 編號, " & SelField(0) & " as 名稱,CP09 as 總收文號,NVL(Decode(SP09,'000',CPM03,CPM04),CP10) AS 案件性質 " & _
'                             "From (Select * From CaseProgress Where " & strSwhSQL1 & "),ServicePractice,CasePropertyMap " & _
'                             "Where CP01=SP01(+) And CP02=SP02(+) And CP03=SP03(+) And CP04=SP04(+) And CP01=CPM01(+) And CP10=CPM02(+) " & strSQL5 & strSubSQL1
'                 strQ = strQ & " Union " & _
'                             "Select CP01||'-'||CP02||'-'||CP03||'-'||CP04 as 編號, " & SelField(1) & " as 名稱,CP09 as 總收文號,NVL(Decode(SP09,'000',CPM03,CPM04),CP10) AS 案件性質 " & _
'                             "From (Select * From CaseProgress Where " & strSwhSQL2 & "),ServicePractice,CasePropertyMap " & _
'                             "Where CP01=SP01(+) And CP02=SP02(+) And CP03=SP03(+) And CP04=SP04(+) And CP01=CPM01(+) And CP10=CPM02(+) " & strSQL5 & strSubSQL2
'
'                If RsQ.State = adStateOpen Then RsQ.Close
'                RsQ.CursorLocation = adUseClient
'                RsQ.Open strQ, cnnConnection, adOpenStatic, adLockReadOnly
'                If RsQ.RecordCount > 0 Then
'                    strTp(0) = "": strTp(1) = ""
'                    Select Case ii
'                        Case 0
'                            strTp(0) = "客戶名稱 中文：" & Trim(textCU04)
'                        Case 1
'                            strTp(0) = "客戶名稱 英文：" & Trim(textCU05) & " " & Trim(textCU88) & " " & Trim(textCU89) & " " & Trim(textCU90)
'                        Case 2
'                            strTp(0) = "客戶名稱 日文：" & Trim(textCU06)
'                    End Select
'                    Select Case jj
'                        Case 0
'                            strTp(1) = " (與對造中文名稱相同) "
'                        Case 1
'                            strTp(1) = " (與對造英文名稱相同) "
'                        Case 2
'                            strTp(1) = " (與對造日文名稱相同) "
'                    End Select
'                    strContext = strContext & strTp(0) & strTp(1) & vbCrLf
'                End If
'            Next jj
'        End If
'    Next ii
'    If strContext <> "" Then
'        strContext = "客戶編號：" & Left(textCU01 & "00000000", 8) & Left(textCU02 & "0", 1) & vbCrLf & _
'                            "客戶名稱：中：" & textCU04 & vbCrLf & _
'                            "　　　　　英：" & ChgSQL(UCase(Trim(textCU05) & IIf(Trim(textCU88) <> MsgText(601), " " & Trim(textCU88), "") & _
'                            IIf(Trim(textCU89) <> MsgText(601), " " & Trim(textCU89), "") & IIf(Trim(textCU90) <> MsgText(601), " " & Trim(textCU90), ""))) & vbCrLf & _
'                            "　　　　　日：" & textCU06 & vbCrLf & vbCrLf & _
'                            strContext & vbCrLf & _
'                            "PS：要副總等級以上主管同意 ：" & vbCrLf & _
'                            "同意：請於客戶備註欄加註何人同意可收文。" & vbCrLf & _
'                            "不同意：請通知專業部及智權人員，刪除客戶及收文案件。"
'        PUB_SendMail strUserNum, "83002", "", "新客戶為對造資料，請協助確認！", strContext
'    End If
End Sub

'Add by Amy 2015/08/24 +檢查客戶狀態未改過且其他資料有改過彈訊息
Private Function ChkDataNotSave() As Boolean
    Dim idx As Integer, bolDifference As Boolean, bolAddrNotSame As Boolean
    
    ChkDataNotSave = False: bolDifference = False: bolAddrNotSame = False
    For idx = 0 To TF_CU - 1
        If idx < 80 Or idx > 85 Then
            'Modify by Amy 2016/04/13 +電腦中心改地址彈訊息
            If m_FieldList(idx).fiOldData <> m_FieldList(idx).fiNewData Then
                '檢查客戶狀態未改過且其他資料有改
                If bolDifference = False Then
                    If m_FieldList(79).fiOldData = m_FieldList(79).fiNewData And m_FieldList(79).fiNewData <> MsgText(601) Then
                        bolDifference = True
                        If Pub_StrUserSt03 <> "M51" Then Exit For
                    End If
                End If
                '檢查客戶地址是否有改
                If Pub_StrUserSt03 = "M51" And bolAddrNotSame = False Then
                    If (idx >= 22 And idx <= 28) Or idx = 101 Then
                        If m_FieldList(idx).fiOldData <> m_FieldList(idx).fiNewData And m_FieldList(idx).fiNewData <> MsgText(idx) Then
                            bolAddrNotSame = True
                        End If
                    End If
                End If
            End If
            'end 2016/04/13
        End If
    Next idx
    'Modify by Amy 2022/06/20 +客戶狀態開放才檢查(客戶狀態,並非操作者權限的下拉選項內容會鎖住)
    If bolDifference = True And cboStatus.Locked = False Then
        'Modify by Amy 2023/06/06 修改訊息
        If MsgBox("修改[非]客戶狀態欄位" & vbCrLf & _
                           "目前客戶狀態為 [" & cboStatus & "] 要修改？" & vbCrLf & _
                           "要修改按「是」，繼續操作按「否」" & vbCrLf & vbCrLf & _
                           "PS.客戶狀態為「" & Replace(stNotModStatus, ",", "、") & "」" & vbCrLf & _
                           "     除非選錯，否則不可任意修改", vbYesNo + vbDefaultButton2 + vbQuestion) = vbYes Then
            ChkDataNotSave = True
            tabCustomer.Tab = 0
            cboStatus.SetFocus
            Exit Function
        End If
    End If
    'Add by Amy 2016/04/13 +電腦中心改地址彈訊息
    If bolAddrNotSame = True Then
        MsgBox "電腦中心人員若代替檔案室修改地址資料，晚上每日批次作業不會檢查此客戶" & vbCrLf & _
                     "是否有CFT案件而通知承辦人及智權人員，請注意！", vbCritical, MsgText(5)
    End If
End Function

Private Sub ChkZipData(ByVal intChoose As Integer, ByRef objTxt As Object, Optional ByRef stZipCode As String = "", Optional ByRef intArea As Integer = 0, Optional ByRef stCountryCode As String = "")
    Dim objZipTxt As Object, objCountryTxt As Object '地址相對應Zip欄位/國籍欄位
    Dim strMsg As String, strAddr As String
    Dim intCount As Integer
    
     Select Case UCase(objTxt.Name)
        Case "TEXTCU23"
            strMsg = "中文地址"
            Set objZipTxt = textCU112
            Set objCountryTxt = textCU10
        Case "TEXTCU31"
            strMsg = "聯絡地址"
            Set objZipTxt = textCU30
            Set objCountryTxt = textCU87
        Case Else
    End Select
    
    Select Case intChoose
        Case 1 'ZipCode多筆(同區/鄉 ZipCode不同)
            '且與畫面上欄位資料前3碼不同或空值,彈郵遞區號查詢畫面
            If InStr(stZipCode, Left(Trim(objZipTxt), 3)) = 0 Or Trim(objZipTxt) = MsgText(601) Then
                If Trim(objZipTxt) <> MsgText(601) Then MsgBox strMsg & "郵遞區號有誤,請選擇正確郵遞區號！"
                Call frm100134.SetParent(Me)
                Me.Hide
                frm100134.BFormZip = objZipTxt.Name
                frm100134.BFormStatus = m_EditMode
                'Add by Amy 2016/10/20 +if
                If m_EditMode = 1 Then
                    frm100134.GetStreet objTxt.Text, 1, intArea, stZipCode
                Else
                    '修改時,原沒區不帶區但仍需判斷 zip是否正確
                    frm100134.GetStreet objTxt.Tag, 1, intArea, stZipCode
                End If
                'end 2016/10/20
                Call frm100134.QueryData
                frm100134.Show
                Exit Sub
            End If
        Case 2 'ZipCode非多筆
            '判斷抓到的郵遞區號是否與畫面上欄位資料前3碼相同
            If Left(objZipTxt.Text, 3) <> stZipCode Then
                If objZipTxt.Text <> MsgText(601) Then MsgBox strMsg & "郵遞區號有誤,系統將自動更正！", , MsgText(5)
                objZipTxt.Text = stZipCode
                Select Case UCase(objZipTxt.Name)
                    Case "TEXTCU30"
                        textCU30_GotFocus
                    Case "TEXTCU112"
                        textCU112_GotFocus
                End Select
            End If
            If objCountryTxt.Text <> MsgText(601) And stCountryCode <> objCountryTxt.Text Then
                MsgBox strMsg & "國籍有誤,系統將自動更正！", , MsgText(5)
                objCountryTxt.Text = stCountryCode
                Select Case UCase(objCountryTxt.Name)
                    Case "TEXTCU10"
                        textCU10_Validate (False)
                    Case "TEXTCU87"
                        textCU87_Validate (False)
                End Select
            End If
            Exit Sub
        Case 3, 4, 5 '抓不到ZipCode-3.區錯/4.只有路且郵遞區號為多筆/5.抓到2個字縣市,但多筆
            tabCustomer.Tab = 2
            MsgBox strMsg & "無法解析郵遞區號，請由下一畫面選取！"
            Call frm100134.SetParent(Me)
            Me.Hide
            frm100134.BFormZip = objZipTxt.Name
            frm100134.BFormStatus = m_EditMode
            'Add by Amy 2016/12/20 +if
            If m_EditMode = 1 Or UCase(objTxt.Name) <> "TEXTCU23" Then
                frm100134.GetStreet objTxt.Text, IIf(intChoose = 6, 2, intChoose), intArea, stZipCode
            Else
                '修改時,原沒區不帶區但仍需判斷 zip是否正確
                frm100134.GetStreet objTxt.Tag, IIf(intChoose = 6, 2, intChoose), intArea, stZipCode
            End If
            'end 2016/12/20
            Call frm100134.QueryData
            frm100134.Show
            Exit Sub
        Case 9 '設定頁籤
            tabCustomer.Tab = 2
            If objTxt.Enabled = True Then
                Select Case UCase(objTxt.Name)
                    Case "TEXTCU23"
                        textCU23.SetFocus
                        textCU23_GotFocus
                    Case "TEXTCU31"
                        textCU31.SetFocus
                        textCU31_GotFocus
                End Select
            End If
    End Select
End Sub

'Added by Lydia 2017/03/31 去掉跳行符號
'Modified by Lydia 2017/04/05 從lostfocus改成change
Private Sub textCU04_Change()
   'Modified by Lydia 2017/04/05 限定在新增或修改
   If m_EditMode = 1 Or m_EditMode = 2 Then textCU04 = PUB_StringFilter(textCU04)
End Sub

'Modified by Lydia 2017/04/05 從lostfocus改成change
Private Sub textCU05_Change()
   'Modified by Lydia 2017/04/05 限定在新增或修改
   If m_EditMode = 1 Or m_EditMode = 2 Then textCU05 = PUB_StringFilter(textCU05)
End Sub

'Modified by Lydia 2017/04/05 從lostfocus改成change
Private Sub textCU06_Change()
   'Modified by Lydia 2017/04/05 限定在新增或修改
   If m_EditMode = 1 Or m_EditMode = 2 Then textCU06 = PUB_StringFilter(textCU06)
End Sub

'Modified by Lydia 2017/04/05 從lostfocus改成change
Private Sub textCU88_Change()
   'Modified by Lydia 2017/04/05 限定在新增或修改
   If m_EditMode = 1 Or m_EditMode = 2 Then textCU88 = PUB_StringFilter(textCU88)
End Sub

'Modified by Lydia 2017/04/05 從lostfocus改成change
Private Sub textCU89_Change()
   'Modified by Lydia 2017/04/05 限定在新增或修改
   If m_EditMode = 1 Or m_EditMode = 2 Then textCU89 = PUB_StringFilter(textCU89)
End Sub

'Modified by Lydia 2017/04/05 從lostfocus改成change
Private Sub textCU90_Change()
   If m_EditMode = 1 Or m_EditMode = 2 Then textCU90 = PUB_StringFilter(textCU90)
End Sub

'Modified by Lydia 2017/04/05 從lostfocus改成change
Private Sub textCU23_Change()
   'Modified by Lydia 2017/04/05 限定在新增或修改
   If m_EditMode = 1 Or m_EditMode = 2 Then textCU23 = PUB_StringFilter(textCU23)
End Sub

'Modified by Lydia 2017/04/05 從lostfocus改成change
Private Sub textCU24_Change()
   If m_EditMode = 1 Or m_EditMode = 2 Then textCU24 = PUB_StringFilter(textCU24)
End Sub

'Modified by Lydia 2017/04/05 從lostfocus改成change
Private Sub textCU25_Change()
   'Modified by Lydia 2017/04/05 限定在新增或修改
   If m_EditMode = 1 Or m_EditMode = 2 Then textCU25 = PUB_StringFilter(textCU25)
End Sub

'Modified by Lydia 2017/04/05 從lostfocus改成change
Private Sub textCU26_Change()
   'Modified by Lydia 2017/04/05 限定在新增或修改
   If m_EditMode = 1 Or m_EditMode = 2 Then textCU26 = PUB_StringFilter(textCU26)
End Sub

'Modified by Lydia 2017/04/05 從lostfocus改成change
Private Sub textCU27_Change()
   'Modified by Lydia 2017/04/05 限定在新增或修改
   If m_EditMode = 1 Or m_EditMode = 2 Then textCU27 = PUB_StringFilter(textCU27)
End Sub

'Modified by Lydia 2017/04/05 從lostfocus改成change
Private Sub textCU28_Change()
   'Modified by Lydia 2017/04/05 限定在新增或修改
   If m_EditMode = 1 Or m_EditMode = 2 Then textCU28 = PUB_StringFilter(textCU28)
End Sub

'Modified by Lydia 2017/04/05 從lostfocus改成change
Private Sub textCU29_Change()
   'Modified by Lydia 2017/04/05 限定在新增或修改
   If m_EditMode = 1 Or m_EditMode = 2 Then textCU29 = PUB_StringFilter(textCU29)
End Sub

'Modified by Lydia 2017/04/05 從lostfocus改成change
Private Sub textCU102_Change()
   'Modified by Lydia 2017/04/05 限定在新增或修改
   If m_EditMode = 1 Or m_EditMode = 2 Then textCU102 = PUB_StringFilter(textCU102)
End Sub
'end 2017/03/31

'Added by Lydia 2017/05/09
Private Sub textCU10_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

'Added by Lydia 2017/06/14
Private Sub textCU59_Validate(Cancel As Boolean)
   If m_EditMode <> 1 And m_EditMode <> 2 Then Exit Sub
   
   If textCU59.Text = "" Then Exit Sub
   If Not CheckLengthIsOK(textCU59, 35) Then
      Cancel = True
   End If
End Sub

Private Sub textCU62_Validate(Cancel As Boolean)
   If m_EditMode <> 1 And m_EditMode <> 2 Then Exit Sub
   
   If textCU62.Text = "" Then Exit Sub
   If Not CheckLengthIsOK(textCU62, 35) Then
      Cancel = True
   End If
End Sub

'Added by Lydia 2019/04/16 檢查代表人欄位輸入順序
Private Function ChkCuPerson() As Boolean
Dim strMsg As String
    ChkCuPerson = False
    
    If Trim(textCU39 & textCU40 & textCU41) = "" And Trim(textCU42 & textCU43 & textCU44 & textCU45 & textCU46 & textCU47 & textCU48 & textCU49 & textCU50 & textCU51 & textCU52 & textCU53 & textCU54 & textCU55 & textCU56) <> "" Then
        strMsg = strMsg & "、代表人1"
    End If
    If Trim(textCU42 & textCU43 & textCU44) = "" And Trim(textCU45 & textCU46 & textCU47 & textCU48 & textCU49 & textCU50 & textCU51 & textCU52 & textCU53 & textCU54 & textCU55 & textCU56) <> "" Then
        strMsg = strMsg & "、代表人2"
    End If
    If Trim(textCU45 & textCU46 & textCU47) = "" And Trim(textCU48 & textCU49 & textCU50 & textCU51 & textCU52 & textCU53 & textCU54 & textCU55 & textCU56) <> "" Then
        strMsg = strMsg & "、代表人3"
    End If
    If Trim(textCU48 & textCU49 & textCU50) = "" And Trim(textCU51 & textCU52 & textCU53 & textCU54 & textCU55 & textCU56) <> "" Then
        strMsg = strMsg & "、代表人4"
    End If
    If Trim(textCU51 & textCU52 & textCU53) = "" And Trim(textCU54 & textCU55 & textCU56) <> "" Then
        strMsg = strMsg & "、代表人5"
    End If
    If strMsg <> "" Then
        MsgBox "請依照順序填入" & Mid(strMsg, 2) & "名稱！", vbCritical, "資料檢核"
        Me.tabCustomer.Tab = 3
        Exit Function
    End If
    
    ChkCuPerson = True
End Function

'Add by Amy 2023/05/03 判斷智權人員與客戶可收文所別(傳入聯絡／中文地址郵區號,一個符合即相同)不同時彈訊息
Private Function ChkPZD07(Optional ByVal IsModify As Boolean = False) As Boolean
    Dim strTpZip(1) As String
    Dim strCU10 As String 'Add by Amy 2024/06/28
    
    If Left(textCU12, 1) = "F" Then ChkPZD07 = True: Exit Function '業務區別為F開頭部門不需判斷
    
    ChkPZD07 = False
    If IsModify = True Then
'*** 修 改 ***
        If m_FieldList(190).fiOldData = MsgText(601) Then
            '若有改 智權人員 以畫面上為主-秀玲
            'Modify by Amy 2024/06/28 原程式搬至ChkAcrossArea,避免有未改到的
            strCU10 = textCU10
            If textCU10 < "010" And InStr(textCU79, "臺灣地址格式不檢查") > 0 Then strCU10 = "999" '國籍為台灣,但中文地址[非]台灣,備註會加註[臺灣地址格式不檢查]
            If ChkAcrossArea(1, Me.Name, textCU13, textCU112, textCU30, m_FieldList(111).fiOldData, m_FieldList(29).fiOldData, strCU10, textCU87, m_FieldList(9).fiOldData, m_FieldList(86).fiOldData) = True _
              And textCU191 = MsgText(601) Then
            'end 2024/06/28
                MsgBox "此為跨所客戶," & Replace(Label41(39), "：", "") & "不可為空！"
                textCU191.SetFocus
                Exit Function
            End If
        End If
    Else
'*** 新 增 ***
        If textCU10 < "010" Then
            strTpZip(0) = textCU112
        End If
        If textCU87 < "010" Then
            strTpZip(1) = textCU30
        End If
        If ChkAcrossArea(0, Me.Name, textCU13, strTpZip(0), strTpZip(1)) = True And textCU191 = MsgText(601) Then
            MsgBox "此為跨所收文," & Replace(Label41(39), "：", "") & "不可為空！"
            textCU191.SetFocus
            Exit Function
        End If
    End If
    
    '有 跨所同意主管 不需再彈訊息
    ChkPZD07 = True
  
End Function

'Mark by Amy 2023/05/03
Private Sub ChkPZD07_Old()
'    'Modify by Amy 2023/04/24 改抓共用function
'    Dim strTpZip(1) As String
''    Dim RsQ As New ADODB.Recordset
''    Dim intQ As Integer
''    Dim strQ As String, strSalesST06 As String
''    Dim bolPZD07Same As Boolean
''
''    strSalesST06 = PUB_GetST06(textCU13)
'    If textCU10 < "010" Then
''        strQ = "Select Distinct PZD07 From PostZipData Where Substr(PZD01,1,3)='" & Left(PUB_ChgNumeralStyle(textCU112), 3) & "' "
'        strTpZip(0) = textCU112
'    End If
'    If textCU87 < "010" Then
''        If strQ <> MsgText(601) Then strQ = strQ & "Union "
''        strQ = strQ & "Select Distinct PZD07 From PostZipData Where Substr(PZD01,1,3)='" & Left(PUB_ChgNumeralStyle(textCU30), 3) & "'"
'        strTpZip(1) = textCU30
'    End If
''
''    'Add by Amy 2020/09/09 若非台灣會error
''    If strQ = MsgText(601) Then Exit Sub
''
''    intQ = 1
''    Set RsQ = ClsLawReadRstMsg(intQ, strQ)
''    If intQ = 1 Then
''        Do While RsQ.EOF = False
''            If IsNull(RsQ.Fields("PZD07")) Then
''                bolPZD07Same = True
''            ElseIf InStr(RsQ.Fields("PZD07"), strSalesST06) > 0 Then
''                bolPZD07Same = True
''            End If
''            RsQ.MoveNext
''        Loop
'        'Modify by Amy 2022/09/01 排除業務區別為F開頭部門
'        'If bolPZD07Same = False And Left(textCU12, 1) <> "F" Then
'        If ChkAcrossArea(0, Me.Name, textCU13, strTpZip(0), strTpZip(1)) = True And Left(textCU12, 1) <> "F" Then
'            MsgBox "若為跨所收文，請記得於 參考備註 欄加註核可主管！"
'        End If
''    End If
''    Set RsQ = Nothing
End Sub

'Add by Amy 2022/10/06 從新客戶建檔進入
Private Sub ShowConsultRecApp()
    Dim RsQ As New ADODB.Recordset, intQ As Integer, strQ As String
    Dim strCRA01 As String, strCra03 As String, strCra11 As String
    
    m_CU07 = "": m_CU15 = ""
    strCRA01 = Replace(m_PrevNo, "Add ", "")
    strCra02 = Mid(strCRA01, InStr(strCRA01, "-") + 1)
    strCRA01 = Replace(strCRA01, "-" & strCra02, "")
    'Modify by Amy 2022/12/21 +crl51 客戶來源
    'Modify by Amy 2023/05/16 +m_Crl49JCmp 前畫面抓到的收據公司欄位
    strQ = "Select cra03,cra08,cra21,cra11,cra07 as CU04,cra09 as CU07,cra10 as CU127,cra12 as CU11,f0316 as CU13,cra13 as CU16,cra14 as CU21 " & _
                ",cra15 as CU18,cra16 as CU19,cra17 as CU20,cra18 as CU22,cra19 as Cu112,cra20 as CU23,cra22 as CU30,cra23 as CU31,cra24 as CU103,crl51 as cu09 " & m_Crl49JCmp & _
                "From ConsultRecApp,Flow003,ConsultRecordList " & _
                "Where cra01='" & strCRA01 & "' And cra02='" & strCra02 & "' And cra01=f0301(+) And f0302='3' And cra01=crl01(+) "
    intQ = 1
    Set RsQ = ClsLawReadRstMsg(intQ, strQ)
    If intQ = 1 Then
        strCra03 = "" & RsQ.Fields("cra03") '關係企業名稱
        strCra08 = "" & RsQ.Fields("cra08") '英文客戶名稱(不需預帶於畫面,因接洽單只有一個欄位不知如何切)
        strCra21 = "" & RsQ.Fields("cra21") '客戶英文地址
        strCra11 = "" & RsQ.Fields("cra11") '客戶國籍中文
        If IsNull(RsQ.Fields("CU04")) = False Then: textCU04 = RsQ.Fields("CU04") '中文客戶名稱
        If IsNull(RsQ.Fields("CU07")) = False Then: textCU07 = RsQ.Fields("CU07"): m_CU07 = RsQ.Fields("CU07") '公司負責人
        textCU10 = GetNationNo(strCra11) '國籍代碼
        'Modify by Amy 2025/03/06 +textCU10.Tag,避免一些不需彈的訊息被觸發
        '              ex:接洽單 1130021562,客戶國籍為香港,存檔時因m_FieldList(9).fiOldData =空,彈「修改客戶國籍,地址國籍是否同時修改...」,而修改成客戶國籍為台灣
        textCU10.Tag = textCU10
        'end 2025/03/06
        If IsNull(RsQ.Fields("CU11")) = False Then: textCU11 = RsQ.Fields("CU11") '申請人ID/統編
        If textCU11 <> MsgText(601) And Len(textCU11) > 8 Then
            optCustomer(0).Value = True '預設個人
            optCustomer(0).Tag = True
        End If
        If IsNull(RsQ.Fields("CU13")) = False Then: textCU13 = RsQ.Fields("CU13") '智權人員
        textCU12 = PUB_GetStaffST15(textCU13, 1)
        If IsNull(RsQ.Fields("CU16")) = False Then: textCU16 = RsQ.Fields("CU16") '電話
        If IsNull(RsQ.Fields("CU21")) = False Then: textCU21 = RsQ.Fields("CU21") 'Line ID
        If IsNull(RsQ.Fields("CU18")) = False Then: textCU18 = RsQ.Fields("CU18") 'Fax1
        If IsNull(RsQ.Fields("CU19")) = False Then: textCU19 = RsQ.Fields("CU19") 'Fax2
        If IsNull(RsQ.Fields("CU20")) = False Then: textCU20 = RsQ.Fields("CU20") '代表信箱
        If IsNull(RsQ.Fields("CU22")) = False Then: textCU22 = RsQ.Fields("CU22") 'Mobile Phone
        'Modify by Amy 2025/06/30 +textCU10,X90619 地址為中國浙江省台州市溫嶺市大溪區高速公路道口一級公路北側-->無法改成大溪鎮
        If IsNull(RsQ.Fields("CU23")) = False Then: textCU23 = ReplaceAddrTW(RsQ.Fields("CU23"), , textCU10) '客戶地址
        If IsNull(RsQ.Fields("CU112")) = False Then: textCU112 = RsQ.Fields("CU112") '客戶地址郵遞區號
        If IsNull(RsQ.Fields("CU30")) = False Then: textCU30 = RsQ.Fields("CU30") '聯絡地址郵遞區號
        'Modify by Amy 2025/06/30 +textCU10 接洽單無聯絡地址國籍,抓接洽單國籍(Cra11)-秀玲,若Cra11[是]台灣,聯絡地址[不是]也會被取代-->因為是少數,故先不管
        If IsNull(RsQ.Fields("CU31")) = False Then: textCU31 = ReplaceAddrTW(RsQ.Fields("CU31"), , textCU10) '聯絡地址
        If IsNull(RsQ.Fields("CU103")) = False Then: textCU103 = RsQ.Fields("CU103") '負責人英文名稱
        If IsNull(RsQ.Fields("CU127")) = False Then: cboContact = "" & RsQ.Fields("CU127") '接洽人
         'Add by Amy 2022/12/21客戶來源
        If IsNull(RsQ.Fields("CU09")) = False Then: textCU09 = RsQ.Fields("CU09")
        textCU09_Validate False
        'end 2022/12/21
        textCU10_Validate False
        textCU13_Validate False
        textCU12_Validate False
        'Add by Amy 2023/05/16 收據公司別欄位
        If m_Crl49JCmp <> MsgText(601) Then
            If "" & RsQ.Fields("Crl49JCmp") <> MsgText(601) Then
                lblCU16X(Val(Right("" & RsQ.Fields("Crl49JCmp"), 1))) = "J"
            End If
        End If
    End If
    Set RsQ = Nothing
End Sub

Private Function GetNationNo(ByVal stName As String) As String
    Dim rsA As New ADODB.Recordset, intA As Integer, strA As String
    
    GetNationNo = ""
    strA = "Select na01 From Nation Where na03='" & stName & "' And Length(na01)=3 "
    intA = 1
    Set rsA = ClsLawReadRstMsg(intA, strA)
    If intA = 1 Then
        GetNationNo = "" & rsA.Fields("na01")
    End If
    Set rsA = Nothing
End Function

'Modify by Amy 2022/10/06
Private Sub SetLock()
    '可能接洽單未輸,故改為有值才鎖(國籍代碼 列外)
    If textCU04 <> MsgText(601) Then textCU04.Locked = True '中文客戶名稱
    If textCU07 <> MsgText(601) Then textCU07.Locked = True '公司負責人
    'If textCU10 <> MsgText(601) Then textCU10.Locked = true  '國籍代碼(可能改4碼,先不鎖)
    If textCU11 <> MsgText(601) Then textCU11.Locked = True  '申請人ID/統編
    If textCU13 <> MsgText(601) Then textCU13.Locked = True '智權人員
    If textCU12 <> MsgText(601) Then textCU12.Locked = True '業務區
    If textCU16 <> MsgText(601) Then textCU16.Locked = True  '電話
    If textCU21 <> MsgText(601) Then textCU21.Locked = True  'Line ID
    If textCU18 <> MsgText(601) Then textCU18.Locked = True  'Fax1
    If textCU19 <> MsgText(601) Then textCU19.Locked = True  'Fax2
    If textCU20 <> MsgText(601) Then textCU20.Locked = True '代表信箱
    If textCU22 <> MsgText(601) Then textCU22.Locked = True 'Mobile Phone
    If textCU23 <> MsgText(601) Then textCU23.Locked = True  '客戶地址郵遞區號
    If textCU30 <> MsgText(601) Then textCU30.Locked = True  '聯絡地址郵遞區號
    If textCU31 <> MsgText(601) Then textCU31.Locked = True '聯絡地址
    If textCU103 <> MsgText(601) Then textCU103.Locked = True  '負責人英文名稱
    If textCU04 <> MsgText(601) Then cboContact.Locked = True '接洽人
End Sub

'Added by Lydia 2022/12/20 改成「FCP提申急件預設組別」
Private Sub Combo4_Validate(Cancel As Boolean)
   If Combo4 <> "" Then
      Combo4 = Left(Combo4, 1) + "." + PUB_GetFCPGrpName(Left(Combo4, 1))
      If Combo4 = Left(Combo4, 1) + "." Then
         Combo4 = Left(Combo4, 1)
         Cancel = True
         Combo4.SetFocus
      End If
   End If
End Sub

'Added by Lydia 2023/01/03
Private Function ChkExistSpec(ByVal pTBL As String, ByVal pNo As String, ByVal pLen As Integer) As Boolean
Dim rsQD As New ADODB.Recordset
Dim strQ1 As String, intQ As Integer
   
   ChkExistSpec = False
   Select Case UCase(pTBL)
       Case "NPMEMO"
           strQ1 = "SELECT NM01 FROM NPMEMO WHERE NM05='" & Left(pNo, pLen) & "' AND NM04 IS NULL "
       Case "APPROVALMEMO2"
           strQ1 = "SELECT AM01 FROM APPROVALMEMO2 WHERE AM05='" & Left(pNo, pLen) & "' AND AM04 IS NULL "
       Case "INCOMMEMO"
           strQ1 = "SELECT IM01 FROM INCOMMEMO WHERE IM05='" & Left(pNo, pLen) & "' AND IM04 IS NULL "
       Case "DEBITNOTEPS"
           strQ1 = "SELECT DNPS01 FROM DEBITNOTEPS WHERE DNPS05='" & Left(pNo, pLen) & "' AND DNPS04 IS NULL "
       Case "FCPEMPBILL"
           strQ1 = "SELECT FEB01 FROM FCPEMPBILL WHERE FEB05='" & Left(pNo, pLen) & "' AND FEB04 IS NULL "
       Case "APPROVALPS"
           strQ1 = "SELECT APS01 FROM APPROVALPS WHERE APS05='" & Left(pNo, pLen) & "' AND APS04 IS NULL "
   End Select
   If strQ1 <> "" Then
       intQ = 1
       Set rsQD = ClsLawReadRstMsg(intQ, strQ1)
       If intQ = 1 Then
          ChkExistSpec = True
       End If
       Set rsQD = Nothing
   End If
   
End Function

'Add by Amy 2023/05/03 檢查只彈訊息,可存檔
'Modify by Amy 2023/06/06 原:sub 改彈訊息選是否
Private Function ChkShowMsg() As Boolean
    Dim strMsg As String
    
    ChkShowMsg = False 'Add by Amy 2023/06/06
    strIDRepeat = "" 'Add by Amy 2023/09/01 身份證/統編 重覆發信內容
    strMsg = ""
    'Modify by Amy 2024/02/29 +有修改統編才做
    'Modify by Amy 2024/07/03 +66666666 ,有加排除的編號,要確認 ChkCU11Same 是否也改
    'Modify by Amy 2024/09/26 bug 統編欄位抓錯(原textCU01)
    If cboStatus <> "不再使用" And textCU11 <> MsgText(601) And textCU11 <> "00000000" And textCU11 <> "66666666" _
      And m_FieldList(10).fiOldData <> textCU11 Then
        '因為學校等特殊客戶會由不同智權人員管控,所以只要提醒即可
        'Modify by Amy 2023/09/01 改成共用
        'Modify by Amy 2024/08/30 客戶為學校,不檢查證號相同,但身份證/統編 重覆發信仍發-秀玲
        If ChkCU11Same(textCU01, textCU02, textCU11, strMsg, m_EditMode, Me.Name, IIf(optCustomer(2).Value = True, "2", ""), strIDRepeat) = True Then
            If strMsg <> MsgText(601) Then
               'Modify by Amy 2023/06/06 改彈是否
               'Modify by Amy 2023/09/05 +顯示[身份證字號/統一編號]
               If MsgBox("身份證字號/統一編號[" & textCU11 & "]與" & IIf(InStr(strMsg, ",") > 0, vbCrLf, "") & strMsg & "相同" & vbCrLf & _
                              "請與智權人員確認不是更名！" & vbCrLf & _
                              "繼續存檔請按「是」，回前畫面請按「否」", vbExclamation + vbYesNo + vbDefaultButton2) = vbNo Then
                     Exit Function
               End If
               'end 2023/09/05
            End If
            If strIDRepeat <> MsgText(601) Then
               strIDRepeat = "身份證字號/統一編號[" & textCU11 & "]與" & IIf(InStr(strIDRepeat, ",") > 0, vbCrLf, "") & strIDRepeat & "相同" & vbCrLf
            End If
        End If
        'end 2024/08/30
    End If
    
    ChkShowMsg = True 'Add by Amy 2023/06/06
End Function
'end 2023/05/03

'Add by Amy 2023/07/06 自動取號搬至此
Private Function GetChkAutoNo() As Boolean
   Dim strTmp As String
   
   GetChkAutoNo = False
   'Added by Lydia 2017/05/09 自動編號移到檢查的後面
   If m_EditMode = 1 And Trim(textCU01) = "" Then
        If ClsPDGetAutoNumber("X", strTmp, True, False) Then
           strTmp = "X" + Right(strTmp, 5)
           textCU01.Text = strTmp
           SetFieldNewData "CU01", textCU01 & String(8 - Len(textCU01), "0")
        Else
           ShowMsg "讀取自動編號檔錯誤，請洽系統管理者 !"
           Exit Function
        End If
   End If
   'end 2017/05/09
   GetChkAutoNo = True
End Function

