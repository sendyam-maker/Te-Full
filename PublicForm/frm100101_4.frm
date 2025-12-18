VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm100101_4 
   BorderStyle     =   1  '單線固定
   Caption         =   "商標案件基本資料查詢"
   ClientHeight    =   6132
   ClientLeft      =   96
   ClientTop       =   996
   ClientWidth     =   9336
   ControlBox      =   0   'False
   LinkTopic       =   "Form8"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6132
   ScaleWidth      =   9336
   Begin VB.CommandButton cmdOK 
      Caption         =   "進度"
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   8.4
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Index           =   13
      Left            =   6180
      Style           =   1  '圖片外觀
      TabIndex        =   4
      Top             =   30
      Width           =   555
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "各項指示"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   8.4
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Index           =   12
      Left            =   0
      Style           =   1  '圖片外觀
      TabIndex        =   13
      Top             =   30
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "申5"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   8.4
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Index           =   11
      Left            =   4170
      TabIndex        =   7
      Top             =   30
      Width           =   405
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "申4"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   8.4
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Index           =   10
      Left            =   3750
      TabIndex        =   8
      Top             =   30
      Width           =   405
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "申3"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   8.4
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Index           =   9
      Left            =   3330
      TabIndex        =   9
      Top             =   30
      Width           =   405
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "申2"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   8.4
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Index           =   8
      Left            =   2910
      TabIndex        =   10
      Top             =   30
      Width           =   405
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "已設定代表圖"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   8.4
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Index           =   7
      Left            =   840
      Style           =   1  '圖片外觀
      TabIndex        =   12
      Top             =   30
      Width           =   1155
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "商品服務"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   8.4
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Index           =   6
      Left            =   2010
      TabIndex        =   11
      Top             =   30
      Width           =   885
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H008080FF&
      Caption         =   "分割案"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   8.4
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Index           =   5
      Left            =   4590
      Style           =   1  '圖片外觀
      TabIndex        =   6
      Top             =   30
      Visible         =   0   'False
      Width           =   675
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H008080FF&
      Caption         =   "相關卷號"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   8.4
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Index           =   4
      Left            =   5280
      Style           =   1  '圖片外觀
      TabIndex        =   5
      Top             =   30
      Visible         =   0   'False
      Width           =   885
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "結束"
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   8.4
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Index           =   3
      Left            =   8850
      Style           =   1  '圖片外觀
      TabIndex        =   1
      Top             =   30
      Width           =   495
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "申請人"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   8.4
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Index           =   1
      Left            =   7440
      Style           =   1  '圖片外觀
      TabIndex        =   2
      Top             =   30
      Width           =   675
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "代理人"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   8.4
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Index           =   0
      Left            =   6750
      Style           =   1  '圖片外觀
      TabIndex        =   3
      Top             =   30
      Width           =   675
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "下一筆"
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   8.4
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Index           =   2
      Left            =   8130
      Style           =   1  '圖片外觀
      TabIndex        =   0
      Top             =   30
      Width           =   705
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFC0C0&
      Height          =   3492
      Left            =   -74880
      MultiLine       =   -1  'True
      ScrollBars      =   2  '垂直捲軸
      TabIndex        =   21
      Top             =   360
      Width           =   7572
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5445
      Left            =   30
      TabIndex        =   14
      Top             =   420
      Width           =   9270
      _ExtentX        =   16341
      _ExtentY        =   9610
      _Version        =   393216
      Tabs            =   9
      TabsPerRow      =   9
      TabHeight       =   741
      TabCaption(0)   =   "基本資料1"
      TabPicture(0)   =   "frm100101_4.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label34"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label33"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label32"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label30"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label23"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label3"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label27"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label25"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label21"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label20(1)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label17"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label16"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label15"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label14"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Label13"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Label12"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Label11"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Label10"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Label9"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Label8(0)"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Label1(1)"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "Label2(1)"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "Label5"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "Label6"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "Label35"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "Label36"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "Label37"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "Label38"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "Label39"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "lbl1(1)"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "lbl1(2)"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "lbl1(7)"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "lbl1(8)"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "lbl1(9)"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "lbl1(10)"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "lbl1(11)"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "lbl1(12)"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "lbl1(13)"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "lbl1(14)"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).Control(39)=   "lbl1(19)"
      Tab(0).Control(39).Enabled=   0   'False
      Tab(0).Control(40)=   "lbl1(20)"
      Tab(0).Control(40).Enabled=   0   'False
      Tab(0).Control(41)=   "lbl1(21)"
      Tab(0).Control(41).Enabled=   0   'False
      Tab(0).Control(42)=   "lbl1(22)"
      Tab(0).Control(42).Enabled=   0   'False
      Tab(0).Control(43)=   "lbl1(23)"
      Tab(0).Control(43).Enabled=   0   'False
      Tab(0).Control(44)=   "lbl1(44)"
      Tab(0).Control(44).Enabled=   0   'False
      Tab(0).Control(45)=   "lbl1(48)"
      Tab(0).Control(45).Enabled=   0   'False
      Tab(0).Control(46)=   "Label4"
      Tab(0).Control(46).Enabled=   0   'False
      Tab(0).Control(47)=   "lbl1(3)"
      Tab(0).Control(47).Enabled=   0   'False
      Tab(0).Control(48)=   "lbl1(15)"
      Tab(0).Control(48).Enabled=   0   'False
      Tab(0).Control(49)=   "Label101"
      Tab(0).Control(49).Enabled=   0   'False
      Tab(0).Control(50)=   "lbl1(55)"
      Tab(0).Control(50).Enabled=   0   'False
      Tab(0).Control(51)=   "lbl1(59)"
      Tab(0).Control(51).Enabled=   0   'False
      Tab(0).Control(52)=   "Label106"
      Tab(0).Control(52).Enabled=   0   'False
      Tab(0).Control(53)=   "Label107"
      Tab(0).Control(53).Enabled=   0   'False
      Tab(0).Control(54)=   "Label1(2)"
      Tab(0).Control(54).Enabled=   0   'False
      Tab(0).Control(55)=   "lbl1(128)"
      Tab(0).Control(55).Enabled=   0   'False
      Tab(0).Control(56)=   "Label111"
      Tab(0).Control(56).Enabled=   0   'False
      Tab(0).Control(57)=   "Label1(116)"
      Tab(0).Control(57).Enabled=   0   'False
      Tab(0).Control(58)=   "lbl1(63)"
      Tab(0).Control(58).Enabled=   0   'False
      Tab(0).Control(59)=   "txt1(0)"
      Tab(0).Control(59).Enabled=   0   'False
      Tab(0).Control(60)=   "txt1(1)"
      Tab(0).Control(60).Enabled=   0   'False
      Tab(0).Control(61)=   "txt1(21)"
      Tab(0).Control(61).Enabled=   0   'False
      Tab(0).Control(62)=   "txt1(59)"
      Tab(0).Control(62).Enabled=   0   'False
      Tab(0).Control(63)=   "txt1(60)"
      Tab(0).Control(63).Enabled=   0   'False
      Tab(0).Control(64)=   "txt1(61)"
      Tab(0).Control(64).Enabled=   0   'False
      Tab(0).Control(65)=   "txt1(62)"
      Tab(0).Control(65).Enabled=   0   'False
      Tab(0).Control(66)=   "Label117"
      Tab(0).Control(66).Enabled=   0   'False
      Tab(0).Control(67)=   "Label118"
      Tab(0).Control(67).Enabled=   0   'False
      Tab(0).Control(68)=   "lbl1(136)"
      Tab(0).Control(68).Enabled=   0   'False
      Tab(0).ControlCount=   69
      TabCaption(1)   =   "基本資料2"
      TabPicture(1)   =   "frm100101_4.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label31"
      Tab(1).Control(1)=   "Label22"
      Tab(1).Control(2)=   "Label40"
      Tab(1).Control(3)=   "Label42"
      Tab(1).Control(4)=   "Label43"
      Tab(1).Control(5)=   "Label45"
      Tab(1).Control(6)=   "Label46"
      Tab(1).Control(7)=   "Label47"
      Tab(1).Control(8)=   "Label48"
      Tab(1).Control(9)=   "Label49(0)"
      Tab(1).Control(10)=   "Label50"
      Tab(1).Control(11)=   "Label51(0)"
      Tab(1).Control(12)=   "Label52"
      Tab(1).Control(13)=   "Label53"
      Tab(1).Control(14)=   "Label54"
      Tab(1).Control(15)=   "lbl1(24)"
      Tab(1).Control(16)=   "lbl1(25)"
      Tab(1).Control(17)=   "lbl1(26)"
      Tab(1).Control(18)=   "lbl1(28)"
      Tab(1).Control(19)=   "lbl1(29)"
      Tab(1).Control(20)=   "lbl1(30)"
      Tab(1).Control(21)=   "lbl1(31)"
      Tab(1).Control(22)=   "lbl1(32)"
      Tab(1).Control(23)=   "lbl1(33)"
      Tab(1).Control(24)=   "lbl1(43)"
      Tab(1).Control(25)=   "Label2(0)"
      Tab(1).Control(26)=   "Label103"
      Tab(1).Control(27)=   "lbl1(57)"
      Tab(1).Control(28)=   "lbl1(58)"
      Tab(1).Control(29)=   "Label104"
      Tab(1).Control(30)=   "Label105"
      Tab(1).Control(31)=   "lbl1(67)"
      Tab(1).Control(32)=   "Label115"
      Tab(1).Control(33)=   "lbl1(66)"
      Tab(1).Control(34)=   "Label116"
      Tab(1).Control(35)=   "txt1(2)"
      Tab(1).Control(36)=   "txt1(3)"
      Tab(1).Control(37)=   "txt1(4)"
      Tab(1).Control(38)=   "txt1(5)"
      Tab(1).Control(39)=   "lbl1(0)"
      Tab(1).Control(40)=   "Label122"
      Tab(1).Control(41)=   "Label123"
      Tab(1).Control(42)=   "lbl1(6)"
      Tab(1).Control(43)=   "Label124"
      Tab(1).Control(44)=   "Label125"
      Tab(1).ControlCount=   45
      TabCaption(2)   =   "申請人1-3"
      TabPicture(2)   =   "frm100101_4.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label19"
      Tab(2).Control(1)=   "Label56"
      Tab(2).Control(2)=   "Label65"
      Tab(2).Control(3)=   "Label66"
      Tab(2).Control(4)=   "lbl1(35)"
      Tab(2).Control(5)=   "lbl1(51)"
      Tab(2).Control(6)=   "Label18"
      Tab(2).Control(7)=   "lbl1(52)"
      Tab(2).Control(8)=   "Label26"
      Tab(2).Control(9)=   "Label84"
      Tab(2).Control(10)=   "Label85"
      Tab(2).Control(11)=   "Label86"
      Tab(2).Control(12)=   "Label87"
      Tab(2).Control(13)=   "Label88"
      Tab(2).Control(14)=   "Label89"
      Tab(2).Control(15)=   "txt1(27)"
      Tab(2).Control(16)=   "txt1(26)"
      Tab(2).Control(17)=   "txt1(25)"
      Tab(2).Control(18)=   "txt1(24)"
      Tab(2).Control(19)=   "txt1(23)"
      Tab(2).Control(20)=   "txt1(22)"
      Tab(2).Control(21)=   "txt1(8)"
      Tab(2).Control(22)=   "txt1(7)"
      Tab(2).Control(23)=   "txt1(6)"
      Tab(2).ControlCount=   24
      TabCaption(3)   =   "申請人4-5 / 延展"
      TabPicture(3)   =   "frm100101_4.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "lbl1(4)"
      Tab(3).Control(1)=   "Label77"
      Tab(3).Control(2)=   "Label24"
      Tab(3).Control(3)=   "Label75"
      Tab(3).Control(4)=   "lbl1(46)"
      Tab(3).Control(5)=   "lbl1(40)"
      Tab(3).Control(6)=   "lbl1(38)"
      Tab(3).Control(7)=   "lbl1(37)"
      Tab(3).Control(8)=   "Label68"
      Tab(3).Control(9)=   "Label67"
      Tab(3).Control(10)=   "Label61"
      Tab(3).Control(11)=   "lbl1(47)"
      Tab(3).Control(12)=   "Label76"
      Tab(3).Control(13)=   "Label83"
      Tab(3).Control(14)=   "lbl1(54)"
      Tab(3).Control(15)=   "Label82"
      Tab(3).Control(16)=   "lbl1(53)"
      Tab(3).Control(17)=   "txt1(33)"
      Tab(3).Control(18)=   "txt1(32)"
      Tab(3).Control(19)=   "txt1(31)"
      Tab(3).Control(20)=   "txt1(30)"
      Tab(3).Control(21)=   "txt1(29)"
      Tab(3).Control(22)=   "txt1(28)"
      Tab(3).Control(23)=   "Label95"
      Tab(3).Control(24)=   "Label94"
      Tab(3).Control(25)=   "Label93"
      Tab(3).Control(26)=   "Label92"
      Tab(3).Control(27)=   "Label91"
      Tab(3).Control(28)=   "Label90"
      Tab(3).ControlCount=   29
      TabCaption(4)   =   "代理人/聯絡人"
      TabPicture(4)   =   "frm100101_4.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Label29"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).Control(1)=   "Label7"
      Tab(4).Control(1).Enabled=   0   'False
      Tab(4).Control(2)=   "lbl1(36)"
      Tab(4).Control(2).Enabled=   0   'False
      Tab(4).Control(3)=   "lbl1(39)"
      Tab(4).Control(3).Enabled=   0   'False
      Tab(4).Control(4)=   "Label80(29)"
      Tab(4).Control(4).Enabled=   0   'False
      Tab(4).Control(5)=   "lbl1(134)"
      Tab(4).Control(5).Enabled=   0   'False
      Tab(4).Control(6)=   "Label80(26)"
      Tab(4).Control(6).Enabled=   0   'False
      Tab(4).Control(7)=   "Label1(0)"
      Tab(4).Control(7).Enabled=   0   'False
      Tab(4).Control(8)=   "lbl1(127)"
      Tab(4).Control(8).Enabled=   0   'False
      Tab(4).Control(9)=   "Label44"
      Tab(4).Control(9).Enabled=   0   'False
      Tab(4).Control(10)=   "lbl1(34)"
      Tab(4).Control(10).Enabled=   0   'False
      Tab(4).Control(11)=   "lbl1(45)"
      Tab(4).Control(11).Enabled=   0   'False
      Tab(4).Control(12)=   "Label41"
      Tab(4).Control(12).Enabled=   0   'False
      Tab(4).Control(13)=   "Label55"
      Tab(4).Control(13).Enabled=   0   'False
      Tab(4).Control(14)=   "lbl1(27)"
      Tab(4).Control(14).Enabled=   0   'False
      Tab(4).Control(15)=   "Label63"
      Tab(4).Control(15).Enabled=   0   'False
      Tab(4).Control(16)=   "txt1(58)"
      Tab(4).Control(16).Enabled=   0   'False
      Tab(4).Control(17)=   "txt1(20)"
      Tab(4).Control(17).Enabled=   0   'False
      Tab(4).Control(18)=   "txt1(19)"
      Tab(4).Control(18).Enabled=   0   'False
      Tab(4).Control(19)=   "txt1(18)"
      Tab(4).Control(19).Enabled=   0   'False
      Tab(4).Control(20)=   "txt1(17)"
      Tab(4).Control(20).Enabled=   0   'False
      Tab(4).Control(21)=   "txt1(16)"
      Tab(4).Control(21).Enabled=   0   'False
      Tab(4).Control(22)=   "txt1(15)"
      Tab(4).Control(22).Enabled=   0   'False
      Tab(4).Control(23)=   "Label100"
      Tab(4).Control(23).Enabled=   0   'False
      Tab(4).Control(24)=   "Label74"
      Tab(4).Control(24).Enabled=   0   'False
      Tab(4).Control(25)=   "Label73"
      Tab(4).Control(25).Enabled=   0   'False
      Tab(4).Control(26)=   "Label72"
      Tab(4).Control(26).Enabled=   0   'False
      Tab(4).Control(27)=   "Label71"
      Tab(4).Control(27).Enabled=   0   'False
      Tab(4).Control(28)=   "Label70"
      Tab(4).Control(28).Enabled=   0   'False
      Tab(4).Control(29)=   "Label69"
      Tab(4).Control(29).Enabled=   0   'False
      Tab(4).Control(30)=   "Combo3(1)"
      Tab(4).Control(30).Enabled=   0   'False
      Tab(4).ControlCount=   31
      TabCaption(5)   =   "優先權資料"
      TabPicture(5)   =   "frm100101_4.frx":008C
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "grdDataList2"
      Tab(5).ControlCount=   1
      TabCaption(6)   =   "代表人"
      TabPicture(6)   =   "frm100101_4.frx":00A8
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "txt1(14)"
      Tab(6).Control(1)=   "txt1(13)"
      Tab(6).Control(2)=   "txt1(12)"
      Tab(6).Control(3)=   "txt1(11)"
      Tab(6).Control(4)=   "txt1(10)"
      Tab(6).Control(5)=   "txt1(9)"
      Tab(6).Control(6)=   "txt1(34)"
      Tab(6).Control(7)=   "txt1(35)"
      Tab(6).Control(8)=   "txt1(36)"
      Tab(6).Control(9)=   "txt1(37)"
      Tab(6).Control(10)=   "txt1(38)"
      Tab(6).Control(11)=   "txt1(39)"
      Tab(6).Control(12)=   "txt1(40)"
      Tab(6).Control(13)=   "txt1(41)"
      Tab(6).Control(14)=   "txt1(42)"
      Tab(6).Control(15)=   "txt1(43)"
      Tab(6).Control(16)=   "txt1(44)"
      Tab(6).Control(17)=   "txt1(45)"
      Tab(6).Control(18)=   "txt1(46)"
      Tab(6).Control(19)=   "txt1(47)"
      Tab(6).Control(20)=   "txt1(48)"
      Tab(6).Control(21)=   "txt1(49)"
      Tab(6).Control(22)=   "txt1(50)"
      Tab(6).Control(23)=   "txt1(51)"
      Tab(6).Control(24)=   "txt1(52)"
      Tab(6).Control(25)=   "txt1(53)"
      Tab(6).Control(26)=   "txt1(54)"
      Tab(6).Control(27)=   "txt1(55)"
      Tab(6).Control(28)=   "txt1(56)"
      Tab(6).Control(29)=   "txt1(57)"
      Tab(6).Control(30)=   "Label60"
      Tab(6).Control(31)=   "Label59"
      Tab(6).Control(32)=   "Label58"
      Tab(6).Control(33)=   "Label57"
      Tab(6).Control(34)=   "Label64"
      Tab(6).Control(35)=   "Label62"
      Tab(6).Control(36)=   "Label96"
      Tab(6).Control(37)=   "Label97"
      Tab(6).Control(38)=   "Label98"
      Tab(6).Control(39)=   "Label99"
      Tab(6).ControlCount=   40
      TabCaption(7)   =   "銷卷資料"
      TabPicture(7)   =   "frm100101_4.frx":00C4
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "Label78"
      Tab(7).Control(0).Enabled=   0   'False
      Tab(7).Control(1)=   "Label79"
      Tab(7).Control(1).Enabled=   0   'False
      Tab(7).Control(2)=   "Label80(0)"
      Tab(7).Control(2).Enabled=   0   'False
      Tab(7).Control(3)=   "Label81"
      Tab(7).Control(3).Enabled=   0   'False
      Tab(7).Control(4)=   "lbl1(16)"
      Tab(7).Control(4).Enabled=   0   'False
      Tab(7).Control(5)=   "lbl1(5)"
      Tab(7).Control(5).Enabled=   0   'False
      Tab(7).Control(6)=   "lbl1(49)"
      Tab(7).Control(6).Enabled=   0   'False
      Tab(7).Control(7)=   "lbl1(50)"
      Tab(7).Control(7).Enabled=   0   'False
      Tab(7).Control(8)=   "lblTFBase"
      Tab(7).Control(8).Enabled=   0   'False
      Tab(7).Control(9)=   "MGrid1"
      Tab(7).Control(9).Enabled=   0   'False
      Tab(7).ControlCount=   10
      TabCaption(8)   =   "其他/商標描述"
      TabPicture(8)   =   "frm100101_4.frx":00E0
      Tab(8).ControlEnabled=   0   'False
      Tab(8).Control(0)=   "Label114"
      Tab(8).Control(1)=   "lbl1(65)"
      Tab(8).Control(2)=   "lbl1(64)"
      Tab(8).Control(3)=   "Label112"
      Tab(8).Control(4)=   "Label113"
      Tab(8).Control(5)=   "Label102"
      Tab(8).Control(6)=   "Label110"
      Tab(8).Control(7)=   "lbl1(60)"
      Tab(8).Control(8)=   "Label108"
      Tab(8).Control(9)=   "lbl1(61)"
      Tab(8).Control(10)=   "Label109"
      Tab(8).Control(11)=   "lbl1(62)"
      Tab(8).Control(12)=   "lbl1(56)"
      Tab(8).Control(13)=   "Label119"
      Tab(8).Control(14)=   "Label120"
      Tab(8).Control(15)=   "txt1(63)"
      Tab(8).Control(16)=   "txt1(64)"
      Tab(8).Control(17)=   "txt1(65)"
      Tab(8).Control(18)=   "Label121"
      Tab(8).ControlCount=   19
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MGrid1 
         Height          =   1788
         Left            =   -74784
         TabIndex        =   288
         Top             =   2280
         Width           =   4788
         _ExtentX        =   8446
         _ExtentY        =   3154
         _Version        =   393216
         FixedCols       =   0
         AllowUserResizing=   3
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.ComboBox Combo3 
         Height          =   276
         Index           =   1
         ItemData        =   "frm100101_4.frx":00FC
         Left            =   -70980
         List            =   "frm100101_4.frx":010F
         Locked          =   -1  'True
         Style           =   2  '單純下拉式
         TabIndex        =   219
         Top             =   1930
         Width           =   1470
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdDataList2 
         Height          =   4815
         Left            =   -74850
         TabIndex        =   214
         Top             =   540
         Width           =   8955
         _ExtentX        =   15790
         _ExtentY        =   8488
         _Version        =   393216
         Cols            =   6
         FixedCols       =   0
         HighLight       =   0
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
         _Band(0).Cols   =   6
      End
      Begin VB.Label lblTFBase 
         Caption         =   "TF基礎案號數："
         Height          =   252
         Left            =   -74790
         TabIndex        =   287
         Top             =   2016
         Width           =   1356
      End
      Begin VB.Label Label125 
         AutoSize        =   -1  'True
         Caption         =   "(％)"
         Height          =   180
         Left            =   -69990
         TabIndex        =   286
         Top             =   1860
         Width           =   300
      End
      Begin VB.Label Label124 
         AutoSize        =   -1  'True
         Caption         =   "延展折扣："
         Height          =   180
         Left            =   -72070
         TabIndex        =   285
         Top             =   1860
         Width           =   900
      End
      Begin MSForms.Label lbl1 
         Height          =   260
         Index           =   6
         Left            =   -70740
         TabIndex        =   284
         Top             =   1830
         Width           =   690
         BackColor       =   -2147483639
         VariousPropertyBits=   27
         Caption         =   " "
         Size            =   "1217;459"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label123 
         AutoSize        =   -1  'True
         Caption         =   "繳註冊費折扣："
         Height          =   180
         Left            =   -74910
         TabIndex        =   283
         Top             =   1860
         Width           =   1260
      End
      Begin VB.Label Label122 
         AutoSize        =   -1  'True
         Caption         =   "(％)"
         Height          =   180
         Left            =   -72870
         TabIndex        =   282
         Top             =   1870
         Width           =   300
      End
      Begin MSForms.Label lbl1 
         Height          =   260
         Index           =   0
         Left            =   -73650
         TabIndex        =   281
         Top             =   1830
         Width           =   690
         BackColor       =   -2147483639
         VariousPropertyBits=   27
         Caption         =   " "
         Size            =   "1217;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label121 
         AutoSize        =   -1  'True
         Caption         =   "商標描述日文："
         Height          =   180
         Left            =   -74880
         TabIndex        =   280
         Top             =   4320
         Width           =   1260
      End
      Begin MSForms.TextBox txt1 
         Height          =   950
         Index           =   65
         Left            =   -73590
         TabIndex        =   279
         Top             =   4290
         Width           =   7680
         VariousPropertyBits=   -1466941409
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "13547;1676"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   950
         Index           =   64
         Left            =   -73590
         TabIndex        =   278
         Top             =   3300
         Width           =   7680
         VariousPropertyBits=   -1466941409
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "13547;1676"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   950
         Index           =   63
         Left            =   -73590
         TabIndex        =   277
         Top             =   2310
         Width           =   7680
         VariousPropertyBits=   -1466941409
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "13547;1676"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label120 
         AutoSize        =   -1  'True
         Caption         =   "商標描述英文："
         Height          =   180
         Left            =   -74880
         TabIndex        =   276
         Top             =   3330
         Width           =   1260
      End
      Begin VB.Label Label119 
         AutoSize        =   -1  'True
         Caption         =   "商標描述中文："
         Height          =   180
         Left            =   -74880
         TabIndex        =   275
         Top             =   2310
         Width           =   1260
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   136
         Left            =   7530
         TabIndex        =   274
         Top             =   3660
         Width           =   285
         BackColor       =   -2147483639
         VariousPropertyBits=   27
         Caption         =   " "
         Size            =   "503;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
         ParagraphAlign  =   3
      End
      Begin VB.Label Label118 
         AutoSize        =   -1  'True
         Caption         =   "註冊證形式："
         Height          =   180
         Left            =   6420
         TabIndex        =   273
         Top             =   3690
         Width           =   1080
      End
      Begin VB.Label Label117 
         AutoSize        =   -1  'True
         Caption         =   "(1:電子 2:紙本)"
         Height          =   180
         Left            =   7890
         TabIndex        =   272
         Top             =   3690
         Width           =   1155
      End
      Begin MSForms.TextBox txt1 
         Height          =   285
         Index           =   62
         Left            =   6120
         TabIndex        =   271
         Top             =   2760
         Width           =   3050
         VariousPropertyBits=   671105055
         BackColor       =   16777215
         Size            =   "5380;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   360
         Index           =   6
         Left            =   -73620
         TabIndex        =   270
         Top             =   1410
         Width           =   7635
         VariousPropertyBits=   -1466941409
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "13467;635"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   360
         Index           =   7
         Left            =   -73620
         TabIndex        =   269
         Top             =   1773
         Width           =   7635
         VariousPropertyBits=   -1466941409
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "13467;635"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   360
         Index           =   8
         Left            =   -73620
         TabIndex        =   268
         Top             =   2136
         Width           =   7635
         VariousPropertyBits=   -1466941409
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "13467;635"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   360
         Index           =   22
         Left            =   -73620
         TabIndex        =   267
         Top             =   3225
         Width           =   7635
         VariousPropertyBits=   -1466941409
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "13467;635"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   360
         Index           =   23
         Left            =   -73620
         TabIndex        =   266
         Top             =   2862
         Width           =   7635
         VariousPropertyBits=   -1466941409
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "13467;635"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   360
         Index           =   24
         Left            =   -73620
         TabIndex        =   265
         Top             =   2499
         Width           =   7635
         VariousPropertyBits=   -1466941409
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "13467;635"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   360
         Index           =   25
         Left            =   -73620
         TabIndex        =   264
         Top             =   4320
         Width           =   7635
         VariousPropertyBits=   -1466941409
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "13467;635"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   360
         Index           =   26
         Left            =   -73620
         TabIndex        =   263
         Top             =   3951
         Width           =   7635
         VariousPropertyBits=   -1466941409
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "13467;635"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   360
         Index           =   27
         Left            =   -73620
         TabIndex        =   262
         Top             =   3588
         Width           =   7635
         VariousPropertyBits=   -1466941409
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "13467;635"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label90 
         AutoSize        =   -1  'True
         Caption         =   "申請地址4(英)："
         Height          =   180
         Left            =   -74880
         TabIndex        =   261
         Top             =   1638
         Width           =   1290
      End
      Begin VB.Label Label91 
         AutoSize        =   -1  'True
         Caption         =   "申請地址4(日)："
         Height          =   180
         Left            =   -74880
         TabIndex        =   260
         Top             =   2046
         Width           =   1290
      End
      Begin VB.Label Label92 
         AutoSize        =   -1  'True
         Caption         =   "申請地址4(中)："
         Height          =   180
         Left            =   -74880
         TabIndex        =   259
         Top             =   1230
         Width           =   1290
      End
      Begin VB.Label Label93 
         AutoSize        =   -1  'True
         Caption         =   "申請地址5(英)："
         Height          =   180
         Left            =   -74880
         TabIndex        =   258
         Top             =   2862
         Width           =   1290
      End
      Begin VB.Label Label94 
         AutoSize        =   -1  'True
         Caption         =   "申請地址5(日)："
         Height          =   180
         Left            =   -74880
         TabIndex        =   257
         Top             =   3270
         Width           =   1290
      End
      Begin VB.Label Label95 
         AutoSize        =   -1  'True
         Caption         =   "申請地址5(中)："
         Height          =   180
         Left            =   -74880
         TabIndex        =   256
         Top             =   2454
         Width           =   1290
      End
      Begin MSForms.TextBox txt1 
         Height          =   360
         Index           =   28
         Left            =   -73530
         TabIndex        =   255
         Top             =   1956
         Width           =   7635
         VariousPropertyBits=   -1466941409
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "13467;635"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   360
         Index           =   29
         Left            =   -73530
         TabIndex        =   254
         Top             =   1548
         Width           =   7635
         VariousPropertyBits=   -1466941409
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "13467;635"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   360
         Index           =   30
         Left            =   -73530
         TabIndex        =   253
         Top             =   1140
         Width           =   7635
         VariousPropertyBits=   -1466941409
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "13467;635"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   360
         Index           =   31
         Left            =   -73530
         TabIndex        =   252
         Top             =   3180
         Width           =   7635
         VariousPropertyBits=   -1466941409
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "13467;635"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   360
         Index           =   32
         Left            =   -73530
         TabIndex        =   251
         Top             =   2772
         Width           =   7635
         VariousPropertyBits=   -1466941409
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "13467;635"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   360
         Index           =   33
         Left            =   -73530
         TabIndex        =   250
         Top             =   2364
         Width           =   7635
         VariousPropertyBits=   -1466941409
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "13467;635"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   53
         Left            =   -73950
         TabIndex        =   249
         Top             =   540
         Width           =   8145
         BackColor       =   -2147483639
         VariousPropertyBits=   27
         Caption         =   " "
         Size            =   "14367;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label82 
         AutoSize        =   -1  'True
         Caption         =   "申請人4："
         Height          =   180
         Left            =   -74880
         TabIndex        =   248
         Top             =   540
         Width           =   810
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   54
         Left            =   -73950
         TabIndex        =   247
         Top             =   825
         Width           =   8145
         BackColor       =   -2147483639
         VariousPropertyBits=   27
         Caption         =   " "
         Size            =   "14367;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label83 
         AutoSize        =   -1  'True
         Caption         =   "申請人5："
         Height          =   180
         Left            =   -74880
         TabIndex        =   246
         Top             =   825
         Width           =   810
      End
      Begin VB.Label Label69 
         AutoSize        =   -1  'True
         Caption         =   "聯絡人1(中)："
         Height          =   180
         Left            =   -74880
         TabIndex        =   245
         Top             =   2280
         Width           =   1110
      End
      Begin VB.Label Label70 
         AutoSize        =   -1  'True
         Caption         =   "聯絡人1(英)："
         Height          =   180
         Left            =   -74880
         TabIndex        =   244
         Top             =   2680
         Width           =   1110
      End
      Begin VB.Label Label71 
         AutoSize        =   -1  'True
         Caption         =   "聯絡人2(日)："
         Height          =   180
         Left            =   -74880
         TabIndex        =   243
         Top             =   4280
         Width           =   1110
      End
      Begin VB.Label Label72 
         AutoSize        =   -1  'True
         Caption         =   "聯絡人2(英)："
         Height          =   180
         Left            =   -74880
         TabIndex        =   242
         Top             =   3880
         Width           =   1110
      End
      Begin VB.Label Label73 
         AutoSize        =   -1  'True
         Caption         =   "聯絡人2(中)："
         Height          =   180
         Left            =   -74880
         TabIndex        =   241
         Top             =   3480
         Width           =   1110
      End
      Begin VB.Label Label74 
         AutoSize        =   -1  'True
         Caption         =   "聯絡人1(日)："
         Height          =   180
         Left            =   -74880
         TabIndex        =   240
         Top             =   3080
         Width           =   1110
      End
      Begin VB.Label Label100 
         AutoSize        =   -1  'True
         Caption         =   "聯絡人部門(日)："
         Height          =   180
         Left            =   -74880
         TabIndex        =   239
         Top             =   4710
         Width           =   1380
      End
      Begin MSForms.TextBox txt1 
         Height          =   360
         Index           =   15
         Left            =   -73485
         TabIndex        =   238
         Top             =   2280
         Width           =   7455
         VariousPropertyBits=   -1466941409
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "13150;635"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   360
         Index           =   16
         Left            =   -73485
         TabIndex        =   237
         Top             =   2680
         Width           =   7455
         VariousPropertyBits=   -1466941409
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "13150;635"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   360
         Index           =   17
         Left            =   -73485
         TabIndex        =   236
         Top             =   3080
         Width           =   7455
         VariousPropertyBits=   -1466941409
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "13150;635"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   360
         Index           =   18
         Left            =   -73485
         TabIndex        =   235
         Top             =   3480
         Width           =   7455
         VariousPropertyBits=   -1466941409
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "13150;635"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   360
         Index           =   19
         Left            =   -73485
         TabIndex        =   234
         Top             =   3880
         Width           =   7455
         VariousPropertyBits=   -1466941409
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "13150;635"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   360
         Index           =   20
         Left            =   -73485
         TabIndex        =   233
         Top             =   4280
         Width           =   7455
         VariousPropertyBits=   -1466941409
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "13150;635"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   360
         Index           =   58
         Left            =   -73485
         TabIndex        =   232
         Top             =   4680
         Width           =   7455
         VariousPropertyBits=   -1466941409
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "13150;635"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label63 
         AutoSize        =   -1  'True
         Caption         =   "D/N固定列印對象："
         Height          =   180
         Left            =   -70380
         TabIndex        =   231
         Top             =   1663
         Width           =   1545
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   27
         Left            =   -73125
         TabIndex        =   230
         Top             =   1626
         Width           =   255
         BackColor       =   -2147483639
         VariousPropertyBits=   27
         Caption         =   " "
         Size            =   "450;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label55 
         AutoSize        =   -1  'True
         Caption         =   "(Y：印)"
         Height          =   180
         Left            =   -72795
         TabIndex        =   229
         Top             =   1663
         Width           =   600
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "D/N是否列印申請人："
         Height          =   180
         Left            =   -74880
         TabIndex        =   228
         Top             =   1663
         Width           =   1725
      End
      Begin MSForms.Label lbl1 
         Height          =   260
         Index           =   45
         Left            =   -68730
         TabIndex        =   227
         Top             =   1630
         Width           =   2780
         BackColor       =   -2147483639
         VariousPropertyBits=   27
         Size            =   "4904;459"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   260
         Index           =   34
         Left            =   -73480
         TabIndex        =   226
         Top             =   1300
         Width           =   7520
         BackColor       =   -2147483639
         VariousPropertyBits=   27
         Caption         =   " "
         Size            =   "13264;459"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label44 
         AutoSize        =   -1  'True
         Caption         =   "固定請款對象："
         Height          =   180
         Left            =   -74880
         TabIndex        =   225
         Top             =   1336
         Width           =   1260
      End
      Begin MSForms.Label lbl1 
         Height          =   260
         Index           =   127
         Left            =   -68560
         TabIndex        =   224
         Top             =   970
         Width           =   2600
         BackColor       =   -2147483639
         VariousPropertyBits=   27
         Caption         =   " "
         Size            =   "4586;459"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "CLIENT_MATTER_ID："
         Height          =   180
         Index           =   0
         Left            =   -70380
         TabIndex        =   223
         Top             =   1010
         Width           =   1860
      End
      Begin VB.Label Label80 
         AutoSize        =   -1  'True
         Caption         =   "請款幣別："
         Height          =   180
         Index           =   26
         Left            =   -74880
         TabIndex        =   222
         Top             =   1990
         Width           =   900
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   134
         Left            =   -73920
         TabIndex        =   221
         Top             =   1953
         Width           =   705
         BackColor       =   -2147483639
         VariousPropertyBits=   27
         Caption         =   "TEXT"
         Size            =   "1244;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label80 
         AutoSize        =   -1  'True
         Caption         =   "請款單列印幣別格式："
         Height          =   180
         Index           =   29
         Left            =   -72810
         TabIndex        =   220
         Top             =   1990
         Width           =   1800
      End
      Begin MSForms.Label lbl1 
         Height          =   260
         Index           =   39
         Left            =   -73920
         TabIndex        =   218
         Top             =   970
         Width           =   3410
         BackColor       =   -2147483639
         VariousPropertyBits=   27
         Caption         =   " "
         Size            =   "6015;459"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   260
         Index           =   36
         Left            =   -73920
         TabIndex        =   217
         Top             =   650
         Width           =   7970
         BackColor       =   -2147483639
         VariousPropertyBits=   27
         Caption         =   " "
         Size            =   "14058;459"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "彼所案號："
         Height          =   180
         Left            =   -74880
         TabIndex        =   216
         Top             =   1009
         Width           =   900
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         Caption         =   "FC代理人:"
         Height          =   180
         Left            =   -74880
         TabIndex        =   215
         Top             =   682
         Width           =   795
      End
      Begin VB.Label Label99 
         AutoSize        =   -1  'True
         Caption         =   "代表人10(中,英,日)："
         Height          =   180
         Left            =   -74940
         TabIndex        =   213
         Top             =   4890
         Width           =   1650
      End
      Begin VB.Label Label98 
         AutoSize        =   -1  'True
         Caption         =   "代表人9(中,英,日)："
         Height          =   180
         Left            =   -74940
         TabIndex        =   212
         Top             =   4395
         Width           =   1560
      End
      Begin VB.Label Label97 
         AutoSize        =   -1  'True
         Caption         =   "代表人8(中,英,日)："
         Height          =   180
         Left            =   -74940
         TabIndex        =   211
         Top             =   3915
         Width           =   1560
      End
      Begin VB.Label Label96 
         AutoSize        =   -1  'True
         Caption         =   "代表人7(中,英,日)："
         Height          =   180
         Left            =   -74940
         TabIndex        =   210
         Top             =   3420
         Width           =   1560
      End
      Begin VB.Label Label62 
         AutoSize        =   -1  'True
         Caption         =   "代表人3(中,英,日)："
         Height          =   180
         Left            =   -74940
         TabIndex        =   209
         Top             =   1485
         Width           =   1560
      End
      Begin VB.Label Label64 
         AutoSize        =   -1  'True
         Caption         =   "代表人1(中,英,日)："
         Height          =   180
         Left            =   -74940
         TabIndex        =   208
         Top             =   510
         Width           =   1560
      End
      Begin VB.Label Label57 
         AutoSize        =   -1  'True
         Caption         =   "代表人4(中,英,日)："
         Height          =   180
         Left            =   -74940
         TabIndex        =   207
         Top             =   1965
         Width           =   1560
      End
      Begin VB.Label Label58 
         AutoSize        =   -1  'True
         Caption         =   "代表人5(中,英,日)："
         Height          =   180
         Left            =   -74940
         TabIndex        =   206
         Top             =   2460
         Width           =   1560
      End
      Begin VB.Label Label59 
         AutoSize        =   -1  'True
         Caption         =   "代表人2(中,英,日)："
         Height          =   180
         Left            =   -74940
         TabIndex        =   205
         Top             =   990
         Width           =   1560
      End
      Begin VB.Label Label60 
         AutoSize        =   -1  'True
         Caption         =   "代表人6(中,英,日)："
         Height          =   180
         Left            =   -74940
         TabIndex        =   204
         Top             =   2940
         Width           =   1560
      End
      Begin MSForms.TextBox txt1 
         Height          =   465
         Index           =   57
         Left            =   -68280
         TabIndex        =   203
         Top             =   4890
         Width           =   2475
         VariousPropertyBits=   -1466941409
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "4366;811"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   465
         Index           =   56
         Left            =   -70785
         TabIndex        =   202
         Top             =   4890
         Width           =   2475
         VariousPropertyBits=   -1466941409
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "4366;811"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   465
         Index           =   55
         Left            =   -73275
         TabIndex        =   201
         Top             =   4890
         Width           =   2475
         VariousPropertyBits=   -1466941409
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "4366;811"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   465
         Index           =   54
         Left            =   -68280
         TabIndex        =   200
         Top             =   4395
         Width           =   2475
         VariousPropertyBits=   -1466941409
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "4366;820"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   465
         Index           =   53
         Left            =   -70785
         TabIndex        =   199
         Top             =   4395
         Width           =   2475
         VariousPropertyBits=   -1466941409
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "4366;820"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   465
         Index           =   52
         Left            =   -73275
         TabIndex        =   198
         Top             =   4395
         Width           =   2475
         VariousPropertyBits=   -1466941409
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "4366;820"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   465
         Index           =   51
         Left            =   -68280
         TabIndex        =   197
         Top             =   3915
         Width           =   2475
         VariousPropertyBits=   -1466941409
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "4366;820"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   465
         Index           =   50
         Left            =   -70785
         TabIndex        =   196
         Top             =   3915
         Width           =   2475
         VariousPropertyBits=   -1466941409
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "4366;820"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   465
         Index           =   49
         Left            =   -73275
         TabIndex        =   195
         Top             =   3915
         Width           =   2475
         VariousPropertyBits=   -1466941409
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "4366;820"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   465
         Index           =   48
         Left            =   -68280
         TabIndex        =   194
         Top             =   3420
         Width           =   2475
         VariousPropertyBits=   -1466941409
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "4366;820"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   465
         Index           =   47
         Left            =   -70785
         TabIndex        =   193
         Top             =   3420
         Width           =   2475
         VariousPropertyBits=   -1466941409
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "4366;820"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   465
         Index           =   46
         Left            =   -73275
         TabIndex        =   192
         Top             =   3420
         Width           =   2475
         VariousPropertyBits=   -1466941409
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "4366;820"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   465
         Index           =   45
         Left            =   -68280
         TabIndex        =   191
         Top             =   2940
         Width           =   2475
         VariousPropertyBits=   -1466941409
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "4366;820"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   465
         Index           =   44
         Left            =   -70785
         TabIndex        =   190
         Top             =   2940
         Width           =   2475
         VariousPropertyBits=   -1466941409
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "4366;820"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   465
         Index           =   43
         Left            =   -73275
         TabIndex        =   189
         Top             =   2940
         Width           =   2475
         VariousPropertyBits=   -1466941409
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "4366;820"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   465
         Index           =   42
         Left            =   -68280
         TabIndex        =   188
         Top             =   2460
         Width           =   2475
         VariousPropertyBits=   -1466941409
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "4366;820"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   465
         Index           =   41
         Left            =   -70785
         TabIndex        =   187
         Top             =   2460
         Width           =   2475
         VariousPropertyBits=   -1466941409
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "4366;820"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   465
         Index           =   40
         Left            =   -73275
         TabIndex        =   186
         Top             =   2460
         Width           =   2475
         VariousPropertyBits=   -1466941409
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "4366;820"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   465
         Index           =   39
         Left            =   -68280
         TabIndex        =   185
         Top             =   1965
         Width           =   2475
         VariousPropertyBits=   -1466941409
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "4366;820"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   465
         Index           =   38
         Left            =   -70785
         TabIndex        =   184
         Top             =   1965
         Width           =   2475
         VariousPropertyBits=   -1466941409
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "4366;820"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   465
         Index           =   37
         Left            =   -73275
         TabIndex        =   183
         Top             =   1965
         Width           =   2475
         VariousPropertyBits=   -1466941409
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "4366;820"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   465
         Index           =   36
         Left            =   -68280
         TabIndex        =   182
         Top             =   1485
         Width           =   2475
         VariousPropertyBits=   -1466941409
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "4366;820"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   465
         Index           =   35
         Left            =   -70785
         TabIndex        =   181
         Top             =   1485
         Width           =   2475
         VariousPropertyBits=   -1466941409
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "4366;820"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   465
         Index           =   34
         Left            =   -73275
         TabIndex        =   180
         Top             =   1485
         Width           =   2475
         VariousPropertyBits=   -1466941409
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "4366;820"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   465
         Index           =   9
         Left            =   -73275
         TabIndex        =   179
         Top             =   510
         Width           =   2475
         VariousPropertyBits=   -1466941409
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "4366;811"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   465
         Index           =   10
         Left            =   -70785
         TabIndex        =   178
         Top             =   510
         Width           =   2475
         VariousPropertyBits=   -1466941409
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "4366;811"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   465
         Index           =   11
         Left            =   -68280
         TabIndex        =   177
         Top             =   510
         Width           =   2475
         VariousPropertyBits=   -1466941409
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "4366;811"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   465
         Index           =   12
         Left            =   -73275
         TabIndex        =   176
         Top             =   990
         Width           =   2475
         VariousPropertyBits=   -1466941409
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "4366;820"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   465
         Index           =   13
         Left            =   -70785
         TabIndex        =   175
         Top             =   990
         Width           =   2475
         VariousPropertyBits=   -1466941409
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "4366;820"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   465
         Index           =   14
         Left            =   -68280
         TabIndex        =   174
         Top             =   990
         Width           =   2475
         VariousPropertyBits=   -1466941409
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "4366;820"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   50
         Left            =   -73470
         TabIndex        =   173
         Top             =   1500
         Width           =   6765
         BackColor       =   -2147483639
         VariousPropertyBits=   27
         Caption         =   " "
         Size            =   "11933;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   49
         Left            =   -73620
         TabIndex        =   172
         Top             =   1200
         Width           =   885
         BackColor       =   -2147483639
         VariousPropertyBits=   27
         Caption         =   " "
         Size            =   "1561;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   5
         Left            =   -73620
         TabIndex        =   171
         Top             =   900
         Width           =   885
         BackColor       =   -2147483639
         VariousPropertyBits=   27
         Caption         =   " "
         Size            =   "1561;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   16
         Left            =   -73620
         TabIndex        =   170
         Top             =   600
         Width           =   885
         BackColor       =   -2147483639
         VariousPropertyBits=   27
         Caption         =   " "
         Size            =   "1561;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label81 
         AutoSize        =   -1  'True
         Caption         =   "北所銷卷日："
         Height          =   180
         Left            =   -74790
         TabIndex        =   169
         Top             =   600
         Width           =   1080
      End
      Begin VB.Label Label80 
         AutoSize        =   -1  'True
         Caption         =   "分所銷卷日："
         Height          =   180
         Index           =   0
         Left            =   -74790
         TabIndex        =   168
         Top             =   915
         Width           =   1080
      End
      Begin VB.Label Label79 
         AutoSize        =   -1  'True
         Caption         =   "分所銷卷員："
         Height          =   180
         Left            =   -74790
         TabIndex        =   167
         Top             =   1215
         Width           =   1080
      End
      Begin VB.Label Label78 
         AutoSize        =   -1  'True
         Caption         =   "分所銷卷備註："
         Height          =   180
         Left            =   -74790
         TabIndex        =   166
         Top             =   1530
         Width           =   1260
      End
      Begin MSForms.Label lbl1 
         Height          =   260
         Index           =   56
         Left            =   -73410
         TabIndex        =   157
         Top             =   930
         Width           =   260
         BackColor       =   -2147483639
         VariousPropertyBits=   27
         Caption         =   " "
         Size            =   "450;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   260
         Index           =   62
         Left            =   -69340
         TabIndex        =   156
         Top             =   930
         Width           =   210
         BackColor       =   -2147483639
         VariousPropertyBits=   27
         Caption         =   " "
         Size            =   "370;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label109 
         AutoSize        =   -1  'True
         Caption         =   "請款單份數："
         Height          =   180
         Left            =   -73060
         TabIndex        =   165
         Top             =   630
         Width           =   1080
      End
      Begin MSForms.Label lbl1 
         Height          =   260
         Index           =   61
         Left            =   -71940
         TabIndex        =   164
         Top             =   600
         Width           =   290
         BackColor       =   -2147483639
         VariousPropertyBits=   27
         Caption         =   " "
         Size            =   "503;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label108 
         AutoSize        =   -1  'True
         Caption         =   "定稿份數："
         Height          =   180
         Left            =   -74520
         TabIndex        =   163
         Top             =   630
         Width           =   900
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   60
         Left            =   -73590
         TabIndex        =   162
         Top             =   600
         Width           =   255
         BackColor       =   -2147483639
         VariousPropertyBits=   27
         Caption         =   " "
         Size            =   "450;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label110 
         AutoSize        =   -1  'True
         Caption         =   "EMail同時寄紙本：       (Y:是)"
         Height          =   180
         Left            =   -70870
         TabIndex        =   161
         Top             =   960
         Width           =   2280
      End
      Begin VB.Label Label102 
         AutoSize        =   -1  'True
         Caption         =   "以電子郵件通知：        (Y：是  D：僅D/N）"
         Height          =   180
         Left            =   -74880
         TabIndex        =   160
         Top             =   960
         Width           =   3375
      End
      Begin VB.Label Label113 
         AutoSize        =   -1  'True
         Caption         =   "特殊出名公司："
         Height          =   180
         Left            =   -74880
         TabIndex        =   159
         Top             =   1290
         Width           =   1260
      End
      Begin VB.Label Label112 
         AutoSize        =   -1  'True
         Caption         =   "(J:智權公司 空白:系統預設)"
         Height          =   180
         Left            =   -73260
         TabIndex        =   158
         Top             =   1290
         Width           =   2115
      End
      Begin MSForms.Label lbl1 
         Height          =   260
         Index           =   64
         Left            =   -73590
         TabIndex        =   155
         Top             =   1260
         Width           =   270
         BackColor       =   -2147483639
         VariousPropertyBits=   27
         Size            =   "476;459"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   650
         Index           =   65
         Left            =   -73590
         TabIndex        =   154
         Top             =   1590
         Width           =   7680
         BackColor       =   -2147483639
         Size            =   "13547;1147"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label114 
         AutoSize        =   -1  'True
         Caption         =   "定稿商標名稱："
         Height          =   180
         Left            =   -74880
         TabIndex        =   153
         Top             =   1620
         Width           =   1260
      End
      Begin MSForms.TextBox txt1 
         Height          =   285
         Index           =   61
         Left            =   1080
         TabIndex        =   146
         Top             =   2445
         Width           =   2175
         VariousPropertyBits=   671105055
         BackColor       =   16777215
         Size            =   "3836;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   285
         Index           =   60
         Left            =   6300
         TabIndex        =   145
         Top             =   2445
         Width           =   2175
         VariousPropertyBits=   671105055
         BackColor       =   16777215
         Size            =   "3836;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   285
         Index           =   59
         Left            =   1080
         TabIndex        =   144
         Top             =   570
         Width           =   2175
         VariousPropertyBits=   671105055
         BackColor       =   16777215
         Size            =   "3836;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   615
         Index           =   21
         Left            =   1080
         TabIndex        =   104
         Top             =   1470
         Width           =   8115
         VariousPropertyBits=   -1466941409
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "14314;1094"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   680
         Index           =   5
         Left            =   -73740
         TabIndex        =   20
         Top             =   4410
         Width           =   7880
         VariousPropertyBits=   -1466941409
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "13891;1199"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   680
         Index           =   4
         Left            =   -73740
         TabIndex        =   19
         Top             =   3420
         Width           =   7880
         VariousPropertyBits=   -1466941409
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "13891;1199"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   680
         Index           =   3
         Left            =   -73740
         TabIndex        =   18
         Top             =   2730
         Width           =   7890
         VariousPropertyBits=   -1466941409
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "13917;1199"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   420
         Index           =   2
         Left            =   -73890
         TabIndex        =   17
         Top             =   480
         Width           =   8070
         VariousPropertyBits=   -1466941409
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "14235;741"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   285
         Index           =   1
         Left            =   6300
         TabIndex        =   16
         Top             =   4800
         Width           =   2685
         VariousPropertyBits=   -1466941409
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "4736;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   285
         Index           =   0
         Left            =   1080
         TabIndex        =   15
         Top             =   2130
         Width           =   8115
         VariousPropertyBits=   -1466941409
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "14314;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label116 
         AutoSize        =   -1  'True
         Caption         =   "國內副本收件人："
         Height          =   180
         Left            =   -74920
         TabIndex        =   152
         Top             =   1260
         Width           =   1440
      End
      Begin MSForms.Label lbl1 
         Height          =   260
         Index           =   66
         Left            =   -73390
         TabIndex        =   151
         Top             =   1230
         Width           =   4080
         BackColor       =   -2147483639
         VariousPropertyBits=   27
         Caption         =   " "
         Size            =   "7197;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label115 
         AutoSize        =   -1  'True
         Caption         =   "國內副本接洽人："
         Height          =   180
         Left            =   -69120
         TabIndex        =   150
         Top             =   1260
         Width           =   1440
      End
      Begin MSForms.Label lbl1 
         Height          =   260
         Index           =   67
         Left            =   -67530
         TabIndex        =   149
         Top             =   1230
         Width           =   1470
         BackColor       =   -2147483639
         VariousPropertyBits=   27
         Caption         =   " "
         Size            =   "2593;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   63
         Left            =   6360
         TabIndex        =   148
         Top             =   5115
         Width           =   375
         BackColor       =   -2147483639
         VariousPropertyBits=   27
         Caption         =   " "
         Size            =   "661;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "不催延展：            (Y:不催)"
         Height          =   180
         Index           =   116
         Left            =   5355
         TabIndex        =   147
         Top             =   5115
         Width           =   2085
      End
      Begin VB.Label Label111 
         AutoSize        =   -1  'True
         Caption         =   "(Y.舊/A.新)"
         Height          =   180
         Left            =   4440
         TabIndex        =   143
         Top             =   4537
         Width           =   855
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   128
         Left            =   3990
         TabIndex        =   142
         Top             =   4500
         Width           =   375
         BackColor       =   -2147483639
         VariousPropertyBits=   27
         Caption         =   " "
         Size            =   "661;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "委任書："
         Height          =   180
         Index           =   2
         Left            =   3270
         TabIndex        =   141
         Top             =   4545
         Width           =   720
      End
      Begin VB.Label Label107 
         AutoSize        =   -1  'True
         Caption         =   "畫面上定稿語文："
         Height          =   180
         Left            =   3960
         TabIndex        =   140
         Top             =   4245
         Width           =   1440
      End
      Begin VB.Label Label106 
         AutoSize        =   -1  'True
         Caption         =   "(N:不印 1:台->各國 2:外->台 3:英文)"
         Height          =   180
         Left            =   5760
         TabIndex        =   139
         Top             =   4245
         Width           =   2745
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   59
         Left            =   5400
         TabIndex        =   138
         Top             =   4215
         Width           =   285
         BackColor       =   -2147483639
         VariousPropertyBits=   27
         Caption         =   " "
         Size            =   "503;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label105 
         AutoSize        =   -1  'True
         Caption         =   "(Y:自動代繳)"
         Height          =   180
         Left            =   -70560
         TabIndex        =   137
         Top             =   4140
         Width           =   1010
      End
      Begin VB.Label Label104 
         AutoSize        =   -1  'True
         Caption         =   "FCT註冊費自動代繳："
         Height          =   180
         Left            =   -72660
         TabIndex        =   136
         Top             =   4140
         Width           =   1760
      End
      Begin MSForms.Label lbl1 
         Height          =   260
         Index           =   58
         Left            =   -70890
         TabIndex        =   135
         Top             =   4110
         Width           =   270
         BackColor       =   -2147483639
         VariousPropertyBits=   27
         Caption         =   " "
         Size            =   "476;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   260
         Index           =   57
         Left            =   -67870
         TabIndex        =   134
         Top             =   1530
         Width           =   1920
         BackColor       =   -2147483639
         VariousPropertyBits=   27
         Caption         =   " "
         Size            =   "3387;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label103 
         AutoSize        =   -1  'True
         Caption         =   "接洽人："
         Height          =   180
         Left            =   -68670
         TabIndex        =   133
         Top             =   1560
         Width           =   720
      End
      Begin VB.Label Label76 
         AutoSize        =   -1  'True
         Caption         =   "延展聯絡人："
         Height          =   180
         Left            =   -74880
         TabIndex        =   132
         Top             =   4927
         Width           =   1080
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   47
         Left            =   -73740
         TabIndex        =   131
         Top             =   4890
         Width           =   2835
         BackColor       =   -2147483639
         VariousPropertyBits=   27
         Caption         =   " "
         Size            =   "5001;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   55
         Left            =   1560
         TabIndex        =   130
         Top             =   5085
         Width           =   3615
         BackColor       =   -2147483639
         VariousPropertyBits=   27
         Caption         =   " "
         Size            =   "6376;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label101 
         AutoSize        =   -1  'True
         Caption         =   "同意書商標號數："
         Height          =   180
         Left            =   150
         TabIndex        =   129
         Top             =   5085
         Width           =   1440
      End
      Begin VB.Label Label89 
         AutoSize        =   -1  'True
         Caption         =   "申請地址3(中)："
         Height          =   180
         Left            =   -74910
         TabIndex        =   128
         Top             =   3678
         Width           =   1290
      End
      Begin VB.Label Label88 
         AutoSize        =   -1  'True
         Caption         =   "申請地址3(日)："
         Height          =   180
         Left            =   -74910
         TabIndex        =   127
         Top             =   4410
         Width           =   1290
      End
      Begin VB.Label Label87 
         AutoSize        =   -1  'True
         Caption         =   "申請地址3(英)："
         Height          =   180
         Left            =   -74910
         TabIndex        =   126
         Top             =   4041
         Width           =   1290
      End
      Begin VB.Label Label86 
         AutoSize        =   -1  'True
         Caption         =   "申請地址2(中)："
         Height          =   180
         Left            =   -74910
         TabIndex        =   125
         Top             =   2589
         Width           =   1290
      End
      Begin VB.Label Label85 
         AutoSize        =   -1  'True
         Caption         =   "申請地址2(日)："
         Height          =   180
         Left            =   -74910
         TabIndex        =   124
         Top             =   3315
         Width           =   1290
      End
      Begin VB.Label Label84 
         AutoSize        =   -1  'True
         Caption         =   "申請地址2(英)："
         Height          =   180
         Left            =   -74910
         TabIndex        =   123
         Top             =   2952
         Width           =   1290
      End
      Begin VB.Label Label61 
         AutoSize        =   -1  'True
         Caption         =   "延展代理人："
         Height          =   180
         Left            =   -74880
         TabIndex        =   122
         Top             =   3652
         Width           =   1080
      End
      Begin VB.Label Label67 
         AutoSize        =   -1  'True
         Caption         =   "延展請款對象："
         Height          =   180
         Left            =   -74880
         TabIndex        =   121
         Top             =   4260
         Width           =   1260
      End
      Begin VB.Label Label68 
         AutoSize        =   -1  'True
         Caption         =   "延展彼所案號："
         Height          =   180
         Left            =   -74880
         TabIndex        =   120
         Top             =   3970
         Width           =   1260
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   37
         Left            =   -73530
         TabIndex        =   119
         Top             =   3615
         Width           =   7710
         BackColor       =   -2147483639
         VariousPropertyBits=   27
         Caption         =   " "
         Size            =   "13600;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   38
         Left            =   -73530
         TabIndex        =   118
         Top             =   4251
         Width           =   7470
         BackColor       =   -2147483639
         VariousPropertyBits=   27
         Caption         =   " "
         Size            =   "13176;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   40
         Left            =   -73530
         TabIndex        =   117
         Top             =   3933
         Width           =   2055
         BackColor       =   -2147483639
         VariousPropertyBits=   27
         Caption         =   " "
         Size            =   "3625;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   46
         Left            =   -73290
         TabIndex        =   116
         Top             =   4575
         Width           =   7155
         BackColor       =   -2147483639
         VariousPropertyBits=   27
         Caption         =   " "
         Size            =   "12621;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label75 
         AutoSize        =   -1  'True
         Caption         =   "延展D/N列印對象："
         Height          =   180
         Left            =   -74880
         TabIndex        =   115
         Top             =   4620
         Width           =   1545
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         Caption         =   "(Y:不跑)"
         Height          =   180
         Left            =   -69660
         TabIndex        =   114
         Top             =   3970
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.Label Label77 
         AutoSize        =   -1  'True
         Caption         =   "延展單筆不跑 :"
         Height          =   180
         Left            =   -71355
         TabIndex        =   113
         Top             =   3970
         Visible         =   0   'False
         Width           =   1170
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   4
         Left            =   -70110
         TabIndex        =   112
         Top             =   3933
         Visible         =   0   'False
         Width           =   405
         BackColor       =   -2147483639
         VariousPropertyBits=   27
         Caption         =   " "
         Size            =   "714;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         Caption         =   "申請人3："
         Height          =   180
         Left            =   -74910
         TabIndex        =   111
         Top             =   1147
         Width           =   810
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   52
         Left            =   -74010
         TabIndex        =   110
         Top             =   1110
         Width           =   8145
         BackColor       =   -2147483639
         VariousPropertyBits=   27
         Caption         =   " "
         Size            =   "14367;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "申請人2："
         Height          =   180
         Left            =   -74910
         TabIndex        =   109
         Top             =   862
         Width           =   810
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   51
         Left            =   -74010
         TabIndex        =   108
         Top             =   825
         Width           =   8145
         BackColor       =   -2147483639
         VariousPropertyBits=   27
         Caption         =   " "
         Size            =   "14367;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   15
         Left            =   6315
         TabIndex        =   81
         Top             =   1185
         Width           =   1200
         BackColor       =   -2147483639
         VariousPropertyBits=   27
         Size            =   "2117;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   3
         Left            =   7530
         TabIndex        =   107
         Top             =   1185
         Width           =   1650
         BackColor       =   -2147483639
         VariousPropertyBits=   27
         Size            =   "2910;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "特殊商標："
         Height          =   180
         Left            =   150
         TabIndex        =   106
         Top             =   4800
         Width           =   900
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   48
         Left            =   1080
         TabIndex        =   105
         Top             =   4800
         Width           =   4095
         BackColor       =   -2147483639
         VariousPropertyBits=   27
         Caption         =   " "
         Size            =   "7223;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   345
         Index           =   44
         Left            =   5910
         TabIndex        =   103
         Top             =   480
         Width           =   2295
         ForeColor       =   255
         BackColor       =   -2147483639
         VariousPropertyBits=   27
         Caption         =   "有相關卷號資料"
         Size            =   "4048;609"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   285
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "放棄專用權："
         Height          =   180
         Index           =   0
         Left            =   -70050
         TabIndex        =   102
         Top             =   5120
         Width           =   1080
      End
      Begin MSForms.Label lbl1 
         Height          =   260
         Index           =   43
         Left            =   -68970
         TabIndex        =   101
         Top             =   5120
         Width           =   3140
         BackColor       =   -2147483639
         VariousPropertyBits=   27
         Caption         =   " "
         Size            =   "5530;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   35
         Left            =   -74010
         TabIndex        =   96
         Top             =   555
         Width           =   8145
         BackColor       =   -2147483639
         VariousPropertyBits=   27
         Caption         =   " "
         Size            =   "14367;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   260
         Index           =   33
         Left            =   -69180
         TabIndex        =   95
         Top             =   2130
         Width           =   3350
         BackColor       =   -2147483639
         VariousPropertyBits=   27
         Caption         =   " "
         Size            =   "5900;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   260
         Index           =   32
         Left            =   -70740
         TabIndex        =   94
         Top             =   1530
         Width           =   690
         BackColor       =   -2147483639
         VariousPropertyBits=   27
         Caption         =   " "
         Size            =   "1217;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   260
         Index           =   31
         Left            =   -69570
         TabIndex        =   93
         Top             =   930
         Width           =   3710
         BackColor       =   -2147483639
         VariousPropertyBits=   27
         Caption         =   " "
         Size            =   "6535;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   260
         Index           =   30
         Left            =   -73450
         TabIndex        =   92
         Top             =   5120
         Width           =   3260
         BackColor       =   -2147483639
         VariousPropertyBits=   27
         Caption         =   " "
         Size            =   "5741;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   260
         Index           =   29
         Left            =   -73290
         TabIndex        =   91
         Top             =   4110
         Width           =   350
         BackColor       =   -2147483639
         VariousPropertyBits=   27
         Caption         =   " "
         Size            =   "609;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   260
         Index           =   28
         Left            =   -73680
         TabIndex        =   90
         Top             =   2430
         Width           =   7700
         BackColor       =   -2147483639
         VariousPropertyBits=   27
         Caption         =   " "
         Size            =   "13573;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   260
         Index           =   26
         Left            =   -73770
         TabIndex        =   89
         Top             =   2130
         Width           =   3350
         BackColor       =   -2147483639
         VariousPropertyBits=   27
         Caption         =   " "
         Size            =   "5900;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   260
         Index           =   25
         Left            =   -73650
         TabIndex        =   88
         Top             =   1530
         Width           =   690
         BackColor       =   -2147483639
         VariousPropertyBits=   27
         Caption         =   " "
         Size            =   "1217;459"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   260
         Index           =   24
         Left            =   -73890
         TabIndex        =   87
         Top             =   930
         Width           =   2870
         BackColor       =   -2147483639
         VariousPropertyBits=   27
         Caption         =   " "
         Size            =   "5054;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   23
         Left            =   6300
         TabIndex        =   86
         Top             =   4500
         Width           =   1545
         BackColor       =   -2147483639
         VariousPropertyBits=   27
         Caption         =   " "
         Size            =   "2725;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   22
         Left            =   5400
         TabIndex        =   85
         Top             =   3930
         Width           =   285
         BackColor       =   -2147483639
         VariousPropertyBits=   27
         Caption         =   " "
         Size            =   "503;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   21
         Left            =   5400
         TabIndex        =   84
         Top             =   3645
         Width           =   285
         BackColor       =   -2147483639
         VariousPropertyBits=   27
         Caption         =   " "
         Size            =   "503;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   20
         Left            =   6690
         TabIndex        =   83
         Top             =   3345
         Width           =   2520
         BackColor       =   -2147483639
         VariousPropertyBits=   27
         Caption         =   " "
         Size            =   "4445;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   19
         Left            =   6510
         TabIndex        =   82
         Top             =   3060
         Width           =   2415
         BackColor       =   -2147483639
         VariousPropertyBits=   27
         Caption         =   " "
         Size            =   "4260;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   14
         Left            =   1080
         TabIndex        =   80
         Top             =   4500
         Width           =   1890
         BackColor       =   -2147483639
         VariousPropertyBits=   27
         Caption         =   " "
         Size            =   "3334;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   13
         Left            =   1080
         TabIndex        =   79
         Top             =   4215
         Width           =   510
         BackColor       =   -2147483639
         VariousPropertyBits=   27
         Caption         =   " "
         Size            =   "900;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   12
         Left            =   1590
         TabIndex        =   78
         Top             =   3930
         Width           =   330
         BackColor       =   -2147483639
         VariousPropertyBits=   27
         Caption         =   " "
         Size            =   "582;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   11
         Left            =   1080
         TabIndex        =   77
         Top             =   3645
         Width           =   360
         BackColor       =   -2147483639
         VariousPropertyBits=   27
         Caption         =   " "
         Size            =   "635;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   10
         Left            =   2760
         TabIndex        =   76
         Top             =   3345
         Width           =   1890
         BackColor       =   -2147483639
         VariousPropertyBits=   27
         Caption         =   " "
         Size            =   "3334;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   9
         Left            =   1080
         TabIndex        =   75
         Top             =   3345
         Width           =   1260
         BackColor       =   -2147483639
         VariousPropertyBits=   27
         Caption         =   " "
         Size            =   "2222;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   8
         Left            =   1080
         TabIndex        =   74
         Top             =   3060
         Width           =   4170
         BackColor       =   -2147483639
         VariousPropertyBits=   27
         Caption         =   " "
         Size            =   "7355;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   7
         Left            =   1080
         TabIndex        =   73
         Top             =   2760
         Width           =   4170
         BackColor       =   -2147483639
         VariousPropertyBits=   27
         Caption         =   " "
         Size            =   "7355;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   2
         Left            =   1080
         TabIndex        =   72
         Top             =   1185
         Width           =   4245
         BackColor       =   -2147483639
         VariousPropertyBits=   27
         Caption         =   " "
         Size            =   "7488;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   1
         Left            =   1080
         TabIndex        =   71
         Top             =   885
         Width           =   4080
         BackColor       =   -2147483639
         VariousPropertyBits=   27
         Caption         =   " "
         Size            =   "7197;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label66 
         AutoSize        =   -1  'True
         Caption         =   "申請地址1(英)："
         Height          =   180
         Left            =   -74910
         TabIndex        =   70
         Top             =   1863
         Width           =   1290
      End
      Begin VB.Label Label65 
         AutoSize        =   -1  'True
         Caption         =   "申請地址1(日)："
         Height          =   180
         Left            =   -74910
         TabIndex        =   69
         Top             =   2226
         Width           =   1290
      End
      Begin VB.Label Label56 
         AutoSize        =   -1  'True
         Caption         =   "申請地址1(中)："
         Height          =   180
         Left            =   -74910
         TabIndex        =   68
         Top             =   1500
         Width           =   1290
      End
      Begin VB.Label Label54 
         AutoSize        =   -1  'True
         Caption         =   "全部折扣："
         Height          =   180
         Left            =   -74910
         TabIndex        =   67
         Top             =   1560
         Width           =   900
      End
      Begin VB.Label Label53 
         AutoSize        =   -1  'True
         Caption         =   "副本收受人："
         Height          =   180
         Left            =   -74920
         TabIndex        =   66
         Top             =   2160
         Width           =   1080
      End
      Begin VB.Label Label52 
         AutoSize        =   -1  'True
         Caption         =   "監視系統案號："
         Height          =   180
         Left            =   -74920
         TabIndex        =   65
         Top             =   2460
         Width           =   1260
      End
      Begin VB.Label Label51 
         AutoSize        =   -1  'True
         Caption         =   "代理人備註："
         Height          =   180
         Index           =   0
         Left            =   -74920
         TabIndex        =   64
         Top             =   3510
         Width           =   1080
      End
      Begin VB.Label Label50 
         AutoSize        =   -1  'True
         Caption         =   "代理人收款後辦案："
         Height          =   180
         Left            =   -74920
         TabIndex        =   63
         Top             =   4140
         Width           =   1620
      End
      Begin VB.Label Label49 
         AutoSize        =   -1  'True
         Caption         =   "客戶備註："
         Height          =   180
         Index           =   0
         Left            =   -74920
         TabIndex        =   62
         Top             =   4470
         Width           =   900
      End
      Begin VB.Label Label48 
         AutoSize        =   -1  'True
         Caption         =   "客戶收款後辦案："
         Height          =   180
         Left            =   -74920
         TabIndex        =   61
         Top             =   5120
         Width           =   1440
      End
      Begin VB.Label Label47 
         AutoSize        =   -1  'True
         Caption         =   "客戶案件案號："
         Height          =   180
         Left            =   -70860
         TabIndex        =   60
         Top             =   960
         Width           =   1260
      End
      Begin VB.Label Label46 
         AutoSize        =   -1  'True
         Caption         =   "申請/翻譯折扣："
         Height          =   180
         Left            =   -72070
         TabIndex        =   59
         Top             =   1560
         Width           =   1310
      End
      Begin VB.Label Label45 
         AutoSize        =   -1  'True
         Caption         =   "副本聯絡人："
         Height          =   180
         Left            =   -70330
         TabIndex        =   58
         Top             =   2160
         Width           =   1080
      End
      Begin VB.Label Label43 
         AutoSize        =   -1  'True
         Caption         =   "(％)"
         Height          =   180
         Left            =   -69990
         TabIndex        =   57
         Top             =   1560
         Width           =   300
      End
      Begin VB.Label Label42 
         AutoSize        =   -1  'True
         Caption         =   "(％)"
         Height          =   180
         Left            =   -72870
         TabIndex        =   56
         Top             =   1570
         Width           =   300
      End
      Begin VB.Label Label40 
         AutoSize        =   -1  'True
         Caption         =   "分所案號："
         Height          =   180
         Left            =   -74920
         TabIndex        =   55
         Top             =   960
         Width           =   900
      End
      Begin VB.Label Label39 
         AutoSize        =   -1  'True
         Caption         =   "閉卷原因："
         Height          =   180
         Left            =   5415
         TabIndex        =   54
         Top             =   4800
         Width           =   900
      End
      Begin VB.Label Label38 
         AutoSize        =   -1  'True
         Caption         =   "(1.中文  2.英文  3.日文)"
         Height          =   180
         Left            =   1710
         TabIndex        =   53
         Top             =   4245
         Width           =   1785
      End
      Begin VB.Label Label37 
         AutoSize        =   -1  'True
         Caption         =   "閉卷日期："
         Height          =   180
         Left            =   150
         TabIndex        =   52
         Top             =   4545
         Width           =   900
      End
      Begin VB.Label Label36 
         AutoSize        =   -1  'True
         Caption         =   "定稿語文："
         Height          =   180
         Left            =   150
         TabIndex        =   51
         Top             =   4245
         Width           =   900
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "申請人1："
         Height          =   180
         Left            =   -74910
         TabIndex        =   50
         Top             =   592
         Width           =   810
      End
      Begin VB.Label Label35 
         AutoSize        =   -1  'True
         Caption         =   "卷宗性質："
         Height          =   180
         Left            =   150
         TabIndex        =   49
         Top             =   930
         Width           =   900
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "商標組群："
         Height          =   180
         Left            =   -74920
         TabIndex        =   48
         Top             =   530
         Width           =   900
      End
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         Caption         =   "案件備註："
         Height          =   180
         Left            =   -74920
         TabIndex        =   47
         Top             =   2820
         Width           =   900
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "商品類別："
         Height          =   180
         Left            =   150
         TabIndex        =   46
         Top             =   2145
         Width           =   900
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "申請國家："
         Height          =   180
         Left            =   150
         TabIndex        =   45
         Top             =   1215
         Width           =   900
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "商標種類："
         Height          =   180
         Index           =   1
         Left            =   5415
         TabIndex        =   44
         Top             =   1222
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "本所案號："
         Height          =   180
         Index           =   1
         Left            =   150
         TabIndex        =   43
         Top             =   615
         Width           =   900
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "申請日："
         Height          =   180
         Index           =   0
         Left            =   150
         TabIndex        =   42
         Top             =   2490
         Width           =   720
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "審定來函日："
         Height          =   180
         Left            =   5415
         TabIndex        =   41
         Top             =   3090
         Width           =   1080
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "公告日："
         Height          =   180
         Left            =   150
         TabIndex        =   40
         Top             =   2805
         Width           =   720
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "正商標審定號："
         Height          =   180
         Left            =   5415
         TabIndex        =   39
         Top             =   3375
         Width           =   1260
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "申請案號："
         Height          =   180
         Left            =   5415
         TabIndex        =   38
         Top             =   2490
         Width           =   900
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "審定號："
         Height          =   180
         Left            =   5415
         TabIndex        =   37
         Top             =   2805
         Width           =   720
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "發證日："
         Height          =   180
         Left            =   150
         TabIndex        =   36
         Top             =   3090
         Width           =   720
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "專用權是否存在："
         Height          =   180
         Left            =   3960
         TabIndex        =   35
         Top             =   3675
         Width           =   1440
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "目前准駁："
         Height          =   180
         Left            =   150
         TabIndex        =   34
         Top             =   3675
         Width           =   900
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "專用期限："
         Height          =   180
         Left            =   150
         TabIndex        =   33
         Top             =   3375
         Width           =   900
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "(1. 准  2. 駁）"
         Height          =   180
         Index           =   1
         Left            =   1560
         TabIndex        =   32
         Top             =   3675
         Width           =   1050
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "(Y/N)"
         Height          =   180
         Left            =   5760
         TabIndex        =   31
         Top             =   3675
         Width           =   405
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         Caption         =   "是否閉卷："
         Height          =   180
         Left            =   5370
         TabIndex        =   30
         Top             =   4537
         Width           =   900
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         Caption         =   "（Y／N）"
         Height          =   180
         Left            =   7965
         TabIndex        =   29
         Top             =   4537
         Width           =   780
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "案件名稱："
         Height          =   180
         Left            =   150
         TabIndex        =   28
         Top             =   1500
         Width           =   900
      End
      Begin VB.Label Label23 
         Caption         =   "--"
         Height          =   180
         Left            =   2460
         TabIndex        =   27
         Top             =   3375
         Width           =   255
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         Caption         =   "是否有救濟程序："
         Height          =   180
         Left            =   150
         TabIndex        =   26
         Top             =   3960
         Width           =   1440
      End
      Begin VB.Label Label32 
         AutoSize        =   -1  'True
         Caption         =   "是否有爭議程序："
         Height          =   180
         Left            =   3960
         TabIndex        =   25
         Top             =   3960
         Width           =   1440
      End
      Begin VB.Label Label33 
         AutoSize        =   -1  'True
         Caption         =   "(Y 有)"
         Height          =   180
         Left            =   2025
         TabIndex        =   24
         Top             =   3960
         Width           =   465
      End
      Begin VB.Label Label34 
         AutoSize        =   -1  'True
         Caption         =   "(Y 有)"
         Height          =   180
         Left            =   5760
         TabIndex        =   23
         Top             =   3960
         Width           =   465
      End
   End
   Begin MSForms.Label lbl1 
      Height          =   255
      Index           =   42
      Left            =   5925
      TabIndex        =   100
      Top             =   5895
      Width           =   3195
      BackColor       =   -2147483639
      VariousPropertyBits=   27
      Caption         =   " "
      Size            =   "5636;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbl1 
      Height          =   255
      Index           =   41
      Left            =   1020
      TabIndex        =   99
      Top             =   5895
      Width           =   3825
      BackColor       =   -2147483639
      VariousPropertyBits=   27
      Caption         =   " "
      Size            =   "6747;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label49 
      Caption         =   "Create ID："
      Height          =   180
      Index           =   1
      Left            =   120
      TabIndex        =   98
      Top             =   5925
      Width           =   855
   End
   Begin VB.Label Label51 
      Caption         =   "Update ID："
      Height          =   180
      Index           =   1
      Left            =   4920
      TabIndex        =   97
      Top             =   5925
      Width           =   975
   End
   Begin VB.Label Label28 
      Caption         =   "備註："
      Height          =   252
      Left            =   -74880
      TabIndex        =   22
      Top             =   120
      Width           =   852
   End
End
Attribute VB_Name = "frm100101_4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2025/10/23 TF基礎案號(TM06,TM07)改成可以輸入多筆(Table: TFBaseNo)，放在銷卷頁籤。
'Memo by Lydia 2021/11/30 改成Form2.0 ; lbl1(index)、txt1(index)、grdDataList2
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/25 員工編號欄已修改
'sonia 2010/8/26 日期欄已修改
Option Explicit

Dim StrTag As String, StrTag1 As String
Dim strTemp As String, intK As Integer
'92.04.16 nick 紀錄作用按鍵
Public cmdState As Integer
'add by nick 2004/10/05 檢查是否已經有商品及服務
Public ChkTG As Boolean
'Add By Sindy 2010/02/04
Dim StrTag2 As String, StrTag3 As String, StrTag4 As String, StrTag5 As String
Public m_pub_QL05 As String 'Add By Sindy 2025/8/28 只記錄於此Form


'92.04.16 nick
Public Sub PubShowNextData()
Select Case cmdState
Case 0
     cmdState = -1
     Me.Enabled = False
     If fnSaveParentForm(Me) = False Then
        Me.Enabled = True
        Exit Sub
     End If
     Screen.MousePointer = vbHourglass
     frm100101_10.Show
     frm100101_10.Tag = StrTag ' StrTag  傳代理人代號
     frm100101_10.m_pub_QL05 = pub_QL05 'Add By Sindy 2025/8/28
     frm100101_10.StrMenu
     Screen.MousePointer = vbDefault
     Me.Enabled = True
Case 1
     cmdState = -1
     Me.Enabled = False
     If fnSaveParentForm(Me) = False Then
        Me.Enabled = True
        Exit Sub
     End If
     Screen.MousePointer = vbHourglass
     frm100101_11.Show
     frm100101_11.Tag = StrTag1 ' StrTag    傳申請人代號
     frm100101_11.m_pub_QL05 = pub_QL05 'Add By Sindy 2025/8/28
     frm100101_11.StrMenu
     Screen.MousePointer = vbDefault
     Me.Enabled = True
Case 2
     tmpBol = fnCancelNowFormAndShowParentForm(Me)
Case 3
     fnCloseAllFrm100
'add by nick 2004/09/15
'相關卷號
Case 4
     cmdState = -1
     Me.Enabled = False
     If fnSaveParentForm(Me) = False Then
        Me.Enabled = True
        Exit Sub
     End If
     Screen.MousePointer = vbHourglass
     frm100108_3.Show
     frm100108_3.Tag = txt1(59).Text
     frm100108_3.Caption = "相關卷號"
     frm100108_3.StrMenu2
     Screen.MousePointer = vbDefault
     Me.Enabled = True
'分割案
Case 5
     cmdState = -1
     Me.Enabled = False
     If fnSaveParentForm(Me) = False Then
        Me.Enabled = True
        Exit Sub
     End If
     Screen.MousePointer = vbHourglass
     frm100108_4.Show
     frm100108_4.frm100108_txt_7 = "3"
     frm100108_4.SetDataListWidth
     frm100108_4.Tag = txt1(59).Text
     frm100108_4.Caption = "分割案"
     frm100108_4.StrMenu1
     Screen.MousePointer = vbDefault
     Me.Enabled = True
'add by nick 2004/10/05
Case 6
    frm03010303_04.Hide
    Set frm03010303_04.UpForm = Me
    frm03010303_04.TGKey = txt1(59).Text
    frm03010303_04.AllClass = txt1(0).Text
    frm03010303_04.cmdOK(0).Visible = False
    frm03010303_04.cmd.Visible = False
    frm03010303_04.Cmd2.Visible = False
    frm03010303_04.txt2(0).Visible = False
    frm03010303_04.Line1.Visible = False
    frm03010303_04.txt2(1).Visible = False
    frm03010303_04.txt2(2).Visible = False
    frm03010303_04.txt2(3).Visible = False
    frm03010303_04.Caption = "商品及服務資料"
    'edit by nickc 2008/02/12 改成可以複製
    'frm03010303_04.TXT1(0).Enabled = False
    'frm03010303_04.TXT1(1).Enabled = False
    'frm03010303_04.TXT1(2).Enabled = False
    frm03010303_04.txt1(0).Locked = True
    frm03010303_04.txt1(1).Locked = True
    frm03010303_04.txt1(2).Locked = True
    frm03010303_04.Label2.Visible = False
    Me.Hide
    frm03010303_04.QueryData
    If Trim(txt1(0).Text) <> "" Then 'Add By Sindy 2014/4/29 +if 有商品類別才show
      frm03010303_04.Show vbModal 'Modify By Sindy 2009/09/17 改為強制回應表單
    End If
'Add By Sindy 98/04/09
Case 7
    frmPic001.oCP01 = SystemNumber(txt1(59), 1)
    frmPic001.oCP02 = SystemNumber(txt1(59), 2)
    frmPic001.oCP03 = SystemNumber(txt1(59), 3)
    frmPic001.oCP04 = SystemNumber(txt1(59), 4)
    frmPic001.StrMenu
    frmPic001.CanScan
    frmPic001.SetSeekCmdok 'Add by Amy 2018/07/16
    frmPic001.Show vbModal
    'add by nickc 2005/12/15 檢查有無代表圖
    'Modify by Amy 2018/07/16  改寫至function
'    strSql = "SELECT ibf01,ibf02 FROM imgbytefile WHERE ibf01='" & SystemNumber(txt1(59), 1) & "' and ibf02='" & SystemNumber(txt1(59), 2) & "' and ibf03='" & SystemNumber(txt1(59), 3) & "' and ibf04='" & SystemNumber(txt1(59), 4) & "' and ibf05='1'"
'    CheckOC2
'    adoRecordset1.CursorLocation = adUseClient
'    adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
'    If adoRecordset1.RecordCount <> 0 And adoRecordset1.RecordCount > 0 Then
    If ChkImgByteFile(SystemNumber(txt1(59), 1), SystemNumber(txt1(59), 2), SystemNumber(txt1(59), 3), SystemNumber(txt1(59), 4)) = True Then
        'Modified by Lydia 2016/11/24
        'cmdOK(7).Caption = "已設定代表圖(&I)"
        cmdOK(7).Caption = "已設定代表圖"
        cmdOK(7).BackColor = &HC0FFC0
    Else
        'Modified by Lydia 2016/11/24
        'cmdOK(7).Caption = "未設定代表圖(&I)"
        cmdOK(7).Caption = "未設定代表圖"
        cmdOK(7).BackColor = &HC0C0FF
    End If
'    CheckOC2
    'end 2018/07/16
'98/04/09 End
'Add By Sindy 2010/02/04
Case 8
     cmdState = -1
     Me.Enabled = False
     If fnSaveParentForm(Me) = False Then
        Me.Enabled = True
        Exit Sub
     End If
     Screen.MousePointer = vbHourglass
     frm100101_11.Show
     frm100101_11.Tag = StrTag2 ' StrTag2    傳申請人2代號
     frm100101_11.m_pub_QL05 = pub_QL05 'Add By Sindy 2025/8/28
     frm100101_11.StrMenu
     Screen.MousePointer = vbDefault
     Me.Enabled = True
Case 9
     cmdState = -1
     Me.Enabled = False
     If fnSaveParentForm(Me) = False Then
        Me.Enabled = True
        Exit Sub
     End If
     Screen.MousePointer = vbHourglass
     frm100101_11.Show
     frm100101_11.Tag = StrTag3 ' StrTag3    傳申請人3代號
     frm100101_11.m_pub_QL05 = pub_QL05 'Add By Sindy 2025/8/28
     frm100101_11.StrMenu
     Screen.MousePointer = vbDefault
     Me.Enabled = True
Case 10
     cmdState = -1
     Me.Enabled = False
     If fnSaveParentForm(Me) = False Then
        Me.Enabled = True
        Exit Sub
     End If
     Screen.MousePointer = vbHourglass
     frm100101_11.Show
     frm100101_11.Tag = StrTag4 ' StrTag4    傳申請人4代號
     frm100101_11.m_pub_QL05 = pub_QL05 'Add By Sindy 2025/8/28
     frm100101_11.StrMenu
     Screen.MousePointer = vbDefault
     Me.Enabled = True
Case 11
     cmdState = -1
     Me.Enabled = False
     If fnSaveParentForm(Me) = False Then
        Me.Enabled = True
        Exit Sub
     End If
     Screen.MousePointer = vbHourglass
     frm100101_11.Show
     frm100101_11.Tag = StrTag5 ' StrTag5    傳申請人5代號
     frm100101_11.m_pub_QL05 = pub_QL05 'Add By Sindy 2025/8/28
     frm100101_11.StrMenu
     Screen.MousePointer = vbDefault
     Me.Enabled = True
'2010/02/04 End
'Added by Lydia 2016/11/23
Case 12 '各項指示
     'Added by Lydia 2020/05/05 各項指示：檢查表單是否開啟中
     If PUB_CheckFormExist("frm12040159") Then
         MsgBox "請先關閉〔申請人/代理人/案件各項指示資料〕的畫面！", vbInformation
         Exit Sub
     End If
     'end 2020/05/05
     
     cmdState = -1
     Me.Enabled = False
     If fnSaveParentForm(Me) = False Then
        Me.Enabled = True
        Exit Sub
     End If
     Screen.MousePointer = vbHourglass
     frm12040159.SetParent "Q", Trim(Replace(txt1(59), "-", "")), Me
     frm12040159.Show
     Screen.MousePointer = vbDefault
     Me.Enabled = True
'end 2016/11/23
'Add By Sindy 2020/7/15
Case 13 '進度
   cmdState = -1
   Me.Enabled = False
   If fnSaveParentForm(Me) = False Then
      Me.Enabled = True
      Exit Sub
   End If
   Screen.MousePointer = vbHourglass
   frm100101_2.Show
   frm100101_2.Tag = txt1(59)
   frm100101_2.StrMenu
   Screen.MousePointer = vbDefault
   Me.Enabled = True
Case Else
End Select
End Sub

Private Sub cmdok_Click(Index As Integer)
'92.04.16 nick 紀錄作用按鍵
cmdState = Index
PubShowNextData
End Sub

Sub StrMenu()
'Add By Cheng 2002/07/08
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
Dim strSK03 As String
Dim strSql  As String
Dim Str01 As String, Str02 As String, Str03 As String, Str04 As String
'edit by nickc 2006/07/12
'Dim strArr(T_TM) As String, i As Integer, StrOkTxt(21) As String
Dim strArr() As String, i As Integer, StrOkTxt(21) As String
ReDim strArr(TF_TM) As String
'Modify By Cheng 2002/12/12
'Dim StrOk(43) As String
Dim StrOk(48) As String
'add by nickc 2007/08/27 秀玲說基本檔怎麼抓，就跟基本檔一樣，不用統一中英日或是英中日
Dim tmp01 As String, tmp02 As String, tmp03 As String
'add by Toni 20080926 控制跨部門權限訊息
Dim strTit As String
Dim strMsg As String
Dim nResponse
'end by Toni 20080926
'Added by Lydia 2025/10/23
Dim intR As Integer, strR1 As String
Dim rsRD As New ADODB.Recordset
'end 2025/10/23

'On Error GoTo ErrorHandler
'
'ReDim strArr(TF_TM) As String
Str01 = ""
Str02 = ""
Str03 = ""
Str04 = ""
If Left(Me.Tag, 1) = "N" Then
   strSql = Right(Me.Tag, Len(Me.Tag) - 1)
Else
   strSql = Me.Tag
End If
Str01 = SystemNumber(strSql, 1)
Str02 = SystemNumber(strSql, 2)
Str03 = SystemNumber(strSql, 3)
Str04 = SystemNumber(strSql, 4)

' 使用者沒有權限
'add by Toni 20080926 控制跨部門權限訊息 for 商標案件基本資料查詢
'2008/10/2 modify by sonia
'If IsUserHasRightOfSystem(strUserNum, Str01) = False Then
'   If IsUserHasRightOfFunction("frm100101_1", strCrossDept, False) = False Then
'      strTit = "檢核資料"
'      strMsg = "您沒有使用該系統類別的權限"
'      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
'      tmpBol = fnCancelNowFormAndShowParentForm(Me)
'      Exit Sub
'   End If
'End If
If CheckSR09(strUserNum, Str01, "Y", , Str01, Str02, Str03, Str04) = False Then
   tmpBol = fnCancelNowFormAndShowParentForm(Me)
   Exit Sub
End If
'2008/10/2 end
'End 20080926

pub_QL05 = ";本所案號：" & Str01 & "-" & Str02 & "-" & Str03 & "-" & Str04 & _
           "(基本資料)" 'Add By Sindy 2025/8/7

'2010/7/30 CANCEL BY SONIA 因內外商欲合併,故取消此控制
''2009/8/19 add by sonia FCT無爭議程序之案件內商人員不可查詢(該案有內商承辦人者為FCT爭議案)
'If Str01 = "FCT" And Mid(PUB_GetST03(strUserNum), 1, 2) = "P2" Then
'   StrSQLa = "Select * From CASEPROGRESS,STAFF Where CP01='" & Str01 & "' AND CP02='" & Str02 & "' AND CP03='" & Str03 & "' AND CP04='" & Str04 & "' AND CP14=ST01(+) AND SUBSTR(ST03,1,2)='P2' "
'   rsA.CursorLocation = adUseClient
'   rsA.Open StrSQLa, cnnConnection, adOpenStatic
'   If rsA.RecordCount = 0 Then
'      If rsA.State <> adStateClosed Then rsA.Close
'      Set rsA = Nothing
'      strTit = "檢核資料"
'      strMsg = "非FCT爭議案，您沒有使用該案號資料的權限"
'      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
'      tmpBol = fnCancelNowFormAndShowParentForm(Me)
'      Exit Sub
'   Else
'      If rsA.State <> adStateClosed Then rsA.Close
'      Set rsA = Nothing
'   End If
'End If
''2009/8/19 END
''2009/9/8 add by sonia T非台灣案非外商收文之案件,外商人員不可查詢
'If Str01 = "T" And Mid(PUB_GetST03(strUserNum), 1, 2) = "F1" Then
'   StrSQLa = "Select * From TRADEMARK Where TM01='" & Str01 & "' AND TM02='" & Str02 & "' AND TM03='" & Str03 & "' AND TM04='" & Str04 & "' "
'   rsA.CursorLocation = adUseClient
'   rsA.Open StrSQLa, cnnConnection, adOpenStatic
'   If rsA.RecordCount = 0 Then
'      If rsA.State <> adStateClosed Then rsA.Close
'      Set rsA = Nothing
'      strTit = "檢核資料"
'      strMsg = "無此商標資料"
'      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
'      tmpBol = fnCancelNowFormAndShowParentForm(Me)
'      Exit Sub
'   Else
'      If rsA.Fields("TM10") <> "000" Then '非台灣案才要控管外商人員權限
'         If rsA.State <> adStateClosed Then rsA.Close
'         Set rsA = Nothing
'         StrSQLa = "Select * From CASEPROGRESS Where CP01='" & Str01 & "' AND CP02='" & Str02 & "' AND CP03='" & Str03 & "' AND CP04='" & Str04 & "' AND SUBSTR(CP12,1,2)='F1' "
'         rsA.CursorLocation = adUseClient
'         rsA.Open StrSQLa, cnnConnection, adOpenStatic
'         If rsA.RecordCount = 0 Then
'            strMsg = "非外商收文之大陸商標案，您沒有使用該案號資料的權限"
'            strTit = "查詢資料"
'            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
'            tmpBol = fnCancelNowFormAndShowParentForm(Me)
'            If rsA.State <> adStateClosed Then rsA.Close
'            Set rsA = Nothing
'            Exit Sub
'         Else
'            If rsA.State <> adStateClosed Then rsA.Close
'            Set rsA = Nothing
'         End If
'      Else
'         If rsA.State <> adStateClosed Then rsA.Close
'         Set rsA = Nothing
'      End If
'   End If
'End If
''2009/9/8 END
'2010/7/30 END
   
'Add By Cheng 2002/07/08
strSK03 = ""
StrSQLa = "Select SK03 From SystemKind Where SK01='" & Str01 & "'"
rsA.CursorLocation = adUseClient
rsA.Open StrSQLa, cnnConnection, adOpenStatic
If rsA.RecordCount > 0 Then
   strSK03 = "" & rsA.Fields(0).Value
End If
If rsA.State <> adStateClosed Then rsA.Close
Set rsA = Nothing

strSql = "SELECT * FROM TRADEMARK WHERE TM01='" & Str01 & "' AND TM02='" & Str02 & "' AND TM03='" & Str03 & "' AND TM04='" & Str04 & "'"
CheckOC
adoRecordset.CursorLocation = adUseClient
adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
m_pub_QL05 = pub_QL05 'Add By Sindy 2025/8/28 記錄此Form的查詢條件
If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
   If pub_QL04 <> "" Then InsertQueryLog (adoRecordset.RecordCount) 'Add By Sindy 2025/8/7
    'For i = 0 To 66
    For i = 0 To (TF_TM - 1) 'edit by nickc 2006/07/12 (T_TM - 1)
       Select Case i
       Case 10, 12, 13, 19, 20, 21, 29, 35, 36, 59, 60, 62
            If IsNull(adoRecordset.Fields(i)) Then
                strArr(i + 1) = ""
            Else
                strArr(i + 1) = str(adoRecordset.Fields(i))
            End If
       Case Else
            If IsNull(adoRecordset.Fields(i)) Then
                 strArr(i + 1) = ""
            Else
                 strArr(i + 1) = adoRecordset.Fields(i)
            End If
       End Select
       'DoEvents Add By Sindy 2019/1/4 Mark,因為會和視窗的function(MenuForFormControl)有ErrCode互影響
    Next i
Else
   If pub_QL04 <> "" Then InsertQueryLog (0) 'Add By Sindy 2025/8/7
   ShowNoData
   Screen.MousePointer = vbDefault
       '920416 nick
     'Me.Hide
     tmpBol = fnCancelNowFormAndShowParentForm(Me)
   Exit Sub
End If
CheckOC
Dim strTemp As String    '暫存
Dim strTemp1 As Variant, strTemp2 As Variant, strTemp3 As Variant
Dim j As Integer
intK = 67
'For i = 0 To 67
For i = 1 To TF_TM 'edit by nickc 2006/07/12 T_TM
    Select Case i
    Case 1
         StrOk(0) = strArr(1) + "-" + strArr(2) + "-" + strArr(3) + "-" + strArr(4)
         txt1(59) = strArr(1) + "-" + strArr(2) + "-" + strArr(3) + "-" + strArr(4) 'Add By Sindy 2013/1/31
         'Modify by Amy 2014/04/10 +存取碼
         'Modified by Lydia 2016/10/19 +本所案號
         'strSql = "select PD05 AS  優先權日,PD06 AS 優先權號,NA03 AS 優先權國家,PD09 as 優先權存取碼 from PRIDATE,NATION WHERE PD01='" & strArr(1) & "' AND PD02='" & strArr(2) & "' AND PD03='" & strArr(3) & "' AND PD04 ='" & strArr(4) & "' AND PD07=NA01(+) ORDER BY PD01,PD02,PD03,PD04"
         'Modified by Lydia 2016/10/19 +本所案號
         'Modified by Sindy 2017/9/29 +商品類別
         strSql = "select PD05 AS  優先權日,PD06 AS 優先權號,NA03 AS 優先權國家,PD09 as 優先權存取碼," & _
                  "NVL(A1.TM01||A1.TM02||A1.TM03||A1.TM04,A2.TM01||A2.TM02||A2.TM03||A2.TM04) AS 本所案號,PD10 AS 商品類別 " & _
                  "from PRIDATE,NATION,TRADEMARK A1,TRADEMARK A2 " & _
                  "WHERE PD01='" & strArr(1) & "' AND PD02='" & strArr(2) & "' AND PD03='" & strArr(3) & "' AND PD04 ='" & strArr(4) & "' " & _
                  "AND PD06=A1.TM12(+) AND PD05=A1.TM11(+) AND PD07=A1.TM10(+) " & _
                  "AND PD06=A2.TM15(+) AND PD05=A2.TM11(+) AND PD07=A2.TM10(+) " & _
                  "AND PD07=NA01(+) " & _
                  "ORDER BY PD01,PD02,PD03,PD04"
         CheckOC
         adoRecordset.CursorLocation = adUseClient
         adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
         Set grdDataList2.Recordset = adoRecordset
         CheckOC
        'Add By Cheng 2002/12/12
         strSql = "SELECT CR05,CR06,CR07,CR08 FROM CASERELATION WHERE CR01='" & Str01 & "' AND CR02='" & Str02 & "' AND CR03='" & Str03 & "' AND CR04='" & Str04 & "'"
         CheckOC
         adoRecordset.CursorLocation = adUseClient
         adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
         If adoRecordset.RecordCount <> 0 Then
              StrOk(44) = "有相關卷號資料"
         Else
              StrOk(44) = ""
         End If
         CheckOC
        'Add By Sindy 98/04/09 檢查有無代表圖
        'Modify by Amy 2018/07/16  改寫至function
'        strSql = "SELECT ibf01,ibf02 FROM imgbytefile WHERE ibf01='" & Str01 & "' and ibf02='" & Str02 & "' and ibf03='" & Str03 & "' and ibf04='" & Str04 & "' and ibf05='1'"
'        CheckOC2
'        adoRecordset1.CursorLocation = adUseClient
'        adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
'        If adoRecordset1.RecordCount <> 0 And adoRecordset1.RecordCount > 0 Then
        If ChkImgByteFile(Str01, Str02, Str03, Str04) = True Then
            'Modified by Lydia 2016/11/24
            'cmdOK(7).Caption = "已設定代表圖(&I)"
            cmdOK(7).Caption = "已設定代表圖"
            cmdOK(7).BackColor = &HC0FFC0
        Else
            'Modified by Lydia 2016/11/24
            'cmdOK(7).Caption = "未設定代表圖(&I)"
            cmdOK(7).Caption = "未設定代表圖"
            cmdOK(7).BackColor = &HC0C0FF
        End If
'        CheckOC2
        'end 2018/07/16
        '98/04/09 End
    Case 28
         If strArr(i) = "1" Then
            StrOk(1) = strArr(i) + "  申請"
         Else
            If strArr(i) = "2" Then
                StrOk(1) = strArr(i) + "  異議"
            Else
                If strArr(i) = "3" Then
                    StrOk(1) = strArr(i) + "  評定"
                Else
                    If strArr(i) = "4" Then
                        StrOk(1) = strArr(i) + "  廢止"
                    Else
                        StrOk(1) = strArr(i) + "  錯誤"
                    End If
                End If
            End If
         End If
    Case 10
         strSql = "SELECT NA03 FROM NATION WHERE NA01='" & strArr(i) & "'"
         CheckOC
         adoRecordset.CursorLocation = adUseClient
         adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
         If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
              If IsNull(adoRecordset.Fields(0)) Then
                  StrOk(2) = strArr(i) + ""
              Else
                  StrOk(2) = strArr(i) + "  " + adoRecordset.Fields(0)
              End If
         Else
              StrOk(2) = ""
         End If
         CheckOC
    Case 5
'         StrOk(3) = strArr(i)
         If Not IsNull(strArr(i)) Then
             StrOkTxt(21) = strArr(i)
         Else
             StrOkTxt(21) = ""
         End If
     'add by nick 2004/11/24
     Case 68
          StrOk(4) = strArr(i)
'    Case 6
'         StrOk(4) = strArr(i)
'    Case 7
'         StrOk(5) = strArr(i)
    

    Case 9
         If Not IsNull(strArr(i)) Then
             StrOkTxt(0) = strArr(i)
         Else
             StrOkTxt(0) = ""
         End If
    Case 11
         If Len(Trim(strArr(i))) = 0 Or strArr(i) = "0" Then
             StrOk(6) = ""
             txt1(61) = "" 'Add By Sindy 2013/1/31
         Else
             StrOk(6) = ChangeWStringToTString(strArr(i))
             txt1(61) = ChangeWStringToTString(strArr(i)) 'Add By Sindy 2013/1/31
         End If
    Case 14
         If Len(Trim(strArr(i))) = 0 Or strArr(i) = "0" Then
             StrOk(7) = ""
         Else
             StrOk(7) = ChangeWStringToTString(strArr(i))
         End If
    Case 20
         If Len(Trim(strArr(i))) = 0 Or strArr(i) = "0" Then
             StrOk(8) = ""
         Else
             StrOk(8) = ChangeWStringToTString(strArr(i))
         End If
    Case 21
         StrOk(9) = strArr(i)
    Case 22
         StrOk(10) = strArr(i)
    Case 16
         StrOk(11) = strArr(i)
    Case 18
         StrOk(12) = strArr(i)
    Case 53
         StrOk(13) = strArr(i)
    Case 30
         If Len(Trim(strArr(i))) = 0 Or strArr(i) = "0" Then
             StrOk(14) = ""
         Else
             StrOk(14) = ChangeWStringToTString(strArr(i))
         End If
    Case 8
         strSql = "SELECT SK02 FROM SYSTEMKIND WHERE SK01='" & strArr(1) & "'"
         CheckOC
         adoRecordset.CursorLocation = adUseClient
         adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
         If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
            If IsNull(adoRecordset.Fields(0)) Then
               strTemp = ""
            Else
               strTemp = str(adoRecordset.Fields(0))
            End If
            CheckOC
            strSql = "SELECT PTM03,PTM04 FROM PATENTTRADEMARKMAP WHERE PTM01='" & Val(strTemp) & "' AND PTM02='" & strArr(i) & "'"
            CheckOC
            adoRecordset.CursorLocation = adUseClient
            adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
            If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
                '91.08.19  nick   cfp 時只抓  ptm03
                If UCase(Str01) = "CFP" Then
                    If IsNull(adoRecordset.Fields(0)) Then
                         StrOk(15) = strArr(i) + ""
                    Else
                         StrOk(15) = strArr(i) + "  " + adoRecordset.Fields(0)
                    End If
                Else
                    If strArr(10) = "000" Then
                        If IsNull(adoRecordset.Fields(0)) Then
                             StrOk(15) = strArr(i) + ""
                        Else
                             StrOk(15) = strArr(i) + "  " + adoRecordset.Fields(0)
                        End If
                    Else
                        If IsNull(adoRecordset.Fields(1)) Then
                             StrOk(15) = strArr(i) = ""
                        Else
                             StrOk(15) = strArr(i) + "  " + adoRecordset.Fields(1)
                        End If
                     End If
                End If
            Else
                StrOk(15) = ""
            End If
            CheckOC
         Else
            CheckOC
            StrOk(15) = ""
         End If
    Case 57
         'edit by nickc 2006/07/12
         'StrOk(16) = strArr(i)
         If Len(Trim(strArr(i))) = 0 Or strArr(i) = "0" Then
             StrOk(16) = ""
         Else
             StrOk(16) = ChangeTStringToTDateString(ChangeWStringToTString(strArr(i)))
         End If
    Case 12
         StrOk(17) = strArr(i)
         txt1(60) = strArr(i) 'Add By Sindy 2013/1/31
    Case 15
         StrOk(18) = strArr(i)
         txt1(62) = strArr(i) 'Add by Amy 2022/09/02
         strSql = "select SP01 from servicepractice where SP32='" & strArr(i) & "'"
         CheckOC
         adoRecordset.CursorLocation = adUseClient
         adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
         If adoRecordset.RecordCount <> 0 Then
              If UCase(Left(adoRecordset.Fields(0), 2)) = "TM" Then
                  StrOk(28) = adoRecordset.Fields(0)
              Else
                  StrOk(28) = ""
              End If
         Else
              StrOk(28) = ""
         End If
         CheckOC
    Case 13
         If Len(Trim(strArr(i))) = 0 Or strArr(i) = "0" Then
             StrOk(19) = ""
         Else
             StrOk(19) = ChangeWStringToTString(strArr(i))
         End If
    Case 27
         StrOk(20) = strArr(i)
    Case 17
         StrOk(21) = strArr(i)
    Case 19
         StrOk(22) = strArr(i)
    Case 29
         StrOk(23) = strArr(i)
    Case 31
         strSql = "SELECT ROR02 FROM REASONOFRELIEF WHERE ROR01='" & strArr(i) & "'"
         CheckOC
         adoRecordset.CursorLocation = adUseClient
         adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
         If adoRecordset.RecordCount <> 0 Then
            If Not IsNull(adoRecordset.Fields(0)) Then
                StrOkTxt(1) = adoRecordset.Fields(0)
            Else
                StrOkTxt(1) = ""
            End If
         Else
            StrOkTxt(1) = ""
         End If
         CheckOC
    Case 32
         StrOkTxt(2) = strArr(i)
    Case 34
         StrOk(24) = strArr(i)
    'Add by Morgan 2004/8/2
    '客戶案件案號
    Case 35
      StrOk(31) = strArr(i)
    Case 36 '全部折扣
         StrOk(25) = strArr(i)
    Case 37 '申請/翻譯折扣
         StrOk(32) = strArr(i)
    'Add By Sindy 2025/3/6
    Case 140 '繳註冊費折扣
         StrOk(0) = strArr(i)
    Case 141 '延展折扣
         StrOk(6) = strArr(i)
    '2025/3/6 END
    Case 54
         If Left$(strArr(i), 1) = "X" Then
'edit by nickc 2007/08/27 秀玲說基本檔怎麼抓，就跟基本檔一樣，不用統一中英日或是英中日
'             If Len(strArr(i)) = 9 Then
'                  strSQL = "SELECT cu05||' '||cu88||' '||cu89||' '||cu90,CU04,CU06 FROM CUSTOMER WHERE CU01='" & Left$(strArr(i), 8) & "' AND CU02='" & Right$(strArr(i), 1) & "'"
'             Else
'                  strSQL = "SELECT cu05||' '||cu88||' '||cu89||' '||cu90,CU04,CU06 FROM CUSTOMER WHERE CU01='" & Left$(strArr(i), 8) & "' AND CU02='0'"
'             End If
            StrOk(26) = GetAgentOrCustName(Trim(strArr(i)))
         Else
'edit by nickc 2007/08/27 秀玲說基本檔怎麼抓，就跟基本檔一樣，不用統一中英日或是英中日
'             If Len(strArr(i)) = 9 Then
'                  strSQL = "SELECT FA05||' '||FA63||' '||FA64||' '||FA65,FA04,FA06,FA29 FROM FAGENT WHERE FA01='" & Left$(Trim(strArr(i)), 8) & "' AND FA02='" & Right$(Trim(strArr(i)), 1) & "'"
'             Else
'                  strSQL = "SELECT FA05||' '||FA63||' '||FA64||' '||FA65,FA04,FA06,FA29 FROM FAGENT WHERE FA01='" & Left$(Trim(strArr(i)), 8) & "' AND FA02='0'"
'             End If
            PUB_GetAgentName Str01, Trim(strArr(i)), tmp02
            StrOk(26) = strArr(i) + "  " + tmp02
         End If
'edit by nickc 2007/08/27 秀玲說基本檔怎麼抓，就跟基本檔一樣，不用統一中英日或是英中日
'         CheckOC
'         adoRecordset.CursorLocation = adUseClient
'         adoRecordset.Open strSQL, cnnConnection, adOpenStatic, adLockReadOnly
'         If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
'            'Modify By Cheng 2002/07/08
''            If IsNull(adoRecordset.Fields(0)) Then
'            If IsNull(adoRecordset.Fields(IIf(Left$(strArr(i), 1) = "X", 0, IIf(strSK03 = "0", 1, 0)))) Then
'                  'Modify By Cheng 2002/07/08
''                If IsNull(adoRecordset.Fields(1)) Then
'                If IsNull(adoRecordset.Fields(IIf(Left$(strArr(i), 1) = "X", 0, IIf(strSK03 = "0", 0, 1)))) Then
'                    If IsNull(adoRecordset.Fields(2)) Then
'                        StrOk(26) = strArr(i) + ""
'                    Else
'                        StrOk(26) = strArr(i) + "  " + adoRecordset.Fields(2)
'                    End If
'                Else
'                     'Modify By Cheng 2002/07/08
''                    StrOk(26) = StrArr(i) + "  " + adoRecordset.Fields(1)
'                    StrOk(26) = strArr(i) + "  " + adoRecordset.Fields(IIf(Left$(strArr(i), 1) = "X", 0, IIf(strSK03 = "0", 0, 1)))
'                End If
'            Else
'               'Modify By Cheng 2002/07/08
''                StrOk(26) = StrArr(i) + "  " + adoRecordset.Fields(0)
'                StrOk(26) = strArr(i) + "  " + adoRecordset.Fields(IIf(Left$(strArr(i), 1) = "X", 0, IIf(strSK03 = "0", 1, 0)))
'            End If
        If StrOk(26) <> strArr(i) Then
            'Add by Morgan 2004/1/14
            Lbl1(26).ForeColor = vbBlack
         Else
            'Modify by Morgan 2004/1/14
            'StrOk(26) = ""
            Lbl1(37).ForeColor = vbBlack
            StrOk(26) = strArr(i)
         End If
         CheckOC
   Case 46
        StrOk(27) = strArr(i)
   Case 56 '固定請款對象
        If Left$(strArr(i), 1) = "X" Then
'edit by nickc 2007/08/27 秀玲說基本檔怎麼抓，就跟基本檔一樣，不用統一中英日或是英中日
'             If Len(strArr(i)) = 9 Then
'                  strSQL = "SELECT cu05||' '||cu88||' '||cu89||' '||cu90,CU04,CU06 FROM CUSTOMER WHERE CU01='" & Left$(strArr(i), 8) & "' AND CU02='" & Right$(strArr(i), 1) & "'"
'             Else
'                  strSQL = "SELECT cu05||' '||cu88||' '||cu89||' '||cu90,CU04,CU06 FROM CUSTOMER WHERE CU01='" & Left$(strArr(i), 8) & "' AND CU02='0'"
'             End If
            'Modified by Lydia 2020/12/31 + 客戶編號
            StrOk(34) = strArr(i) + "  " + GetAgentOrCustName(Trim(strArr(i)))
         Else
'edit by nickc 2007/08/27 秀玲說基本檔怎麼抓，就跟基本檔一樣，不用統一中英日或是英中日
'             If Len(strArr(i)) = 9 Then
'                  strSQL = "SELECT FA05||' '||FA63||' '||FA64||' '||FA65,FA04,FA06,FA29 FROM FAGENT WHERE FA01='" & Left$(Trim(strArr(i)), 8) & "' AND FA02='" & Right$(Trim(strArr(i)), 1) & "'"
'             Else
'                  strSQL = "SELECT FA05||' '||FA63||' '||FA64||' '||FA65,FA04,FA06,FA29 FROM FAGENT WHERE FA01='" & Left$(Trim(strArr(i)), 8) & "' AND FA02='0'"
'             End If
            PUB_GetAgentName Str01, Trim(strArr(i)), tmp02
            StrOk(34) = strArr(i) + "  " + tmp02
         End If
'edit by nickc 2007/08/27 秀玲說基本檔怎麼抓，就跟基本檔一樣，不用統一中英日或是英中日
'         CheckOC
'         adoRecordset.CursorLocation = adUseClient
'         adoRecordset.Open strSQL, cnnConnection, adOpenStatic, adLockReadOnly
'         If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
'            If IsNull(adoRecordset.Fields(IIf(Left$(strArr(i), 1) = "X", 0, IIf(strSK03 = "0", 1, 0)))) Then
'                If IsNull(adoRecordset.Fields(IIf(Left$(strArr(i), 1) = "X", 0, IIf(strSK03 = "0", 0, 1)))) Then
'                    If IsNull(adoRecordset.Fields(2)) Then
'                        StrOk(34) = strArr(i) + ""
'                    Else
'                        StrOk(34) = strArr(i) + "  " + adoRecordset.Fields(2)
'                    End If
'                Else
'                    StrOk(34) = strArr(i) + "  " + adoRecordset.Fields(IIf(Left$(strArr(i), 1) = "X", 0, IIf(strSK03 = "0", 0, 1)))
'                End If
'            Else
'                StrOk(34) = strArr(i) + "  " + adoRecordset.Fields(IIf(Left$(strArr(i), 1) = "X", 0, IIf(strSK03 = "0", 1, 0)))
'            End If
         If StrOk(34) <> strArr(i) Then
            'Add by Morgan 2004/1/14
            Lbl1(34).ForeColor = vbBlack
         Else
            'Add by Morgan 2004/1/14
            'StrOk(34) = ""
            Lbl1(34).ForeColor = vbBlack
            StrOk(34) = strArr(i)
         End If
         CheckOC
   Case 44
        If Len(strArr(i)) = 9 Then
              strSql = "SELECT FA05||' '||FA63||' '||FA64||' '||FA65,FA04,FA06,FA29,FA39 FROM FAGENT WHERE FA01='" & Left$(Trim(strArr(i)), 8) & "' AND FA02='" & Right$(Trim(strArr(i)), 1) & "'"
         Else
              strSql = "SELECT FA05||' '||FA63||' '||FA64||' '||FA65,FA04,FA06,FA29,FA39 FROM FAGENT WHERE FA01='" & Left$(Trim(strArr(i)), 8) & "' AND FA02='0'"
         End If
         CheckOC
        adoRecordset.CursorLocation = adUseClient
        adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
         If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
            'Modify By Cheng 2002/07/08
'            If Trim(adoRecordset.Fields(0)) = "" Then
            '2005/9/14 MODIFY BY SONIA
            'If IsNull(adoRecordset.Fields(IIf(strSK03 = "0", 1, 0))) Then
'edit by nickc 2007/08/27 秀玲說基本檔怎麼抓，就跟基本檔一樣，不用統一中英日或是英中日
'
'            If IsNull(Trim(adoRecordset.Fields(IIf(strSK03 = "0", 1, 0)))) Then
'             '2005/9/14 END
'              'Modify By Cheng 2002/07/08
''               If IsNull(adoRecordset.Fields(1)) Then
'               If IsNull(Trim(adoRecordset.Fields(IIf(strSK03 = "0", 0, 1)))) Then
'                   If IsNull(Trim(adoRecordset.Fields(2))) Then
'                          StrOk(36) = strArr(i) + ""
'                   Else
'                         StrOk(36) = strArr(i) + "  " + adoRecordset.Fields(2)
'                   End If
'               Else
'                  'Modify By Cheng 2002/07/08
''                   StrOk(36) = StrArr(i) + "  " + adoRecordset.Fields(1)
'                   StrOk(36) = strArr(i) + "  " + adoRecordset.Fields(IIf(strSK03 = "0", 0, 1))
'               End If
'            Else
'               'Modify By Cheng 2002/07/08
''               StrOk(36) = StrArr(i) + "  " + adoRecordset.Fields(0)
'               StrOk(36) = strArr(i) + "  " + adoRecordset.Fields(IIf(strSK03 = "0", 1, 0))
'
'            End If
            PUB_GetAgentName Str01, Trim(strArr(i)), tmp02
            StrOk(36) = strArr(i) + "  " + tmp02
            
            If IsNull(adoRecordset.Fields(3)) Then
                StrOkTxt(4) = ""
            Else
                StrOkTxt(4) = adoRecordset.Fields(3)
            End If
            If IsNull(adoRecordset.Fields(4)) Then
                 StrOk(29) = ""
            Else
                 StrOk(29) = adoRecordset.Fields(4)
            End If
            'Add by Morgan 2004/1/14
            Lbl1(36).ForeColor = vbBlack
         Else
            'Modify by Morgan 2004/1/14
            'StrOk(36) = ""
            Lbl1(36).ForeColor = vbRed
            StrOk(36) = strArr(i)
            
            StrOkTxt(4) = ""
            StrOk(29) = ""
         End If
         CheckOC
   Case 23
        If Len(strArr(i)) = 9 Then
              'edit by nickc 2007/08/27 秀玲說基本檔怎麼抓，就跟基本檔一樣，不用統一中英日或是英中日
              'strSQL = "SELECT CU04,cu05||' '||cu88||' '||cu89||' '||cu90,CU06,CU79,CU72 FROM CUSTOMER WHERE CU01='" & Left$(strArr(i), 8) & "' AND CU02='" & Right$(strArr(i), 1) & "'"
              strSql = "SELECT CU04,cu05,CU06,CU79,CU72 FROM CUSTOMER WHERE CU01='" & Left$(strArr(i), 8) & "' AND CU02='" & Right$(strArr(i), 1) & "'"
         Else
              'edit by nickc 2007/08/27 秀玲說基本檔怎麼抓，就跟基本檔一樣，不用統一中英日或是英中日
              'strSQL = "SELECT CU04,cu05||' '||cu88||' '||cu89||' '||cu90,CU06,CU79,CU72 FROM CUSTOMER WHERE CU01='" & Left$(strArr(i), 8) & "' AND CU02='0'"
              strSql = "SELECT CU04,cu05,CU06,CU79,CU72 FROM CUSTOMER WHERE CU01='" & Left$(strArr(i), 8) & "' AND CU02='0'"
         End If
         CheckOC
         adoRecordset.CursorLocation = adUseClient
         adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
         If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
'edit by nickc 2007/08/27 秀玲說基本檔怎麼抓，就跟基本檔一樣，不用統一中英日或是英中日
'             If IsNull(adoRecordset.Fields(0)) Then
'                  If IsNull(adoRecordset.Fields(1)) Then
'                     If IsNull(adoRecordset.Fields(2)) Then
'                          StrOk(35) = strArr(i) + ""
'                     Else
'                          StrOk(35) = strArr(i) + "  " + adoRecordset.Fields(2)
'                     End If
'                  Else
'                     StrOk(35) = strArr(i) + "  " + adoRecordset.Fields(1)
'                  End If
'             Else
'                  StrOk(35) = strArr(i) + "  " + adoRecordset.Fields(0)
'             End If
             If IsNull(adoRecordset.Fields("CU04")) = False Then
                StrOk(35) = strArr(i) + "  " + adoRecordset.Fields("CU04")
             ElseIf IsNull(adoRecordset.Fields("CU05")) = False Then
                StrOk(35) = strArr(i) + "  " + adoRecordset.Fields("CU05")
             ElseIf IsNull(adoRecordset.Fields("CU06")) = False Then
                StrOk(35) = strArr(i) + "  " + adoRecordset.Fields("CU06")
             End If
             
             If IsNull(adoRecordset.Fields(3)) Then
                  StrOkTxt(5) = ""
             Else
                  StrOkTxt(5) = adoRecordset.Fields(3)
             End If
             If IsNull(adoRecordset.Fields(4)) Then
                 StrOk(30) = ""
             Else
                 StrOk(30) = adoRecordset.Fields(4)
             End If
            'Add by Morgan 2004/1/14
            Lbl1(35).ForeColor = vbBlack
         Else
            'Modify by Morgan 2004/1/14
            Lbl1(35).ForeColor = vbRed
            'StrOk(35) = ""
             StrOk(35) = strArr(i)
             
             StrOkTxt(5) = ""
             StrOk(30) = ""
         End If
         CheckOC
   Case 58
         StrOkTxt(3) = strArr(i)
         'Add by Morgan 2003/12/01
         If strArr(i) <> "" Then
            If InStr(1, strArr(i), "原為聯合商標") > 0 Then
               StrOk(3) = "(原為聯合商標)"
            ElseIf InStr(1, strArr(i), "原為服務標章") > 0 Then
               StrOk(3) = "(原為服務標章)"
            ElseIf InStr(1, strArr(i), "原為聯合服務標章") > 0 Then
               StrOk(3) = "(原為聯合服務標章)"
            End If
         End If
         'End 2003/12/01
   Case 24
         StrOkTxt(6) = strArr(i)
   Case 25
         StrOkTxt(7) = strArr(i)
   Case 26
         StrOkTxt(8) = strArr(i)
   Case 45
         StrOk(39) = strArr(i)
   Case 47
         StrOkTxt(9) = strArr(i)
   Case 48
         StrOkTxt(10) = strArr(i)
   Case 49
         StrOkTxt(11) = strArr(i)
   Case 50
         StrOkTxt(12) = strArr(i)
   Case 51
         StrOkTxt(13) = strArr(i)
   Case 52
         StrOkTxt(14) = strArr(i)
   Case 33
        If Left$(strArr(i), 1) = "X" Then
'edit by nickc 2007/08/27 秀玲說基本檔怎麼抓，就跟基本檔一樣，不用統一中英日或是英中日
'             If Len(strArr(i)) = 9 Then
'                  strSQL = "SELECT cu05||' '||cu88||' '||cu89||' '||cu90,CU04,CU06 FROM CUSTOMER WHERE CU01='" & Left$(strArr(i), 8) & "' AND CU02='" & Right$(strArr(i), 1) & "'"
'             Else
'                  strSQL = "SELECT cu05||' '||cu88||' '||cu89||' '||cu90,CU04,CU06 FROM CUSTOMER WHERE CU01='" & Left$(strArr(i), 8) & "' AND CU02='0'"
'             End If
             StrOk(37) = strArr(i) + "  " + GetAgentOrCustName(Trim(strArr(i)))
         Else
'edit by nickc 2007/08/27 秀玲說基本檔怎麼抓，就跟基本檔一樣，不用統一中英日或是英中日
'             If Len(strArr(i)) = 9 Then
'                  strSQL = "SELECT FA05||' '||FA63||' '||FA64||' '||FA65,FA04,FA06,FA29 FROM FAGENT WHERE FA01='" & Left$(Trim(strArr(i)), 8) & "' AND FA02='" & Right$(Trim(strArr(i)), 1) & "'"
'             Else
'                  strSQL = "SELECT FA05||' '||FA63||' '||FA64||' '||FA65,FA04,FA06,FA29 FROM FAGENT WHERE FA01='" & Left$(Trim(strArr(i)), 8) & "' AND FA02='0'"
'             End If
             If PUB_GetAgentName(Str01, Trim(strArr(i)), tmp02) Then
                StrOk(37) = strArr(i) + "  " + tmp02
             Else
                StrOk(37) = strArr(i)
             End If
         End If
'edit by nickc 2007/08/27 秀玲說基本檔怎麼抓，就跟基本檔一樣，不用統一中英日或是英中日
'         CheckOC
'         adoRecordset.CursorLocation = adUseClient
'         adoRecordset.Open strSQL, cnnConnection, adOpenStatic, adLockReadOnly
'         If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
            'Modify By Cheng 2002/07/08
'            If IsNull(adoRecordset.Fields(0)) Then
'edit by nickc 2007/08/27 秀玲說基本檔怎麼抓，就跟基本檔一樣，不用統一中英日或是英中日
'            If IsNull(adoRecordset.Fields(IIf(Left$(strArr(i), 1) = "X", 0, IIf(strSK03 = "0", 1, 0)))) Then
'               'Modify By Cheng 2002/07/08
''                If IsNull(adoRecordset.Fields(1)) Then
'                If IsNull(adoRecordset.Fields(IIf(Left$(strArr(i), 1) = "X", 0, IIf(strSK03 = "0", 0, 1)))) Then
'                    If IsNull(adoRecordset.Fields(2)) Then
'                        StrOk(37) = strArr(i) + ""
'                    Else
'                        StrOk(37) = strArr(i) + "  " + adoRecordset.Fields(2)
'                    End If
'                Else
'                  'Modify By Cheng 2002/07/08
''                    StrOk(37) = StrArr(i) + "  " + adoRecordset.Fields(1)
'                    StrOk(37) = strArr(i) + "  " + adoRecordset.Fields(IIf(Left$(strArr(i), 1) = "X", 0, IIf(strSK03 = "0", 0, 1)))
'                End If
'            Else
'               'Modify By Cheng 2002/07/08
''                StrOk(37) = StrArr(i) + "  " + adoRecordset.Fields(0)
'                StrOk(37) = strArr(i) + "  " + adoRecordset.Fields(IIf(Left$(strArr(i), 1) = "X", 0, IIf(strSK03 = "0", 1, 0)))
'            End If
         If StrOk(37) <> strArr(i) Then
            'Add by Morgan 2004/1/14
            Lbl1(37).ForeColor = vbBlack
         Else
            'Modify by Morgan 2004/1/14
            'StrOk(37) = ""
            Lbl1(37).ForeColor = vbRed
            StrOk(37) = strArr(i)
         End If
         CheckOC
   Case 66
         If Left$(strArr(i), 1) = "X" Then
'edit by nickc 2007/08/27 秀玲說基本檔怎麼抓，就跟基本檔一樣，不用統一中英日或是英中日
'             If Len(strArr(i)) = 9 Then
'                  strSQL = "SELECT cu05||' '||cu88||' '||cu89||' '||cu90,CU04,CU06 FROM CUSTOMER WHERE CU01='" & Left$(strArr(i), 8) & "' AND CU02='" & Right$(strArr(i), 1) & "'"
'             Else
'                  strSQL = "SELECT cu05||' '||cu88||' '||cu89||' '||cu90,CU04,CU06 FROM CUSTOMER WHERE CU01='" & Left$(strArr(i), 8) & "' AND CU02='0'"
'             End If
             StrOk(38) = strArr(i) + "  " + GetAgentOrCustName(Trim(strArr(i)))
         Else
'edit by nickc 2007/08/27 秀玲說基本檔怎麼抓，就跟基本檔一樣，不用統一中英日或是英中日
'             If Len(strArr(i)) = 9 Then
'                  strSQL = "SELECT FA05||' '||FA63||' '||FA64||' '||FA65,FA04,FA06,FA29 FROM FAGENT WHERE FA01='" & Left$(Trim(strArr(i)), 8) & "' AND FA02='" & Right$(Trim(strArr(i)), 1) & "'"
'             Else
'                  strSQL = "SELECT FA05||' '||FA63||' '||FA64||' '||FA65,FA04,FA06,FA29 FROM FAGENT WHERE FA01='" & Left$(Trim(strArr(i)), 8) & "' AND FA02='0'"
'             End If
             If PUB_GetAgentName(Str01, Trim(strArr(i)), tmp02) Then
                 StrOk(38) = strArr(i) + "  " + tmp02
             Else
                StrOk(38) = strArr(i)
             End If
         End If
'edit by nickc 2007/08/27 秀玲說基本檔怎麼抓，就跟基本檔一樣，不用統一中英日或是英中日
'         CheckOC
'         adoRecordset.CursorLocation = adUseClient
'         adoRecordset.Open strSQL, cnnConnection, adOpenStatic, adLockReadOnly
'         If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
'            'Modify By Cheng 2002/07/08
''            If IsNull(adoRecordset.Fields(0)) Then
'            If IsNull(adoRecordset.Fields(IIf(Left$(strArr(i), 1) = "X", 0, IIf(strSK03 = "0", 1, 0)))) Then
'               'Modify By Cheng 2002/07/08
''                If IsNull(adoRecordset.Fields(1)) Then
'                If IsNull(adoRecordset.Fields(IIf(Left$(strArr(i), 1) = "X", 0, IIf(strSK03 = "0", 0, 1)))) Then
'                    If IsNull(adoRecordset.Fields(2)) Then
'                        StrOk(38) = strArr(i) + ""
'                    Else
'                        StrOk(38) = strArr(i) + "  " + adoRecordset.Fields(2)
'                    End If
'                Else
'                  'Modify By Cheng 2002/07/08
''                    StrOk(38) = StrArr(i) + "  " + adoRecordset.Fields(1)
'                    StrOk(38) = strArr(i) + "  " + adoRecordset.Fields(IIf(Left$(strArr(i), 1) = "X", 0, IIf(strSK03 = "0", 0, 1)))
'                End If
'            Else
'               'Modify By Cheng 2002/07/08
''                StrOk(38) = StrArr(i) + "  " + adoRecordset.Fields(0)
'                StrOk(38) = strArr(i) + "  " + adoRecordset.Fields(IIf(Left$(strArr(i), 1) = "X", 0, IIf(strSK03 = "0", 1, 0)))
'            End If
          If StrOk(38) <> strArr(i) Then
            'Add by Morgan 2004/1/14
            Lbl1(38).ForeColor = vbBlack
         Else
            'Modify by Morgan 2004/1/14
            'StrOk(38) = ""
            Lbl1(38).ForeColor = vbBlack
            StrOk(38) = strArr(i)
         End If
         CheckOC
   Case 65
         StrOk(40) = strArr(i)
   Case 38
         StrOkTxt(15) = strArr(i)
   Case 39
         StrOkTxt(16) = strArr(i)
   Case 40
         StrOkTxt(17) = strArr(i)
   Case 41
         StrOkTxt(18) = strArr(i)
   Case 42
         StrOkTxt(19) = strArr(i)
   Case 43
         StrOkTxt(20) = strArr(i)
   Case 59
         'edit by nick 2004/10/05
         'StrOk(41) = GetPrjSalesNM(strArr(i)) & " " & strArr(60) & " " & strArr(61)
         StrOk(41) = GetPrjSalesNM(strArr(i)) & " " & ChangeTStringToTDateString(ChangeWStringToTString(strArr(60))) & " " & Format(strArr(61), "##:##")
   Case 62
         'edit by nick 2004/10/05
         'StrOk(42) = GetPrjSalesNM(strArr(i)) & " " & strArr(63) & " " & strArr(64)
         StrOk(42) = GetPrjSalesNM(strArr(i)) & " " & ChangeTStringToTDateString(ChangeWStringToTString(strArr(63))) & " " & Format(strArr(64), "##:##")
   Case 67
         StrOk(43) = strArr(i)
   Case 69 'D/N固定列印對象
        If Left$(strArr(i), 1) = "X" Then
'edit by nickc 2007/08/27 秀玲說基本檔怎麼抓，就跟基本檔一樣，不用統一中英日或是英中日
'             If Len(strArr(i)) = 9 Then
'                  strSQL = "SELECT cu05||' '||cu88||' '||cu89||' '||cu90,CU04,CU06 FROM CUSTOMER WHERE CU01='" & Left$(strArr(i), 8) & "' AND CU02='" & Right$(strArr(i), 1) & "'"
'             Else
'                  strSQL = "SELECT cu05||' '||cu88||' '||cu89||' '||cu90,CU04,CU06 FROM CUSTOMER WHERE CU01='" & Left$(strArr(i), 8) & "' AND CU02='0'"
'             End If
             StrOk(45) = strArr(i) + "  " + GetAgentOrCustName(Trim(strArr(i)))
         Else
'edit by nickc 2007/08/27 秀玲說基本檔怎麼抓，就跟基本檔一樣，不用統一中英日或是英中日
'             If Len(strArr(i)) = 9 Then
'                  strSQL = "SELECT FA05||' '||FA63||' '||FA64||' '||FA65,FA04,FA06,FA29 FROM FAGENT WHERE FA01='" & Left$(Trim(strArr(i)), 8) & "' AND FA02='" & Right$(Trim(strArr(i)), 1) & "'"
'             Else
'                  strSQL = "SELECT FA05||' '||FA63||' '||FA64||' '||FA65,FA04,FA06,FA29 FROM FAGENT WHERE FA01='" & Left$(Trim(strArr(i)), 8) & "' AND FA02='0'"
'             End If
             If PUB_GetAgentName(Str01, Trim(strArr(i)), tmp02) Then
                StrOk(45) = strArr(i) + "  " + tmp02
             Else
                StrOk(45) = strArr(i)
             End If
         End If
'edit by nickc 2007/08/27 秀玲說基本檔怎麼抓，就跟基本檔一樣，不用統一中英日或是英中日
'         CheckOC
'         adoRecordset.CursorLocation = adUseClient
'         adoRecordset.Open strSQL, cnnConnection, adOpenStatic, adLockReadOnly
'         If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
'            If IsNull(adoRecordset.Fields(IIf(Left$(strArr(i), 1) = "X", 0, IIf(strSK03 = "0", 1, 0)))) Then
'                If IsNull(adoRecordset.Fields(IIf(Left$(strArr(i), 1) = "X", 0, IIf(strSK03 = "0", 0, 1)))) Then
'                    If IsNull(adoRecordset.Fields(2)) Then
'                        StrOk(45) = strArr(i) + ""
'                    Else
'                        StrOk(45) = strArr(i) + "  " + adoRecordset.Fields(2)
'                    End If
'                Else
'                    StrOk(45) = strArr(i) + "  " + adoRecordset.Fields(IIf(Left$(strArr(i), 1) = "X", 0, IIf(strSK03 = "0", 0, 1)))
'                End If
'            Else
'                StrOk(45) = strArr(i) + "  " + adoRecordset.Fields(IIf(Left$(strArr(i), 1) = "X", 0, IIf(strSK03 = "0", 1, 0)))
'            End If
         If StrOk(45) <> strArr(i) Then
            'Add by Morgan 2004/1/14
            Lbl1(45).ForeColor = vbBlack
         Else
            'Modify by Morgan 2004/1/14
            'StrOk(45) = ""
            Lbl1(45).ForeColor = vbBlack
            StrOk(45) = strArr(i)
         End If
         CheckOC
   Case 70 '延展D/N列印對象
        If Left$(strArr(i), 1) = "X" Then
'edit by nickc 2007/08/27 秀玲說基本檔怎麼抓，就跟基本檔一樣，不用統一中英日或是英中日
'             If Len(strArr(i)) = 9 Then
'                  strSQL = "SELECT cu05||' '||cu88||' '||cu89||' '||cu90,CU04,CU06 FROM CUSTOMER WHERE CU01='" & Left$(strArr(i), 8) & "' AND CU02='" & Right$(strArr(i), 1) & "'"
'             Else
'                  strSQL = "SELECT cu05||' '||cu88||' '||cu89||' '||cu90,CU04,CU06 FROM CUSTOMER WHERE CU01='" & Left$(strArr(i), 8) & "' AND CU02='0'"
'             End If
             StrOk(46) = strArr(i) + "  " + GetAgentOrCustName(Trim(strArr(i)))
         Else
'edit by nickc 2007/08/27 秀玲說基本檔怎麼抓，就跟基本檔一樣，不用統一中英日或是英中日
'             If Len(strArr(i)) = 9 Then
'                  strSQL = "SELECT FA05||' '||FA63||' '||FA64||' '||FA65,FA04,FA06,FA29 FROM FAGENT WHERE FA01='" & Left$(Trim(strArr(i)), 8) & "' AND FA02='" & Right$(Trim(strArr(i)), 1) & "'"
'             Else
'                  strSQL = "SELECT FA05||' '||FA63||' '||FA64||' '||FA65,FA04,FA06,FA29 FROM FAGENT WHERE FA01='" & Left$(Trim(strArr(i)), 8) & "' AND FA02='0'"
'             End If
             If PUB_GetAgentName(Str01, Trim(strArr(i)), tmp02) Then
                StrOk(46) = strArr(i) + "  " + tmp02
             Else
                StrOk(46) = strArr(i)
             End If
         End If
'edit by nickc 2007/08/27 秀玲說基本檔怎麼抓，就跟基本檔一樣，不用統一中英日或是英中日
'         CheckOC
'         adoRecordset.CursorLocation = adUseClient
'         adoRecordset.Open strSQL, cnnConnection, adOpenStatic, adLockReadOnly
'         If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
'            If IsNull(adoRecordset.Fields(IIf(Left$(strArr(i), 1) = "X", 0, IIf(strSK03 = "0", 1, 0)))) Then
'                If IsNull(adoRecordset.Fields(IIf(Left$(strArr(i), 1) = "X", 0, IIf(strSK03 = "0", 0, 1)))) Then
'                    If IsNull(adoRecordset.Fields(2)) Then
'                        StrOk(46) = strArr(i) + ""
'                    Else
'                        StrOk(46) = strArr(i) + "  " + adoRecordset.Fields(2)
'                    End If
'                Else
'                    StrOk(46) = strArr(i) + "  " + adoRecordset.Fields(IIf(Left$(strArr(i), 1) = "X", 0, IIf(strSK03 = "0", 0, 1)))
'                End If
'            Else
'                StrOk(46) = strArr(i) + "  " + adoRecordset.Fields(IIf(Left$(strArr(i), 1) = "X", 0, IIf(strSK03 = "0", 1, 0)))
'            End If
         If StrOk(46) <> strArr(i) Then
            'Add by Morgan 2004/1/14
            Lbl1(46).ForeColor = vbBlack
         Else
            'Modify by Morgan 2004/1/14
            'StrOk(46) = ""
            Lbl1(46).ForeColor = vbBlack
            StrOk(46) = strArr(i)
         End If
         CheckOC
   Case 71 '延展聯絡人
        StrOk(47) = strArr(i)
   Case 72 '特殊商標
        If strArr(i) <> "" Then
            StrOk(48) = strArr(i) & "  " & PUB_GetSpecialPTName("2", strArr(i))
        Else
            StrOk(48) = ""
        End If
    'add by nickc 2006/07/12
    Case 73
         If Len(Trim(strArr(i))) = 0 Or strArr(i) = "0" Then
             Lbl1(5) = ""
         Else
             Lbl1(5) = ChangeTStringToTDateString(ChangeWStringToTString(strArr(i)))
         End If
    Case 74
         strSql = "SELECT nvl(ST02,'" & strArr(i) & "') FROM STAFF WHERE ST01='" & strArr(i) & "'"
         CheckOC
         adoRecordset.CursorLocation = adUseClient
         adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
         If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
            If IsNull(adoRecordset.Fields(0)) Then
               Lbl1(49) = strArr(i) + ""
            Else
               Lbl1(49) = strArr(i) + "  " + adoRecordset.Fields(0)
            End If
         Else
            Lbl1(49) = ""
         End If
         CheckOC
    Case 75
         Lbl1(50) = strArr(i)
    'add by nickc 2006/12/07
    Case 78
         If Len(strArr(i)) = 9 Then
              'edit by nickc 2007/08/27 秀玲說基本檔怎麼抓，就跟基本檔一樣，不用統一中英日或是英中日
              'strSQL = "SELECT CU04,cu05||' '||cu88||' '||cu89||' '||cu90,CU06,CU79 FROM CUSTOMER WHERE CU01='" & Left$(strArr(i), 8) & "' AND CU02='" & Right$(strArr(i), 1) & "'"
              strSql = "SELECT CU04,cu05,CU06,CU79 FROM CUSTOMER WHERE CU01='" & Left$(strArr(i), 8) & "' AND CU02='" & Right$(strArr(i), 1) & "'"
         Else
              'edit by nickc 2007/08/27 秀玲說基本檔怎麼抓，就跟基本檔一樣，不用統一中英日或是英中日
              'strSQL = "SELECT CU04,cu05||' '||cu88||' '||cu89||' '||cu90,CU06,CU79 FROM CUSTOMER WHERE CU01='" & Left$(strArr(i), 8) & "' AND CU02='0'"
              strSql = "SELECT CU04,cu05,CU06,CU79 FROM CUSTOMER WHERE CU01='" & Left$(strArr(i), 8) & "' AND CU02='0'"
         End If
         CheckOC
         adoRecordset.CursorLocation = adUseClient
         adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
         If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
'edit by nickc 2007/08/27 秀玲說基本檔怎麼抓，就跟基本檔一樣，不用統一中英日或是英中日
'             If IsNull(adoRecordset.Fields(0)) Then
'                  If IsNull(adoRecordset.Fields(1)) Then
'                     If IsNull(adoRecordset.Fields(2)) Then
'                          lbl1(51).Caption = strArr(i) + ""
'                     Else
'                          lbl1(51).Caption = strArr(i) + "  " + adoRecordset.Fields(2)
'                     End If
'                  Else
'                     lbl1(51).Caption = strArr(i) + "  " + adoRecordset.Fields(1)
'                  End If
'             Else
'                  lbl1(51).Caption = strArr(i) + "  " + adoRecordset.Fields(0)
'             End If
            If IsNull(adoRecordset.Fields("CU04")) = False Then
               Lbl1(51).Caption = strArr(i) + "  " + adoRecordset.Fields("CU04")
            ElseIf IsNull(adoRecordset.Fields("CU05")) = False Then
               Lbl1(51).Caption = strArr(i) + "  " + adoRecordset.Fields("CU05")
            ElseIf IsNull(adoRecordset.Fields("CU06")) = False Then
               Lbl1(51).Caption = strArr(i) + "  " + adoRecordset.Fields("CU06")
            End If
            Lbl1(51).ForeColor = vbBlack
         Else
            Lbl1(51).ForeColor = vbRed
            Lbl1(51).Caption = strArr(i)
         End If
         CheckOC
    Case 79
         If Len(strArr(i)) = 9 Then
              'edit by nickc 2007/08/27 秀玲說基本檔怎麼抓，就跟基本檔一樣，不用統一中英日或是英中日
              'strSQL = "SELECT CU04,cu05||' '||cu88||' '||cu89||' '||cu90,CU06,CU79 FROM CUSTOMER WHERE CU01='" & Left$(strArr(i), 8) & "' AND CU02='" & Right$(strArr(i), 1) & "'"
              strSql = "SELECT CU04,cu05,CU06,CU79 FROM CUSTOMER WHERE CU01='" & Left$(strArr(i), 8) & "' AND CU02='" & Right$(strArr(i), 1) & "'"
         Else
              'edit by nickc 2007/08/27 秀玲說基本檔怎麼抓，就跟基本檔一樣，不用統一中英日或是英中日
              'strSQL = "SELECT CU04,cu05||' '||cu88||' '||cu89||' '||cu90,CU06,CU79 FROM CUSTOMER WHERE CU01='" & Left$(strArr(i), 8) & "' AND CU02='0'"
              strSql = "SELECT CU04,cu05,CU06,CU79 FROM CUSTOMER WHERE CU01='" & Left$(strArr(i), 8) & "' AND CU02='0'"
         End If
         CheckOC
         adoRecordset.CursorLocation = adUseClient
         adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
         If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
'edit by nickc 2007/08/27 秀玲說基本檔怎麼抓，就跟基本檔一樣，不用統一中英日或是英中日
'             If IsNull(adoRecordset.Fields(0)) Then
'                  If IsNull(adoRecordset.Fields(1)) Then
'                     If IsNull(adoRecordset.Fields(2)) Then
'                          lbl1(52).Caption = strArr(i) + ""
'                     Else
'                          lbl1(52).Caption = strArr(i) + "  " + adoRecordset.Fields(2)
'                     End If
'                  Else
'                     lbl1(52).Caption = strArr(i) + "  " + adoRecordset.Fields(1)
'                  End If
'             Else
'                  lbl1(52).Caption = strArr(i) + "  " + adoRecordset.Fields(0)
'             End If
            If IsNull(adoRecordset.Fields("CU04")) = False Then
               Lbl1(52).Caption = strArr(i) + "  " + adoRecordset.Fields("CU04")
            ElseIf IsNull(adoRecordset.Fields("CU05")) = False Then
               Lbl1(52).Caption = strArr(i) + "  " + adoRecordset.Fields("CU05")
            ElseIf IsNull(adoRecordset.Fields("CU06")) = False Then
               Lbl1(52).Caption = strArr(i) + "  " + adoRecordset.Fields("CU06")
            End If
            Lbl1(52).ForeColor = vbBlack
         Else
            Lbl1(52).ForeColor = vbRed
            Lbl1(52).Caption = strArr(i)
         End If
         CheckOC
    Case 80
         If Len(strArr(i)) = 9 Then
              'edit by nickc 2007/08/27 秀玲說基本檔怎麼抓，就跟基本檔一樣，不用統一中英日或是英中日
              'strSQL = "SELECT CU04,cu05||' '||cu88||' '||cu89||' '||cu90,CU06,CU79 FROM CUSTOMER WHERE CU01='" & Left$(strArr(i), 8) & "' AND CU02='" & Right$(strArr(i), 1) & "'"
              strSql = "SELECT CU04,cu05,CU06,CU79 FROM CUSTOMER WHERE CU01='" & Left$(strArr(i), 8) & "' AND CU02='" & Right$(strArr(i), 1) & "'"
         Else
              'edit by nickc 2007/08/27 秀玲說基本檔怎麼抓，就跟基本檔一樣，不用統一中英日或是英中日
              'strSQL = "SELECT CU04,cu05||' '||cu88||' '||cu89||' '||cu90,CU06,CU79 FROM CUSTOMER WHERE CU01='" & Left$(strArr(i), 8) & "' AND CU02='0'"
              strSql = "SELECT CU04,cu05,CU06,CU79 FROM CUSTOMER WHERE CU01='" & Left$(strArr(i), 8) & "' AND CU02='0'"
         End If
         CheckOC
         adoRecordset.CursorLocation = adUseClient
         adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
         If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
'edit by nickc 2007/08/27 秀玲說基本檔怎麼抓，就跟基本檔一樣，不用統一中英日或是英中日
'             If IsNull(adoRecordset.Fields(0)) Then
'                  If IsNull(adoRecordset.Fields(1)) Then
'                     If IsNull(adoRecordset.Fields(2)) Then
'                          lbl1(53).Caption = strArr(i) + ""
'                     Else
'                          lbl1(53).Caption = strArr(i) + "  " + adoRecordset.Fields(2)
'                     End If
'                  Else
'                     lbl1(53).Caption = strArr(i) + "  " + adoRecordset.Fields(1)
'                  End If
'             Else
'                  lbl1(53).Caption = strArr(i) + "  " + adoRecordset.Fields(0)
'             End If
            If IsNull(adoRecordset.Fields("CU04")) = False Then
               Lbl1(53).Caption = strArr(i) + "  " + adoRecordset.Fields("CU04")
            ElseIf IsNull(adoRecordset.Fields("CU05")) = False Then
               Lbl1(53).Caption = strArr(i) + "  " + adoRecordset.Fields("CU05")
            ElseIf IsNull(adoRecordset.Fields("CU06")) = False Then
               Lbl1(53).Caption = strArr(i) + "  " + adoRecordset.Fields("CU06")
            End If
            Lbl1(53).ForeColor = vbBlack
         Else
            Lbl1(53).ForeColor = vbRed
            Lbl1(53).Caption = strArr(i)
         End If
         CheckOC
    Case 81
         If Len(strArr(i)) = 9 Then
              'edit by nickc 2007/08/27 秀玲說基本檔怎麼抓，就跟基本檔一樣，不用統一中英日或是英中日
              'strSQL = "SELECT CU04,cu05||' '||cu88||' '||cu89||' '||cu90,CU06,CU79 FROM CUSTOMER WHERE CU01='" & Left$(strArr(i), 8) & "' AND CU02='" & Right$(strArr(i), 1) & "'"
              strSql = "SELECT CU04,cu05,CU06,CU79 FROM CUSTOMER WHERE CU01='" & Left$(strArr(i), 8) & "' AND CU02='" & Right$(strArr(i), 1) & "'"
         Else
              'edit by nickc 2007/08/27 秀玲說基本檔怎麼抓，就跟基本檔一樣，不用統一中英日或是英中日
              'strSQL = "SELECT CU04,cu05||' '||cu88||' '||cu89||' '||cu90,CU06,CU79 FROM CUSTOMER WHERE CU01='" & Left$(strArr(i), 8) & "' AND CU02='0'"
              strSql = "SELECT CU04,cu05,CU06,CU79 FROM CUSTOMER WHERE CU01='" & Left$(strArr(i), 8) & "' AND CU02='0'"
         End If
         CheckOC
         adoRecordset.CursorLocation = adUseClient
         adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
         If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
'edit by nickc 2007/08/27 秀玲說基本檔怎麼抓，就跟基本檔一樣，不用統一中英日或是英中日
'             If IsNull(adoRecordset.Fields(0)) Then
'                  If IsNull(adoRecordset.Fields(1)) Then
'                     If IsNull(adoRecordset.Fields(2)) Then
'                          lbl1(54).Caption = strArr(i) + ""
'                     Else
'                          lbl1(54).Caption = strArr(i) + "  " + adoRecordset.Fields(2)
'                     End If
'                  Else
'                     lbl1(54).Caption = strArr(i) + "  " + adoRecordset.Fields(1)
'                  End If
'             Else
'                  lbl1(54).Caption = strArr(i) + "  " + adoRecordset.Fields(0)
'             End If
            If IsNull(adoRecordset.Fields("CU04")) = False Then
               Lbl1(54).Caption = strArr(i) + "  " + adoRecordset.Fields("CU04")
            ElseIf IsNull(adoRecordset.Fields("CU05")) = False Then
               Lbl1(54).Caption = strArr(i) + "  " + adoRecordset.Fields("CU05")
            ElseIf IsNull(adoRecordset.Fields("CU06")) = False Then
               Lbl1(54).Caption = strArr(i) + "  " + adoRecordset.Fields("CU06")
            End If
            Lbl1(54).ForeColor = vbBlack
         Else
            Lbl1(54).ForeColor = vbRed
            Lbl1(54).Caption = strArr(i)
         End If
         CheckOC
    Case 82
         txt1(24) = strArr(i)
    Case 83
         txt1(27) = strArr(i)
    Case 84
         txt1(30) = strArr(i)
    Case 85
         txt1(33) = strArr(i)
    Case 86
         txt1(23) = strArr(i)
    Case 87
         txt1(26) = strArr(i)
    Case 88
         txt1(29) = strArr(i)
    Case 89
         txt1(32) = strArr(i)
    Case 90
         txt1(22) = strArr(i)
    Case 91
         txt1(25) = strArr(i)
    Case 92
         txt1(28) = strArr(i)
    Case 93
         txt1(31) = strArr(i)
    Case 94
         txt1(34) = strArr(i)
    Case 95
         txt1(35) = strArr(i)
    Case 96
         txt1(36) = strArr(i)
    Case 97
         txt1(37) = strArr(i)
    Case 98
         txt1(38) = strArr(i)
    Case 99
         txt1(39) = strArr(i)
    Case 100
         txt1(40) = strArr(i)
    Case 101
         txt1(41) = strArr(i)
    Case 102
         txt1(42) = strArr(i)
    Case 103
         txt1(43) = strArr(i)
    Case 104
         txt1(44) = strArr(i)
    Case 105
         txt1(45) = strArr(i)
    Case 106
         txt1(46) = strArr(i)
    Case 107
         txt1(47) = strArr(i)
    Case 108
         txt1(48) = strArr(i)
    Case 109
         txt1(49) = strArr(i)
    Case 110
         txt1(50) = strArr(i)
    Case 111
         txt1(51) = strArr(i)
    Case 112
         txt1(52) = strArr(i)
    Case 113
         txt1(53) = strArr(i)
    Case 114
         txt1(54) = strArr(i)
    Case 115
         txt1(55) = strArr(i)
    Case 116
         txt1(56) = strArr(i)
    Case 117
         txt1(57) = strArr(i)
    Case 76
         txt1(58) = strArr(i)
     'add by nickc 2007/04/20
     Case 118
        Lbl1(55) = strArr(i)
    'Add by Morgan 2008/5/26
    Case 121
        Lbl1(56) = strArr(i)
    'Add by Morgan 2008/8/4
    Case 123
        Lbl1(57) = PUB_GetContact(strArr(23), strArr(i))
    'add by Toni 2008/10/21
    Case 122
      Lbl1(58) = strArr(i)
      'end 2008/10/21
    'Add By Sindy 2009/09/09
    Case 77
      Lbl1(59) = strArr(i)
    Case 124
      Lbl1(60) = strArr(i)
    Case 125
      Lbl1(61) = strArr(i)
    Case 126
      Lbl1(62) = strArr(i)
    '2009/09/09 End
    Case 127 'Add by Morgan 2010/11/5
      Lbl1(127) = strArr(i)
    Case 128 'Add by Sindy 2012/2/8
      Lbl1(128) = strArr(i)
    Case 129 'Add by Sindy 2013/8/26
      Lbl1(63) = strArr(i)
    Case 130 'Add by Sindy 2013/12/13
      Lbl1(64) = strArr(i)
    Case 131 'Add by Sindy 2015/6/30
      Lbl1(65) = strArr(i)
    'Add by Sindy 2024/6/14
    Case 137
      txt1(63) = strArr(i)
    Case 138
      txt1(64) = strArr(i)
    Case 139
      txt1(65) = strArr(i)
    '2024/6/14 END
    'Added by Morgan 2016/12/8
    Case 132 '國內副本收件人
      Lbl1(66) = strArr(i)
      If strArr(i) <> "" Then
         If ClsLawLawGetName(strArr(i), strExc(9)) = True Then
            Lbl1(66) = Lbl1(66) + "  " + strExc(9)
         End If
      End If
    Case 133 '國內副本接洽人
      If strArr(132) <> "" And strArr(i) <> "" Then
         Lbl1(67) = PUB_GetContact(strArr(132), strArr(i))
      Else
         Lbl1(67) = ""
      End If
    'end 2016/12/8
    'Add By Sindy 2016/11/24
    Case 134
      Lbl1(134) = strArr(i)
    Case 135
      Combo3(1).ListIndex = Val(strArr(i))
    '2016/11/24 END
    'Added by Morgan 2022/12/1
    Case 136
      Lbl1(136) = strArr(i)
    Case Else
    End Select
    'DoEvents Add By Sindy 2019/1/4 Mark,因為會和視窗的function(MenuForFormControl)有ErrCode互影響
Next i
Call SetTM72forCol(strArr(72)) 'Add By Sindy 2025/7/31

'Modify By Cheng 2002/12/12
'For i = 0 To 43
For i = 0 To UBound(StrOk)                             '2006/07/12 加備註，以後新增欄位，直接在上面修改，此2段迴圈
   'Modify by Morgan 2003/12/01                        '不可修改，不然會影響資料顯現，而且陣列的宣告也不用一直的修改
    'If i <> 3 And i <> 4 And i <> 5 Then
    'edit by nick 2004/11/24
    'If i <> 4 And i <> 5 Then
    'Modify by Amy 2022/09/02 審定號 strOk(18) 原label 改textbox
    If i <> 5 And i <> 17 And i <> 18 Then
   'End 2003/12/01
        Lbl1(i) = StrOk(i)
    End If
Next i
For i = 0 To UBound(StrOkTxt)
    txt1(i) = StrOkTxt(i)
Next i '傳參數　　　代理人
StrTag = strArr(44)
'傳參數　　　申請人
StrTag1 = strArr(23)
'Add By Sindy 2010/02/04
cmdOK(8).Visible = False
cmdOK(9).Visible = False
cmdOK(10).Visible = False
cmdOK(11).Visible = False
StrTag2 = strArr(78)
StrTag3 = strArr(79)
StrTag4 = strArr(80)
StrTag5 = strArr(81)
If Trim(StrTag2) <> "" Then cmdOK(8).Visible = True
If Trim(StrTag3) <> "" Then cmdOK(9).Visible = True
If Trim(StrTag4) <> "" Then cmdOK(10).Visible = True
If Trim(StrTag5) <> "" Then cmdOK(11).Visible = True
'2010/02/04 End
'add by nickc 2005/05/30  檢查有無分割或相關卷號
cmdOK(5).Visible = ChkDataBy308(txt1(59).Text)
cmdOK(4).Visible = ChkDataByCR(txt1(59).Text)

'Added by Lydia 2025/10/23 TF基礎案號(TM06,TM07)改成可以輸入多筆(Table: TFBaseNo)，放在銷卷頁籤。
lblTFBase.Visible = False
MGrid1.Visible = False
SSTab1.TabCaption(7) = "銷卷資料"
If strArr(1) = "TF" And Mid(strArr(2), 6, 1) = "0" And strArr(3) = "0" And strArr(4) = "00" Then
   SSTab1.TabCaption(7) = "銷卷/TF基礎案號數"
   Call SetGrid(True)
   lblTFBase.Visible = True
   MGrid1.Visible = True
   strR1 = "SELECT  tfbn05,tfbn06 ,na03 as tfbn06n,decode(t1.tm01||t2.tm01,null,'非本所案件',decode(t1.tm01,null,decode(t2.tm28,'1',null,'N')||rtrim(t2.tm01||'-'||t2.tm02||'-'||t2.tm03||'-'||t2.tm04),decode(t1.tm28,'1',null,'N')||rtrim(t1.tm01||'-'||t1.tm02||'-'||t1.tm03||'-'||t1.tm04)) ) as tcase " & _
           " FROM tfbaseno,nation,trademark t1, trademark t2" & _
           " WHERE tfbn01='" & strArr(1) & "' and tfbn02='" & strArr(2) & "' and tfbn03='" & strArr(3) & "' and tfbn04='" & strArr(4) & "' " & _
           " and tfbn06=na01(+) AND tfbn05=t1.tm15(+) AND tfbn06=t1.tm10(+) AND tfbn05=t2.tm12(+) AND tfbn06=t2.tm10(+) order by tfbn08,tfbn09 "
   intR = 1
   Set rsRD = ClsLawReadRstMsg(intR, strR1)
   If intR = 1 Then
       Set MGrid1.Recordset = rsRD
       Call SetGrid(False)
   End If
End If
Set rsRD = Nothing
'end 2025/10/23

'ErrorHandler:
'   If Err.Number <> 0 Then
'      MsgBox "(" & Err.Number & ")" & Err.Description
'   End If
'   Me.Enabled = True
'   Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Activate()
   pub_QL05 = m_pub_QL05 'Add By Sindy 2025/8/28 還原此Form的查詢條件記錄
End Sub

'Private Sub cmdRef_Click()
''    Dim stTmp As String
''    stTmp = Right(Space(2) & txt1(59), 15)
''    Where1103ComeFrom Me, Trim(Left(stTmp, 3)), Mid(stTmp, 5, 6), Mid(stTmp, 12, 1), Mid(stTmp, 14, 2)
'End Sub

Private Sub Form_Load()
Dim Lbl As Object

For Each Lbl In Me.Lbl1
    Lbl.BackColor = &H8000000F
Next
'Added by Lydia 2025/09/12
'Modified by Lydia 2025/10/23
lblTFBase.Visible = False
MGrid1.Visible = False
'end 2025/09/12

bolToEndByNick = False
MoveFormToCenter Me
SSTab1.Tab = 0 'Add by Amy 2014/04/10

GRIDHEAND 'Added by Lydia 2016/10/19

If bolFNation = False Then
    Label29.Visible = False
    Lbl1(36).Visible = False
    cmdOK(0).Value = False
End If
'92.04.16 nick
cmdState = -1

'Added by Lydia 2020/05/05 各項指示：顯示按鈕
If strSrvDate(1) >= 各項指示啟用日 Then
   cmdOK(12).Visible = True
Else
   cmdOK(12).Visible = False
End If
'end 2020/05/05
'Memo by Amy 2024/03/08 隱藏延展單筆不跑
End Sub

Private Sub Form_Unload(Cancel As Integer)
pub_QL05 = m_pub_QL05 'Add By Sindy 2025/9/12 還原此Form的查詢條件記錄 (多筆查詢有影響)
Set frm100101_4 = Nothing
End Sub

'copy from frm020501 by nickc 2007/08/27
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

'Added by Lydia 2016/10/19
Private Function GRIDHEAND()
    With grdDataList2
    .row = 0
    .col = 0
    .ColWidth(0) = 800
    .Text = "優先權日"
    .col = 1
    .ColWidth(1) = 2500
    .Text = "優先權號"
    .col = 2
    .ColWidth(2) = 1000
    .Text = "優先權國家"
    .col = 3
    .ColWidth(3) = 1300
    .Text = "優先權存取碼"
    .col = 4
    .ColWidth(4) = 1300
    .Text = "本所案號"
    'Add By Sindy 2017/9/29
    .col = 5
    .ColWidth(5) = 2000
    .Text = "商品類別"
    '2017/9/29 END
    End With
End Function

'Added by Lydia 2016/10/26 修正Win7 輸入法問題
Private Sub txt1_GotFocus(Index As Integer)
   TextInverse txt1(Index) 'Added by Lydia 2016/12/6
   OpenIme
End Sub

'Add By Sindy 2025/7/31
Private Sub SetTM72forCol(strTM72)
   If Trim(strTM72) = "" Then
      Label119.Visible = False
      txt1(63).Visible = False
      Label120.Visible = False
      txt1(64).Visible = False
      Label121.Visible = False
      txt1(65).Visible = False
   Else
      Label119.Visible = True
      txt1(63).Visible = True
      Label120.Visible = True
      txt1(64).Visible = True
      Label121.Visible = True
      txt1(65).Visible = True
   End If
End Sub

'Added by Lydia 2025/10/23
Private Sub SetGrid(Optional ByVal pReset As Boolean = False)
Dim arrGridHeadText, arrGridHeadWidth
Dim iRow As Integer
   
   '                        0             1           2         3
   arrGridHeadText = Array("TF基礎案號", "TFBN06", "申請國家", "本所案號")
   arrGridHeadWidth = Array(1800, 0, 1000, 1600)
   MGrid1.Visible = False
   MGrid1.Cols = UBound(arrGridHeadText) + 1
   If pReset = True Then
       MGrid1.Clear
       MGrid1.Rows = 2
   End If
       
   For iRow = 0 To MGrid1.Cols - 1
       MGrid1.row = 0
       MGrid1.col = iRow
       MGrid1.Text = arrGridHeadText(iRow)
       MGrid1.CellAlignment = flexAlignCenterCenter
       MGrid1.ColWidth(iRow) = arrGridHeadWidth(iRow)
   Next

   MGrid1.Visible = True
End Sub
