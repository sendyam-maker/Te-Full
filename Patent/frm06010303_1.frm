VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm06010303_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "各式申請書-變更"
   ClientHeight    =   7420
   ClientLeft      =   170
   ClientTop       =   840
   ClientWidth     =   9440
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7420
   ScaleWidth      =   9440
   Begin VB.Frame FraPA174 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   0  '沒有框線
      Caption         =   "Frame1"
      Height          =   525
      Left            =   8250
      TabIndex        =   189
      Top             =   1320
      Visible         =   0   'False
      Width           =   825
      Begin VB.CommandButton CmdPA174 
         BackColor       =   &H00C0FFFF&
         Caption         =   "特殊字"
         Height          =   280
         Left            =   0
         Style           =   1  '圖片外觀
         TabIndex        =   190
         Top             =   210
         Width           =   800
      End
      Begin VB.Label lblPA174 
         Caption         =   "有特殊字"
         ForeColor       =   &H000000FF&
         Height          =   225
         Left            =   35
         TabIndex        =   191
         Top             =   0
         Width           =   765
      End
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   1080
      MaxLength       =   3
      TabIndex        =   109
      Top             =   105
      Width           =   550
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   1632
      MaxLength       =   6
      TabIndex        =   108
      Top             =   105
      Width           =   855
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   2472
      MaxLength       =   1
      TabIndex        =   107
      Top             =   105
      Width           =   255
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   2712
      MaxLength       =   2
      TabIndex        =   106
      Top             =   105
      Width           =   375
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   300
      Index           =   0
      Left            =   7050
      TabIndex        =   99
      Top             =   30
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "回前畫面(&U)"
      CausesValidation=   0   'False
      Height          =   300
      Index           =   1
      Left            =   7920
      TabIndex        =   100
      Top             =   30
      Width           =   1200
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5745
      Left            =   30
      TabIndex        =   101
      Top             =   1620
      Width           =   9375
      _ExtentX        =   16528
      _ExtentY        =   10142
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      TabCaption(0)   =   "申請人/代表人"
      TabPicture(0)   =   "frm06010303_1.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label4(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label23(0)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label23(1)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label23(2)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label23(3)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label23(4)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label23(5)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label23(6)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label23(7)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label23(8)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label23(9)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label23(10)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label23(11)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Label23(12)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Label23(13)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Label23(14)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Label4(1)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Label4(2)"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Label4(3)"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Label4(4)"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "Label55"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "Label52"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "lblNameAgent"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "Label23(55)"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "Label6"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "lstNameAgent"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "Text6(0)"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "Text6(1)"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "Text6(2)"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "Text6(3)"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "Text6(4)"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "Text8(0)"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "Text8(1)"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "Text8(2)"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "Text8(3)"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "Text8(4)"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "Text8(9)"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "Text8(8)"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).Control(39)=   "Text8(7)"
      Tab(0).Control(39).Enabled=   0   'False
      Tab(0).Control(40)=   "Text8(6)"
      Tab(0).Control(40).Enabled=   0   'False
      Tab(0).Control(41)=   "Text8(5)"
      Tab(0).Control(41).Enabled=   0   'False
      Tab(0).Control(42)=   "Text7(0)"
      Tab(0).Control(42).Enabled=   0   'False
      Tab(0).Control(43)=   "Text5"
      Tab(0).Control(43).Enabled=   0   'False
      Tab(0).Control(44)=   "Text7(1)"
      Tab(0).Control(44).Enabled=   0   'False
      Tab(0).Control(45)=   "Text7(2)"
      Tab(0).Control(45).Enabled=   0   'False
      Tab(0).Control(46)=   "Text7(3)"
      Tab(0).Control(46).Enabled=   0   'False
      Tab(0).Control(47)=   "Text7(4)"
      Tab(0).Control(47).Enabled=   0   'False
      Tab(0).Control(48)=   "Check1(0)"
      Tab(0).Control(48).Enabled=   0   'False
      Tab(0).Control(49)=   "Check1(1)"
      Tab(0).Control(49).Enabled=   0   'False
      Tab(0).Control(50)=   "Check1(2)"
      Tab(0).Control(50).Enabled=   0   'False
      Tab(0).Control(51)=   "Text41"
      Tab(0).Control(51).Enabled=   0   'False
      Tab(0).Control(52)=   "Text62"
      Tab(0).Control(52).Enabled=   0   'False
      Tab(0).Control(53)=   "txtCP84"
      Tab(0).Control(53).Enabled=   0   'False
      Tab(0).ControlCount=   54
      TabCaption(1)   =   " 其 他 "
      TabPicture(1)   =   "frm06010303_1.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Check1(20)"
      Tab(1).Control(1)=   "Check1(19)"
      Tab(1).Control(2)=   "Check1(18)"
      Tab(1).Control(3)=   "Check1(17)"
      Tab(1).Control(4)=   "Check1(16)"
      Tab(1).Control(5)=   "Check1(15)"
      Tab(1).Control(6)=   "Check1(14)"
      Tab(1).Control(7)=   "Check1(13)"
      Tab(1).Control(8)=   "Check1(12)"
      Tab(1).Control(9)=   "Check1(11)"
      Tab(1).Control(10)=   "Check1(10)"
      Tab(1).Control(11)=   "Check1(9)"
      Tab(1).Control(12)=   "Check1(8)"
      Tab(1).Control(13)=   "Check1(7)"
      Tab(1).Control(14)=   "Check1(6)"
      Tab(1).Control(15)=   "Check1(5)"
      Tab(1).Control(16)=   "Text56"
      Tab(1).Control(17)=   "Text54"
      Tab(1).Control(18)=   "Text52"
      Tab(1).Control(19)=   "Text48"
      Tab(1).Control(20)=   "Text46"
      Tab(1).Control(21)=   "Text44"
      Tab(1).Control(22)=   "Text36"
      Tab(1).Control(23)=   "Text34"
      Tab(1).Control(24)=   "Text38(2)"
      Tab(1).Control(25)=   "Text38(1)"
      Tab(1).Control(26)=   "Text60"
      Tab(1).Control(27)=   "Text58"
      Tab(1).Control(28)=   "Text50"
      Tab(1).Control(29)=   "Text38(0)"
      Tab(1).Control(30)=   "Label3(8)"
      Tab(1).Control(31)=   "Label42(2)"
      Tab(1).Control(32)=   "Label42(1)"
      Tab(1).Control(33)=   "Label42(0)"
      Tab(1).ControlCount=   34
      TabCaption(2)   =   "代表人"
      TabPicture(2)   =   "frm06010303_1.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Combo2(1)"
      Tab(2).Control(1)=   "Combo2(0)"
      Tab(2).Control(2)=   "Combo2(9)"
      Tab(2).Control(3)=   "Combo2(8)"
      Tab(2).Control(4)=   "Combo2(7)"
      Tab(2).Control(5)=   "Combo2(6)"
      Tab(2).Control(6)=   "Combo2(5)"
      Tab(2).Control(7)=   "Combo2(4)"
      Tab(2).Control(8)=   "Combo2(3)"
      Tab(2).Control(9)=   "Combo2(2)"
      Tab(2).Control(10)=   "Check1(3)"
      Tab(2).Control(11)=   "Text9(3)"
      Tab(2).Control(12)=   "Text9(0)"
      Tab(2).Control(13)=   "Text9(1)"
      Tab(2).Control(14)=   "Text9(2)"
      Tab(2).Control(15)=   "Text9(24)"
      Tab(2).Control(16)=   "Text9(21)"
      Tab(2).Control(17)=   "Text9(18)"
      Tab(2).Control(18)=   "Text9(15)"
      Tab(2).Control(19)=   "Text9(12)"
      Tab(2).Control(20)=   "Text9(9)"
      Tab(2).Control(21)=   "Text9(6)"
      Tab(2).Control(22)=   "Text9(4)"
      Tab(2).Control(23)=   "Text9(5)"
      Tab(2).Control(24)=   "Text9(7)"
      Tab(2).Control(25)=   "Text9(8)"
      Tab(2).Control(26)=   "Text9(10)"
      Tab(2).Control(27)=   "Text9(11)"
      Tab(2).Control(28)=   "Text9(13)"
      Tab(2).Control(29)=   "Text9(14)"
      Tab(2).Control(30)=   "Text9(16)"
      Tab(2).Control(31)=   "Text9(17)"
      Tab(2).Control(32)=   "Text9(19)"
      Tab(2).Control(33)=   "Text9(20)"
      Tab(2).Control(34)=   "Text9(22)"
      Tab(2).Control(35)=   "Text9(23)"
      Tab(2).Control(36)=   "Text9(25)"
      Tab(2).Control(37)=   "Text9(26)"
      Tab(2).Control(38)=   "Text9(27)"
      Tab(2).Control(39)=   "Text9(28)"
      Tab(2).Control(40)=   "Text9(29)"
      Tab(2).Control(41)=   "Label23(54)"
      Tab(2).Control(42)=   "Label23(53)"
      Tab(2).Control(43)=   "Label23(52)"
      Tab(2).Control(44)=   "Label23(51)"
      Tab(2).Control(45)=   "Label23(50)"
      Tab(2).Control(46)=   "Label23(49)"
      Tab(2).Control(47)=   "Label23(48)"
      Tab(2).Control(48)=   "Label23(47)"
      Tab(2).Control(49)=   "Label23(46)"
      Tab(2).Control(50)=   "Label23(45)"
      Tab(2).Control(51)=   "Label23(44)"
      Tab(2).Control(52)=   "Label23(43)"
      Tab(2).Control(53)=   "Label23(42)"
      Tab(2).Control(54)=   "Label23(41)"
      Tab(2).Control(55)=   "Label23(40)"
      Tab(2).Control(56)=   "Label23(39)"
      Tab(2).Control(57)=   "Label23(38)"
      Tab(2).Control(58)=   "Label23(37)"
      Tab(2).Control(59)=   "Label23(36)"
      Tab(2).Control(60)=   "Label23(35)"
      Tab(2).Control(61)=   "Label23(34)"
      Tab(2).Control(62)=   "Label23(33)"
      Tab(2).Control(63)=   "Label23(32)"
      Tab(2).Control(64)=   "Label23(31)"
      Tab(2).Control(65)=   "Label23(15)"
      Tab(2).Control(66)=   "Label23(16)"
      Tab(2).Control(67)=   "Label23(17)"
      Tab(2).Control(68)=   "Label23(18)"
      Tab(2).Control(69)=   "Label23(19)"
      Tab(2).Control(70)=   "Label23(20)"
      Tab(2).ControlCount=   71
      Begin VB.CheckBox Check1 
         Caption         =   "申請人國籍"
         ForeColor       =   &H000000C0&
         Height          =   180
         Index           =   20
         Left            =   -74760
         TabIndex        =   57
         Top             =   5130
         Width           =   1485
      End
      Begin VB.CheckBox Check1 
         Caption         =   "追加、刪除或更正發明人/創作人/設計人"
         ForeColor       =   &H000000C0&
         Height          =   180
         Index           =   19
         Left            =   -74760
         TabIndex        =   56
         Top             =   4860
         Width           =   3585
      End
      Begin VB.CheckBox Check1 
         Caption         =   "變更發明人/創作人/設計人之姓名"
         ForeColor       =   &H000000C0&
         Height          =   180
         Index           =   18
         Left            =   -74760
         TabIndex        =   55
         Top             =   4620
         Width           =   3045
      End
      Begin VB.CheckBox Check1 
         Caption         =   "變更發明人/創作人/設計人之國籍"
         ForeColor       =   &H000000C0&
         Height          =   180
         Index           =   17
         Left            =   -74760
         TabIndex        =   54
         Top             =   4380
         Width           =   3825
      End
      Begin VB.TextBox txtCP84 
         Height          =   270
         Left            =   6120
         MaxLength       =   7
         TabIndex        =   27
         Top             =   2310
         Width           =   1140
      End
      Begin VB.TextBox Text62 
         Height          =   270
         Left            =   4260
         MaxLength       =   1
         TabIndex        =   1
         Top             =   372
         Width           =   375
      End
      Begin VB.TextBox Text41 
         Height          =   270
         Left            =   7260
         MaxLength       =   1
         TabIndex        =   2
         Top             =   372
         Width           =   375
      End
      Begin VB.ComboBox Combo2 
         Height          =   276
         Index           =   1
         Left            =   -73920
         Style           =   2  '單純下拉式
         TabIndex        =   63
         Top             =   1435
         Width           =   3255
      End
      Begin VB.ComboBox Combo2 
         Height          =   276
         Index           =   0
         Left            =   -73920
         Style           =   2  '單純下拉式
         TabIndex        =   59
         Top             =   372
         Width           =   3255
      End
      Begin VB.ComboBox Combo2 
         Height          =   276
         Index           =   9
         Left            =   -69720
         Style           =   2  '單純下拉式
         TabIndex        =   95
         Top             =   4624
         Width           =   3255
      End
      Begin VB.ComboBox Combo2 
         Height          =   276
         Index           =   8
         Left            =   -69720
         Style           =   2  '單純下拉式
         TabIndex        =   91
         Top             =   3561
         Width           =   3255
      End
      Begin VB.ComboBox Combo2 
         Height          =   276
         Index           =   7
         Left            =   -69720
         Style           =   2  '單純下拉式
         TabIndex        =   87
         Top             =   2498
         Width           =   3255
      End
      Begin VB.ComboBox Combo2 
         Height          =   276
         Index           =   6
         Left            =   -69720
         Style           =   2  '單純下拉式
         TabIndex        =   83
         Top             =   1435
         Width           =   3255
      End
      Begin VB.ComboBox Combo2 
         Height          =   276
         Index           =   5
         Left            =   -69720
         Style           =   2  '單純下拉式
         TabIndex        =   79
         Top             =   372
         Width           =   3255
      End
      Begin VB.ComboBox Combo2 
         Height          =   276
         Index           =   4
         Left            =   -73920
         Style           =   2  '單純下拉式
         TabIndex        =   75
         Top             =   4624
         Width           =   3255
      End
      Begin VB.ComboBox Combo2 
         Height          =   276
         Index           =   3
         Left            =   -73920
         Style           =   2  '單純下拉式
         TabIndex        =   71
         Top             =   3561
         Width           =   3255
      End
      Begin VB.ComboBox Combo2 
         Height          =   276
         Index           =   2
         Left            =   -73920
         Style           =   2  '單純下拉式
         TabIndex        =   67
         Top             =   2498
         Width           =   3255
      End
      Begin VB.CheckBox Check1 
         Caption         =   "代表人"
         Height          =   180
         Index           =   3
         Left            =   -74760
         TabIndex        =   58
         Top             =   372
         Width           =   855
      End
      Begin VB.CheckBox Check1 
         Caption         =   "商品組群"
         Height          =   180
         Index           =   16
         Left            =   -74760
         TabIndex        =   52
         Top             =   4065
         Width           =   1215
      End
      Begin VB.CheckBox Check1 
         Caption         =   "商品類別"
         Height          =   180
         Index           =   15
         Left            =   -74760
         TabIndex        =   50
         Top             =   3795
         Width           =   1215
      End
      Begin VB.CheckBox Check1 
         Caption         =   "減縮商品"
         Height          =   180
         Index           =   14
         Left            =   -74760
         TabIndex        =   48
         Top             =   3525
         Width           =   1575
      End
      Begin VB.CheckBox Check1 
         Caption         =   "圖樣"
         Height          =   180
         Index           =   13
         Left            =   -74760
         TabIndex        =   46
         Top             =   3255
         Width           =   855
      End
      Begin VB.CheckBox Check1 
         Caption         =   "正商標號數"
         Height          =   180
         Index           =   12
         Left            =   -74760
         TabIndex        =   44
         Top             =   3015
         Width           =   1455
      End
      Begin VB.CheckBox Check1 
         Caption         =   "其他"
         Height          =   180
         Index           =   11
         Left            =   -74760
         TabIndex        =   42
         Top             =   2760
         Width           =   855
      End
      Begin VB.CheckBox Check1 
         Caption         =   "代理人"
         Height          =   180
         Index           =   10
         Left            =   -74760
         TabIndex        =   40
         Top             =   2400
         Width           =   855
      End
      Begin VB.CheckBox Check1 
         Caption         =   "代表人印鑑"
         Height          =   180
         Index           =   9
         Left            =   -74760
         TabIndex        =   38
         Top             =   2170
         Width           =   1575
      End
      Begin VB.CheckBox Check1 
         Caption         =   "申請人印鑑"
         Height          =   180
         Index           =   8
         Left            =   -74760
         TabIndex        =   36
         Top             =   1932
         Width           =   1695
      End
      Begin VB.CheckBox Check1 
         Caption         =   "案件名稱"
         Height          =   180
         Index           =   7
         Left            =   -74760
         TabIndex        =   32
         Top             =   885
         Width           =   1695
      End
      Begin VB.CheckBox Check1 
         Caption         =   "專利商標種類代號"
         Height          =   180
         Index           =   6
         Left            =   -74760
         TabIndex        =   30
         Top             =   612
         Width           =   1815
      End
      Begin VB.CheckBox Check1 
         Caption         =   "申請日"
         Height          =   180
         Index           =   5
         Left            =   -74760
         TabIndex        =   28
         Top             =   372
         Width           =   855
      End
      Begin VB.CheckBox Check1 
         Caption         =   "申請人地址"
         Height          =   180
         Index           =   2
         Left            =   90
         TabIndex        =   15
         Top             =   3765
         Width           =   1245
      End
      Begin VB.CheckBox Check1 
         Caption         =   "申請人中譯文"
         Height          =   180
         Index           =   1
         Left            =   90
         TabIndex        =   9
         Top             =   2385
         Visible         =   0   'False
         Width           =   3300
      End
      Begin VB.CheckBox Check1 
         Caption         =   "申請人"
         Height          =   180
         Index           =   0
         Left            =   90
         TabIndex        =   3
         Top             =   672
         Width           =   1575
      End
      Begin VB.TextBox Text7 
         Height          =   285
         Index           =   4
         Left            =   390
         MaxLength       =   9
         TabIndex        =   8
         Top             =   2025
         Width           =   1125
      End
      Begin VB.TextBox Text7 
         Height          =   285
         Index           =   3
         Left            =   390
         MaxLength       =   9
         TabIndex        =   7
         Top             =   1755
         Width           =   1125
      End
      Begin VB.TextBox Text7 
         Height          =   285
         Index           =   2
         Left            =   390
         MaxLength       =   9
         TabIndex        =   6
         Top             =   1485
         Width           =   1125
      End
      Begin VB.TextBox Text7 
         Height          =   285
         Index           =   1
         Left            =   390
         MaxLength       =   9
         TabIndex        =   5
         Top             =   1215
         Width           =   1125
      End
      Begin VB.TextBox Text5 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   270
         Left            =   1170
         MaxLength       =   7
         TabIndex        =   0
         Top             =   372
         Width           =   1095
      End
      Begin VB.TextBox Text56 
         Height          =   270
         Left            =   -72720
         MaxLength       =   200
         TabIndex        =   49
         Top             =   3525
         Width           =   375
      End
      Begin VB.TextBox Text54 
         Height          =   270
         Left            =   -72720
         MaxLength       =   1
         TabIndex        =   47
         Top             =   3255
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox Text52 
         Height          =   270
         Left            =   -72720
         MaxLength       =   20
         TabIndex        =   45
         Top             =   3015
         Width           =   1575
      End
      Begin VB.TextBox Text48 
         Height          =   270
         Left            =   -72720
         MaxLength       =   1
         TabIndex        =   41
         Top             =   2430
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox Text46 
         Height          =   270
         Left            =   -72720
         MaxLength       =   1
         TabIndex        =   39
         Top             =   2172
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox Text44 
         Height          =   270
         Left            =   -72720
         MaxLength       =   1
         TabIndex        =   37
         Top             =   1932
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox Text36 
         Height          =   270
         Left            =   -72720
         MaxLength       =   1
         TabIndex        =   31
         Top             =   645
         Width           =   375
      End
      Begin VB.TextBox Text34 
         Height          =   270
         Left            =   -72720
         MaxLength       =   8
         TabIndex        =   29
         Top             =   372
         Width           =   975
      End
      Begin VB.TextBox Text7 
         Height          =   285
         Index           =   0
         Left            =   390
         MaxLength       =   9
         TabIndex        =   4
         Top             =   930
         Width           =   1125
      End
      Begin MSForms.TextBox Text9 
         Height          =   285
         Index           =   3
         Left            =   -73920
         TabIndex        =   64
         Top             =   1712
         Width           =   3255
         VariousPropertyBits=   679493661
         BackColor       =   14737632
         MaxLength       =   50
         Size            =   "5741;503"
         FontName        =   "新細明體-ExtB"
         FontEffects     =   1073750016
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text9 
         Height          =   285
         Index           =   0
         Left            =   -73920
         TabIndex        =   60
         Top             =   649
         Width           =   3255
         VariousPropertyBits=   679493661
         BackColor       =   14737632
         MaxLength       =   50
         Size            =   "5741;503"
         FontName        =   "新細明體-ExtB"
         FontEffects     =   1073750016
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text9 
         Height          =   285
         Index           =   1
         Left            =   -73920
         TabIndex        =   61
         Top             =   911
         Width           =   3255
         VariousPropertyBits=   679493661
         BackColor       =   14737632
         MaxLength       =   80
         Size            =   "5741;503"
         FontName        =   "新細明體-ExtB"
         FontEffects     =   1073750016
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text9 
         Height          =   285
         Index           =   2
         Left            =   -73920
         TabIndex        =   62
         Top             =   1173
         Width           =   3255
         VariousPropertyBits=   679493661
         BackColor       =   14737632
         MaxLength       =   40
         Size            =   "5741;503"
         FontName        =   "新細明體-ExtB"
         FontEffects     =   1073750016
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text9 
         Height          =   285
         Index           =   24
         Left            =   -69720
         TabIndex        =   92
         Top             =   3838
         Width           =   3255
         VariousPropertyBits=   679493661
         BackColor       =   14737632
         MaxLength       =   50
         Size            =   "5741;503"
         FontName        =   "新細明體-ExtB"
         FontEffects     =   1073750016
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text9 
         Height          =   285
         Index           =   21
         Left            =   -69720
         TabIndex        =   88
         Top             =   2775
         Width           =   3255
         VariousPropertyBits=   679493661
         BackColor       =   14737632
         MaxLength       =   50
         Size            =   "5741;503"
         FontName        =   "新細明體-ExtB"
         FontEffects     =   1073750016
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text9 
         Height          =   285
         Index           =   18
         Left            =   -69720
         TabIndex        =   84
         Top             =   1712
         Width           =   3255
         VariousPropertyBits=   679493661
         BackColor       =   14737632
         MaxLength       =   50
         Size            =   "5741;503"
         FontName        =   "新細明體-ExtB"
         FontEffects     =   1073750016
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text9 
         Height          =   285
         Index           =   15
         Left            =   -69720
         TabIndex        =   80
         Top             =   649
         Width           =   3255
         VariousPropertyBits=   679493661
         BackColor       =   14737632
         MaxLength       =   50
         Size            =   "5741;503"
         FontName        =   "新細明體-ExtB"
         FontEffects     =   1073750016
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text9 
         Height          =   285
         Index           =   12
         Left            =   -73920
         TabIndex        =   76
         Top             =   4901
         Width           =   3255
         VariousPropertyBits=   679493661
         BackColor       =   14737632
         MaxLength       =   50
         Size            =   "5741;503"
         FontName        =   "新細明體-ExtB"
         FontEffects     =   1073750016
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text9 
         Height          =   285
         Index           =   9
         Left            =   -73920
         TabIndex        =   72
         Top             =   3838
         Width           =   3255
         VariousPropertyBits=   679493661
         BackColor       =   14737632
         MaxLength       =   50
         Size            =   "5741;503"
         FontName        =   "新細明體-ExtB"
         FontEffects     =   1073750016
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text9 
         Height          =   285
         Index           =   6
         Left            =   -73920
         TabIndex        =   68
         Top             =   2775
         Width           =   3255
         VariousPropertyBits=   679493661
         BackColor       =   14737632
         MaxLength       =   50
         Size            =   "5741;503"
         FontName        =   "新細明體-ExtB"
         FontEffects     =   1073750016
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text9 
         Height          =   285
         Index           =   4
         Left            =   -73920
         TabIndex        =   65
         Top             =   1974
         Width           =   3255
         VariousPropertyBits=   679493661
         BackColor       =   14737632
         MaxLength       =   80
         Size            =   "5741;503"
         FontName        =   "新細明體-ExtB"
         FontEffects     =   1073750016
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text38 
         Height          =   285
         Index           =   2
         Left            =   -72720
         TabIndex        =   35
         Top             =   1650
         Width           =   6465
         VariousPropertyBits=   671105051
         MaxLength       =   160
         Size            =   "11404;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text38 
         Height          =   285
         Index           =   1
         Left            =   -72720
         TabIndex        =   34
         Top             =   1365
         Width           =   6465
         VariousPropertyBits=   671105051
         MaxLength       =   180
         Size            =   "11404;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text60 
         Height          =   285
         Left            =   -72720
         TabIndex        =   53
         Top             =   4065
         Width           =   6465
         VariousPropertyBits=   671105051
         MaxLength       =   300
         Size            =   "11404;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text58 
         Height          =   285
         Left            =   -72720
         TabIndex        =   51
         Top             =   3795
         Width           =   6465
         VariousPropertyBits=   671105051
         MaxLength       =   395
         Size            =   "11404;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text50 
         Height          =   285
         Left            =   -72720
         TabIndex        =   43
         Top             =   2730
         Width           =   6420
         VariousPropertyBits=   671105051
         MaxLength       =   2000
         Size            =   "11324;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text38 
         Height          =   285
         Index           =   0
         Left            =   -72720
         TabIndex        =   33
         Top             =   1065
         Width           =   6465
         VariousPropertyBits=   671105051
         MaxLength       =   160
         Size            =   "11404;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text9 
         Height          =   285
         Index           =   5
         Left            =   -73920
         TabIndex        =   66
         Top             =   2236
         Width           =   3255
         VariousPropertyBits=   679493661
         BackColor       =   14737632
         MaxLength       =   40
         Size            =   "5741;503"
         FontName        =   "新細明體-ExtB"
         FontEffects     =   1073750016
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text9 
         Height          =   285
         Index           =   7
         Left            =   -73920
         TabIndex        =   69
         Top             =   3037
         Width           =   3255
         VariousPropertyBits=   679493661
         BackColor       =   14737632
         MaxLength       =   80
         Size            =   "5741;503"
         FontName        =   "新細明體-ExtB"
         FontEffects     =   1073750016
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text9 
         Height          =   285
         Index           =   8
         Left            =   -73920
         TabIndex        =   70
         Top             =   3299
         Width           =   3255
         VariousPropertyBits=   679493661
         BackColor       =   14737632
         MaxLength       =   40
         Size            =   "5741;503"
         FontName        =   "新細明體-ExtB"
         FontEffects     =   1073750016
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text9 
         Height          =   285
         Index           =   10
         Left            =   -73920
         TabIndex        =   73
         Top             =   4100
         Width           =   3255
         VariousPropertyBits=   679493661
         BackColor       =   14737632
         MaxLength       =   80
         Size            =   "5741;503"
         FontName        =   "新細明體-ExtB"
         FontEffects     =   1073750016
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text9 
         Height          =   285
         Index           =   11
         Left            =   -73920
         TabIndex        =   74
         Top             =   4365
         Width           =   3255
         VariousPropertyBits=   679493661
         BackColor       =   14737632
         MaxLength       =   40
         Size            =   "5741;503"
         FontName        =   "新細明體-ExtB"
         FontEffects     =   1073750016
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text9 
         Height          =   285
         Index           =   13
         Left            =   -73920
         TabIndex        =   77
         Top             =   5163
         Width           =   3255
         VariousPropertyBits=   679493661
         BackColor       =   14737632
         MaxLength       =   80
         Size            =   "5741;503"
         FontName        =   "新細明體-ExtB"
         FontEffects     =   1073750016
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text9 
         Height          =   285
         Index           =   14
         Left            =   -73920
         TabIndex        =   78
         Top             =   5415
         Width           =   3255
         VariousPropertyBits=   679493661
         BackColor       =   14737632
         MaxLength       =   40
         Size            =   "8555;529"
         FontName        =   "新細明體-ExtB"
         FontEffects     =   1073750016
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text9 
         Height          =   285
         Index           =   16
         Left            =   -69720
         TabIndex        =   81
         Top             =   911
         Width           =   3255
         VariousPropertyBits=   679493661
         BackColor       =   14737632
         MaxLength       =   80
         Size            =   "5741;503"
         FontName        =   "新細明體-ExtB"
         FontEffects     =   1073750016
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text9 
         Height          =   285
         Index           =   17
         Left            =   -69720
         TabIndex        =   82
         Top             =   1173
         Width           =   3255
         VariousPropertyBits=   679493661
         BackColor       =   14737632
         MaxLength       =   40
         Size            =   "5741;503"
         FontName        =   "新細明體-ExtB"
         FontEffects     =   1073750016
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text9 
         Height          =   285
         Index           =   19
         Left            =   -69720
         TabIndex        =   85
         Top             =   1974
         Width           =   3255
         VariousPropertyBits=   679493661
         BackColor       =   14737632
         MaxLength       =   80
         Size            =   "5741;503"
         FontName        =   "新細明體-ExtB"
         FontEffects     =   1073750016
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text9 
         Height          =   285
         Index           =   20
         Left            =   -69720
         TabIndex        =   86
         Top             =   2236
         Width           =   3255
         VariousPropertyBits=   679493661
         BackColor       =   14737632
         MaxLength       =   40
         Size            =   "5741;503"
         FontName        =   "新細明體-ExtB"
         FontEffects     =   1073750016
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text9 
         Height          =   285
         Index           =   22
         Left            =   -69720
         TabIndex        =   89
         Top             =   3037
         Width           =   3255
         VariousPropertyBits=   679493661
         BackColor       =   14737632
         MaxLength       =   80
         Size            =   "5741;503"
         FontName        =   "新細明體-ExtB"
         FontEffects     =   1073750016
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text9 
         Height          =   285
         Index           =   23
         Left            =   -69720
         TabIndex        =   90
         Top             =   3299
         Width           =   3255
         VariousPropertyBits=   679493661
         BackColor       =   14737632
         MaxLength       =   40
         Size            =   "5741;503"
         FontName        =   "新細明體-ExtB"
         FontEffects     =   1073750016
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text9 
         Height          =   285
         Index           =   25
         Left            =   -69720
         TabIndex        =   93
         Top             =   4100
         Width           =   3255
         VariousPropertyBits=   679493661
         BackColor       =   14737632
         MaxLength       =   80
         Size            =   "5741;503"
         FontName        =   "新細明體-ExtB"
         FontEffects     =   1073750016
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text9 
         Height          =   285
         Index           =   26
         Left            =   -69720
         TabIndex        =   94
         Top             =   4365
         Width           =   3255
         VariousPropertyBits=   679493661
         BackColor       =   14737632
         MaxLength       =   40
         Size            =   "5741;503"
         FontName        =   "新細明體-ExtB"
         FontEffects     =   1073750016
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text9 
         Height          =   285
         Index           =   27
         Left            =   -69720
         TabIndex        =   96
         Top             =   4901
         Width           =   3255
         VariousPropertyBits=   679493661
         BackColor       =   14737632
         MaxLength       =   50
         Size            =   "5741;503"
         FontName        =   "新細明體-ExtB"
         FontEffects     =   1073750016
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text9 
         Height          =   285
         Index           =   28
         Left            =   -69720
         TabIndex        =   97
         Top             =   5163
         Width           =   3255
         VariousPropertyBits=   679493661
         BackColor       =   14737632
         MaxLength       =   80
         Size            =   "5741;503"
         FontName        =   "新細明體-ExtB"
         FontEffects     =   1073750016
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text9 
         Height          =   285
         Index           =   29
         Left            =   -69720
         TabIndex        =   98
         Top             =   5415
         Width           =   3255
         VariousPropertyBits=   679493661
         BackColor       =   14737632
         MaxLength       =   40
         Size            =   "5741;503"
         FontName        =   "新細明體-ExtB"
         FontEffects     =   1073750016
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text8 
         Height          =   285
         Index           =   5
         Left            =   4830
         TabIndex        =   17
         Top             =   4005
         Width           =   4425
         VariousPropertyBits=   671105055
         MaxLength       =   154
         Size            =   "7805;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text8 
         Height          =   285
         Index           =   6
         Left            =   4830
         TabIndex        =   19
         Top             =   4290
         Width           =   4425
         VariousPropertyBits=   671105055
         MaxLength       =   154
         Size            =   "7805;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text8 
         Height          =   285
         Index           =   7
         Left            =   4830
         TabIndex        =   21
         Top             =   4575
         Width           =   4425
         VariousPropertyBits=   671105055
         MaxLength       =   154
         Size            =   "7805;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text8 
         Height          =   285
         Index           =   8
         Left            =   4830
         TabIndex        =   23
         Top             =   4860
         Width           =   4425
         VariousPropertyBits=   671105055
         MaxLength       =   154
         Size            =   "7805;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text8 
         Height          =   285
         Index           =   9
         Left            =   4830
         TabIndex        =   25
         Top             =   5145
         Width           =   4425
         VariousPropertyBits=   671105055
         MaxLength       =   154
         Size            =   "7805;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text8 
         Height          =   285
         Index           =   4
         Left            =   390
         TabIndex        =   24
         Top             =   5145
         Width           =   4425
         VariousPropertyBits=   671105055
         MaxLength       =   80
         Size            =   "7805;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text8 
         Height          =   285
         Index           =   3
         Left            =   390
         TabIndex        =   22
         Top             =   4860
         Width           =   4425
         VariousPropertyBits=   671105055
         MaxLength       =   80
         Size            =   "7805;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text8 
         Height          =   285
         Index           =   2
         Left            =   390
         TabIndex        =   20
         Top             =   4575
         Width           =   4425
         VariousPropertyBits=   671105055
         MaxLength       =   80
         Size            =   "7805;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text8 
         Height          =   285
         Index           =   1
         Left            =   390
         TabIndex        =   18
         Top             =   4290
         Width           =   4425
         VariousPropertyBits=   671105055
         MaxLength       =   80
         Size            =   "7805;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text8 
         Height          =   285
         Index           =   0
         Left            =   390
         TabIndex        =   16
         Top             =   4005
         Width           =   4425
         VariousPropertyBits=   671105055
         MaxLength       =   80
         Size            =   "7805;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text6 
         Height          =   285
         Index           =   4
         Left            =   390
         TabIndex        =   14
         Top             =   3432
         Visible         =   0   'False
         Width           =   2895
         VariousPropertyBits=   671105051
         MaxLength       =   60
         Size            =   "7223;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text6 
         Height          =   285
         Index           =   3
         Left            =   390
         TabIndex        =   13
         Top             =   3192
         Visible         =   0   'False
         Width           =   2895
         VariousPropertyBits=   671105051
         MaxLength       =   60
         Size            =   "7223;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text6 
         Height          =   285
         Index           =   2
         Left            =   390
         TabIndex        =   12
         Top             =   2952
         Visible         =   0   'False
         Width           =   2895
         VariousPropertyBits=   671105051
         MaxLength       =   60
         Size            =   "7223;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text6 
         Height          =   285
         Index           =   1
         Left            =   390
         TabIndex        =   11
         Top             =   2712
         Visible         =   0   'False
         Width           =   2895
         VariousPropertyBits=   671105051
         MaxLength       =   60
         Size            =   "7223;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text6 
         Height          =   285
         Index           =   0
         Left            =   390
         TabIndex        =   10
         Top             =   2490
         Visible         =   0   'False
         Width           =   2895
         VariousPropertyBits=   671105051
         MaxLength       =   60
         Size            =   "7223;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.ListBox lstNameAgent 
         Height          =   315
         Left            =   6120
         TabIndex        =   26
         Top             =   780
         Width           =   1500
         VariousPropertyBits=   746586139
         ScrollBars      =   2
         DisplayStyle    =   2
         Size            =   "2646;556"
         MatchEntry      =   0
         ListStyle       =   1
         MultiSelect     =   1
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "繳費金額:"
         Height          =   180
         Left            =   5325
         TabIndex        =   188
         Top             =   2310
         Width           =   765
      End
      Begin VB.Label Label23 
         Caption         =   "(請先至客戶資料維護修改地址再產生申請書)    中文地址 / 英文地址"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   55
         Left            =   1320
         TabIndex        =   187
         Top             =   3780
         Width           =   6105
      End
      Begin VB.Label lblNameAgent 
         AutoSize        =   -1  'True
         Caption         =   "出名代理人"
         Height          =   180
         Left            =   5190
         TabIndex        =   186
         Top             =   810
         Visible         =   0   'False
         Width           =   900
      End
      Begin VB.Label Label52 
         AutoSize        =   -1  'True
         Caption         =   "是否修改申請書內容              (Y:WORD)"
         Height          =   180
         Left            =   2580
         TabIndex        =   185
         Top             =   420
         Width           =   3090
      End
      Begin VB.Label Label55 
         AutoSize        =   -1  'True
         Caption         =   "是否列印申請書            (N:不印)"
         Height          =   180
         Left            =   5880
         TabIndex        =   184
         Top             =   420
         Width           =   2445
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "10 (中)"
         Height          =   180
         Index           =   54
         Left            =   -70440
         TabIndex        =   183
         Top             =   4901
         Width           =   528
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "10 (英)"
         Height          =   180
         Index           =   53
         Left            =   -70440
         TabIndex        =   182
         Top             =   5163
         Width           =   528
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "10 (日)"
         Height          =   180
         Index           =   52
         Left            =   -70440
         TabIndex        =   181
         Top             =   5415
         Width           =   528
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "9 (中)"
         Height          =   180
         Index           =   51
         Left            =   -70440
         TabIndex        =   180
         Top             =   3838
         Width           =   432
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "9 (英)"
         Height          =   180
         Index           =   50
         Left            =   -70440
         TabIndex        =   179
         Top             =   4100
         Width           =   432
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "9 (日)"
         Height          =   180
         Index           =   49
         Left            =   -70440
         TabIndex        =   178
         Top             =   4365
         Width           =   432
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "8 (中)"
         Height          =   180
         Index           =   48
         Left            =   -70440
         TabIndex        =   177
         Top             =   2775
         Width           =   432
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "8 (英)"
         Height          =   180
         Index           =   47
         Left            =   -70440
         TabIndex        =   176
         Top             =   3037
         Width           =   432
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "8 (日)"
         Height          =   180
         Index           =   46
         Left            =   -70440
         TabIndex        =   175
         Top             =   3299
         Width           =   432
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "7 (中)"
         Height          =   180
         Index           =   45
         Left            =   -70440
         TabIndex        =   174
         Top             =   1712
         Width           =   432
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "7 (英)"
         Height          =   180
         Index           =   44
         Left            =   -70440
         TabIndex        =   173
         Top             =   1974
         Width           =   432
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "7 (日)"
         Height          =   180
         Index           =   43
         Left            =   -70440
         TabIndex        =   172
         Top             =   2236
         Width           =   432
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "6 (中)"
         Height          =   180
         Index           =   42
         Left            =   -70440
         TabIndex        =   171
         Top             =   649
         Width           =   432
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "6 (英)"
         Height          =   180
         Index           =   41
         Left            =   -70440
         TabIndex        =   170
         Top             =   911
         Width           =   432
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "6 (日)"
         Height          =   180
         Index           =   40
         Left            =   -70440
         TabIndex        =   169
         Top             =   1173
         Width           =   432
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "5 (中)"
         Height          =   180
         Index           =   39
         Left            =   -74640
         TabIndex        =   168
         Top             =   4901
         Width           =   432
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "5 (英)"
         Height          =   180
         Index           =   38
         Left            =   -74640
         TabIndex        =   167
         Top             =   5163
         Width           =   432
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "5 (日)"
         Height          =   180
         Index           =   37
         Left            =   -74640
         TabIndex        =   166
         Top             =   5415
         Width           =   432
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "3 (中)"
         Height          =   180
         Index           =   36
         Left            =   -74640
         TabIndex        =   165
         Top             =   2775
         Width           =   432
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "3 (英)"
         Height          =   180
         Index           =   35
         Left            =   -74640
         TabIndex        =   164
         Top             =   3037
         Width           =   432
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "3 (日)"
         Height          =   180
         Index           =   34
         Left            =   -74640
         TabIndex        =   163
         Top             =   3299
         Width           =   432
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "4 (中)"
         Height          =   180
         Index           =   33
         Left            =   -74640
         TabIndex        =   162
         Top             =   3838
         Width           =   432
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "4 (英)"
         Height          =   180
         Index           =   32
         Left            =   -74640
         TabIndex        =   161
         Top             =   4100
         Width           =   432
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "4 (日)"
         Height          =   180
         Index           =   31
         Left            =   -74640
         TabIndex        =   160
         Top             =   4365
         Width           =   432
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "1 (中)"
         Height          =   180
         Index           =   15
         Left            =   -74640
         TabIndex        =   159
         Top             =   649
         Width           =   432
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "1 (英)"
         Height          =   180
         Index           =   16
         Left            =   -74640
         TabIndex        =   158
         Top             =   911
         Width           =   432
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "1 (日)"
         Height          =   180
         Index           =   17
         Left            =   -74640
         TabIndex        =   157
         Top             =   1173
         Width           =   432
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "2 (中)"
         Height          =   180
         Index           =   18
         Left            =   -74640
         TabIndex        =   156
         Top             =   1712
         Width           =   432
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "2 (英)"
         Height          =   180
         Index           =   19
         Left            =   -74640
         TabIndex        =   155
         Top             =   1974
         Width           =   432
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "2 (日)"
         Height          =   180
         Index           =   20
         Left            =   -74640
         TabIndex        =   154
         Top             =   2236
         Width           =   432
      End
      Begin MSForms.Label Label3 
         Height          =   285
         Index           =   8
         Left            =   -72300
         TabIndex        =   147
         Top             =   690
         Width           =   5760
         VariousPropertyBits=   27
         Caption         =   "LblFM2"
         Size            =   "10160;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label4 
         Height          =   285
         Index           =   4
         Left            =   1575
         TabIndex        =   146
         Top             =   2025
         Width           =   3480
         VariousPropertyBits=   27
         Caption         =   "LblFM2"
         Size            =   "6138;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label4 
         Height          =   285
         Index           =   3
         Left            =   1575
         TabIndex        =   145
         Top             =   1755
         Width           =   825
         VariousPropertyBits=   27
         Caption         =   "LblFM2"
         Size            =   "1455;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label4 
         Height          =   285
         Index           =   2
         Left            =   1575
         TabIndex        =   144
         Top             =   1485
         Width           =   825
         VariousPropertyBits=   27
         Caption         =   "LblFM2"
         Size            =   "1455;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label4 
         Height          =   285
         Index           =   1
         Left            =   1575
         TabIndex        =   143
         Top             =   1215
         Width           =   825
         VariousPropertyBits=   27
         Caption         =   "LblFM2"
         Size            =   "1455;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label42 
         AutoSize        =   -1  'True
         Caption         =   "英文"
         Height          =   180
         Index           =   2
         Left            =   -74490
         TabIndex        =   134
         Top             =   1365
         Width           =   360
      End
      Begin VB.Label Label42 
         AutoSize        =   -1  'True
         Caption         =   "中文"
         Height          =   180
         Index           =   1
         Left            =   -74490
         TabIndex        =   133
         Top             =   1095
         Width           =   360
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "5:"
         Height          =   180
         Index           =   14
         Left            =   210
         TabIndex        =   132
         Top             =   5145
         Width           =   135
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "4:"
         Height          =   180
         Index           =   13
         Left            =   210
         TabIndex        =   131
         Top             =   4860
         Width           =   135
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "3:"
         Height          =   180
         Index           =   12
         Left            =   210
         TabIndex        =   130
         Top             =   4575
         Width           =   135
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "2:"
         Height          =   180
         Index           =   11
         Left            =   210
         TabIndex        =   129
         Top             =   4290
         Width           =   135
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "1:"
         Height          =   180
         Index           =   10
         Left            =   210
         TabIndex        =   128
         Top             =   4005
         Width           =   135
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "5:"
         Height          =   180
         Index           =   9
         Left            =   210
         TabIndex        =   127
         Top             =   2025
         Width           =   135
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "4:"
         Height          =   180
         Index           =   8
         Left            =   210
         TabIndex        =   126
         Top             =   1755
         Width           =   135
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "3:"
         Height          =   180
         Index           =   7
         Left            =   210
         TabIndex        =   125
         Top             =   1485
         Width           =   135
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "2:"
         Height          =   180
         Index           =   6
         Left            =   210
         TabIndex        =   124
         Top             =   1215
         Width           =   135
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "1:"
         Height          =   180
         Index           =   5
         Left            =   210
         TabIndex        =   123
         Top             =   930
         Width           =   135
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "5:"
         Height          =   180
         Index           =   4
         Left            =   210
         TabIndex        =   122
         Top             =   3435
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "4:"
         Height          =   180
         Index           =   3
         Left            =   210
         TabIndex        =   121
         Top             =   3195
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "3:"
         Height          =   180
         Index           =   2
         Left            =   210
         TabIndex        =   120
         Top             =   2955
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "2:"
         Height          =   180
         Index           =   1
         Left            =   210
         TabIndex        =   119
         Top             =   2715
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label Label2 
         Caption         =   "申請書日期:"
         Height          =   255
         Left            =   90
         TabIndex        =   105
         Top             =   375
         Width           =   975
      End
      Begin VB.Label Label42 
         AutoSize        =   -1  'True
         Caption         =   "日文"
         Height          =   180
         Index           =   0
         Left            =   -74490
         TabIndex        =   104
         Top             =   1635
         Width           =   360
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "1:"
         Height          =   180
         Index           =   0
         Left            =   210
         TabIndex        =   103
         Top             =   2475
         Visible         =   0   'False
         Width           =   135
      End
      Begin MSForms.Label Label4 
         Height          =   285
         Index           =   0
         Left            =   1575
         TabIndex        =   102
         Top             =   930
         Width           =   825
         VariousPropertyBits=   27
         Caption         =   "LblFM2"
         Size            =   "1455;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   3105
      Left            =   8604
      TabIndex        =   192
      Top             =   396
      Visible         =   0   'False
      Width           =   3105
      Begin VB.CheckBox Check1 
         Caption         =   "代表人中譯文"
         Height          =   180
         Index           =   4
         Left            =   60
         TabIndex        =   203
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox Text10 
         Height          =   270
         Index           =   0
         Left            =   765
         MaxLength       =   60
         TabIndex        =   202
         Top             =   555
         Width           =   3120
      End
      Begin VB.TextBox Text10 
         Height          =   270
         Index           =   1
         Left            =   765
         MaxLength       =   60
         TabIndex        =   201
         Top             =   795
         Width           =   3120
      End
      Begin VB.TextBox Text10 
         Height          =   270
         Index           =   2
         Left            =   765
         MaxLength       =   60
         TabIndex        =   200
         Top             =   1035
         Width           =   3120
      End
      Begin VB.TextBox Text10 
         Height          =   270
         Index           =   3
         Left            =   765
         MaxLength       =   60
         TabIndex        =   199
         Top             =   1275
         Width           =   3120
      End
      Begin VB.TextBox Text10 
         Height          =   270
         Index           =   4
         Left            =   765
         MaxLength       =   60
         TabIndex        =   198
         Top             =   1515
         Width           =   3120
      End
      Begin VB.TextBox Text10 
         Height          =   270
         Index           =   5
         Left            =   765
         MaxLength       =   60
         TabIndex        =   197
         Top             =   1740
         Width           =   3120
      End
      Begin VB.TextBox Text10 
         Height          =   270
         Index           =   6
         Left            =   765
         MaxLength       =   60
         TabIndex        =   196
         Top             =   1995
         Width           =   3120
      End
      Begin VB.TextBox Text10 
         Height          =   270
         Index           =   7
         Left            =   765
         MaxLength       =   60
         TabIndex        =   195
         Top             =   2235
         Width           =   3120
      End
      Begin VB.TextBox Text10 
         Height          =   270
         Index           =   8
         Left            =   765
         MaxLength       =   60
         TabIndex        =   194
         Top             =   2475
         Width           =   3120
      End
      Begin VB.TextBox Text10 
         Height          =   270
         Index           =   9
         Left            =   765
         MaxLength       =   60
         TabIndex        =   193
         Top             =   2715
         Width           =   3120
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "1:"
         Height          =   180
         Index           =   21
         Left            =   525
         TabIndex        =   213
         Top             =   555
         Width           =   135
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "2:"
         Height          =   180
         Index           =   22
         Left            =   525
         TabIndex        =   212
         Top             =   795
         Width           =   135
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "4:"
         Height          =   180
         Index           =   23
         Left            =   525
         TabIndex        =   211
         Top             =   1275
         Width           =   135
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "3:"
         Height          =   180
         Index           =   24
         Left            =   525
         TabIndex        =   210
         Top             =   1035
         Width           =   135
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "6:"
         Height          =   180
         Index           =   25
         Left            =   525
         TabIndex        =   209
         Top             =   1755
         Width           =   135
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "5:"
         Height          =   180
         Index           =   26
         Left            =   525
         TabIndex        =   208
         Top             =   1515
         Width           =   135
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "8:"
         Height          =   180
         Index           =   27
         Left            =   525
         TabIndex        =   207
         Top             =   2235
         Width           =   135
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "7:"
         Height          =   180
         Index           =   28
         Left            =   525
         TabIndex        =   206
         Top             =   1995
         Width           =   135
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "10:"
         Height          =   180
         Index           =   29
         Left            =   525
         TabIndex        =   205
         Top             =   2715
         Width           =   225
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "9:"
         Height          =   180
         Index           =   30
         Left            =   525
         TabIndex        =   204
         Top             =   2475
         Width           =   135
      End
   End
   Begin MSForms.ComboBox Combo1 
      Height          =   300
      Left            =   1050
      TabIndex        =   135
      Top             =   1320
      Width           =   7170
      VariousPropertyBits=   746604571
      DisplayStyle    =   3
      Size            =   "12647;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label3 
      Height          =   285
      Index           =   11
      Left            =   4200
      TabIndex        =   153
      Top             =   1080
      Width           =   2010
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "3545;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label3 
      Height          =   285
      Index           =   10
      Left            =   1080
      TabIndex        =   152
      Top             =   1080
      Width           =   2130
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "3757;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label3 
      Height          =   285
      Index           =   9
      Left            =   4200
      TabIndex        =   151
      Top             =   105
      Width           =   2130
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "3757;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label57 
      AutoSize        =   -1  'True
      Caption         =   "法定期限:"
      Height          =   180
      Index           =   3
      Left            =   3360
      TabIndex        =   150
      Top             =   1080
      Width           =   768
   End
   Begin VB.Label Label57 
      AutoSize        =   -1  'True
      Caption         =   "本所期限:"
      Height          =   180
      Index           =   2
      Left            =   120
      TabIndex        =   149
      Top             =   1080
      Width           =   768
   End
   Begin VB.Label Label57 
      AutoSize        =   -1  'True
      Caption         =   "申請國家:"
      Height          =   180
      Index           =   1
      Left            =   3360
      TabIndex        =   148
      Top             =   105
      Width           =   765
   End
   Begin MSForms.Label Label3 
      Height          =   285
      Index           =   7
      Left            =   4200
      TabIndex        =   142
      Top             =   420
      Width           =   2130
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "3757;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label3 
      Height          =   285
      Index           =   6
      Left            =   1080
      TabIndex        =   141
      Top             =   420
      Width           =   2130
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "3757;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label3 
      Height          =   285
      Index           =   5
      Left            =   7290
      TabIndex        =   140
      Top             =   750
      Width           =   1650
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "2910;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label3 
      Height          =   285
      Index           =   4
      Left            =   4200
      TabIndex        =   139
      Top             =   750
      Width           =   2130
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "3757;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label3 
      Height          =   285
      Index           =   3
      Left            =   1080
      TabIndex        =   138
      Top             =   750
      Width           =   2130
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "3757;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label3 
      Height          =   285
      Index           =   2
      Left            =   7290
      TabIndex        =   137
      Top             =   420
      Width           =   1650
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "2910;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label3 
      Height          =   285
      Index           =   1
      Left            =   7290
      TabIndex        =   136
      Top             =   1080
      Width           =   1650
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "2910;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label57 
      AutoSize        =   -1  'True
      Caption         =   "案件性質:"
      Height          =   180
      Index           =   0
      Left            =   6480
      TabIndex        =   118
      Top             =   420
      Width           =   765
   End
   Begin VB.Label Label56 
      AutoSize        =   -1  'True
      Caption         =   "機關文號:"
      Height          =   180
      Left            =   6480
      TabIndex        =   117
      Top             =   750
      Width           =   765
   End
   Begin MSForms.Label Label30 
      Height          =   285
      Left            =   6300
      TabIndex        =   116
      Top             =   1080
      Width           =   945
      VariousPropertyBits=   27
      Caption         =   "來函收文日:"
      Size            =   "1667;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label29 
      AutoSize        =   -1  'True
      Caption         =   "智權人員:"
      Height          =   180
      Left            =   3360
      TabIndex        =   115
      Top             =   750
      Width           =   768
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      Caption         =   "承辦人　:"
      Height          =   180
      Left            =   120
      TabIndex        =   114
      Top             =   750
      Width           =   768
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      Caption         =   "申請案號:"
      Height          =   180
      Left            =   120
      TabIndex        =   113
      Top             =   420
      Width           =   768
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "專利號數:"
      Height          =   180
      Left            =   3360
      TabIndex        =   112
      Top             =   420
      Width           =   768
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "專利名稱:"
      Height          =   180
      Left            =   120
      TabIndex        =   111
      Top             =   1350
      Width           =   765
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "本所案號:"
      Height          =   180
      Left            =   120
      TabIndex        =   110
      Top             =   105
      Width           =   765
   End
End
Attribute VB_Name = "frm06010303_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/11/4 Form2.0已修改
'Modified by Morgan 2015/10/29
'申請人中譯文相關欄位隱藏, 統一勾選 申請人 (都先做客戶變更名稱作業然後選新的編號)
'end 2015/10/29

'Memo By Morgan 2012/12/10 智權人員欄已修改
'Memo by Morgan2010/12/27 申請案號欄已修改
'Memo by Morgan2010/8/12 日期欄已修改
Option Explicit

'Modify by Morgan 2005/8/1 改用動態陣列
'Dim pa(1 To T_PA) As String
Dim pa() As String
Dim m_CP110 As String, m_AgentName As String
Dim m_CP22 As String

Dim strReceiveNo As String
Dim intWhere As Integer, intGo As Integer
Dim bolSaveOK As Boolean 'Add by Morgan 2006/7/4
Public oParent As Form 'Add by Morgan 2011/10/5
'Add By Sindy 2018/8/10
Public m_CP118isY As String '是否為電子送件申請書:Y.是 N.不是 "".發文作業呼叫的
Dim m_CaseNo As String
Dim cp() As String
'2018/8/10 END
Dim m_strApplNum As String 'Add By Sindy 2018/10/19 申請人編號
Dim m_Representative As String 'Add By Sindy 2018/10/19 代表人
Public m_CP20isN As Boolean 'Add By Sindy 2023/6/27 True=不請款


Private Sub StartLetter(ByVal ET01 As String, ByVal ET03 As String)
 Dim strTxt(5) As String
 Dim ii As Integer
 Dim strTmp As String
   
   ii = 0
   EndLetter ET01, strReceiveNo, ET03, strUserNum
    'Add By Cheng 2003/01/21
    '傳入例外欄位
    Select Case intWhere
    Case 國內
        Select Case ET03
        Case "02"
            '若有規費金額
            If GetCP17(strReceiveNo) <> "0" Then
                ii = ii + 1
                strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                   "('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','規費','三、變更規費" & Format(GetCP17(strReceiveNo), "#,##0") & "元整。')"
            End If
        Case "06"
            '若有規費金額
            If GetCP17(strReceiveNo) <> "0" Then
                ii = ii + 1
                strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                   "('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','規費','三、變更規費" & Format(GetCP17(strReceiveNo), "#,##0") & "元整。')"
            End If
        End Select
    End Select
    '若有專利號數
    If Label3(7) <> "" Then
        ii = ii + 1
        strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                       "('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','列印備註','專利第" & Label3(7) & "號證書正本乙紙。')"
    '若無專利號數
    Else
        If Me.Text1.Text = "FCP" Then
            'Modify By Sindy 2018/3/8 原:說明書首頁一式三份
            ii = ii + 1
            strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                           "('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','列印備註','專利申請書首頁一份。')"
        End If
    End If
    
    'Add By Sindy 2018/7/25
    '今欲變更為「<變更事項檔-申請地址1中>[ <變更事項檔-申請地址2中>][ <變更事項檔-申請地址3中>][ <變更事項檔-申請地址4中>][ <變更事項檔-申請地址5中>]
    '」，代表人為「<變更事項檔-代表人1中><變更事項檔-代表人中譯文1>[ <變更事項檔-代表人2中><變更事項檔-代表人中譯文2>][ <變更事項檔-代表人3中><變更事項檔-代表人3中譯文>][ <變更事項檔-代表人4中><變更事項檔-代表人4中譯文>][ <變更事項檔-代表人5中><變更事項檔-代表人5中譯文>][ <變更事項檔-代表人6中><變更事項檔-代表人6中譯文>][ <變更事項檔-代表人7中><變更事項檔-代表人7中譯文>][ <變更事項檔-代表人8中><變更事項檔-代表人8中譯文>][ <變更事項檔-代表人9中><變更事項檔-代表人9中譯文>][ <變更事項檔-代表人10中><變更事項檔-代表人10中譯文>]
    '」，謹請　鈞局准予變更
    strExc(0) = "select * from changeevent where ce01='" & strReceiveNo & "'"
    intI = 1
    Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
    If intI = 1 Then
       strTmp = ""
       '申請人1中文、英文地址
       If "" & RsTemp.Fields("ce23") <> "" Then strTmp = strTmp & RsTemp.Fields("ce23")
       If "" & RsTemp.Fields("ce24") <> "" Then strTmp = strTmp & "（" & "" & RsTemp.Fields("ce24") & "）"
       '申請人2中文、英文地址
       If "" & RsTemp.Fields("ce26") <> "" Or "" & RsTemp.Fields("ce27") <> "" Then strTmp = strTmp & " "
       If "" & RsTemp.Fields("ce26") <> "" Then strTmp = strTmp & RsTemp.Fields("ce26")
       If "" & RsTemp.Fields("ce27") <> "" Then strTmp = strTmp & "（" & "" & RsTemp.Fields("ce27") & "）"
       '申請人3中文、英文地址
       If "" & RsTemp.Fields("ce29") <> "" Or "" & RsTemp.Fields("ce30") <> "" Then strTmp = strTmp & " "
       If "" & RsTemp.Fields("ce29") <> "" Then strTmp = strTmp & RsTemp.Fields("ce29")
       If "" & RsTemp.Fields("ce30") <> "" Then strTmp = strTmp & "（" & "" & RsTemp.Fields("ce30") & "）"
       '申請人4中文、英文地址
       If "" & RsTemp.Fields("ce32") <> "" Or "" & RsTemp.Fields("ce33") <> "" Then strTmp = strTmp & " "
       If "" & RsTemp.Fields("ce32") <> "" Then strTmp = strTmp & RsTemp.Fields("ce32")
       If "" & RsTemp.Fields("ce33") <> "" Then strTmp = strTmp & "（" & "" & RsTemp.Fields("ce33") & "）"
       '申請人5中文、英文地址
       If "" & RsTemp.Fields("ce35") <> "" Or "" & RsTemp.Fields("ce36") <> "" Then strTmp = strTmp & " "
       If "" & RsTemp.Fields("ce35") <> "" Then strTmp = strTmp & RsTemp.Fields("ce35")
       If "" & RsTemp.Fields("ce36") <> "" Then strTmp = strTmp & "（" & "" & RsTemp.Fields("ce36") & "）"
       ii = ii + 1
       strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                      "('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','變更後地址','" & strTmp & "')"
       strTmp = ""
       '代表人1中文、中譯文、英文
       If "" & RsTemp.Fields("ce10") <> "" Then strTmp = strTmp & RsTemp.Fields("ce10")
       If "" & RsTemp.Fields("ce63") <> "" Then strTmp = strTmp & RsTemp.Fields("ce63")
       If "" & RsTemp.Fields("ce11") <> "" Then strTmp = strTmp & "（" & "" & RsTemp.Fields("ce11") & "）"
       '代表人2中文、中譯文、英文
       If "" & RsTemp.Fields("ce13") <> "" Or "" & RsTemp.Fields("ce64") <> "" Or "" & RsTemp.Fields("ce14") <> "" Then strTmp = strTmp & " "
       If "" & RsTemp.Fields("ce13") <> "" Then strTmp = strTmp & RsTemp.Fields("ce13")
       If "" & RsTemp.Fields("ce64") <> "" Then strTmp = strTmp & RsTemp.Fields("ce64")
       If "" & RsTemp.Fields("ce14") <> "" Then strTmp = strTmp & "（" & "" & RsTemp.Fields("ce14") & "）"
       '代表人3中文、中譯文、英文
       If "" & RsTemp.Fields("ce68") <> "" Or "" & RsTemp.Fields("ce92") <> "" Or "" & RsTemp.Fields("ce69") <> "" Then strTmp = strTmp & " "
       If "" & RsTemp.Fields("ce68") <> "" Then strTmp = strTmp & RsTemp.Fields("ce68")
       If "" & RsTemp.Fields("ce92") <> "" Then strTmp = strTmp & RsTemp.Fields("ce92")
       If "" & RsTemp.Fields("ce69") <> "" Then strTmp = strTmp & "（" & "" & RsTemp.Fields("ce69") & "）"
       '代表人4中文、中譯文、英文
       If "" & RsTemp.Fields("ce71") <> "" Or "" & RsTemp.Fields("ce93") <> "" Or "" & RsTemp.Fields("ce72") <> "" Then strTmp = strTmp & " "
       If "" & RsTemp.Fields("ce71") <> "" Then strTmp = strTmp & RsTemp.Fields("ce71")
       If "" & RsTemp.Fields("ce93") <> "" Then strTmp = strTmp & RsTemp.Fields("ce93")
       If "" & RsTemp.Fields("ce72") <> "" Then strTmp = strTmp & "（" & "" & RsTemp.Fields("ce72") & "）"
       '代表人5中文、中譯文、英文
       If "" & RsTemp.Fields("ce74") <> "" Or "" & RsTemp.Fields("ce94") <> "" Or "" & RsTemp.Fields("ce75") <> "" Then strTmp = strTmp & " "
       If "" & RsTemp.Fields("ce74") <> "" Then strTmp = strTmp & RsTemp.Fields("ce74")
       If "" & RsTemp.Fields("ce94") <> "" Then strTmp = strTmp & RsTemp.Fields("ce94")
       If "" & RsTemp.Fields("ce75") <> "" Then strTmp = strTmp & "（" & "" & RsTemp.Fields("ce75") & "）"
       '代表人6中文、中譯文、英文
       If "" & RsTemp.Fields("ce77") <> "" Or "" & RsTemp.Fields("ce95") <> "" Or "" & RsTemp.Fields("ce78") <> "" Then strTmp = strTmp & " "
       If "" & RsTemp.Fields("ce77") <> "" Then strTmp = strTmp & RsTemp.Fields("ce77")
       If "" & RsTemp.Fields("ce95") <> "" Then strTmp = strTmp & RsTemp.Fields("ce95")
       If "" & RsTemp.Fields("ce78") <> "" Then strTmp = strTmp & "（" & "" & RsTemp.Fields("ce78") & "）"
       '代表人7中文、中譯文、英文
       If "" & RsTemp.Fields("ce80") <> "" Or "" & RsTemp.Fields("ce96") <> "" Or "" & RsTemp.Fields("ce81") <> "" Then strTmp = strTmp & " "
       If "" & RsTemp.Fields("ce80") <> "" Then strTmp = strTmp & RsTemp.Fields("ce80")
       If "" & RsTemp.Fields("ce96") <> "" Then strTmp = strTmp & RsTemp.Fields("ce96")
       If "" & RsTemp.Fields("ce81") <> "" Then strTmp = strTmp & "（" & "" & RsTemp.Fields("ce81") & "）"
       '代表人8中文、中譯文、英文
       If "" & RsTemp.Fields("ce83") <> "" Or "" & RsTemp.Fields("ce97") <> "" Or "" & RsTemp.Fields("ce84") <> "" Then strTmp = strTmp & " "
       If "" & RsTemp.Fields("ce83") <> "" Then strTmp = strTmp & RsTemp.Fields("ce83")
       If "" & RsTemp.Fields("ce97") <> "" Then strTmp = strTmp & RsTemp.Fields("ce97")
       If "" & RsTemp.Fields("ce84") <> "" Then strTmp = strTmp & "（" & "" & RsTemp.Fields("ce84") & "）"
       '代表人9中文、中譯文、英文
       If "" & RsTemp.Fields("ce86") <> "" Or "" & RsTemp.Fields("ce98") <> "" Or "" & RsTemp.Fields("ce87") <> "" Then strTmp = strTmp & " "
       If "" & RsTemp.Fields("ce86") <> "" Then strTmp = strTmp & RsTemp.Fields("ce86")
       If "" & RsTemp.Fields("ce98") <> "" Then strTmp = strTmp & RsTemp.Fields("ce98")
       If "" & RsTemp.Fields("ce87") <> "" Then strTmp = strTmp & "（" & "" & RsTemp.Fields("ce87") & "）"
       '代表人10中文、中譯文、英文
       If "" & RsTemp.Fields("ce89") <> "" Or "" & RsTemp.Fields("ce99") <> "" Or "" & RsTemp.Fields("ce90") <> "" Then strTmp = strTmp & " "
       If "" & RsTemp.Fields("ce89") <> "" Then strTmp = strTmp & RsTemp.Fields("ce89")
       If "" & RsTemp.Fields("ce99") <> "" Then strTmp = strTmp & RsTemp.Fields("ce99")
       If "" & RsTemp.Fields("ce90") <> "" Then strTmp = strTmp & "（" & "" & RsTemp.Fields("ce90") & "）"
       ii = ii + 1
       strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                      "('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','變更後代表人','" & strTmp & "')"
    End If
    '2018/7/25 END
    
    'edit by nickc 2007/02/05 不用 dll 了
    'If Not objLawDll.ExecSQL(ii, strTxt) Then
    If Not ClsLawExecSQL(ii, strTxt) Then
       MsgBox "儲存例外欄位失敗，請洽系統管理員 !", vbCritical
    End If
End Sub

'申請書
Private Function StartLetter2(ByVal ET01 As String, ByVal ET03 As String) As Boolean
Dim strTxt(200) As String, strTmp As String, strTemp1 As String, strTemp2 As String
Dim ii As Integer, jj As Integer
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
Dim nIndex As Integer, intSeqno As Integer
'Dim bolFee As Boolean 'Add By Sindy 2019/10/7
   
   ii = 0
   EndLetter ET01, strReceiveNo, ET03, strUserNum
   
   ii = ii + 1
   strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
      " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','本所案號','" & m_CaseNo & "')"
   
   'Modify By Sindy 2018/10/19 判斷是否有變更申請人,若有要傳入其資料,讀資料用 + , m_strApplNum
   Call PUB_GetApplPA_EData(ET01, ET03, strReceiveNo, pa(), , m_strApplNum)

   '出名代理人
'   strExc(0) = "select oa05,ST26,st02 from ouragent,staff where oa01='" & pa(1) & "' and instr('" & cp(110) & "',oa02)>0 and st01(+)=oa02 order by OA03"
'   intI = 1
'   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'   If intI = 1 Then
'      With RsTemp
'      jj = 1
'      Do While Not .EOF
'         ii = ii + 1
'         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'            " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','代理人" & jj & "-中文姓名','" & PUB_ConvertNameFormat("" & .Fields("st02")) & "')"
'         jj = jj + 1
'         .MoveNext
'      Loop
'      End With
'   End If
   'Modify By Sindy 2020/4/7 申請書:出名代理人
   Call PUB_ReadPToAppBaseData(pa(1), pa(2), pa(3), pa(4), 1, cp(110), ET01, strReceiveNo, ET03, ii, strTxt())
   
'   '變更事項:
'   'Modify By Sindy 2019/10/7 僅只有變更申請人地址,代表人中譯文,代表人時,無規費
'   Dim objCheck As Object
'   bolFee = False '2,4,3
'   For Each objCheck In Check1
'      If objCheck.Index <> 2 And objCheck.Index <> 4 And objCheck.Index <> 3 Then
'         If objCheck.Value = 1 Then
'            bolFee = True
'            Exit For
'         End If
'      End If
'   Next
'   '2019/10/7 END
   
   '變更申請人地址
   If Check1(2).Value = 1 Then
      strTmp = ""
      'Modify By Sindy 2018/9/17 + ChgSQL
      '申請人1
      If IsEmptyText(Text8(0)) = False Or IsEmptyText(Text8(5)) = False Then
         strTmp = strTmp & "申請人1原「" & pa(31) & IIf(pa(36) = "", "", "(" & ChgSQL(pa(36)) & ")") & "」變更為「" & Text8(0) & IIf(Text8(5) = "", "", "(" & ChgSQL(Text8(5)) & ")") & "」"
      End If
      '申請人2
      If IsEmptyText(Text8(1)) = False Or IsEmptyText(Text8(6)) = False Then
         strTmp = strTmp & "申請人2原「" & pa(32) & IIf(pa(37) = "", "", "(" & ChgSQL(pa(37)) & ")") & "」變更為「" & Text8(1) & IIf(Text8(6) = "", "", "(" & ChgSQL(Text8(6)) & ")") & "」"
      End If
      '申請人3
      If IsEmptyText(Text8(2)) = False Or IsEmptyText(Text8(7)) = False Then
         strTmp = strTmp & "申請人3原「" & pa(33) & IIf(pa(38) = "", "", "(" & ChgSQL(pa(38)) & ")") & "」變更為「" & Text8(2) & IIf(Text8(7) = "", "", "(" & ChgSQL(Text8(7)) & ")") & "」"
      End If
      '申請人4
      If IsEmptyText(Text8(3)) = False Or IsEmptyText(Text8(8)) = False Then
         strTmp = strTmp & "申請人4原「" & pa(34) & IIf(pa(39) = "", "", "(" & ChgSQL(pa(39)) & ")") & "」變更為「" & Text8(3) & IIf(Text8(8) = "", "", "(" & ChgSQL(Text8(8)) & ")") & "」"
      End If
      '申請人5
      If IsEmptyText(Text8(4)) = False Or IsEmptyText(Text8(9)) = False Then
         strTmp = strTmp & "申請人5原「" & pa(35) & IIf(pa(40) = "", "", "(" & ChgSQL(pa(40)) & ")") & "」變更為「" & Text8(4) & IIf(Text8(9) = "", "", "(" & ChgSQL(Text8(9)) & ")") & "」"
      End If
      If strTmp = "" Then strTmp = "原「　　」變更為「　　」"
      ii = ii + 1
      'Modified by Lydia 2022/10/24 +ChgSQL
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','變更申請人地址','" & ChgSQL(strTmp) & "')"
   End If
   '變更公司代表人
   If Check1(3).Value = 1 Then
      strTmp = ""
      'Modify By Sindy 2021/8/2 敏莉,Landy:日文名稱都不出現
      '代表人1
      If IsEmptyText(Text9(0)) = False Or IsEmptyText(Text9(1)) = False Or IsEmptyText(Text9(2)) = False Or _
         IsEmptyText(pa(79)) = False Or IsEmptyText(pa(80)) = False Or IsEmptyText(pa(81)) = False Then
         'strTmp = strTmp & "代表人1原「" & pa(79) & IIf(pa(80) = "", "", "(" & pa(80) & ")") & IIf(pa(81) = "", "", "(" & pa(81) & ")") & "」變更為「" & Text9(0) & IIf(Text9(1) = "", "", "(" & Text9(1) & ")") & IIf(Text9(2) = "", "", "(" & Text9(2) & ")") & "」"
         strTmp = strTmp & "代表人1原「" & pa(79) & IIf(pa(80) = "", "", "(" & pa(80) & ")") & "」變更為「" & Text9(0) & IIf(Text9(1) = "", "", "(" & Text9(1) & ")") & "」"
      End If
      '代表人2
      If IsEmptyText(Text9(3)) = False Or IsEmptyText(Text9(4)) = False Or IsEmptyText(Text9(5)) = False Or _
         IsEmptyText(pa(82)) = False Or IsEmptyText(pa(83)) = False Or IsEmptyText(pa(84)) = False Then
         'strTmp = strTmp & "代表人2原「" & pa(82) & IIf(pa(83) = "", "", "(" & pa(83) & ")") & IIf(pa(84) = "", "", "(" & pa(84) & ")") & "」變更為「" & Text9(3) & IIf(Text9(4) = "", "", "(" & Text9(4) & ")") & IIf(Text9(5) = "", "", "(" & Text9(5) & ")") & "」"
         strTmp = strTmp & "代表人2原「" & pa(82) & IIf(pa(83) = "", "", "(" & pa(83) & ")") & "」變更為「" & Text9(3) & IIf(Text9(4) = "", "", "(" & Text9(4) & ")") & "」"
      End If
      '代表人3
      If IsEmptyText(Text9(6)) = False Or IsEmptyText(Text9(7)) = False Or IsEmptyText(Text9(8)) = False Or _
         IsEmptyText(pa(109)) = False Or IsEmptyText(pa(110)) = False Or IsEmptyText(pa(111)) = False Then
         'strTmp = strTmp & "代表人3原「" & pa(109) & IIf(pa(110) = "", "", "(" & pa(110) & ")") & IIf(pa(111) = "", "", "(" & pa(111) & ")") & "」變更為「" & Text9(6) & IIf(Text9(7) = "", "", "(" & Text9(7) & ")") & IIf(Text9(8) = "", "", "(" & Text9(8) & ")") & "」"
         strTmp = strTmp & "代表人3原「" & pa(109) & IIf(pa(110) = "", "", "(" & pa(110) & ")") & "」變更為「" & Text9(6) & IIf(Text9(7) = "", "", "(" & Text9(7) & ")") & "」"
      End If
      '代表人4
      If IsEmptyText(Text9(9)) = False Or IsEmptyText(Text9(10)) = False Or IsEmptyText(Text9(11)) = False Or _
         IsEmptyText(pa(112)) = False Or IsEmptyText(pa(113)) = False Or IsEmptyText(pa(114)) = False Then
         'strTmp = strTmp & "代表人4原「" & pa(112) & IIf(pa(113) = "", "", "(" & pa(113) & ")") & IIf(pa(114) = "", "", "(" & pa(114) & ")") & "」變更為「" & Text9(9) & IIf(Text9(10) = "", "", "(" & Text9(10) & ")") & IIf(Text9(11) = "", "", "(" & Text9(11) & ")") & "」"
         strTmp = strTmp & "代表人4原「" & pa(112) & IIf(pa(113) = "", "", "(" & pa(113) & ")") & "」變更為「" & Text9(9) & IIf(Text9(10) = "", "", "(" & Text9(10) & ")") & "」"
      End If
      '代表人5
      If IsEmptyText(Text9(12)) = False Or IsEmptyText(Text9(13)) = False Or IsEmptyText(Text9(14)) = False Or _
         IsEmptyText(pa(115)) = False Or IsEmptyText(pa(116)) = False Or IsEmptyText(pa(117)) = False Then
         strTmp = strTmp & "代表人5原「" & pa(115) & IIf(pa(116) = "", "", "(" & pa(116) & ")") & "」變更為「" & Text9(12) & IIf(Text9(13) = "", "", "(" & Text9(13) & ")") & "」"
      End If
      '代表人6
      If IsEmptyText(Text9(15)) = False Or IsEmptyText(Text9(16)) = False Or IsEmptyText(Text9(17)) = False Or _
         IsEmptyText(pa(118)) = False Or IsEmptyText(pa(119)) = False Or IsEmptyText(pa(120)) = False Then
         'strTmp = strTmp & "代表人6原「" & pa(118) & IIf(pa(119) = "", "", "(" & pa(119) & ")") & IIf(pa(120) = "", "", "(" & pa(120) & ")") & "」變更為「" & Text9(15) & IIf(Text9(16) = "", "", "(" & Text9(16) & ")") & IIf(Text9(17) = "", "", "(" & Text9(17) & ")") & "」"
         strTmp = strTmp & "代表人6原「" & pa(118) & IIf(pa(119) = "", "", "(" & pa(119) & ")") & "」變更為「" & Text9(15) & IIf(Text9(16) = "", "", "(" & Text9(16) & ")") & "」"
      End If
      '代表人7
      If IsEmptyText(Text9(18)) = False Or IsEmptyText(Text9(19)) = False Or IsEmptyText(Text9(20)) = False Or _
         IsEmptyText(pa(121)) = False Or IsEmptyText(pa(122)) = False Or IsEmptyText(pa(123)) = False Then
         'strTmp = strTmp & "代表人7原「" & pa(121) & IIf(pa(122) = "", "", "(" & pa(122) & ")") & IIf(pa(123) = "", "", "(" & pa(123) & ")") & "」變更為「" & Text9(18) & IIf(Text9(19) = "", "", "(" & Text9(19) & ")") & IIf(Text9(20) = "", "", "(" & Text9(20) & ")") & "」"
         strTmp = strTmp & "代表人7原「" & pa(121) & IIf(pa(122) = "", "", "(" & pa(122) & ")") & "」變更為「" & Text9(18) & IIf(Text9(19) = "", "", "(" & Text9(19) & ")") & "」"
      End If
      '代表人8
      If IsEmptyText(Text9(21)) = False Or IsEmptyText(Text9(22)) = False Or IsEmptyText(Text9(23)) = False Or _
         IsEmptyText(pa(124)) = False Or IsEmptyText(pa(125)) = False Or IsEmptyText(pa(126)) = False Then
         strTmp = strTmp & "代表人8原「" & pa(124) & IIf(pa(125) = "", "", "(" & pa(125) & ")") & "」變更為「" & Text9(21) & IIf(Text9(22) = "", "", "(" & Text9(22) & ")") & "」"
      End If
      '代表人9
      If IsEmptyText(Text9(24)) = False Or IsEmptyText(Text9(25)) = False Or IsEmptyText(Text9(26)) = False Or _
         IsEmptyText(pa(127)) = False Or IsEmptyText(pa(128)) = False Or IsEmptyText(pa(129)) = False Then
         'strTmp = strTmp & "代表人9原「" & pa(127) & IIf(pa(128) = "", "", "(" & pa(128) & ")") & IIf(pa(129) = "", "", "(" & pa(129) & ")") & "」變更為「" & Text9(24) & IIf(Text9(25) = "", "", "(" & Text9(25) & ")") & IIf(Text9(26) = "", "", "(" & Text9(26) & ")") & "」"
         strTmp = strTmp & "代表人9原「" & pa(127) & IIf(pa(128) = "", "", "(" & pa(128) & ")") & "」變更為「" & Text9(24) & IIf(Text9(25) = "", "", "(" & Text9(25) & ")") & "」"
      End If
      '代表人10
      If IsEmptyText(Text9(27)) = False Or IsEmptyText(Text9(28)) = False Or IsEmptyText(Text9(29)) = False Or _
         IsEmptyText(pa(130)) = False Or IsEmptyText(pa(131)) = False Or IsEmptyText(pa(132)) = False Then
         'strTmp = strTmp & "代表人10原「" & pa(130) & IIf(pa(131) = "", "", "(" & pa(131) & ")") & IIf(pa(132) = "", "", "(" & pa(132) & ")") & "」變更為「" & Text9(27) & IIf(Text9(28) = "", "", "(" & Text9(28) & ")") & IIf(Text9(29) = "", "", "(" & Text9(29) & ")") & "」"
         strTmp = strTmp & "代表人10原「" & pa(130) & IIf(pa(131) = "", "", "(" & pa(131) & ")") & "」變更為「" & Text9(27) & IIf(Text9(28) = "", "", "(" & Text9(28) & ")") & "」"
      End If
      If strTmp = "" Then strTmp = "原「　　」變更為「　　」"
      ii = ii + 1
      'Modified by Lydia 2022/10/24 +ChgSQL
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','變更公司代表人','" & ChgSQL(strTmp) & "')"
   End If
   
   '變更申請人姓名
   If Check1(0).Value = 1 Or Check1(1).Value = 1 Then
      strTmp = ""
      If Check1(0).Value = 1 Then
         For nIndex = 0 To 4
            If IsEmptyText(Text7(nIndex)) = False Then
               'Modify By Sindy 2018/10/19 加顯示英文名稱
               '變更前
               'Modify By Sindy 2021/7/29 加外商國名 => , , True
               strTemp1 = GetPrjPeople1(ChangeCustomerL(pa(26 + nIndex)), , True) '中
               strTemp2 = GetPrjPeople1(ChangeCustomerL(pa(26 + nIndex)), 2) '英
               strTmp = strTmp & "申請人" & nIndex + 1 & "原「" & strTemp1 & IIf(strTemp1 <> strTemp2 And strTemp2 <> "", "(" & strTemp2 & ")", "") & "」"
               '變更後
               'Modify By Sindy 2021/7/29 加外商國名 => , , True
               strTemp1 = GetPrjPeople1(Text7(nIndex).Text, , True) '中
               strTemp2 = GetPrjPeople1(Text7(nIndex).Text, 2) '英
               strTmp = strTmp & "變更為「" & strTemp1 & IIf(strTemp1 <> strTemp2 And strTemp2 <> "", "(" & strTemp2 & ")", "") & "」"
               '2018/10/19 END
            End If
         Next nIndex
      End If
      If Check1(1).Value = 1 Then
         For nIndex = 0 To 4
            If IsEmptyText(Text6(nIndex)) = False Then
               'Modify By Sindy 2021/7/29 加外商國名 => , , True
               strTmp = strTmp & "申請人" & nIndex + 1 & "原「" & GetPrjPeople1(ChangeCustomerL(pa(26 + nIndex)), , True) & "」變更為「" & Text6(nIndex).Text & "」"
            End If
         Next nIndex
      End If
      If strTmp = "" Then strTmp = "原「　　」變更為「　　」"
      ii = ii + 1
      'Modified by Lydia 2022/10/24 +ChgSQL
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','變更申請人姓名','" & ChgSQL(strTmp) & "')"
   End If
   '變更專利名稱
   If Check1(7).Value = 1 Then
      strTmp = ""
      '中文
      If Text38(0).Text <> "" Then
         strTmp = strTmp & "專利中文名稱原「" & pa(5) & "」變更為「" & Text38(0).Text & "」"
      End If
      '英文
      If Text38(1).Text <> "" Then
         strTmp = strTmp & "專利英文名稱原「" & pa(6) & "」變更為「" & Text38(1).Text & "」"
      End If
      '日文
      If Text38(2).Text <> "" Then
         strTmp = strTmp & "專利日文名稱原「" & pa(7) & "」變更為「" & Text38(2).Text & "」"
      End If
      If strTmp = "" Then strTmp = "原「　　」變更為「　　」"
      ii = ii + 1
      'Modified by Lydia 2022/10/24 +ChgSQL
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','變更專利名稱','" & ChgSQL(strTmp) & "')"
   End If
   '變更代理人
   If Check1(10).Value = 1 Then
      strTmp = "": strExc(9) = "": strExc(10) = ""
      '舊的出名代理人
      strExc(0) = "select cp05,cp09,cp110 from caseprogress" & _
                  " where cp01='" & pa(1) & "' and cp02='" & pa(2) & "' and cp03='" & pa(3) & "' and cp04='" & pa(4) & "'" & _
                  " and cp110 is not null and cp09<>'" & cp(9) & "'" & _
                  " order by cp66 desc,cp67 desc"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
'         strExc(0) = "select oa05,ST26,st02 from ouragent,staff where oa01='" & pa(1) & "' and instr('" & RsTemp.Fields("cp110") & "',oa02)>0 and st01(+)=oa02 order by OA03"
'         intI = 1
'         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'         If intI = 1 Then
'            RsTemp.MoveFirst
'            With RsTemp
'            Do While Not .EOF
'               'Modify By Sindy 2020/4/6
'               '專利處
'               If Left(Pub_StrUserSt03, 2) = "P1" Then
'                  strExc(9) = strExc(9) & "、" & .Fields("st02")
'               Else
'               '2020/4/6 END
'                  strExc(9) = strExc(9) & "、" & PUB_ConvertNameFormat("" & .Fields("st02"))
'               End If
'               RsTemp.MoveNext
'            Loop
'            End With
'            If strExc(9) <> "" Then strExc(9) = Mid(strExc(9), 2)
'         End If
         'Modify By Sindy 2020/4/7 舊的出名代理人
         strExc(9) = PUB_ReadPToAppBaseData(pa(1), pa(2), pa(3), pa(4), 0, "" & RsTemp.Fields("cp110"))
      End If
      '新的出名代理人
'      strExc(0) = "select oa05,ST26,st02 from ouragent,staff where oa01='" & pa(1) & "' and instr('" & cp(110) & "',oa02)>0 and st01(+)=oa02 order by OA03"
'      intI = 1
'      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'      If intI = 1 Then
'         RsTemp.MoveFirst
'         With RsTemp
'         Do While Not .EOF
'            'Modify By Sindy 2020/4/6
'            '專利處
'            If Left(Pub_StrUserSt03, 2) = "P1" Then
'               strExc(10) = strExc(10) & "、" & .Fields("st02")
'            Else
'            '2020/4/6 END
'               strExc(10) = strExc(10) & "、" & PUB_ConvertNameFormat("" & .Fields("st02"))
'            End If
'            RsTemp.MoveNext
'         Loop
'         End With
'         If strExc(10) <> "" Then strExc(10) = Mid(strExc(10), 2)
'      End If
      'Modify By Sindy 2020/4/7 新的出名代理人
      strExc(10) = PUB_ReadPToAppBaseData(pa(1), pa(2), pa(3), pa(4), 0, cp(110))
      
      strTmp = "原「" & IIf(strExc(9) <> "", strExc(9), "　　") & "」變更為「" & IIf(strExc(10) <> "", strExc(10), "　　") & "」"
      ii = ii + 1
      'Modified by Lydia 2022/10/24 +ChgSQL
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','變更代理人','" & ChgSQL(strTmp) & "')"
   End If
   '變更其他
   If Check1(11).Value = 1 Then
      ii = ii + 1
       'Modified by Lydia 2022/10/24 +ChgSQL
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','變更其他','原「　　」變更為「" & ChgSQL(Text50.Text) & "」')"
   End If
   
   'Add By Sindy 2018/9/17
   '變更發明人或創作人或設計人之國籍
   If Check1(17).Value = 1 Then
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','變更發明人或創作人或設計人之國籍','♀')"
   End If
   '變更發明人或創作人或設計人之姓名
   If Check1(18).Value = 1 Then
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','變更發明人或創作人或設計人之姓名','♀')"
   End If
   '追加刪除或更正發明人或創作人或設計人
   If Check1(19).Value = 1 Then
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','追加刪除或更正發明人或創作人或設計人','♀')"
   End If
   '2018/9/17
   'Add By Sindy 2023/2/20
   If Check1(20).Value = 1 Then
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','變更申請人國籍','♀')"
   End If
   '2023/2/20 END
   
'   'Modify By Sindy 2019/10/7
'   '僅只有變更申請人地址,代表人中譯文,代表人時,無規費
'   If bolFee = False Then txtCP84 = 0
'   '2019/10/7 END
   ii = ii + 1
   strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
      " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','繳費金額','" & Val(txtCP84) & "')"
   
   ii = ii + 1
   strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
      " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','附件-基本資料表','" & m_CaseNo & ".contact.pdf')"
   
   If Not ClsLawExecSQL(ii, strTxt) Then
      MsgBox "儲存例外欄位失敗，請洽系統管理員 !", vbCritical
   Else
      StartLetter2 = True
   End If
End Function

Private Sub Check1_Click(Index As Integer)
Dim i As Integer
   initCheck Index
   Select Case Index
      Case 1
         If Check1(Index).Value = 1 Then
            For i = 0 To 4
               If Text7(i) <> "" Then
                  'edit by nickc 2007/02/02 不用 dll 了
                  'If objPublicData.GetCustomerNameAndAddress(Text7(i).Text, strExc(1), strExc(2), strExc(3), strExc(4)) Then
                  If ClsPDGetCustomerNameAndAddress(Text7(i).Text, strExc(1), strExc(2), strExc(3), strExc(4)) Then
                     Text6(i) = strExc(1)
                  Else
                     Text6(i) = ""
                  End If
               'Modify By Sindy 2018/5/10 Mark
'               Else
'                  Text6(i) = ""
               End If
            Next
         End If
      Case 2
         'Modify By Sindy 2018/5/10
         'If Check1(0).Value = 1 And Check1(Index).Value = 1 Then
         If Check1(Index).Value = 1 Then
         '2018/5/10 END
            For i = 0 To 4
               If Text7(i) <> "" Then
                  'edit by nickc 2007/02/02 不用 dll 了
                  'If objPublicData.GetCustomerNameAndAddress(Text7(i).Text, strExc(1), strExc(2), strExc(3), strExc(4)) Then
                  If ClsPDGetCustomerNameAndAddress(Text7(i).Text, strExc(1), strExc(2), strExc(3), strExc(4)) Then
                     Text8(i) = strExc(2)
                     Text8(i + 5) = strExc(3) 'Add By Sindy 2018/5/25
                  Else
                     Text8(i) = ""
                     Text8(i + 5) = "" 'Add By Sindy 2018/5/25
                  End If
               'Modify By Sindy 2018/5/10 Mark
'               Else
'                  Text8(i) = ""
               End If
            Next
         End If
      Case 3
         If Check1(Index).Value = 1 Then
            IntoCombol
         End If
      'Add By Sindy 2018/9/17 須繳納變更規費新台幣300元
      Case 18, 19
'         If Check1(18).Value = 1 Or Check1(19).Value = 1 Then
'            Label12.Visible = True
'         Else
'            Label12.Visible = False
'         End If
         '2018/9/17 END
   End Select
End Sub

Private Sub initCheck(ByVal iSitu As Integer)
 Dim i As Integer
   Select Case iSitu
      Case 0
         If Check1(0).Value = 1 Then
            For i = 3 To 7
               Text7(i - 3).Enabled = True
            Next
         Else
            For i = 3 To 7
               'Add By Sindy 2018/5/25
               'Text7(i - 3).Enabled = False
               Text7(i - 3).Enabled = True
               '2018/5/25 END
               Text7(i - 3).Text = ""
               Label4(i - 3).Caption = "" 'Add By Sindy 2018/5/25
            Next
         End If
      Case 1
         If Check1(1).Value = 1 Then
            For i = 16 To 20
               Text6(i - 16).Enabled = True
            Next
         Else
            For i = 16 To 20
               Text6(i - 16).Enabled = False
               Text6(i - 16).Text = ""
            Next
         End If
      Case 2
         If Check1(2).Value = 1 Then
            For i = 0 To 9 '4
               Text8(i).Enabled = True
            Next
         Else
            For i = 0 To 9 '4
               Text8(i).Enabled = False
               Text8(i).Text = ""
            Next
         End If
      Case 3
         If Check1(3).Value = 1 Then
            For i = 0 To 9
               Combo2(i).Enabled = True
            Next
            For i = 0 To 29
               Text9(i).Enabled = True
            Next
         Else
            For i = 0 To 9
               Combo2(i).Clear
               Combo2(i).Enabled = False
            Next
            For i = 0 To 29
               Text9(i).Enabled = False
               Text9(i).Text = ""
            Next
         End If
      Case 4
         If Check1(4).Value = 1 Then
            For i = 0 To 9
               Text10(i).Enabled = True
            Next
         Else
            For i = 0 To 9
               Text10(i).Enabled = False
               Text10(i).Text = ""
            Next
         End If
      Case 5
         If Check1(5).Value = 1 Then
            Text34.Enabled = True
         Else
            Text34.Enabled = False
            Text34.Text = ""
         End If
      Case 6
         If Check1(6).Value = 1 Then
            Text36.Enabled = True
         Else
            Text36.Enabled = False
            Text36.Text = ""
         End If
      Case 7
         If Check1(7).Value = 1 Then
            For i = 0 To 2
               Text38(i).Enabled = True
            Next
         Else
            For i = 0 To 2
               Text38(i).Enabled = False
               Text38(i).Text = ""
            Next
         End If
      Case 8
         If Check1(8).Value = 1 Then
            Text44.Enabled = True
         Else
            Text44.Enabled = False
            Text44.Text = ""
         End If
      Case 9
         If Check1(9).Value = 1 Then
            Text46.Enabled = True
         Else
            Text46.Enabled = False
            Text46.Text = ""
         End If
      Case 10
         If Check1(10).Value = 1 Then
            Text48.Enabled = True
         Else
            Text48.Enabled = False
            Text48.Text = ""
         End If
      Case 11
         If Check1(11).Value = 1 Then
            Text50.Enabled = True
         Else
            Text50.Enabled = False
            Text50.Text = ""
         End If
      Case 12
         If Check1(12).Value = 1 Then
            Text52.Enabled = True
         Else
            Text52.Enabled = False
            Text52.Text = ""
         End If
      Case 13
         If Check1(13).Value = 1 Then
            Text54.Enabled = True
         Else
            Text54.Enabled = False
            Text54.Text = ""
         End If
      Case 14
         If Check1(14).Value = 1 Then
            Text56.Enabled = True
         Else
            Text56.Enabled = False
            Text56.Text = ""
         End If
      Case 15
         If Check1(15).Value = 1 Then
            Text58.Enabled = True
         Else
            Text58.Enabled = False
            Text58.Text = ""
         End If
      Case 16
         If Check1(16).Value = 1 Then
            Text60.Enabled = True
         Else
            Text60.Enabled = False
            Text60.Text = ""
         End If
   End Select
End Sub

Private Sub cmdok_Click(Index As Integer)
Dim bolChk As Boolean, strTmp As String
Dim bPrt As Boolean
Dim strFolder As String, strFileName As String 'Add By Sindy 2018/7/4
Dim nIndex As Integer
Dim bolFee As Boolean
   
   If Index = 0 Then
      ' 90.07.18 modify by louis (申請人補滿九碼)
      For nIndex = 0 To 4
         If IsEmptyText(Text7(nIndex)) = False Then
            Text7(nIndex) = Text7(nIndex) & String(9 - Len(Text7(nIndex)), "0")
         End If
      Next nIndex
      
      'Add By Cheng 2002/05/22
      '重新檢查欄位有效性
      If TxtValidate = False Then Exit Sub
      'Added by Lydia 2020/02/21 產生各式申請書時，若基本檔「名稱有特殊字」已勾選，彈訊息提醒，並一併開啟原始檔。
      If (pa(1) = "FCP" Or pa(1) = "P") And pa(174) = "Y" Then
          MsgBox MsgText(1111), vbInformation
          If Pub_GetPA174toFile("0", pa(1), pa(2), pa(3), pa(4), Me, frm100101_M_1) = False Then
              Exit Sub
          End If
      End If
      'end 2020/02/21
      
      'Modify By Sindy 2023/7/19 移到存檔前檢查
      '變更事項:
      'Modify By Sindy 2019/10/7 僅只有變更申請人地址,代表人中譯文,代表人時,無規費
      Dim objCheck As Object
      bolFee = False '2,4,3
      For Each objCheck In Check1
         If objCheck.Index <> 2 And objCheck.Index <> 4 And objCheck.Index <> 3 Then
            If objCheck.Value = 1 Then
               bolFee = True
               Exit For
            End If
         End If
      Next
      '2019/10/7 END
      If bolFee = False Then
         If Check1(2).Value = 1 Or Check1(4).Value = 1 Or Check1(3).Value = 1 Then
            txtCP84 = 0
         End If
      End If
      '2023/7/19 END
      
      'Add by Sindy 2021/11/4 檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
      If PUB_ChkUniText(Me) = False Then
         Exit Sub
      End If
      '2021/10/8 END
      
      If FormSave = False Then MsgBox "存檔失敗，請洽系統管理員 !", vbCritical: Exit Sub
      bolSaveOK = True 'Add by Morgan 2006/7/4
      oParent.Tag = True 'Added by Morgan 2012/4/11
      
      'Add By Sindy 2018/8/10 電子送件申請書
      If m_CP118isY = "Y" Then
         m_CaseNo = PUB_FCPCaseNo2FileName(pa(1), pa(2), pa(3), pa(4))
         'Modify By Sindy 2019/1/17
         If Pub_StrUserSt15 = "P12" Then
            strFolder = PUB_Getdesktop
            strFolder = strFolder & "\" & m_CaseNo
            If Dir(strFolder, vbDirectory) = "" Then
               MkDir strFolder
            End If
         Else
         '2019/1/17 END
            If UCase(pub_DbTerminalName) <> UCase(正式資料庫電腦名稱) Or Pub_StrUserSt03 = "M51" Then
               strFolder = PUB_Getdesktop
            Else
               strFolder = FCP電子送件檔案存放路徑
            End If
            strFolder = strFolder & "\" & m_CaseNo
            If Dir(strFolder, vbDirectory) = "" Then
               MkDir strFolder
            End If
         End If
         
         '1.基本資料
         'Modify By Sindy 2018/10/19 判斷是否有變更申請人,若有,要傳入其資料,讀資料用
         If Check1(0).Value = 1 Then
            m_strApplNum = Text7(0) & "@" & Text7(1) & "@" & Text7(2) & "@" & Text7(3) & "@" & Text7(4) & "@"
         Else
            m_strApplNum = ""
         End If
         'Modify By Sindy 2018/10/19 判斷是否有變更代表人,若有,要傳入其資料,讀資料用
         If Check1(3).Value = 1 Then
            m_Representative = Text9(0) & "@" & Text9(1) & "@" & Text9(2) & "@" & _
                               Text9(3) & "@" & Text9(4) & "@" & Text9(5) & "@" & _
                               Text9(6) & "@" & Text9(7) & "@" & Text9(8) & "@" & _
                               Text9(9) & "@" & Text9(10) & "@" & Text9(11) & "@" & _
                               Text9(12) & "@" & Text9(13) & "@" & Text9(14) & "@" & _
                               Text9(15) & "@" & Text9(16) & "@" & Text9(17) & "@" & _
                               Text9(18) & "@" & Text9(19) & "@" & Text9(20) & "@" & _
                               Text9(21) & "@" & Text9(22) & "@" & Text9(23) & "@" & _
                               Text9(24) & "@" & Text9(25) & "@" & Text9(26) & "@" & _
                               Text9(27) & "@" & Text9(28) & "@" & Text9(29) & "@"
         Else
            m_Representative = ""
         End If
         'Modify By Sindy 2019/2/21
         If Pub_StrUserSt15 = "P12" Then
            '2.申請書
            If StartLetter2("01", "22") = False Then Exit Sub
            NowPrint strReceiveNo, "01", "22", False, strUserNum, , , True, strExc(9)
            strFileName = strFolder & "\" & m_CaseNo & ".data"
            'Call PUB_MakeDoc(strExc(9), strFileName)
            '1.基本資料
            'Modify By Sindy 2019/1/30 P121395.contact要顯示發明人資料
            'Modify By Sindy 2020/3/11 IIf(Check1(2).Value = 1, True, False) => True : 要抓客戶檔資料 ex:P-123217 代表人抓公司負責人
            StartLetterPA_EData "01", "14", strReceiveNo, pa, cp, True, True, m_strApplNum, m_Representative
            NowPrint strReceiveNo, "01", "14", False, strUserNum, , , True, strExc(10)
            'strFileName = strFolder & "\" & m_CaseNo & ".contact"
            Call PUB_MakeDoc(strExc(9) & Chr(12) & strExc(10), strFileName, False)
         Else
         '2019/2/21 END
            'Modify By Sindy 2018/9/5 若變更地址時,則抓客戶檔地址 + IIf(Check1(2).Value = 1, True, False)
            'Modify By Sindy 2018/10/19 判斷是否有變更申請人,若有,要傳入其資料,讀資料用 + , m_strApplNum
            '                           判斷是否有變更代表人,若有,要傳入其資料,讀資料用 + , m_Representative
            'Modify By Sindy 2019/3/20 目前產生變更的電子送件申請書的"基本資料表"會將發明人資訊帶出，請修改程式使其不會將發明人資訊帶出
            StartLetterPA_EData "01", "14", strReceiveNo, pa, cp, False, IIf(Check1(2).Value = 1, True, False), m_strApplNum, m_Representative
            NowPrint strReceiveNo, "01", "14", False, strUserNum, , , True, strExc(9)
            strFileName = strFolder & "\" & m_CaseNo & ".contact"
            Call PUB_MakeDoc(strExc(9), strFileName)
            '2.申請書
            If StartLetter2("01", "22") = False Then Exit Sub
            NowPrint strReceiveNo, "01", "22", False, strUserNum, , , True, strExc(9)
            strFileName = strFolder & "\" & "專利申請案變更事項申請書"
            Call PUB_MakeDoc(strExc(9), strFileName)
         End If
         
      Else
      '2018/8/10 END
         strLetterDate = Text5.Text
         ' 90.07.18 modify by louis
         bPrt = False
         If Text41 <> "N" Then
            If Text62 = "Y" Then
               bolChk = True
            Else
               bolChk = False
            End If
            
            Select Case intWhere
               Case 國內
                  If Check1(10).Value = 1 And bPrt = False Then
                     '變更代理人 01
                     strTmp = "01"
                     bPrt = True
                  End If
                  If Check1(3).Value = 1 And bPrt = False Then
                     '變更代表人 02
                     strTmp = "02"
                     bPrt = True
                  End If
                   'Add By Cheng 2003/01/21
                   '變更地址
                   If Check1(2).Value = vbChecked And bPrt = False Then
                     '變更地址 06
                     strTmp = "06"
                     bPrt = True
                   End If
                  If Check1(8).Value = 1 And bPrt = False Then
                     strExc(0) = "SELECT CU15 FROM CUSTOMER WHERE (CU01,CU02) IN " & _
                        "(SELECT SUBSTR(PA26,1,8),SUBSTR(PA26,9,1) FROM PATENT WHERE " & ChgPatent(pa(1) & pa(2) & pa(3) & pa(4)) & ")"
                     intI = 1
                     Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
                     If intI = 1 Then
   '                     If RsTemp.Fields(0) = "0" Then
   '                        '變更公司印鑑 04
   '                        strTmp = "04"
   '                     ElseIf RsTemp.Fields(0) = "1" Then
   '                        '變更個人印鑑 03
   '                        strTmp = "03"
   '                     End If
                        'Modify By Sindy 2012/5/24
                        If RsTemp.Fields(0) = "0" Then
                           '變更個人印鑑 03
                           strTmp = "03"
                        Else
                           '變更公司印鑑 04
                           strTmp = "04"
                        End If
                        '2012/5/24 End
                     End If
                     bPrt = True
                  End If
                  If Check1(9).Value = 1 And bPrt = False Then
                     '變更公司印鑑 4
                     strTmp = "04"
                     bPrt = True
                  End If
                  If Check1(0).Value = 1 And bPrt = False Then
                     '變更申請人
                     strTmp = "05"
                     bPrt = True
                  End If
                  
               Case 國外_FC
                  
                  If Check1(10).Value = 1 And bPrt = False Then
                     '變更代理人 08
                     strTmp = "08"
                     bPrt = True
                  End If
                  If (Check1(0).Value = 1 Or Me.Check1(1).Value = vbChecked) And Check1(2).Value = 1 And (Check1(3).Value = 1 Or Me.Check1(4).Value = vbChecked) And bPrt = False Then
                     '變更申請人(或中譯文)及地址及代表人(或中譯文) 7
                     strTmp = "07"
                     bPrt = True
                  ElseIf (Check1(0).Value = 1 Or Me.Check1(1).Value = vbChecked) And Check1(2).Value = 1 And bPrt = False Then
                     '變更申請人(或中譯文)及地址 4
                     strTmp = "04"
                     bPrt = True
                  ElseIf (Check1(0).Value = 1 Or Me.Check1(1).Value = vbChecked) And (Check1(3).Value = 1 Or Me.Check1(4).Value = vbChecked) And bPrt = False Then
                     '變更申請人(或中譯文)及代表人(或中譯文) 5
                     strTmp = "05"
                     bPrt = True
                  ElseIf Check1(2).Value = 1 And (Check1(3).Value = 1 Or Me.Check1(4).Value = vbChecked) And bPrt = False Then
                     '變更公司地址及代表人(或中譯文) 6
                     strTmp = "06"
                     bPrt = True
                  ElseIf (Check1(0).Value = 1 Or Me.Check1(1).Value = vbChecked) And bPrt = False Then
                     '變更申請人(或中譯文) 1
                     strTmp = "01"
                     bPrt = True
                  ElseIf Check1(2).Value = 1 And bPrt = False Then
                     '變更公司地址 2
                     strTmp = "02"
                     bPrt = True
                  ElseIf (Check1(3).Value = 1 Or Me.Check1(4).Value = vbChecked) And bPrt = False Then
                     '變更代表人(或中譯文) 03
                     strTmp = "03"
                     bPrt = True
                  End If
                  
                  If Check1(8).Value = 1 And bPrt = False Then
                     strExc(0) = "SELECT CU15 FROM CUSTOMER WHERE (CU01,CU02) IN " & _
                        "(SELECT SUBSTR(PA26,1,8),SUBSTR(PA26,9,1) FROM PATENT WHERE " & ChgPatent(pa(1) & pa(2) & pa(3) & pa(4)) & ")"
                     intI = 1
                     Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
                     If intI = 1 Then
   '                     If RsTemp.Fields(0) = "0" Then
   '                        '變更公司印鑑 10
   '                        strTmp = "10"
   '                     ElseIf RsTemp.Fields(0) = "1" Then
   '                        '變更個人印鑑 9
   '                        strTmp = "09"
   '                     End If
                        'Modify By Sindy 2012/5/24
                        If RsTemp.Fields(0) = "0" Then
                           '變更個人印鑑 9
                           strTmp = "09"
                        Else
                           '變更公司印鑑 10
                           strTmp = "10"
                        End If
                        '2012/5/24 End
                     End If
                     bPrt = True
                  End If
                  If Check1(9).Value = 1 And bPrt = False Then
                     '變更公司印鑑 10
                     strTmp = "10"
                     bPrt = True
                  End If
                  
                  'Add by Morgan 2010/6/10 其他
                  If Check1(11).Value = 1 And bPrt = False Then
                     strTmp = 11
                     bPrt = True
                  End If
            End Select
            
            If bPrt = True Then
               StartLetter "01", strTmp
               NowPrint strReceiveNo, "01", strTmp, bolChk, strUserNum
            End If
         End If
      End If
   End If
   Unload Me
End Sub

Public Sub LoadMe(ByVal RecNo As String, ByVal txt1 As String, ByVal txt2 As String, _
   ByVal txt3 As String, ByVal txt4 As String, ByVal iGo As Integer)
   Text1 = txt1
   Text2 = txt2
   Text3 = txt3
   Text4 = txt4
   strReceiveNo = RecNo
   If Left(Format(iGo), 1) = "4" Then
      intWhere = 國內
   Else
      intWhere = 國外_FC
   End If
   intGo = iGo
   Text5.MaxLength = 7
   Text34.MaxLength = 7
   Text5 = strSrvDate(2)
   If intGo = 41 Or intGo = 61 Then
      Label52.Visible = True
      Label55.Visible = True
      Text41.Visible = True
   Else
      Label52.Visible = False
      Label55.Visible = False
      Text41.Visible = False
      Text41.Text = "N"
      Text62.Visible = False
   End If
   
   'Add by Morgan 2005/8/8
   ReDim pa(TF_PA)
   ReDim cp(TF_CP) 'Add By Sindy 2018/8/10
   ReadPatent
   'Add by Morgan 2005/8/8
   '加出名代理人清單供勾選
   lstNameAgent.Clear
   If (intGo = 41 Or intGo = 61) And pa(9) = 台灣國家代號 Then
      PUB_SetOurAgent lstNameAgent, pa(), m_CP110, , True
      lstNameAgent.Visible = True
      lblNameAgent.Visible = True
   Else
      lstNameAgent.Visible = False
      lblNameAgent.Visible = False
   End If
   'Added by Sindy 2021/5/10 如果一開始將ListBox拉到需要的大小，字型會自動放大；所以畫面預設為一列高度，Form_Load才放大到需要的大小
   lstNameAgent.Height = 1100
   lstNameAgent.Width = 1300

   Combo1.ListIndex = 0
   Text5.Text = strSrvDate(2)
   
   'Added by Morgan 2011/11/23 考慮第二次進來但未存檔保留原來狀態
   If intGo = 64 Then
      bolSaveOK = oParent.m_bolSaveChgEvent
   End If
End Sub

Private Sub Combo2_Click(Index As Integer)
   Dim strCust As String
   Dim strCU01 As String
   Dim strCU02 As String
   Dim strCUField As String
   Dim rsTmp As ADODB.Recordset
   Dim strSql As String
   
   ' 90.06.22 modify by louis
   strCust = Combo2(Index).List(Combo2(Index).ListIndex)
   If IsEmptyText(strCust) = False Then
      strCU01 = Mid(strCust, 1, 8)
      strCU02 = Mid(strCust, 9, 1)
      Select Case Mid(strCust, 11, 1)
         Case "1": strCUField = "CU39,CU40,CU41"
         Case "2": strCUField = "CU42,CU43,CU44"
         Case "3": strCUField = "CU45,CU46,CU47"
         Case "4": strCUField = "CU48,CU49,CU50"
         Case "5": strCUField = "CU51,CU52,CU53"
         Case "6": strCUField = "CU54,CU55,CU56"
      End Select
      Set rsTmp = New ADODB.Recordset
      strSql = "SELECT " & strCUField & " FROM CUSTOMER " & _
               "WHERE CU01 = '" & strCU01 & "' AND " & _
                     "CU02 = '" & strCU02 & "' "
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         Text9(Index * 3).Text = "" & rsTmp.Fields(0) 'Modified by Morgan 2015/10/22 也可能沒中文
         Text9(Index * 3 + 1).Text = "" & rsTmp.Fields(1)
         Text9(Index * 3 + 2).Text = "" & rsTmp.Fields(2)
      End If
      rsTmp.Close
      Set rsTmp = Nothing
   Else
      Text9(Index * 3).Text = Empty
      Text9(Index * 3 + 1).Text = Empty
      Text9(Index * 3 + 2).Text = Empty
   End If
   
   'Dim i As Integer, strTmp As String
   'If Combo2(Index) = "" Then
   '   For i = 0 To 2
   '      Text9(i + Index * 3) = ""
   '   Next
   '   Exit Sub
   'End If
   '
   'strTmp = Mid(Combo2(Index).Text, InStr(Combo2(Index).Text, "-") + 1, 1)
   'strExc(1) = "CU" & 39 + (Val(strTmp) - 1) * 3 & ",CU" & 40 + (Val(strTmp) - 1) * 3 & ",CU" & 41 + (Val(strTmp) - 1) * 3
   ' strExc(0) = "SELECT " & strExc(1) & " FROM CUSTOMER WHERE " & ChgCustomer(Left(Combo2(Index).Text, InStr(Combo2(Index).Text, "-") - 1))
   'intI = 1
   'Set rsTemp = clslawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
   'If intI = 1 Then
   '   For i = 0 To 2
   '      If Not IsNull(rsTemp.Fields(i)) Then
   '         Text9(i + Index * 3) = rsTemp.Fields(i)
   '      Else
   '         Text9(i + Index * 3) = ""
   '      End If
   '   Next
   'End If
End Sub

Private Sub Form_Load()
    'Added by Lydia 2016/09/10 設定代表人中文名稱和英文名稱長度
    Text9(0).MaxLength = Pub_MaxCEL10
    Text9(1).MaxLength = Pub_MaxCEL11
    Text9(3).MaxLength = Pub_MaxCEL10
    Text9(4).MaxLength = Pub_MaxCEL11
    Text9(6).MaxLength = Pub_MaxCEL10
    Text9(7).MaxLength = Pub_MaxCEL11
    Text9(9).MaxLength = Pub_MaxCEL10
    Text9(10).MaxLength = Pub_MaxCEL11
    Text9(12).MaxLength = Pub_MaxCEL10
    Text9(13).MaxLength = Pub_MaxCEL11
    Text9(15).MaxLength = Pub_MaxCEL10
    Text9(16).MaxLength = Pub_MaxCEL11
    Text9(18).MaxLength = Pub_MaxCEL10
    Text9(19).MaxLength = Pub_MaxCEL11
    Text9(21).MaxLength = Pub_MaxCEL10
    Text9(22).MaxLength = Pub_MaxCEL11
    Text9(24).MaxLength = Pub_MaxCEL10
    Text9(25).MaxLength = Pub_MaxCEL11
    Text9(27).MaxLength = Pub_MaxCEL10
    Text9(28).MaxLength = Pub_MaxCEL11
    'end 2016/09/10
    
   MoveFormToCenter Me
   Text62 = "Y"
   'Add by Morgan 2003/12/07
   Check1_Click (3)
   'End 2003/12/07
   
   SSTab1.Tab = 0 'Add By Sindy 2018/6/12
   
   FraPA174.BackColor = &H8000000F 'Added by Lydia 2020/02/21
   
End Sub

Private Sub ReadPatent()
Dim i As Integer, j As Integer
Dim rsTemp1 As New ADODB.Recordset, Lbl As Object, bolChk As Boolean

   For Each Lbl In Label3
      Lbl = ""
   Next
   For Each Lbl In Label4
      Lbl = ""
   Next
   pa(1) = Text1
   pa(2) = Text2
   pa(3) = Text3
   pa(4) = Text4
   If ClsPDReadPatentDatabase(pa(), intWhere) Then 'edit by nickc 2007/02/02 不用 dll 了 If objPublicData.ReadPatentDatabase(pA(), intWhere) Then
      If pa(9) <> "" Then
         'edit by nickc 2007/02/02 不用 dll 了
         'If objPublicData.GetNation(pA(9), strExc(0)) Then Label3(9) = strExc(0)
         If ClsPDGetNation(pa(9), strExc(0)) Then Label3(9) = strExc(0)
      End If
      AddCboName Combo1, pa(5), pa(6), pa(7)
      Label3(6) = pa(11)
      Label3(7) = pa(22)
   End If
   
   If pa(9) = 台灣國家代號 Then
      strExc(1) = "CPM03,"
   Else
      strExc(1) = "CPM04,"
   End If
   
   'Add By Sindy 2018/8/10
   cp(9) = strReceiveNo
   If PUB_ReadCaseProgressDatabase(cp(), intWhere) Then
      'Add By Sindy 2023/7/19 True:更改(因客戶/本所誤繕需更改-不請款)
      'Modify By Sindy 2023/8/4
      'If m_CP20isN = True Then
      If cp(10) = 更改 Then
      '2023/8/4 END
         txtCP84.Tag = 300
      Else
      '2023/7/19 END
         txtCP84.Tag = cp(17)
      End If
      txtCP84.Text = txtCP84.Tag
   End If
   '2018/8/10 END
   
   strExc(0) = "select " & strExc(1) & "staff.st02 as st1,staff1.st02 as st2,cp43,cp06,cp07,CP110 from caseprogress," & _
      "casepropertymap,staff,staff staff1 where cp09='" & strReceiveNo & "' and cp01=cpm01(+) and cp10=cpm02(+) and " & _
      "cp14=staff.st01(+) and cp13=staff1.st01(+)"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
   With RsTemp
   If intI = 1 Then
      m_CP110 = "" & .Fields("CP110")
      If Not IsNull(.Fields(0)) Then Label3(2) = .Fields(0)
      If Not IsNull(.Fields(1)) Then Label3(3) = .Fields(1)
      If Not IsNull(.Fields(2)) Then Label3(4) = .Fields(2)
      If Not IsNull(.Fields(3)) Then
         strExc(0) = "SELECT CP05,CP08 FROM CASEPROGRESS WHERE CP09='" & .Fields(3) & "'"
         intI = 1
         Set rsTemp1 = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            If Not IsNull(rsTemp1.Fields(0)) Then Label3(1) = TransDate(rsTemp1.Fields(0), 1)
            If Not IsNull(rsTemp1.Fields(1)) Then Label3(5) = rsTemp1.Fields(1)
         End If
      End If
      If Not IsNull(.Fields(4)) Then Label3(10) = .Fields(4)
      If Not IsNull(.Fields(5)) Then Label3(11) = .Fields(5)
   End If
   End With
   
   strExc(0) = "select * from changeevent where ce01 = '" & strReceiveNo & "'"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
   With RsTemp
   If intI = 1 Then
      
      '申請人
'      If .Fields("ce09").Value = "1" Then Check1(0).Value = 1
      bolChk = False
      For i = 3 To 7
         If IsNull(.Fields(i)) = False Then
            Text7(i - 3) = .Fields(i)
            ChgType (i - 3)
            bolChk = True
         End If
      Next
      If bolChk Then Check1(0).Value = 1
      
      IntoCombol
      
      '代表人
'      If .Fields("ce16").Value = "1" Then Check1(3).Value = 1
      bolChk = False
      For i = 9 To 14
         If IsNull(.Fields(i)) = False Then
            Text9(i - 9) = .Fields(i)
            bolChk = True
         End If
      Next
      For i = 67 To 90
         If IsNull(.Fields(i)) = False Then
            Text9(i - 61) = .Fields(i)
            bolChk = True
         End If
      Next
      If bolChk Then Check1(3).Value = 1

      '申請人中譯文
'      If .Fields("ce22").Value = "1" Then Check1(1).Value = 1
      bolChk = False
      For i = 16 To 20
         If IsNull(.Fields(i)) = False Then
            Text6(i - 16) = .Fields(i)
            bolChk = True
         End If
         Text6(i - 16).Visible = False
      Next
      If bolChk Then Check1(1).Value = 1
      
      '代表人中議文
'      If .Fields("ce65").Value = "1" Then Check1(4).Value = 1
      bolChk = False
      For i = 62 To 63
         If IsNull(.Fields(i)) = False Then
            Text10(i - 62) = .Fields(i)
            bolChk = True
         End If
      Next
      For i = 91 To 98
         If IsNull(.Fields(i)) = False Then
            Text10(i - 89) = .Fields(i)
            bolChk = True
         End If
      Next
      
      If bolChk Then Check1(4).Value = 1
            
      '申請人地址
      j = 0
'      If .Fields("ce38").Value = "1" Then Check1(2).Value = 1
      bolChk = False
      For i = 22 To 35 Step 3
         If IsNull(.Fields(i)) = False Then
            Text8(j) = .Fields(i)
            bolChk = True
         End If
         'Add By Sindy 2018/5/25
         If IsNull(.Fields(i + 1)) = False Then
            Text8(j + 5) = .Fields(i + 1)
            bolChk = True
         End If
         '2018/5/25 END
         j = j + 1
      Next
      If bolChk Then Check1(2).Value = 1
      
'      If .Fields("ce03").Value = "1" Then Check1(5).Value = 1
      If IsNull(.Fields("ce02")) = False Then
         Text34 = .Fields("ce02")
         Check1(5).Value = 1
      End If
      
'      If .Fields("ce40").Value = "1" Then Check1(6).Value = 1
      If IsNull(.Fields("ce39")) = False Then
         Text36 = .Fields("ce39")
         ChgType (36)
         Check1(6).Value = 1
      End If
      
'      If .Fields("ce44").Value = "1" Then Check1(7).Value = 1
      bolChk = False
      For i = 40 To 42
         If IsNull(.Fields(i)) = False Then
            Text38(i - 40) = .Fields(i)
            bolChk = True
         End If
      Next
      If bolChk Then Check1(7).Value = 1
      
'      If .Fields("ce46").Value = "1" Then Check1(14).Value = 1
      If IsNull(.Fields("ce45")) = False Then
         Text56 = .Fields("ce45")
         Check1(14).Value = 1
      End If

'      If .Fields("ce48").Value = "1" Then Check1(15).Value = 1
      If IsNull(.Fields("ce47")) = False Then
         Text58 = .Fields("ce47")
         Check1(15).Value = 1
      End If

'      If .Fields("ce50").Value = "1" Then Check1(16).Value = 1
      If IsNull(.Fields("ce49")) = False Then
         Text60 = .Fields("ce49")
         Check1(16).Value = 1
      End If

'      If .Fields("ce52").Value = "1" Then Check1(8).Value = 1
      If IsNull(.Fields("ce51")) = False Then
         Text44 = .Fields("ce51")
         Check1(8).Value = 1
      End If

'      If .Fields("ce54").Value = "1" Then Check1(9).Value = 1
      If IsNull(.Fields("ce53")) = False Then
         Text46 = .Fields("ce53")
         Check1(9).Value = 1
      End If

'      If .Fields("ce56").Value = "1" Then Check1(10).Value = 1
      If IsNull(.Fields("ce55")) = False Then
         Text48 = .Fields("ce55")
         Check1(10).Value = 1
      End If

'      If .Fields("ce58").Value = "1" Then Check1(12).Value = 1
      If IsNull(.Fields("ce57")) = False Then
         Text52 = .Fields("ce57")
         Check1(12).Value = 1
      End If

'      If .Fields("ce60").Value = "1" Then Check1(13).Value = 1
      If IsNull(.Fields("ce59")) = False Then
         Text54 = .Fields("ce59")
         Check1(13).Value = 1
      End If
      
'      If .Fields("ce62").Value = "1" Then Check1(11).Value = 1
      'Add By Sindy 2018/9/17 記錄到其他欄位
      If IsNull(.Fields("ce61")) = False Then
         Text50 = .Fields("ce61")
         '變更發明人/創作人/設計人之國籍
         If InStr(Text50, Trim(Check1(17).Caption) & ";") > 0 Then
            Text50 = Replace(Text50, Trim(Check1(17).Caption) & ";", "")
            Check1(17).Value = 1
         End If
         '變更發明人/創作人/設計人之姓名
         If InStr(Text50, Trim(Check1(18).Caption) & ";") > 0 Then
            Text50 = Replace(Text50, Trim(Check1(18).Caption) & ";", "")
            Check1(18).Value = 1
         End If
         '追加、刪除或更正發明人/創作人/設計人
         If InStr(Text50, Trim(Check1(19).Caption) & ";") > 0 Then
            Text50 = Replace(Text50, Trim(Check1(19).Caption) & ";", "")
            Check1(19).Value = 1
         End If
         'Add By Sindy 2023/2/20
         '申請人國籍
         If InStr(Text50, Trim(Check1(20).Caption) & ";") > 0 Then
            Text50 = Replace(Text50, Trim(Check1(20).Caption) & ";", "")
            Check1(20).Value = 1
         End If
         '2023/2/20 END
         'Check1(11).Value = 1
         If Text50 <> "" Then Check1(11).Value = 1
      End If
      '2018/9/17 END
   End If
   End With
   
   'Added by Lydia 2020/02/21 預設「名稱有特殊字」
   FraPA174.Visible = False
   If pa(1) = "FCP" Or pa(1) = "P" Then
       If pa(174) = "Y" Then
          FraPA174.Visible = True
       End If
   End If
   'end 2020/02/21
   
End Sub

Private Sub IntoCombol()
   Dim i As Integer, j As Integer
   Dim strSql As String
   Dim rsTmp As ADODB.Recordset
   Dim strCP01 As String
   Dim strCP02 As String
   Dim strCP03 As String
   Dim strCP04 As String
   Dim strCU(5) As String
   Dim nIndex As Integer
   
   For nIndex = 0 To 9
      Combo2(nIndex).Clear
      Combo2(nIndex).AddItem ""
   Next
   
   strCU(0) = Empty
   strCU(1) = Empty
   strCU(2) = Empty
   strCU(3) = Empty
   strCU(4) = Empty
   
   ' 90.06.22 modify by louis 申請人不打勾時, 代表人資料從基本檔帶入
   If Check1(0).Value = 1 Then
      If IsEmptyText(Text7(0)) And IsEmptyText(Text7(1)) And IsEmptyText(Text7(2)) And IsEmptyText(Text7(3)) And IsEmptyText(Text7(4)) Then: Exit Sub
      If IsEmptyText(Text7(0)) = False Then
         strCU(0) = Text7(0) & String(9 - Len(Text7(0)), "0")
      End If
      If IsEmptyText(Text7(1)) = False Then
         strCU(1) = Text7(1) & String(9 - Len(Text7(1)), "0")
      End If
      If IsEmptyText(Text7(2)) = False Then
         strCU(2) = Text7(2) & String(9 - Len(Text7(2)), "0")
      End If
      If IsEmptyText(Text7(3)) = False Then
         strCU(3) = Text7(3) & String(9 - Len(Text7(3)), "0")
      End If
      If IsEmptyText(Text7(4)) = False Then
         strCU(4) = Text7(4) & String(9 - Len(Text7(4)), "0")
      End If
      For nIndex = 0 To 4
         If IsEmptyText(strCU(nIndex)) = False Then
            'Modify By Sindy 2018/6/12
            'strExc(0) = "SELECT CU40,CU43,CU46,CU49,CU52,CU55 FROM CUSTOMER WHERE " & ChgCustomer(strCU(nIndex))
            strExc(0) = "SELECT nvl(CU40,nvl(CU39,CU41))" & _
                              ",nvl(CU43,nvl(CU42,CU44))" & _
                              ",nvl(CU46,nvl(CU45,CU47))" & _
                              ",nvl(CU49,nvl(CU48,CU50))" & _
                              ",nvl(CU52,nvl(CU51,CU53))" & _
                              ",nvl(CU55,nvl(CU54,CU56)) FROM CUSTOMER WHERE " & ChgCustomer(strCU(nIndex))
            '2018/6/12 END
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
            For i = 0 To 5
               If IsNull(RsTemp.Fields(i)) = False Then
                  If IsEmptyText(RsTemp.Fields(i)) = False Then
                     Combo2(0).AddItem strCU(nIndex) & "-" & CStr(i + 1) & " " & RsTemp.Fields(i)
                     Combo2(1).AddItem strCU(nIndex) & "-" & CStr(i + 1) & " " & RsTemp.Fields(i)
                     Combo2(2).AddItem strCU(nIndex) & "-" & CStr(i + 1) & " " & RsTemp.Fields(i)
                     Combo2(3).AddItem strCU(nIndex) & "-" & CStr(i + 1) & " " & RsTemp.Fields(i)
                     Combo2(4).AddItem strCU(nIndex) & "-" & CStr(i + 1) & " " & RsTemp.Fields(i)
                     Combo2(5).AddItem strCU(nIndex) & "-" & CStr(i + 1) & " " & RsTemp.Fields(i)
                     Combo2(6).AddItem strCU(nIndex) & "-" & CStr(i + 1) & " " & RsTemp.Fields(i)
                     Combo2(7).AddItem strCU(nIndex) & "-" & CStr(i + 1) & " " & RsTemp.Fields(i)
                     Combo2(8).AddItem strCU(nIndex) & "-" & CStr(i + 1) & " " & RsTemp.Fields(i)
                     Combo2(9).AddItem strCU(nIndex) & "-" & CStr(i + 1) & " " & RsTemp.Fields(i)
                  End If
               End If
            Next i
         End If
      Next
   Else
      Set rsTmp = New ADODB.Recordset
      strSql = "SELECT CP01,CP02,CP03,CP04 FROM CASEPROGRESS " & _
               "WHERE CP09 = '" & strReceiveNo & "' "
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         rsTmp.MoveFirst
         strCP01 = rsTmp.Fields("CP01")
         strCP02 = rsTmp.Fields("CP02")
         strCP03 = rsTmp.Fields("CP03")
         strCP04 = rsTmp.Fields("CP04")
      End If
      rsTmp.Close

      If IsEmptyText(strCP01) = False And IsEmptyText(strCP01) = False Then
         strSql = "SELECT * FROM PATENT " & _
                  "WHERE PA01 = '" & strCP01 & "' AND " & _
                        "PA02 = '" & strCP02 & "' AND " & _
                        "PA03 = '" & strCP03 & "' AND " & _
                        "PA04 = '" & strCP04 & "' "
         rsTmp.CursorLocation = adUseClient
         rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
         If rsTmp.RecordCount > 0 Then
            nIndex = 0
            If IsNull(rsTmp.Fields("PA26")) = False Then
               If IsEmptyText(rsTmp.Fields("PA26")) = False Then
                  strCU(nIndex) = rsTmp.Fields("PA26")
                  nIndex = nIndex + 1
               End If
            End If
            If IsNull(rsTmp.Fields("PA27")) = False Then
               If IsEmptyText(rsTmp.Fields("PA27")) = False Then
                  strCU(nIndex) = rsTmp.Fields("PA27")
                  nIndex = nIndex + 1
               End If
            End If
            If IsNull(rsTmp.Fields("PA28")) = False Then
               If IsEmptyText(rsTmp.Fields("PA28")) = False Then
                  strCU(nIndex) = rsTmp.Fields("PA28")
                  nIndex = nIndex + 1
               End If
            End If
            If IsNull(rsTmp.Fields("PA29")) = False Then
               If IsEmptyText(rsTmp.Fields("PA29")) = False Then
                  strCU(nIndex) = rsTmp.Fields("PA29")
                  nIndex = nIndex + 1
               End If
            End If
            If IsNull(rsTmp.Fields("PA40")) = False Then
               If IsEmptyText(rsTmp.Fields("PA30")) = False Then
                  strCU(nIndex) = rsTmp.Fields("PA30")
                  nIndex = nIndex + 1
               End If
            End If
         End If
         rsTmp.Close
      End If
      Set rsTmp = Nothing
      
      For nIndex = 0 To 4
         If IsEmptyText(strCU(nIndex)) = False Then
            'Modify By Sindy 2018/6/12
            'strExc(0) = "SELECT CU40,CU43,CU46,CU49,CU52,CU55 FROM CUSTOMER WHERE " & ChgCustomer(strCU(nIndex))
            strExc(0) = "SELECT nvl(CU40,nvl(CU39,CU41))" & _
                              ",nvl(CU43,nvl(CU42,CU44))" & _
                              ",nvl(CU46,nvl(CU45,CU47))" & _
                              ",nvl(CU49,nvl(CU48,CU50))" & _
                              ",nvl(CU52,nvl(CU51,CU53))" & _
                              ",nvl(CU55,nvl(CU54,CU56)) FROM CUSTOMER WHERE " & ChgCustomer(strCU(nIndex))
            '2018/6/12 END
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
            For i = 0 To 5
               If IsNull(RsTemp.Fields(i)) = False Then
                  If IsEmptyText(RsTemp.Fields(i)) = False Then
                     Combo2(0).AddItem strCU(nIndex) & "-" & CStr(i + 1) & " " & RsTemp.Fields(i)
                     Combo2(1).AddItem strCU(nIndex) & "-" & CStr(i + 1) & " " & RsTemp.Fields(i)
                     Combo2(2).AddItem strCU(nIndex) & "-" & CStr(i + 1) & " " & RsTemp.Fields(i)
                     Combo2(3).AddItem strCU(nIndex) & "-" & CStr(i + 1) & " " & RsTemp.Fields(i)
                     Combo2(4).AddItem strCU(nIndex) & "-" & CStr(i + 1) & " " & RsTemp.Fields(i)
                     Combo2(5).AddItem strCU(nIndex) & "-" & CStr(i + 1) & " " & RsTemp.Fields(i)
                     Combo2(6).AddItem strCU(nIndex) & "-" & CStr(i + 1) & " " & RsTemp.Fields(i)
                     Combo2(7).AddItem strCU(nIndex) & "-" & CStr(i + 1) & " " & RsTemp.Fields(i)
                     Combo2(8).AddItem strCU(nIndex) & "-" & CStr(i + 1) & " " & RsTemp.Fields(i)
                     Combo2(9).AddItem strCU(nIndex) & "-" & CStr(i + 1) & " " & RsTemp.Fields(i)
                  End If
               End If
            Next i
         End If
      Next nIndex
   End If
End Sub

Private Function FormSave() As Boolean
   Dim i As Integer, strTmp As String
   Dim nIndex As Integer
   'Add By Cheng 2003/03/04
   Dim StrSQLa As String
   'Add By Cheng 2003/03/18
   Dim strCe(99) As String, bolChk As Boolean
   Dim strTmpA(1 To 5) As String
   Dim intStep As Integer
   Dim strTxt(1 To 20) As String
   
   'Add by Morgan 2011/6/13
   Dim bolChkMemo605 As Boolean, bolChkMemo416 As Boolean
   Dim strOldMemo605 As String, strOldMemo416 As String
   Dim strNewMemo605 As String, strNewMemo416 As String
   
   ' 90.07.18 modify by louis (申請人補滿九碼)
   For nIndex = 0 To 4
      If IsEmptyText(Text7(nIndex)) = False Then
         Text7(nIndex) = Text7(nIndex) & String(9 - Len(Text7(nIndex)), "0")
      End If
   Next nIndex
   
   strTmp = ""
   
   '變更申請人
   If Check1(0).Value = 1 Then
      strTmp = strTmp & "CE04=" & CNULL(ChangeCustomerL(Text7(0))) & ",CE05=" & CNULL(ChangeCustomerL(Text7(1))) & ",CE06=" & _
         CNULL(ChangeCustomerL(Text7(2))) & ",CE07=" & CNULL(ChangeCustomerL(Text7(3))) & ",CE08=" & CNULL(ChangeCustomerL(Text7(4))) & ","

   End If
   '變更代表人
   If Check1(3).Value = 1 Then
      strTmp = strTmp & "CE10=" & CNULL(Text9(0)) & ",CE11=" & CNULL(Text9(1)) & ",CE12=" & CNULL(Text9(2)) & _
         ",CE13=" & CNULL(Text9(3)) & ",CE14=" & CNULL(Text9(4)) & ",CE15=" & CNULL(Text9(5)) & ",CE68=" & CNULL(Text9(6)) & _
         ",CE69=" & CNULL(Text9(7)) & ",CE70=" & CNULL(Text9(8)) & ",CE71=" & CNULL(Text9(9)) & ",CE72=" & CNULL(Text9(10)) & _
         ",CE73=" & CNULL(Text9(11)) & ",CE74=" & CNULL(Text9(12)) & ",CE75=" & CNULL(Text9(13)) & ",CE76=" & CNULL(Text9(14)) & _
         ",CE77=" & CNULL(Text9(15)) & ",CE78=" & CNULL(Text9(16)) & ",CE79=" & CNULL(Text9(17)) & ",CE80=" & CNULL(Text9(18)) & _
         ",CE81=" & CNULL(Text9(19)) & ",CE82=" & CNULL(Text9(20)) & ",CE83=" & CNULL(Text9(21)) & ",CE84=" & CNULL(Text9(22)) & _
         ",CE85=" & CNULL(Text9(23)) & ",CE86=" & CNULL(Text9(24)) & ",CE87=" & CNULL(Text9(25)) & ",CE88=" & CNULL(Text9(26)) & _
         ",CE89=" & CNULL(Text9(27)) & ",CE90=" & CNULL(Text9(28)) & ",CE91=" & CNULL(Text9(29)) & ","
   End If
    '變更申請人中譯文
   If Check1(1).Value = 1 Then
      strTmp = strTmp & "CE17=" & CNULL(Text6(0)) & ",CE18=" & CNULL(Text6(1)) & ",CE19=" & _
         CNULL(Text6(2)) & ",CE20=" & CNULL(Text6(3)) & ",CE21=" & CNULL(Text6(4)) & ","
   End If
   
   If Check1(4).Value = 1 Then
      strTmp = strTmp & "CE63=" & CNULL(Text10(0)) & ",CE64=" & CNULL(Text10(1)) & ",CE92=" & CNULL(Text10(2)) & _
         ",CE93=" & CNULL(Text10(3)) & ",CE94=" & CNULL(Text10(4)) & ",CE95=" & CNULL(Text10(5)) & ",CE96=" & CNULL(Text10(6)) & _
         ",CE97=" & CNULL(Text10(7)) & ",CE98=" & CNULL(Text10(8)) & ",CE99=" & CNULL(Text10(9)) & ","
   End If
   
   If Check1(2).Value = 1 Then
      'Modify By Sindy 2018/5/25 + 申請人英文地址
'      strTmp = strTmp & "CE23=" & CNULL(Text8(0)) & ",CE26=" & CNULL(Text8(1)) & ",CE29=" & _
'         CNULL(Text8(2)) & ",CE32=" & CNULL(Text8(3)) & ",CE35=" & CNULL(Text8(4)) & ","
      'Modify By Sindy 2018/9/17 + ChgSQL ex.FCP-59585
      strTmp = strTmp & "CE23=" & CNULL(Text8(0)) & ",CE24=" & CNULL(ChgSQL(Text8(5))) & ",CE26=" & CNULL(Text8(1)) & ",CE27=" & CNULL(ChgSQL(Text8(6))) & _
         ",CE29=" & CNULL(Text8(2)) & ",CE30=" & CNULL(ChgSQL(Text8(7))) & ",CE32=" & CNULL(Text8(3)) & ",CE33=" & CNULL(ChgSQL(Text8(8))) & _
         ",CE35=" & CNULL(Text8(4)) & ",CE36=" & CNULL(ChgSQL(Text8(9))) & ","
      '2018/5/25 END
   End If
   
   If Check1(5).Value = 1 Then
      strTmp = strTmp & "CE02=" & CNULL(Text34) & ","
   End If
   
   If Check1(6).Value = 1 Then
      strTmp = strTmp & "CE39=" & CNULL(Text36) & ","
   End If
   
   If Check1(7).Value = 1 Then
      strTmp = strTmp & "CE41=" & CNULL(Text38(0)) & ",CE42=" & CNULL(Text38(1)) & ",CE43=" & CNULL(Text38(2)) & ","
   End If
   
   If Check1(14).Value = 1 Then
      strTmp = strTmp & "CE45=" & CNULL(Text56) & ","
   End If

   If Check1(15).Value = 1 Then
      strTmp = strTmp & "CE47=" & CNULL(Text58) & ","
   End If

   If Check1(16).Value = 1 Then
      strTmp = strTmp & "CE49=" & CNULL(Text60) & ","
   End If

   If Check1(8).Value = 1 Then
      strTmp = strTmp & "CE51='v',"
   End If

   If Check1(9).Value = 1 Then
      strTmp = strTmp & "CE53='v',"
   End If

   If Check1(10).Value = 1 Then
      strTmp = strTmp & "CE55='v',"
   End If

   If Check1(12).Value = 1 Then
      strTmp = strTmp & "CE57=" & CNULL(Text52) & ","
   End If

   If Check1(13).Value = 1 Then
      strTmp = strTmp & "CE59='v',"
   End If
   
   'Add By Sindy 2018/9/17 記錄到其他欄位
   'Modify By Sindy 2023/2/20 + Or Check1(20).Value = 1
   If Check1(17).Value = 1 Or Check1(18).Value = 1 Or Check1(19).Value = 1 Or Check1(20).Value = 1 Then
      Text50 = Trim(Text50)
      If Text50 <> "" And Right(Text50, 1) <> ";" Then Text50 = Text50 & ";"
      If Check1(17).Value = 1 Then
         Text50 = Text50 & "變更發明人/創作人/設計人之國籍;"
      End If
      If Check1(18).Value = 1 Then
         Text50 = Text50 & "變更發明人/創作人/設計人之姓名;"
      End If
      If Check1(19).Value = 1 Then
         Text50 = Text50 & "追加、刪除或更正發明人/創作人/設計人;"
      End If
      'Add By Sindy 2023/2/20
      If Check1(20).Value = 1 Then
         Text50 = Text50 & "申請人國籍;"
      End If
      '2023/2/20 END
   End If
   'If Check1(11).Value = 1 Then
   'Modify By Sindy 2023/2/20 + Or Check1(20).Value = 1
   If Check1(11).Value = 1 Or Check1(17).Value = 1 Or Check1(18).Value = 1 Or Check1(19).Value = 1 Or Check1(20).Value = 1 Then
   '2018/9/17 END
      strTmp = strTmp & "CE61=" & CNULL(Text50) & ","
   End If
   
   If strTmp <> "" Then
      strTmp = Left(strTmp, Len(strTmp) - 1)
   Else
      FormSave = True
      Exit Function
   End If
   
 FormSave = True
 On Error GoTo CheckingErr
 cnnConnection.BeginTrans
   
   strExc(1) = "DELETE FROM CHANGEEVENT WHERE CE01='" & strReceiveNo & "'"
   strExc(2) = "INSERT INTO CHANGEEVENT (CE01) VALUES ('" & strReceiveNo & "')"
   strExc(3) = "UPDATE CHANGEEVENT SET " & strTmp & " WHERE CE01='" & strReceiveNo & "'"
   cnnConnection.Execute strExc(1)
   cnnConnection.Execute strExc(2)
   cnnConnection.Execute strExc(3)
   
   strExc(0) = "SELECT * FROM CHANGEEVENT WHERE CE01='" & strReceiveNo & "'"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      With RsTemp
         For i = 1 To 99
            If IsNull(.Fields(i - 1)) Then
               strCe(i) = ""
            Else
               strCe(i) = .Fields(i - 1)
            End If
         Next
      End With
      strExc(1) = ""
      strExc(2) = ""
      strExc(3) = ""

      '申請日 10
      If strCe(2) <> "" Then
         'Modify by Morgan 2006/7/6 所有欄位名稱前面加[原],放變更前資料
         'strExc(1) = strExc(1) & "申請日 : " & strCe(2) & ","
         'strExc(2) = strExc(2) & "PA10=" & strCe(2) & ","
         If TransDate(strCe(2), 2) <> TransDate(pa(10), 2) Then
            strExc(1) = strExc(1) & "原申請日 : " & pa(10) & ","
         End If
         strExc(2) = strExc(2) & "PA10=" & TransDate(strCe(2), 2) & ","
         
         strExc(3) = strExc(3) & "CE03='1',"
      End If
      
      '申請人 26-30
      bolChk = False
      For i = 4 To 8
         If strCe(i) <> "" Then
            bolChk = True
            Exit For
         End If
      Next
      If bolChk = True Then
         'Modify by Morgan 2006/7/6
         'strExc(1) = strExc(1) & "原申請人 : "
         strTmp = ""
         
         For i = 4 To 8
            If strCe(i) <> "" Then
               'Modify by Morgan 2006/7/6
               'strExc(1) = strExc(1) & strCe(i) & ","
               'add by sonia 2015/9/25 P-111611 多申請人時,原申請人要記錄所有人
               'If ChangeCustomerL(strCe(i)) <> ChangeCustomerL(pa(22 + i)) Then
               '   strTmp = strTmp & pa(22 + i) & ","
               'End If
               If ChangeCustomerL(pa(22 + i)) <> "" Then strTmp = strTmp & pa(22 + i) & ","
               'end 2015/9/25
               'edit by nickc 2007/02/02 不用 dll 了
               'If objPublicData.GetCustomerNameAndAddress(strCe(i), strTmpA(5), strTmpA(1), strTmpA(2), strTmpA(3)) Then
               If ClsPDGetCustomerNameAndAddress(strCe(i), strTmpA(5), strTmpA(1), strTmpA(2), strTmpA(3)) Then
                  strExc(2) = strExc(2) & "PA" & i + 27 & "=" & CNULL(ChgSQL(strTmpA(1))) & ",PA" & i + 32 & "=" & CNULL(ChgSQL(strTmpA(2))) & ",PA" & i + 37 & "=" & CNULL(ChgSQL(strTmpA(3))) & ","
               End If
            Else
                strExc(2) = strExc(2) & "PA" & i + 27 & "=Null ,PA" & i + 32 & "=Null ,PA" & i + 37 & "=Null ,"
                'add by sonia 2015/9/25 P-111611 二申請人變更為一個
                If ChangeCustomerL(pa(22 + i)) <> "" Then strTmp = strTmp & pa(22 + i) & ","
            End If
            strExc(2) = strExc(2) & "PA" & i + 22 & "=" & CNULL(ChangeCustomerL(strCe(i))) & ","
         Next
         strExc(3) = strExc(3) & "CE09='1',"
         
         'Add by Morgan 2006/7/6
         If strTmp <> "" Then
            strExc(1) = strExc(1) & "原申請人 :" & strTmp
         End If
         
      Else
         '申請地址 31-45
         bolChk = False
         For i = 23 To 37
            If strCe(i) <> "" Then
               bolChk = True
               Exit For
            End If
         Next
         If bolChk = True Then
            'Modify by Morgan 2006/7/6
            'strExc(1) = strExc(1) & "申請地址 : "
            strTmp = ""
            
            For i = 0 To 4
               If strCe(23 + 3 * i) <> "" Then
                  'Modify by Morgan 2006/7/6
                  'strExc(1) = strExc(1) & strCe(23 + 3 * i) & ","
                  If strCe(23 + 3 * i) <> pa(i + 31) Then
                     strTmp = strTmp & pa(i + 31) & ","
                  End If
               End If
               'Modify By Sindy 2018/12/4 + ChgSQL()
               strExc(2) = strExc(2) & "PA" & i + 31 & "=" & CNULL(ChgSQL(strCe(23 + 3 * i))) & ","
            Next
            'Add By Sindy 2018/5/25 + 申請人英文地址
            For i = 0 To 4
               If strCe(24 + 3 * i) <> "" Then
                  If strCe(24 + 3 * i) <> pa(i + 36) Then
                     strTmp = strTmp & pa(i + 36) & ","
                  End If
               End If
               'Modify By Sindy 2018/12/4 + ChgSQL()
               strExc(2) = strExc(2) & "PA" & i + 36 & "=" & CNULL(ChgSQL(strCe(24 + 3 * i))) & ","
            Next
            '2018/5/25 END
            strExc(3) = strExc(3) & "CE38='1',"
            'Add by Morgan 2006/7/6
            If strTmp <> "" Then
               strExc(1) = strExc(1) & "原申請地址 :" & strTmp
            End If
         End If
      End If
      
      '專利商標種類代號 08
      If strCe(39) <> "" Then
         'Modify by Morgan 2006/7/6
         'strExc(1) = strExc(1) & "專利商標種類代號 : " & strCe(39) & ","
         If strCe(39) <> pa(8) Then
            strExc(1) = strExc(1) & "原專利商標種類代號 : " & pa(8) & ","
         End If
         
         strExc(2) = strExc(2) & "PA08='" & strCe(39) & "',"
         strExc(3) = strExc(3) & "CE40='1',"
      End If

      '案件名稱 05-07
      bolChk = False
      For i = 41 To 43
         If strCe(i) <> "" Then
            bolChk = True
            Exit For
         End If
      Next
      If bolChk = True Then
         'Modify by Morgan 2006/7/6
         'strExc(1) = strExc(1) & "案件名稱 : "
         strTmp = ""
         
         For i = 41 To 43
            If strCe(i) <> "" Then
               'Modify by Morgan 2006/7/6
               'strExc(1) = strExc(1) & strCe(i) & ","
               If strCe(i) <> pa(i - 36) Then
                  strTmp = strTmp & pa(i - 36) & ","
               End If
            End If
            'Modify by Morgan 2006/7/6
            'strExc(2) = strExc(2) & "PA" & i - 36 & "=" & CNULL(strCe(i)) & ","
            strExc(2) = strExc(2) & "PA" & Format(i - 36, "00") & "=" & CNULL(strCe(i)) & ","
         Next
         strExc(3) = strExc(3) & "CE44='1',"
         'Add by Morgan 2006/7/6
         If strTmp <> "" Then
            strExc(1) = strExc(1) & "原案件名稱:" & strTmp
         End If
      End If

      '代表人 79-84
      bolChk = False
      For i = 10 To 15
         If strCe(i) <> "" Then
            bolChk = True
            Exit For
         End If
      Next
      If Not bolChk Then
         For i = 68 To 91
            If strCe(i) <> "" Then
               bolChk = True
               Exit For
            End If
         Next
      End If

      If bolChk Then
         'Modify by Morgan 2006/7/6
         'strExc(1) = strExc(1) & "代表人 : "
         strTmp = ""
         
         For i = 10 To 15
            If strCe(i) <> "" Then
               'Modify by Morgan 2006/7/6
               'strExc(1) = strExc(1) & strCe(i) & ","
               If strCe(i) <> pa(i + 69) Then
                  strTmp = strTmp & pa(i + 69) & ","
               End If
            End If
            strExc(2) = strExc(2) & "PA" & i + 69 & "=" & CNULL(strCe(i)) & ","
         Next
         For i = 68 To 91
            If strCe(i) <> "" Then
               'Modify by Morgan 2006/7/6
               'strExc(1) = strExc(1) & strCe(i) & ","
               If strCe(i) <> pa(i + 41) Then
                  strTmp = strTmp & pa(i + 41) & ","
               End If
            End If
            strExc(2) = strExc(2) & "PA" & i + 41 & "=" & CNULL(strCe(i)) & ","
         Next
         strExc(3) = strExc(3) & "CE16='1',"
         'Add by Morgan 2006/7/6
         If strTmp <> "" Then
            strExc(1) = strExc(1) & "原代表人:" & strTmp
         End If
      End If
      
      '代表人中譯文
      If Not bolChk Then
         bolChk = False
         For i = 63 To 64
            If strCe(i) <> "" Then
               bolChk = True
               Exit For
            End If
         Next
         If Not bolChk Then
            For i = 92 To 99
               If strCe(i) <> "" Then
                  bolChk = True
                  Exit For
               End If
            Next
         End If
         If bolChk Then
            'Modify by Morgan 2006/7/6
            'strExc(1) = strExc(1) & "代表人中譯文 : "
            strTmp = ""
            
            strExc(2) = strExc(2) & "PA79=" & CNULL(strCe(63)) & ",PA82=" & CNULL(strCe(64)) & "," & _
               "PA109=" & CNULL(strCe(92)) & ",PA112=" & CNULL(strCe(93)) & ",PA115=" & CNULL(strCe(94)) & "," & _
               "PA118=" & CNULL(strCe(95)) & ",PA121=" & CNULL(strCe(96)) & ",PA124=" & CNULL(strCe(97)) & "," & _
               "PA127=" & CNULL(strCe(98)) & ",PA130=" & CNULL(strCe(99)) & ","
            For i = 63 To 64
               
               If strCe(i) <> "" Then
                  'Modify by Morgan 2006/7/6
                  'strExc(1) = strExc(1) & strCe(i) & ","
                  If strCe(i) <> pa(79 + 3 * (i - 63)) Then
                     strTmp = strTmp & pa(79 + 3 * (i - 63)) & ","
                  End If
               End If
            Next
            For i = 92 To 99
               If strCe(i) <> "" Then
                  'Modify by Morgan 2006/7/6
                  'strExc(1) = strExc(1) & strCe(i) & ","
                  If strCe(i) <> pa(109 + 3 * (i - 92)) Then
                     strTmp = strTmp & pa(109 + 17 * (i - 92)) & ","
                  End If
               End If
            Next
            strExc(3) = strExc(3) & "CE65='1',"
            'Add by Morgan 2006/7/6
            If strTmp <> "" Then
               strExc(1) = strExc(1) & "原代表人中譯文:" & strTmp
            End If
         End If
      End If
      'Add By Sindy 2018/9/18 其他
      If strCe(61) <> "" Then
         strExc(1) = strExc(1) & Text50 & ","
      End If
      '2018/9/18 END
      If strExc(1) <> "" Then
         For i = 2 To 3
            If Right(strExc(i), 1) = "," Then strExc(i) = Left(strExc(i), Len(strExc(i)) - 1)
         Next
         intStep = intStep + 1
         
         'Added by Morgan 2011/11/23 發文呼叫才更新
         If Not (oParent Is Nothing) Then
            If Left(oParent.Name, 9) = "frm060104" Or Left(oParent.Name, 9) = "frm040104" Then
         'end 2011/11/23
         
               '若有要更新的欄位
               If strExc(1) <> "" Then
                    'Modified by Morgan 2018/11/22 +ChgSQL(地址有單引號) FCP-49960
                    strTxt(intStep) = "UPDATE CASEPROGRESS SET CP64=CP64||'" & ChgSQL(strExc(1)) & "' WHERE CP09='" & strReceiveNo & "'"
                   cnnConnection.Execute strTxt(intStep)
                    intStep = intStep + 1
               End If
               '若有要更新的欄位
               If strExc(2) <> "" Then
    
                  'Add by Morgan 2011/6/13
                  '若有期限則於更新資料前紀錄原來的備註
                  'Modified by Morgan 2012/2/2 +pa75
                  strExc(0) = "select pa26,pa27,pa28,pa29,pa30,pa75,np07 from nextprogress,patent" & _
                     " where " & ChgNextProgress(pa(1) & pa(2) & pa(3) & pa(4)) & " and np06 is null and np07 in (416,605)" & _
                     " and pa01(+)=np02 and pa02(+)=np03 and pa03(+)=np04 and pa04(+)=np05"
                  intI = 1
                  Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                  If intI = 1 Then
                     If RsTemp("np07") = "416" Then
                        bolChkMemo416 = True
                        'Modified by Lydia 2022/08/02 整合模組：修改為複數新規則
                        'For intI = 0 To 4
                        '   If Not IsNull(RsTemp(intI)) Then
                        '      'Modified by Morgan 2012/2/2 +pa75
                        '      'Modified by Morgan 2013/9/11 改抓設定檔
                        '      'strOldMemo416 = PUB_Get416Memo(ChangeCustomerL(RsTemp(intI)), ChangeCustomerL("" & RsTemp("pa75")))
                        '      strOldMemo416 = PUB_GetNpMemo(pa(1) & pa(2) & pa(3) & pa(4), "416", ChangeCustomerL("" & RsTemp("pa75")), ChangeCustomerL(RsTemp(intI)))
                        '      If strOldMemo416 <> "" Then Exit For
                        '   End If
                        'Next
                        strOldMemo416 = PUB_GetNpMemo2("1", pa(1) & pa(2) & pa(3) & pa(4), "416", ChangeCustomerL("" & RsTemp.Fields("pa75")), RsTemp.Fields("PA26") & "," & RsTemp.Fields("PA27") & "," & RsTemp.Fields("PA28") & "," & RsTemp.Fields("PA29") & "," & RsTemp.Fields("PA30"))
                        'end 2022/08/02
                     ElseIf RsTemp("np07") = "605" Then
                        bolChkMemo605 = True
                        strExc(9) = PUB_GetReceiver(pa(1), pa(2), pa(3), pa(4), "605", "1")
                        'Modified by Morgan 2012/6/4 +pa26
                        'Modified by Morgan 2013/9/11 改抓設定檔
                        'strOldMemo605 = PUB_Get605Memo(strExc(9), RsTemp("pa26"), pa(1) & pa(2) & pa(3) & pa(4))
                        'Modified by Lydia 2022/08/02 整合模組：修改為複數新規則
                        'strOldMemo605 = PUB_GetNpMemo(pa(1) & pa(2) & pa(3) & pa(4), "605", strExc(9), RsTemp("pa26"))
                        strOldMemo605 = PUB_GetNpMemo2("1", pa(1) & pa(2) & pa(3) & pa(4), "605", strExc(9), RsTemp.Fields("PA26") & "," & RsTemp.Fields("PA27") & "," & RsTemp.Fields("PA28") & "," & RsTemp.Fields("PA29") & "," & RsTemp.Fields("PA30"))
                     End If
                  End If
                  'end 2011/6/13
   
                  strTxt(intStep) = "UPDATE PATENT SET " & strExc(2) & " WHERE " & ChgPatent(pa(1) & pa(2) & pa(3) & pa(4))
                  cnnConnection.Execute strTxt(intStep)
                  intStep = intStep + 1
                  
                  'Add by Morgan 2011/6/13
                  '若有期限則於更新資料後清除原來的備註並加入新的備註
                  'Modified by Morgan 2012/2/2 +pa75
                  strExc(0) = "select pa26,pa27,pa28,pa29,pa30,pa75 from patent where " & ChgPatent(pa(1) & pa(2) & pa(3) & pa(4))
                  intI = 1
                  Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                  If intI = 1 Then
                     If bolChkMemo416 Then
                        'Modified by Lydia 2022/08/02 整合模組：修改為複數新規則
                        'For intI = 0 To 4
                        '   If Not IsNull(RsTemp(intI)) Then
                        '      'Modified by Morgan 2012/2/2 +pa75
                        '      'Modified by Morgan 2013/9/11 改抓設定檔
                        '      'strNewMemo416 = PUB_Get416Memo(ChangeCustomerL(RsTemp(intI)), ChangeCustomerL("" & RsTemp("pa75")))
                        '      strNewMemo416 = PUB_GetNpMemo(pa(1) & pa(2) & pa(3) & pa(4), "416", ChangeCustomerL("" & RsTemp("pa75")), ChangeCustomerL(RsTemp(intI)))
                        '      If strNewMemo416 <> "" Then Exit For
                        '   End If
                        'Next
                        strNewMemo416 = PUB_GetNpMemo2("1", pa(1) & pa(2) & pa(3) & pa(4), "416", ChangeCustomerL("" & RsTemp.Fields("pa75")), RsTemp.Fields("PA26") & "," & RsTemp.Fields("PA27") & "," & RsTemp.Fields("PA28") & "," & RsTemp.Fields("PA29") & "," & RsTemp.Fields("PA30"))
                        'end 2022/08/02
                        
                        If strNewMemo416 <> strOldMemo416 Then
                           If strOldMemo416 <> "" Then 'Added by Lydia 2022/08/02
                             strSql = "update nextprogress set np15=replace(np15,'" & ChgSQL(strOldMemo416) & "','')" & _
                                " where " & ChgNextProgress(pa(1) & pa(2) & pa(3) & pa(4)) & " and np06 is null and np07=416"
                             cnnConnection.Execute strSql, intI
                           End If 'Added by Lydia 2022/08/02
                           If strNewMemo416 <> "" Then 'Added by Lydia 2022/08/02
                             strSql = "update nextprogress set np15='" & ChgSQL(strNewMemo416) & "'||';'||np15" & _
                                " where " & ChgNextProgress(pa(1) & pa(2) & pa(3) & pa(4)) & " and np06 is null and np07=416"
                             cnnConnection.Execute strSql, intI
                           End If  'Added by Lydia 2022/08/02
                        End If
                       
                     ElseIf bolChkMemo605 = True Then
                        strExc(9) = PUB_GetReceiver(pa(1), pa(2), pa(3), pa(4), "605", "1")
                        'Modified by Morgan 2012/6/4 +pa26
                        'Modified by Morgan 2013/9/11 改抓設定檔
                        'strNewMemo605 = PUB_Get605Memo(strExc(9), RsTemp("pa26"), pa(1) & pa(2) & pa(3) & pa(4))
                        'Modified by Lydia 2022/08/02 整合模組：修改為複數新規則
                        'strNewMemo605 = PUB_GetNpMemo(pa(1) & pa(2) & pa(3) & pa(4), "605", strExc(9), RsTemp("pa26"))
                        strNewMemo605 = PUB_GetNpMemo2("1", pa(1) & pa(2) & pa(3) & pa(4), "605", strExc(9), RsTemp.Fields("PA26") & "," & RsTemp.Fields("PA27") & "," & RsTemp.Fields("PA28") & "," & RsTemp.Fields("PA29") & "," & RsTemp.Fields("PA30"))

                        If strNewMemo605 <> strOldMemo605 Then
                          If strOldMemo605 <> "" Then 'Added by Lydia 2022/08/02
                             strSql = "update nextprogress set np15=replace(np15,'" & ChgSQL(strOldMemo605) & "','')" & _
                                " where " & ChgNextProgress(pa(1) & pa(2) & pa(3) & pa(4)) & " and np06 is null and np07=605"
                             cnnConnection.Execute strSql, intI
                          End If 'Added by Lydia 2022/08/02
                          If strNewMemo605 <> "" Then 'Added by Lydia 2022/08/02
                            strSql = "update nextprogress set np15='" & ChgSQL(strNewMemo605) & "'||';'||np15" & _
                               " where " & ChgNextProgress(pa(1) & pa(2) & pa(3) & pa(4)) & " and np06 is null and np07=605"
                            cnnConnection.Execute strSql, intI
                          End If 'Added by Lydia 2022/08/02
                        End If
                     End If
                  End If
                  'end 2011/6/13
               End If
            End If 'Addedd 2011/11/23
         End If 'Addedd 2011/11/23

      End If
   End If
   
   'Add by Morgan 2005/8/1
   'Modify By Sindy 2018/8/10
   cp(110) = m_CP110
   Dim strCon As String
   'If lstNameAgent.Visible = True Then
      'Modify By Sindy 2019/3/4
      If Pub_StrUserSt15 <> "P12" Then
         cp(84) = Val(txtCP84)
         strCon = strCon & ",cp84=" & cp(84)
      End If
      '2019/3/4 END
      'Modify By Sindy 2020/3/12 m_CP118isY=空白時,代表從發文作業呼叫此視窗,不需更新CP118欄位
      If m_CP118isY <> "" Then
      '2020/3/12 END
         If m_CP118isY = "Y" Then
            cp(118) = "A"
         Else
            cp(118) = ""
         End If
      End If
      strCon = strCon & ",cp118=" & CNULL(cp(118))
      'Modify By Sindy 2023/6/27
      If m_CP20isN = True And Val(cp(84)) = 300 Then '不請款且有發文規費:則更改的進度檔的相關資料改如下:費用0→規費300→點數-0.3
         strCon = strCon & ",cp20='N',cp16=0,cp17=300,cp18=-0.3"
      End If
      '2023/6/27 END
      strSql = " UPDATE CASEPROGRESS SET cp22=" & CNULL(m_CP22) & ",cp110=" & CNULL(m_CP110) & strCon & " WHERE CP09='" & strReceiveNo & "'"
      cnnConnection.Execute strSql
   'End If
   '2018/8/10 END
   
   cnnConnection.CommitTrans
   FormSave = True
   
CheckingErr:
   If Err.NUMBER <> 0 Then
      cnnConnection.RollbackTrans
      MsgBox Err.Description, vbCritical
   End If
   
End Function

Private Sub Form_Unload(Cancel As Integer)
   Select Case intGo
      Case 41
         'Modify by Morgan 2011/10/5
         'frm040103_1.Show
         oParent.Show
      Case 42
         'Modify by Morgan 2011/10/5
         'frm04010302_1.Show
         oParent.Show
      Case 43
         'Modify by Morgan 2011/10/5
         'frm040104_3.Show
         oParent.Show
      Case 44
         'Modify by Morgan 2011/10/5
         'frm040104_a.Show
         oParent.Show
      Case 45
         'Modify by Morgan 2011/10/5
         'frm040104_7.Show
         oParent.Show
         
      Case 46
         'Modify by Morgan 2011/10/5
         'frm040104_c.Show
         oParent.Show
         
      'Add by Morgan 2004/6/21
      Case 47
         'Modify by Morgan 2011/10/5
         'frm040104_e.Show
         oParent.Show
         
      Case 61
         'Modify by Morgan 2011/10/5
         'frm060103_1.Show
         'frm060103_1.ClearForm
         oParent.Show
         oParent.ClearForm
         
      Case 62
         'Modify by Morgan 2011/10/5
         'frm06010302_1.Show
         oParent.Show
         
      'Remove by Morgan 2011/10/5 不可能自己呼叫自己
      'Case 63
      '   frm06010303_1.Show
         
      'Add by Morgan 2006/7/4
      Case 64
         'Modify by Morgan 2011/10/5
         'frm060104_3.Show
         'frm060104_3.m_bolSaveChgEvent = bolSaveOK 'Add by Morgan 2006/7/4
         oParent.Show
         oParent.m_bolSaveChgEvent = bolSaveOK
         
   End Select
   Set frm06010303_1 = Nothing
End Sub

Private Sub Text10_GotFocus(Index As Integer)
   'edit by nickc 2007/07/11 切換輸入法改用API
   'Text10(Index).IMEMode = "1"
   OpenIme
   TextInverse Text10(Index)
End Sub

Private Sub Text10_Validate(Index As Integer, Cancel As Boolean)
   'edit by nickc 2007/07/11 切換輸入法改用API
   'Text10(Index).IMEMode = "2"
   CloseIme
End Sub

Private Sub Text34_GotFocus()
  TextInverse Text34
End Sub

Private Sub Text34_Validate(Cancel As Boolean)
   If Text34 <> "" Then
      If Not ChkDate(Text34) Then Cancel = True
   End If
   If Cancel = True Then TextInverse Text34
End Sub

Private Sub Text36_GotFocus()
  TextInverse Text36
End Sub

Private Sub Text36_Validate(Cancel As Boolean)
   If Check1(6).Value = 1 Then
      If ChgType(36) = False Then Cancel = True
   End If
End Sub

Private Sub Text38_GotFocus(Index As Integer)
  TextInverse Text38(Index)
End Sub

Private Sub Text38_Validate(Index As Integer, Cancel As Boolean)
   If Check1(7).Value = 1 And Index = 2 And (Text38(0) = "" And Text38(1) = "" And Text38(2) = "") Then
      MsgBox "案件名稱不可同時空白 !", vbCritical
      Text38(0).SetFocus
   End If
End Sub

Private Sub Text41_GotFocus()
  TextInverse Text41
End Sub

Private Sub Text41_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 78 And KeyAscii <> 8 Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub Text44_GotFocus()
  TextInverse Text44
End Sub

Private Sub Text46_GotFocus()
  TextInverse Text46
End Sub

Private Sub Text48_GotFocus()
  TextInverse Text48
End Sub

Private Sub Text5_GotFocus()
  TextInverse Text5
End Sub

Private Sub Text5_Validate(Cancel As Boolean)
   Cancel = Not ChkLetterDate(Text5.Text)
   If Cancel = True Then TextInverse Text5
End Sub

Private Sub Text50_GotFocus()
  TextInverse Text50
End Sub

Private Sub Text52_GotFocus()
  TextInverse Text52
End Sub

Private Sub Text54_GotFocus()
  TextInverse Text54
End Sub

Private Sub Text56_GotFocus()
  TextInverse Text56
End Sub

Private Sub Text58_GotFocus()
  TextInverse Text58
End Sub

Private Sub Text6_GotFocus(Index As Integer)
   'edit by nickc 2007/07/11 切換輸入法改用API
   'Text6(Index).IMEMode = "1"
   OpenIme
     TextInverse Text6(Index)
End Sub

Private Sub Text6_Validate(Index As Integer, Cancel As Boolean)
   'edit by nickc 2007/07/11 切換輸入法改用API
   'Text6(Index).IMEMode = "2"
   CloseIme
End Sub

Private Sub Text60_GotFocus()
  TextInverse Text60
End Sub

Private Sub Text62_GotFocus()
  TextInverse Text62
End Sub

Private Sub Text62_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 89 And KeyAscii <> 8 Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub Text7_GotFocus(Index As Integer)
  TextInverse Text7(Index)
End Sub

Private Sub Text7_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text7_Validate(Index As Integer, Cancel As Boolean)
   If Text7(Index) <> "" Then
      If ChgType(Index) = False Then
         Cancel = True
         TextInverse Text7(Index)
      End If
   End If
   If Index = 4 Then IntoCombol
End Sub

Private Function ChgType(i As Integer) As Boolean
 Dim strTempName As String
   ChgType = False
   Select Case i
      Case 0, 1, 2, 3, 4
         'edit by nickc 2007/02/05 不用 dll 了
         'If objLawDll.LawGetName(Text7(i), strTempName) Then
         If ClsLawLawGetName(Text7(i), strTempName) Then
            Label4(i) = strTempName
            ChgType = True
         End If
      Case 36
         'edit by nickc 2007/02/02 不用 dll 了
         'If objPublicData.GetPatentTrademarkKind(專利, Text36.Text, strTempName, False, 台灣國家代號) = 1 Then
         If ClsPDGetPatentTrademarkKind(專利, Text36.Text, strTempName, False, 台灣國家代號) = 1 Then
            Label3(8) = strTempName
            'Add By Cheng 2002/05/17
            ChgType = True
         End If
   End Select
End Function

Private Sub Text8_GotFocus(Index As Integer)
  TextInverse Text8(Index)
End Sub

Private Sub Text9_GotFocus(Index As Integer)
  TextInverse Text9(Index)
End Sub

'Add By Cheng 2002/05/22
Private Function TxtValidate() As Boolean
Dim objTxt As Object
Dim ii As Integer  '2015/9/25 add by sonia
Dim Cancel As Boolean
   
   TxtValidate = False
   If Me.Text34.Enabled = True Then
      Cancel = False
      Text34_Validate Cancel
      If Cancel = True Then
         Me.Text34.SetFocus
         Text34_GotFocus
         Exit Function
      End If
   End If
   
   If Me.Text36.Enabled = True Then
      Cancel = False
      Text36_Validate Cancel
      If Cancel = True Then
         Me.Text36.SetFocus
         Text36_GotFocus
         Exit Function
      End If
   End If
   
   For Each objTxt In Text38
      If objTxt.Enabled = True Then
         Cancel = False
         Text38_Validate objTxt.Index, Cancel
         If Cancel = True Then
            Me.Text38(objTxt.Index).SetFocus
            Text38_GotFocus objTxt.Index
            Exit Function
         End If
      End If
   Next
   
   If Me.Text5.Enabled = True Then
      Cancel = False
      Text5_Validate Cancel
      If Cancel = True Then
         Me.Text5.SetFocus
         Text5_GotFocus
         Exit Function
      End If
   End If
   
   For Each objTxt In Text7
      If objTxt.Enabled = True Then
         Cancel = False
         Text7_Validate objTxt.Index, Cancel
         If Cancel = True Then
            Me.Text7(objTxt.Index).SetFocus
            Text7_GotFocus objTxt.Index
            Exit Function
         End If
      End If
   Next
   
   'Add by Morgan 2003/12/07
   Dim objCheck As Object, bolCheck As Boolean
   bolCheck = False
   For Each objCheck In Check1
      If objCheck.Value = 1 Then
         bolCheck = True
         Exit For
      End If
   Next
   If bolCheck = False Then
      MsgBox "至少需勾選一項變更！", vbCritical
      Exit Function
   End If
   'End 2003/12/07

   If Check1(0).Value = 1 Then
      If Text7(0) = "" And Text7(1) = "" And Text7(2) = "" And Text7(3) = "" And Text7(4) = "" Then
         MsgBox "有勾選申請人時，申請人不可空白 !", vbCritical
         SSTab1.Tab = 0
         Text7(0).SetFocus
         Exit Function
      End If
      'Add by Morgan 2007/1/11
      'modify by sonia 2015/9/25 P-111611 二申請人變更為一申請人,故改為提醒
      'For intI = 0 To 4
      '   If pa(26 + intI) <> "" And Text7(intI) = "" Then
      '      MsgBox "本案有申請人" & Format(1 + intI) & "，不可空白！"
      '      SSTab1.Tab = 0
      '      Text7(intI).SetFocus
      '      Exit Function
      '   End If
      'Next
      ii = 1
      For intI = 1 To 4
         If pa(26 + intI) <> "" And Text7(intI) = "" Then
            ii = ii + 1
         End If
      Next
      If ii > 1 Then
         If MsgBox("本案原有" & ii & "個申請人，確定人數減少？", vbYesNo + vbDefaultButton2 + vbInformation) = vbNo Then
            SSTab1.Tab = 0
            Text7(ii - 1).SetFocus
            Exit Function
         End If
      End If
      'end 2015/9/25
      
      'Add by Morgan 2011/8/2
      '若有勾申請人時需檢查輸入的申請人編號不可與原來的完全相同(九碼)
      If ChangeCustomerL(Text7(0)) = ChangeCustomerL(pa(26)) And _
          ChangeCustomerL(Text7(1)) = ChangeCustomerL(pa(27)) And _
           ChangeCustomerL(Text7(2)) = ChangeCustomerL(pa(28)) And _
            ChangeCustomerL(Text7(3)) = ChangeCustomerL(pa(29)) And _
             ChangeCustomerL(Text7(4)) = ChangeCustomerL(pa(30)) Then
         MsgBox "新申請人編號與目前相同 !", vbCritical
         SSTab1.Tab = 0
         Text7(0).SetFocus
         Exit Function
      End If
   End If

   'Add by Morgan 2005/8/1
   If lstNameAgent.Visible = True Then
      Cancel = False
      lstNameAgent_Validate Cancel
      If Cancel = True Then
         SSTab1.Tab = 0
         lstNameAgent.SetFocus
         Exit Function
      End If
   End If
   
   'Add By Sindy 2018/9/17
'   If Label12.Visible = True Then
'      txtCP84 = Val(txtCP84) + 300
'      txtCP84.Tag = txtCP84.Text
'   End If
   If Val(txtCP84) > 0 And txtCP84.Enabled Then
      Cancel = False
      txtCP84_Validate Cancel
      If Cancel = True Then
         SSTab1.Tab = 0
'         If Label12.Visible = True Then txtCP84 = Val(txtCP84) - 300
         txtCP84.SetFocus
         Exit Function
      End If
   End If
   '2018/9/17 END
   
   TxtValidate = True
End Function

'Add By Cheng 2003/01/21
'取得規費金額
Private Function GetCP17(strCP09 As String) As String
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
    GetCP17 = "0"
    StrSQLa = "Select CP17 From CaseProgress Where CP09 ='" & strCP09 & "' "
    rsA.CursorLocation = adUseClient
    rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
    If rsA.RecordCount > 0 Then
        GetCP17 = "" & rsA.Fields(0).Value
    End If
    If GetCP17 = "" Then GetCP17 = "0"
    If rsA.State <> adStateClosed Then rsA.Close
    Set rsA = Nothing
End Function
'Add by Morgan 2005/8/1
'檢查並設定cp110資料
Private Sub lstNameAgent_Validate(Cancel As Boolean)
   Dim ii As Integer, bolCheck As Boolean
   bolCheck = False
   m_CP110 = "": m_AgentName = ""
   For ii = 0 To lstNameAgent.ListCount - 1
      If lstNameAgent.Selected(ii) = True Then
         'modify by sonia 2016/10/7 員工編號已可非數字需做轉換
         'm_CP110 = m_CP110 & "," & lstNameAgent.ItemData(ii)
         'Modify By Sindy 2021/5/10
         'm_CP110 = m_CP110 & "," & PUB_Num2Id(lstNameAgent.ItemData(ii))
         m_CP110 = m_CP110 & "," & PUB_GetItemData(lstNameAgent.Tag, ii)
         m_AgentName = m_AgentName & "、" & lstNameAgent.List(ii)
         '2021/5/10 END
         bolCheck = True
      End If
   Next
   If Left(m_CP110, 1) = "," Then m_CP110 = Mid(m_CP110, 2)
   If Left(m_AgentName, 1) = "、" Then m_AgentName = Mid(m_AgentName, 2) 'Add By Sindy 2021/5/10
   If bolCheck = True Then
      m_CP22 = ""
   Else
      m_CP22 = "N"
      If MsgBox("未勾選代理人，確定不出名？", vbYesNo + vbDefaultButton2 + vbInformation) = vbNo Then
         Cancel = True
      End If
   End If
End Sub

Private Sub txtCP84_GotFocus()
   TextInverse txtCP84
   CloseIme
End Sub

Private Sub txtCP84_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 8 And Not IsNumeric(Chr(KeyAscii)) Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub txtCP84_Validate(Cancel As Boolean)
   'Modify By Sindy 2019/3/4
   If Pub_StrUserSt15 <> "P12" Then
   '2019/3/4 END
      '台灣
      If pa(9) = "000" Then
         If Val(txtCP84.Text) <> Val(cp(17)) And Val(txtCP84.Text) <> Val(txtCP84.Tag) Then
            If MsgBox("發文規費【" & txtCP84.Text & "】與收文規費【" & cp(17) & "】不同，確定要繼續！", vbInformation + vbYesNo + vbDefaultButton2) = vbYes Then
               txtCP84.Tag = txtCP84.Text
            Else
               txtCP84_GotFocus
               Cancel = True
            End If
         End If
      End If
   End If
End Sub

'Added by Lydia 2020/02/21 外專：案件名稱有特殊字，開啟FCP0xxxxx.新案性質.案件名稱.doc
Private Sub CmdPA174_Click()

    If pa(1) = "" Or pa(2) = "" Or pa(3) = "" Or pa(4) = "" Then Exit Sub
    If Pub_GetPA174toFile("0", pa(1), pa(2), pa(3), pa(4), Me, frm100101_M_1) = True Then
    End If
    
End Sub
