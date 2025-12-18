VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm03020604_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "各式申請書-更正"
   ClientHeight    =   5424
   ClientLeft      =   72
   ClientTop       =   996
   ClientWidth     =   9168
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5424
   ScaleWidth      =   9168
   Begin VB.TextBox txtFee 
      Height          =   270
      Left            =   7560
      Locked          =   -1  'True
      MaxLength       =   6
      TabIndex        =   100
      Top             =   3195
      Width           =   735
   End
   Begin VB.Frame Frame1 
      Caption         =   "附送書件"
      Height          =   1760
      Left            =   6960
      TabIndex        =   101
      Top             =   3540
      Width           =   1875
      Begin VB.CheckBox chkAtt1 
         Caption         =   "註冊證、核准函"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   109
         Tag             =   ".ATT.pdf"
         Top             =   1470
         Value           =   1  '核取
         Width           =   1695
      End
      Begin VB.CheckBox chkAtt1 
         Caption         =   "更名證明文件"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   106
         Tag             =   ".change.pdf"
         Top             =   1245
         Width           =   1695
      End
      Begin VB.CheckBox chkAtt1 
         Caption         =   "移轉契約"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   105
         Tag             =   ".asasignment.pdf"
         Top             =   990
         Width           =   1695
      End
      Begin VB.CheckBox chkAtt1 
         Caption         =   "優先權證明文件"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   104
         Tag             =   ".PRI.pdf"
         Top             =   735
         Width           =   1695
      End
      Begin VB.CheckBox chkAtt1 
         Caption         =   "委任書"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   103
         Tag             =   ".poa.pdf"
         Top             =   495
         Width           =   1215
      End
      Begin VB.CheckBox chkAtt1 
         Caption         =   "基本資料表"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   102
         Tag             =   ".contact.pdf"
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   270
      Left            =   1230
      MaxLength       =   7
      TabIndex        =   5
      Top             =   2550
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   2
      Left            =   8250
      TabIndex        =   36
      Top             =   75
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   6300
      TabIndex        =   34
      Top             =   75
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "回前畫面(&U)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   7140
      TabIndex        =   35
      Top             =   75
      Width           =   1080
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      ItemData        =   "frm03020604_1.frx":0000
      Left            =   1260
      List            =   "frm03020604_1.frx":000D
      Style           =   2  '單純下拉式
      TabIndex        =   4
      Top             =   1185
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   1020
      MaxLength       =   3
      TabIndex        =   0
      Top             =   510
      Width           =   550
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   1575
      MaxLength       =   6
      TabIndex        =   1
      Top             =   510
      Width           =   855
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   2415
      MaxLength       =   1
      TabIndex        =   2
      Top             =   510
      Width           =   255
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   2655
      MaxLength       =   2
      TabIndex        =   3
      Top             =   510
      Width           =   375
   End
   Begin VB.TextBox Text7 
      Height          =   270
      Left            =   4200
      MaxLength       =   1
      TabIndex        =   6
      Top             =   2550
      Width           =   300
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   2355
      Left            =   180
      TabIndex        =   59
      Top             =   3000
      Width           =   6225
      _ExtentX        =   10986
      _ExtentY        =   4149
      _Version        =   393216
      Style           =   1
      Tabs            =   5
      TabsPerRow      =   5
      TabHeight       =   520
      TabCaption(0)   =   "註冊證"
      TabPicture(0)   =   "frm03020604_1.frx":001D
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Check1(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Check1(1)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Check1(2)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Check1(3)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Check1(4)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Check1(5)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "註冊變更核准"
      TabPicture(1)   =   "frm03020604_1.frx":0039
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Check2(5)"
      Tab(1).Control(1)=   "Check2(4)"
      Tab(1).Control(2)=   "Check2(3)"
      Tab(1).Control(3)=   "Check2(2)"
      Tab(1).Control(4)=   "Check2(1)"
      Tab(1).Control(5)=   "Check2(0)"
      Tab(1).Control(6)=   "Check2(6)"
      Tab(1).ControlCount=   7
      TabCaption(2)   =   "延展核准"
      TabPicture(2)   =   "frm03020604_1.frx":0055
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Check3(3)"
      Tab(2).Control(1)=   "Check3(2)"
      Tab(2).Control(2)=   "Check3(1)"
      Tab(2).Control(3)=   "Check3(0)"
      Tab(2).ControlCount=   4
      TabCaption(3)   =   "移轉核准"
      TabPicture(3)   =   "frm03020604_1.frx":0071
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Check4(4)"
      Tab(3).Control(1)=   "Check4(3)"
      Tab(3).Control(2)=   "Check4(2)"
      Tab(3).Control(3)=   "Check4(1)"
      Tab(3).Control(4)=   "Check4(0)"
      Tab(3).ControlCount=   5
      TabCaption(4)   =   "授權核准"
      TabPicture(4)   =   "frm03020604_1.frx":008D
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Check5(4)"
      Tab(4).Control(1)=   "Check5(3)"
      Tab(4).Control(2)=   "Check5(2)"
      Tab(4).Control(3)=   "Check5(1)"
      Tab(4).Control(4)=   "Check5(0)"
      Tab(4).ControlCount=   5
      Begin VB.CheckBox Check5 
         Caption         =   "商標名稱："
         Height          =   255
         Index           =   0
         Left            =   -74640
         TabIndex        =   29
         Top             =   420
         Width           =   3795
      End
      Begin VB.CheckBox Check5 
         Caption         =   "被授權人中文名稱："
         Height          =   255
         Index           =   1
         Left            =   -74640
         TabIndex        =   30
         Top             =   690
         Width           =   4845
      End
      Begin VB.CheckBox Check5 
         Caption         =   "被授權人英文名稱："
         Height          =   255
         Index           =   2
         Left            =   -74640
         TabIndex        =   31
         Top             =   960
         Width           =   3735
      End
      Begin VB.CheckBox Check5 
         Caption         =   "授權期間："
         Height          =   255
         Index           =   3
         Left            =   -74640
         TabIndex        =   32
         Top             =   1245
         Width           =   3735
      End
      Begin VB.CheckBox Check5 
         Caption         =   "授權商品："
         Height          =   255
         Index           =   4
         Left            =   -74640
         TabIndex        =   33
         Top             =   1500
         Width           =   3735
      End
      Begin VB.CheckBox Check4 
         Caption         =   "商標名稱："
         Height          =   255
         Index           =   0
         Left            =   -74670
         TabIndex        =   24
         Top             =   420
         Width           =   3795
      End
      Begin VB.CheckBox Check4 
         Caption         =   "受讓人中文名稱："
         Height          =   255
         Index           =   1
         Left            =   -74670
         TabIndex        =   25
         Top             =   690
         Width           =   4845
      End
      Begin VB.CheckBox Check4 
         Caption         =   "受讓人英文名稱："
         Height          =   255
         Index           =   2
         Left            =   -74670
         TabIndex        =   26
         Top             =   960
         Width           =   3735
      End
      Begin VB.CheckBox Check4 
         Caption         =   "受讓人中文地址："
         Height          =   255
         Index           =   3
         Left            =   -74670
         TabIndex        =   27
         Top             =   1245
         Width           =   3735
      End
      Begin VB.CheckBox Check4 
         Caption         =   "受讓人英文地址："
         Height          =   255
         Index           =   4
         Left            =   -74670
         TabIndex        =   28
         Top             =   1500
         Width           =   3735
      End
      Begin VB.CheckBox Check3 
         Caption         =   "商標權人："
         Height          =   255
         Index           =   0
         Left            =   -74640
         TabIndex        =   20
         Top             =   420
         Width           =   3795
      End
      Begin VB.CheckBox Check3 
         Caption         =   "商標名稱："
         Height          =   255
         Index           =   1
         Left            =   -74640
         TabIndex        =   21
         Top             =   690
         Width           =   4845
      End
      Begin VB.CheckBox Check3 
         Caption         =   "核准延展之權利期限："
         Height          =   255
         Index           =   2
         Left            =   -74640
         TabIndex        =   22
         Top             =   960
         Width           =   3735
      End
      Begin VB.CheckBox Check3 
         Caption         =   "核准延展之商品："
         Height          =   255
         Index           =   3
         Left            =   -74640
         TabIndex        =   23
         Top             =   1245
         Width           =   3735
      End
      Begin VB.CheckBox Check2 
         Caption         =   "代表人英文名稱："
         Height          =   255
         Index           =   6
         Left            =   -74640
         TabIndex        =   19
         Top             =   2010
         Width           =   3735
      End
      Begin VB.CheckBox Check2 
         Caption         =   "商標名稱："
         Height          =   255
         Index           =   0
         Left            =   -74640
         TabIndex        =   13
         Top             =   390
         Width           =   3795
      End
      Begin VB.CheckBox Check2 
         Caption         =   "商標權人中文名稱："
         Height          =   255
         Index           =   1
         Left            =   -74640
         TabIndex        =   14
         Top             =   660
         Width           =   4845
      End
      Begin VB.CheckBox Check2 
         Caption         =   "商標權人英文名稱："
         Height          =   255
         Index           =   2
         Left            =   -74640
         TabIndex        =   15
         Top             =   930
         Width           =   3735
      End
      Begin VB.CheckBox Check2 
         Caption         =   "申請人中文地址："
         Height          =   255
         Index           =   3
         Left            =   -74640
         TabIndex        =   16
         Top             =   1215
         Width           =   3735
      End
      Begin VB.CheckBox Check2 
         Caption         =   "申請人英文地址："
         Height          =   255
         Index           =   4
         Left            =   -74640
         TabIndex        =   17
         Top             =   1470
         Width           =   3735
      End
      Begin VB.CheckBox Check2 
         Caption         =   "代表人中文名稱："
         Height          =   255
         Index           =   5
         Left            =   -74640
         TabIndex        =   18
         Top             =   1740
         Width           =   3735
      End
      Begin VB.CheckBox Check1 
         Caption         =   "寄存證明正本一份"
         Height          =   255
         Index           =   13
         Left            =   -74760
         TabIndex        =   97
         Top             =   1740
         Width           =   3735
      End
      Begin VB.CheckBox Check2 
         Caption         =   "容後補呈"
         Height          =   255
         Index           =   13
         Left            =   -69600
         TabIndex        =   96
         Top             =   1740
         Width           =   1095
      End
      Begin VB.CheckBox Check2 
         Caption         =   "容後補呈"
         Height          =   255
         Index           =   22
         Left            =   -70500
         TabIndex        =   95
         Top             =   1410
         Width           =   1215
      End
      Begin VB.CheckBox Check1 
         Caption         =   "修正規費壹仟元整"
         Height          =   255
         Index           =   22
         Left            =   -74760
         TabIndex        =   94
         Top             =   1410
         Width           =   2775
      End
      Begin VB.CheckBox Check2 
         Caption         =   "容後補呈"
         Height          =   255
         Index           =   26
         Left            =   -70500
         TabIndex        =   93
         Top             =   2385
         Width           =   1215
      End
      Begin VB.CheckBox Check2 
         Caption         =   "容後補呈"
         Height          =   255
         Index           =   25
         Left            =   -70500
         TabIndex        =   92
         Top             =   2145
         Width           =   1215
      End
      Begin VB.CheckBox Check2 
         Caption         =   "容後補呈"
         Height          =   255
         Index           =   24
         Left            =   -70500
         TabIndex        =   91
         Top             =   1905
         Width           =   1215
      End
      Begin VB.CheckBox Check2 
         Caption         =   "容後補呈"
         Height          =   255
         Index           =   23
         Left            =   -70500
         TabIndex        =   90
         Top             =   1665
         Width           =   1215
      End
      Begin VB.CheckBox Check2 
         Caption         =   "容後補呈"
         Height          =   255
         Index           =   21
         Left            =   -70500
         TabIndex        =   89
         Top             =   1152
         Width           =   1215
      End
      Begin VB.CheckBox Check1 
         Caption         =   "終止授權契約書正本一份"
         Height          =   255
         Index           =   26
         Left            =   -74760
         TabIndex        =   88
         Top             =   2385
         Width           =   2775
      End
      Begin VB.CheckBox Check1 
         Caption         =   "授權契約書正本一份"
         Height          =   255
         Index           =   25
         Left            =   -74760
         TabIndex        =   87
         Top             =   2145
         Width           =   2775
      End
      Begin VB.CheckBox Check1 
         Caption         =   "讓與契約書正本一份"
         Height          =   255
         Index           =   24
         Left            =   -74760
         TabIndex        =   86
         Top             =   1905
         Width           =   2775
      End
      Begin VB.CheckBox Check1 
         Caption         =   "證書正本一份"
         Height          =   255
         Index           =   23
         Left            =   -74760
         TabIndex        =   85
         Top             =   1665
         Width           =   2775
      End
      Begin VB.CheckBox Check1 
         Caption         =   "修正規費參佰元整"
         Height          =   255
         Index           =   21
         Left            =   -74760
         TabIndex        =   84
         Top             =   1152
         Width           =   2775
      End
      Begin VB.CheckBox Check2 
         Caption         =   "容後補呈"
         Height          =   255
         Index           =   20
         Left            =   -70500
         TabIndex        =   83
         Top             =   912
         Width           =   1215
      End
      Begin VB.CheckBox Check2 
         Caption         =   "容後補呈"
         Height          =   255
         Index           =   19
         Left            =   -70500
         TabIndex        =   82
         Top             =   672
         Width           =   1215
      End
      Begin VB.CheckBox Check2 
         Caption         =   "容後補呈"
         Height          =   255
         Index           =   18
         Left            =   -70500
         TabIndex        =   81
         Top             =   432
         Width           =   1215
      End
      Begin VB.CheckBox Check2 
         Caption         =   "容後補呈"
         Height          =   255
         Index           =   17
         Left            =   -69600
         TabIndex        =   80
         Top             =   2715
         Width           =   1095
      End
      Begin VB.CheckBox Check2 
         Caption         =   "容後補呈"
         Height          =   255
         Index           =   16
         Left            =   -69600
         TabIndex        =   79
         Top             =   2475
         Width           =   1095
      End
      Begin VB.CheckBox Check2 
         Caption         =   "容後補呈"
         Height          =   255
         Index           =   15
         Left            =   -69600
         TabIndex        =   78
         Top             =   2235
         Width           =   1095
      End
      Begin VB.CheckBox Check2 
         Caption         =   "容後補呈"
         Height          =   255
         Index           =   14
         Left            =   -69600
         TabIndex        =   77
         Top             =   1995
         Width           =   1095
      End
      Begin VB.CheckBox Check2 
         Caption         =   "容後補呈"
         Height          =   255
         Index           =   12
         Left            =   -69600
         TabIndex        =   76
         Top             =   1485
         Width           =   1095
      End
      Begin VB.CheckBox Check2 
         Caption         =   "容後補呈"
         Height          =   255
         Index           =   11
         Left            =   -69600
         TabIndex        =   75
         Top             =   1245
         Width           =   1095
      End
      Begin VB.CheckBox Check2 
         Caption         =   "容後補呈"
         Height          =   255
         Index           =   10
         Left            =   -69600
         TabIndex        =   74
         Top             =   1005
         Width           =   1095
      End
      Begin VB.CheckBox Check2 
         Caption         =   "容後補呈"
         Height          =   255
         Index           =   9
         Left            =   -69600
         TabIndex        =   73
         Top             =   765
         Width           =   1095
      End
      Begin VB.CheckBox Check2 
         Caption         =   "容後補呈"
         Height          =   255
         Index           =   8
         Left            =   -69600
         TabIndex        =   72
         Top             =   525
         Width           =   1095
      End
      Begin VB.CheckBox Check1 
         Caption         =   "切結書正本一份"
         Height          =   255
         Index           =   20
         Left            =   -74760
         TabIndex        =   71
         Top             =   912
         Width           =   2775
      End
      Begin VB.CheckBox Check1 
         Caption         =   "圖式修正本一式三份"
         Height          =   255
         Index           =   19
         Left            =   -74760
         TabIndex        =   70
         Top             =   672
         Width           =   2775
      End
      Begin VB.CheckBox Check1 
         Caption         =   "原文說明書一式二份"
         Height          =   255
         Index           =   18
         Left            =   -74760
         TabIndex        =   69
         Top             =   432
         Width           =   2775
      End
      Begin VB.CheckBox Check1 
         Caption         =   "國籍證明書正本一份"
         Height          =   255
         Index           =   17
         Left            =   -74760
         TabIndex        =   68
         Top             =   2715
         Width           =   3735
      End
      Begin VB.CheckBox Check1 
         Caption         =   "法人地位證明書正本一份"
         Height          =   255
         Index           =   16
         Left            =   -74760
         TabIndex        =   67
         Top             =   2475
         Width           =   3735
      End
      Begin VB.CheckBox Check1 
         Caption         =   "繼承證明正本一份"
         Height          =   255
         Index           =   15
         Left            =   -74760
         TabIndex        =   66
         Top             =   2235
         Width           =   3735
      End
      Begin VB.CheckBox Check1 
         Caption         =   "死亡證明正本一份"
         Height          =   255
         Index           =   14
         Left            =   -74760
         TabIndex        =   65
         Top             =   1995
         Width           =   3735
      End
      Begin VB.CheckBox Check1 
         Caption         =   "僱傭契約或經認證之讓與文件正本一份"
         Height          =   255
         Index           =   12
         Left            =   -74760
         TabIndex        =   64
         Top             =   1485
         Width           =   3735
      End
      Begin VB.CheckBox Check1 
         Caption         =   "發明人拒簽之切結書正本一份"
         Height          =   255
         Index           =   11
         Left            =   -74760
         TabIndex        =   63
         Top             =   1245
         Width           =   3735
      End
      Begin VB.CheckBox Check1 
         Caption         =   "圖卡一式二份"
         Height          =   255
         Index           =   10
         Left            =   -74760
         TabIndex        =   62
         Top             =   1005
         Width           =   3735
      End
      Begin VB.CheckBox Check1 
         Caption         =   "掛號回執影本一份"
         Height          =   255
         Index           =   9
         Left            =   -74760
         TabIndex        =   61
         Top             =   765
         Width           =   3735
      End
      Begin VB.CheckBox Check1 
         Caption         =   "原文申請專利範圍修正本一式三份"
         Height          =   255
         Index           =   8
         Left            =   -74760
         TabIndex        =   60
         Top             =   525
         Width           =   3735
      End
      Begin VB.CheckBox Check1 
         Caption         =   "商品名稱："
         Height          =   255
         Index           =   5
         Left            =   360
         TabIndex        =   12
         Top             =   1755
         Width           =   3735
      End
      Begin VB.CheckBox Check1 
         Caption         =   "權利期間："
         Height          =   255
         Index           =   4
         Left            =   360
         TabIndex        =   11
         Top             =   1485
         Width           =   3735
      End
      Begin VB.CheckBox Check1 
         Caption         =   "商標圖樣："
         Height          =   255
         Index           =   3
         Left            =   360
         TabIndex        =   10
         Top             =   1230
         Width           =   3735
      End
      Begin VB.CheckBox Check1 
         Caption         =   "商標名稱："
         Height          =   255
         Index           =   2
         Left            =   360
         TabIndex        =   9
         Top             =   945
         Width           =   3735
      End
      Begin VB.CheckBox Check1 
         Caption         =   "商標權人："
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   8
         Top             =   675
         Width           =   4845
      End
      Begin VB.CheckBox Check1 
         Caption         =   "商標註冊號數："
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   7
         Top             =   405
         Width           =   3795
      End
   End
   Begin MSForms.ListBox lstNameAgent 
      Height          =   315
      Left            =   7530
      TabIndex        =   108
      Top             =   2220
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
   Begin VB.Label lblFee 
      Caption         =   "規費 :"
      Height          =   180
      Left            =   6990
      TabIndex        =   107
      Top             =   3240
      Width           =   570
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "申請書日期 :"
      Height          =   180
      Left            =   180
      TabIndex        =   99
      Top             =   2595
      Visible         =   0   'False
      Width           =   990
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "出名代理人"
      Height          =   180
      Left            =   6570
      TabIndex        =   98
      Top             =   2220
      Width           =   900
   End
   Begin MSForms.Label Label12 
      Height          =   285
      Index           =   10
      Left            =   4920
      TabIndex        =   58
      Top             =   2220
      Width           =   1500
      VariousPropertyBits=   27
      Caption         =   "Label12"
      Size            =   "2646;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label12 
      Height          =   285
      Index           =   9
      Left            =   1260
      TabIndex        =   57
      Top             =   2220
      Width           =   1500
      VariousPropertyBits=   27
      Caption         =   "Label12"
      Size            =   "2646;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label29 
      AutoSize        =   -1  'True
      Caption         =   "案件性質:"
      Height          =   180
      Left            =   4020
      TabIndex        =   56
      Top             =   510
      Width           =   765
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      Caption         =   "機關文號:"
      Height          =   180
      Left            =   4020
      TabIndex        =   55
      Top             =   1938
      Width           =   765
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      Caption         =   "來函收文日:"
      Height          =   180
      Left            =   210
      TabIndex        =   54
      Top             =   1938
      Width           =   945
   End
   Begin MSForms.Label Label12 
      Height          =   285
      Index           =   0
      Left            =   4920
      TabIndex        =   53
      Top             =   510
      Width           =   1500
      VariousPropertyBits=   27
      Caption         =   "Label12"
      Size            =   "2646;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "智權人員:"
      Height          =   180
      Left            =   4020
      TabIndex        =   52
      Top             =   1536
      Width           =   765
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "承辦人　:"
      Height          =   180
      Left            =   210
      TabIndex        =   51
      Top             =   1536
      Width           =   765
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "本所案號:"
      Height          =   180
      Left            =   210
      TabIndex        =   50
      Top             =   510
      Width           =   765
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "申請案號:"
      Height          =   180
      Left            =   210
      TabIndex        =   49
      Top             =   852
      Width           =   765
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "審定號數:"
      Height          =   180
      Left            =   4020
      TabIndex        =   48
      Top             =   852
      Width           =   765
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "商標名稱:"
      Height          =   180
      Left            =   210
      TabIndex        =   47
      Top             =   1185
      Width           =   765
   End
   Begin MSForms.Label Label12 
      Height          =   285
      Index           =   1
      Left            =   1260
      TabIndex        =   46
      Top             =   852
      Width           =   1500
      VariousPropertyBits=   27
      Caption         =   "Label12"
      Size            =   "2646;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label12 
      Height          =   285
      Index           =   2
      Left            =   4920
      TabIndex        =   45
      Top             =   852
      Width           =   1500
      VariousPropertyBits=   27
      Caption         =   "Label12"
      Size            =   "2646;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label12 
      Height          =   285
      Index           =   3
      Left            =   1980
      TabIndex        =   44
      Top             =   1185
      Width           =   6930
      VariousPropertyBits=   27
      Caption         =   "Label12"
      Size            =   "12224;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label12 
      Height          =   285
      Index           =   4
      Left            =   1260
      TabIndex        =   43
      Top             =   1530
      Width           =   1500
      VariousPropertyBits=   27
      Caption         =   "Label12"
      Size            =   "2646;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label12 
      Height          =   285
      Index           =   5
      Left            =   4920
      TabIndex        =   42
      Top             =   1536
      Width           =   1500
      VariousPropertyBits=   27
      Caption         =   "Label12"
      Size            =   "2646;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label12 
      Height          =   285
      Index           =   6
      Left            =   1260
      TabIndex        =   41
      Top             =   1875
      Width           =   1500
      VariousPropertyBits=   27
      Caption         =   "Label12"
      Size            =   "2646;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label12 
      Height          =   285
      Index           =   7
      Left            =   4920
      TabIndex        =   40
      Top             =   1875
      Width           =   3960
      VariousPropertyBits=   27
      Caption         =   "Label12"
      Size            =   "6985;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      Caption         =   "是否修改申請書內容          (Y:WORD)"
      Height          =   180
      Index           =   1
      Left            =   2490
      TabIndex        =   39
      Top             =   2595
      Width           =   2880
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      Caption         =   "法定期限:"
      Height          =   180
      Index           =   0
      Left            =   4020
      TabIndex        =   38
      Top             =   2220
      Width           =   765
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      Caption         =   "本所期限:"
      Height          =   180
      Left            =   210
      TabIndex        =   37
      Top             =   2220
      Width           =   765
   End
End
Attribute VB_Name = "frm03020604_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Lydia 2021/08/04 Form2.0已修改; Label2(index)、lstNameAgent
'Memo By Sindy 2012/12/4 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/29 員工編號欄已修改
'Memo By Sindy 2010/8/11 日期欄已修改
Option Explicit

Dim strReceiveNo As String
Dim tm() As String, m_CP110 As String, m_AgentName As String
Dim intWhere As Integer, intLastRow As Integer
Dim m_strNPReceiveNo As String '點選未收的期限的收文號
Dim m_CP43 As String, m2_CP43 As String '相關總收文號
Dim m_CP10 As String, m2_CP10 As String, m3_CP10 As String '案件性質
Dim m2_CP27 As String '相關總收文號-發文日期 Add By Sindy 2012/6/1
Dim strCaseType As String
'Added by Lydia 2020/12/31
Dim m_CP118  As String '是否電子送件
Dim m_CaseNo As String '電子送件-本所案號
Dim m_F21st07 As String 'FCT程序分機
Dim strAppDetail As String '申請內容

Private Sub cmdok_Click(Index As Integer)
Dim bolChk As Boolean, strTmp As String
Dim i As Integer
'Added by Lydia 2020/12/31
Dim strFolder As String, strFileName As String
Dim strContent As String

   Select Case Index
      Case 0 '確定
         strCaseType = ""
         If m_CP10 = "302" And m2_CP10 = "1701" Then '更正註冊證
            strTmp = "01"
         ElseIf m2_CP10 = "1001" And m3_CP10 = "301" Then '註冊變更核准
            strTmp = "02"
         ElseIf m2_CP10 = "1001" And m3_CP10 = "102" Then '延展核准
            strTmp = "03"
         ElseIf m2_CP10 = "1001" And m3_CP10 = "501" Then '移轉核准
            strTmp = "04"
         ElseIf m2_CP10 = "1001" And m3_CP10 = "502" Then '授權核准
            strTmp = "05"
         'Add By Sindy 2012/6/1 ＋更正其他狀況時
         Else
            strTmp = "06"
            'Add By Sindy 2014/11/5
            If m3_CP10 = "101" Then '申請
               strCaseType = "註冊"
            Else
               TmSt = "TM01='" & tm(1) & "' AND TM02='" & tm(2) & "' AND TM03='" & _
                       tm(3) & "' AND TM04='" & tm(4) & "'"
               strCaseType = ExceptFieldData("商標狀況")
               If strCaseType = "申請" Then strCaseType = ""
            End If
            '2014/11/5 END
         '2012/6/1 End
         End If
         'Added by Lydia 2020/12/31 電子送件申請書=補正申請書
         If m_CP118 = "Y" Then
             strTmp = "10"
         End If
         'end 2020/12/31
         
         strLetterDate = Text5.Text
         If strTmp = "" Then
            MsgBox "該性質並無申請書！"
         'Added by Lydia 2020/12/31 電子送件申請書
         ElseIf m_CP118 = "Y" Then
            m_CaseNo = PUB_FCPCaseNo2FileName(tm(1), tm(2), tm(3), tm(4))
            '桌面上建立案號資料夾
            strFolder = PUB_Getdesktop
            strFolder = strFolder & "\" & m_CaseNo
            If Dir(strFolder, vbDirectory) = "" Then
                MkDir strFolder
            End If
            '申請書
            If StartLetter2("90", strTmp, strReceiveNo) = False Then Exit Sub
            '判斷要基本資料表,先不存檔
            If chkAtt1(0).Value = 1 Then
                 NowPrint strReceiveNo, "90", strTmp, False, strUserNum, , , True, strContent
                 strFileName = strFolder & "\" & m_CaseNo & ".補正申請書-商簡A"
            Else
                 NowPrint strReceiveNo, "90", strTmp, False, strUserNum, , , True, strContent
                 strFileName = strFolder & "\" & m_CaseNo & ".補正申請書-商簡A"
                 Call PUB_MakeDoc(strContent, strFileName)
            End If
            
            '基本資料表
            If chkAtt1(0).Value = 1 Then '若不勾選基本資料表不用產生.contact檔案
                If StartLetter2("90", "11", strReceiveNo) = False Then Exit Sub
                '統一將基本資料表要和申請書放在同一份文件
                NowPrint strReceiveNo, "90", "11", False, strUserNum, , strContent, True, strContent
                If strFileName = "" Then strFileName = strFolder & "\" & m_CaseNo & ".contact"
                strContent = Replace(strContent, vbCrLf & Chr(12), vbCrLf & "|#(分節)#|")    '換頁符號Chr(12)替換為分節符號 "|#(分節)#|"
                Call PUB_MakeDoc(strContent, strFileName, , , , , True)  '分節處理頁碼
            End If
            
         'end 2020/12/31
         Else
            bolChk = False
            If m_CP10 = "302" And m2_CP10 = "1701" Then '更正註冊證
               For i = 0 To 5
                  If Check1(i).Value = 1 Then
                     bolChk = True
                     Exit For
                  End If
               Next
            ElseIf m2_CP10 = "1001" And m3_CP10 = "301" Then '註冊變更核准
               For i = 0 To 6
                  If Check2(i).Value = 1 Then
                     bolChk = True
                     Exit For
                  End If
               Next
            ElseIf m2_CP10 = "1001" And m3_CP10 = "102" Then '延展核准
               For i = 0 To 3
                  If Check3(i).Value = 1 Then
                     bolChk = True
                     Exit For
                  End If
               Next
            ElseIf m2_CP10 = "1001" And m3_CP10 = "501" Then '移轉核准
               For i = 0 To 4
                  If Check4(i).Value = 1 Then
                     bolChk = True
                     Exit For
                  End If
               Next
            ElseIf m2_CP10 = "1001" And m3_CP10 = "502" Then '授權核准
               For i = 0 To 4
                  If Check5(i).Value = 1 Then
                     bolChk = True
                     Exit For
                  End If
               Next
            End If
            'Modify By Sindy 2012/6/1 Mark
'            If bolChk = False Then
'               MsgBox "請選擇欲補之文件 !", vbCritical
'               Exit Sub
'            End If
            
            If TxtValidate = False Then Exit Sub
            If FormSave = False Then MsgBox "存檔失敗，請洽系統管理員 !", vbCritical: Exit Sub
            
            If Text7 = "Y" Then
               bolChk = True
            Else
               bolChk = False
            End If
            
            'StartLetter "90", Text1 & Text2 & Text3 & Text4 & "&302", strTmp
            'NowPrint Text1 & Text2 & Text3 & Text4 & "&302", "90", strTmp, bolChk, strUserNum
            StartLetter "90", strReceiveNo, strTmp
            NowPrint strReceiveNo, "90", strTmp, bolChk, strUserNum, 0
         End If
         frm030206_1.Show
         '回到原畫面要清除畫面
         frm030206_1.ClearForm
      Case 1 '回前畫面
         frm030206_1.Show
      Case 2 '結束
         Unload frm030206_1
   End Select
   Unload Me
End Sub

Private Sub StartLetter(ByVal ET01 As String, ByVal ET02 As String, ByVal ET03 As String)
Dim strTxt(1 To 10) As String, strTmp As String
Dim ii As Integer, i As Integer, j As Integer
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
   
   EndLetter ET01, ET02, ET03, strUserNum
   ii = 0: j = 0
   
   If m_CP10 = "302" And m2_CP10 = "1701" Then '更正註冊證
      For i = 0 To 5
         If Check1(i).Value = 1 Then
            j = j + 1
            strTmp = CStr(j) & "." & Check1(i).Caption
            ii = ii + 1
            strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                 "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
                 "','補文件 V " & Format(j) & "','" & strTmp & "')"
         End If
      Next
   ElseIf m2_CP10 = "1001" And m3_CP10 = "301" Then '註冊變更核准
      For i = 0 To 6
         If Check2(i).Value = 1 Then
            j = j + 1
            strTmp = Check2(i).Caption
            ii = ii + 1
            strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                 "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
                 "','補文件 V " & Format(j) & "','" & strTmp & "')"
         End If
      Next
   ElseIf m2_CP10 = "1001" And m3_CP10 = "102" Then '延展核准
      For i = 0 To 3
         If Check3(i).Value = 1 Then
            j = j + 1
            strTmp = Check3(i).Caption
            ii = ii + 1
            strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                 "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
                 "','補文件 V " & Format(j) & "','" & strTmp & "')"
         End If
      Next
   ElseIf m2_CP10 = "1001" And m3_CP10 = "501" Then '移轉核准
      For i = 0 To 4
         If Check4(i).Value = 1 Then
            j = j + 1
            strTmp = Check4(i).Caption
            ii = ii + 1
            strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                 "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
                 "','補文件 V " & Format(j) & "','" & strTmp & "')"
         End If
      Next
   ElseIf m2_CP10 = "1001" And m3_CP10 = "502" Then '授權核准
      For i = 0 To 4
         If Check5(i).Value = 1 Then
            j = j + 1
            strTmp = Check5(i).Caption
            ii = ii + 1
            strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                 "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
                 "','補文件 V " & Format(j) & "','" & strTmp & "')"
         End If
      Next
   'Add By Sindy 2012/6/1 ＋更正其他狀況時
   Else
      'Add By Sindy 2014/11/5
      If strCaseType <> "" Then
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
              "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
              "','案件種類','" & ChgSQL(strCaseType) & "')"
      End If
      '2014/11/5 END
      '相關總收文號-發文日期
      If m2_CP27 <> "" Then
         m2_CP27 = Val(m2_CP27) - 19110000
         If Len(m2_CP27) = 6 Then m2_CP27 = "0" & Trim(m2_CP27)
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
             "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
             "','說明一','本案業於" & Val(Left(m2_CP27, 3)) & "年" & Mid(m2_CP27, 4, 2) & "月" & Right(m2_CP27, 2) & "日提出申請在案。')"
      Else
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
               "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
               "','說明一','本案業於　年　月　日提出申請在案。')"
      End If
   '2012/6/1 End
   End If
   
   'Add By Sindy 2016/5/31
   If tm(8) = "7" Then '7.證明標章
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                   "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
                   "','證明標章','證明標章')"
   End If
   '2016/5/31 END
   
   ii = ii + 1
   strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
       "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
       "','機關文號','" & Label12(7) & "')"
   
   If ii <> 0 Then
      If Not ClsLawExecSQL(ii, strTxt) Then
         MsgBox "儲存例外欄位失敗，請洽系統管理員 !", vbCritical
      End If
   End If
End Sub

Private Sub Combo1_Click()
   Select Case Combo1
      Case "中"
         Label12(3) = tm(5)
      Case "英"
         Label12(3) = tm(6)
      Case "日"
         Label12(3) = tm(7)
   End Select
End Sub

Private Sub Form_Activate()
Me.Text7.SetFocus
End Sub

Private Sub Form_Load()
Dim tKind As String 'Added by Lydia 2020/12/31 特殊申請書

   MoveFormToCenter Me
   intWhere = 國外_FC
   With frm030206_1
      Text1 = .Text1
      Text2 = .Text2
      Text3 = .Text3
      Text4 = .Text4
      'Added by Lydia 2020/12/31
      tKind = .Text6
      If tKind = "2" Then m_CP118 = "Y"
      'end 2020/12/31
      strReceiveNo = .Tag
   End With
   ReDim tm(TF_TM)
   ReadTradeMark
   '加出名代理人清單供勾選
   lstNameAgent.Clear
   'Modified by Lydia 2021/08/04 傳入案件性質、Form 2.0
   'PUB_SetOurAgent lstNameAgent, tm(), m_CP110
   PUB_SetOurAgent lstNameAgent, tm(), m_CP110, m_CP10, True
   'Added by Lydia 2021/08/04 如果一開始將ListBox拉到需要的大小，字型會自動放大；所以畫面預設為一列高度，Form_Load才放大到需要的大小
   lstNameAgent.Height = 930
   lstNameAgent.Width = 1500
   
   Combo1.ListIndex = 0
   Text5.Text = strSrvDate(2)
   
   'Added by Lydia 2020/12/31 電子送件
   SSTab1.Tab = 0  '預設
   If tKind = "2" Then
       m_CP118 = "Y"
       Frame1.Visible = True
       txtFee = Format(txtFee, "#####0") 'Added by Lydia 2021/09/07
       lblFee.Visible = True
       txtFee.Visible = True
   Else
       Frame1.Visible = False
       lblFee.Visible = False
       txtFee.Visible = False
       SSTab1.Width = 8235
   End If
   'end 2020/12/31
   'Memo by Amy 2023/02/08 原預勾「基本資料表」取消,增加「證冊證、核准函」並預勾-陳金蓮
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm03020604_1 = Nothing
End Sub

Private Sub ReadTradeMark()
Dim rsTemp1 As New ADODB.Recordset
'Modified by Lydia 2021/08/04
'Dim Lbl As LABEL
Dim Lbl As Object
   
   m_CP10 = "": m2_CP10 = "": m3_CP10 = ""
   m_CP43 = "": m2_CP43 = "": m2_CP27 = ""
   For Each Lbl In Label12
      Lbl = ""
   Next
   tm(1) = Text1
   tm(2) = Text2
   tm(3) = Text3
   tm(4) = Text4
   If ClsPDReadTrademarkDatabase(tm(), intWhere) Then
      Text5 = tm(11)
      Label12(1) = tm(12)
      Label12(2) = tm(15)
      Label12(3) = tm(5)
   End If
   'Modified by Lydia 2020/12/31 +cp118,cp17,FCT程序分機
   'strExc(0) = "select cpm03,staff.st02 as st1,staff1.st02 as st2,cp43,cp10,CP06,CP07,CP84,CP110 " & _
      "from caseprogress,casepropertymap,staff,staff staff1 " & _
      "where cp09='" & strReceiveNo & "' " & _
      "AND cp01=cpm01(+) and cp10=cpm02(+) and cp14=staff.st01(+) " & _
      "and cp13=staff1.st01(+) "
   strExc(0) = "select cpm03,s1.st02 as st1,s2.st02 as st2,cp43,cp10,cp06,cp07,cp84,cp110,cp118,cp17,s3.st07 " & _
                    "from caseprogress,casepropertymap,staff s1 ,staff s2,staff s3 " & _
                    "where cp09='" & strReceiveNo & "' " & _
                    "and cp01=cpm01(+) and cp10=cpm02(+) and cp14=s1.st01(+) " & _
                    "and cp13=s2.st01(+) and s2.st57=s3.st01(+) "
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   With RsTemp
   If intI = 1 Then
      m_CP110 = "" & .Fields("CP110")
      m_CP10 = "" & .Fields("CP10")
      If Not IsNull(.Fields(0)) Then
         Label12(0) = .Fields(0) '案件性質
      End If
      If Not IsNull(.Fields(1)) Then Label12(4) = .Fields(1) '承辦人
      If Not IsNull(.Fields(2)) Then Label12(5) = .Fields(2) '智權人員
      'Added by Lydia 2020/12/31
      txtFee = Format(Val("" & .Fields("CP17")), "#,##0")
      m_CP118 = "" & .Fields("CP118")
      If m_CP118 <> "" Then
         m_CP118 = "Y"
         txtFee = Val("" & .Fields("CP17"))
      End If
      m_F21st07 = "" & .Fields("st07") 'FCT程序分機
      'end 2020/12/31
      m_CP43 = "" & .Fields(3) '相關總收文號
      If Not IsNull(.Fields(3)) Then
         '取得相關總收文號資料
         strExc(0) = "SELECT * FROM CASEPROGRESS WHERE CP09='" & .Fields(3) & "'"
         intI = 1
         Set rsTemp1 = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            m2_CP10 = "" & rsTemp1.Fields("CP10")
            m2_CP43 = "" & rsTemp1.Fields("CP43")
            m2_CP27 = "" & rsTemp1.Fields("CP27") 'Add By Sindy 2012/6/1
            If Not IsNull(rsTemp1.Fields("CP05")) Then Label12(6) = TransDate(rsTemp1.Fields("CP05"), 1) '來函收文日
            If Not IsNull(rsTemp1.Fields("CP08")) Then Label12(7) = rsTemp1.Fields("CP08") '機關文號
         End If
         '再推回一層取得其相關總收文號資料
         strExc(0) = "SELECT * FROM CASEPROGRESS WHERE CP09='" & m2_CP43 & "'"
         intI = 1
         Set rsTemp1 = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            m3_CP10 = "" & rsTemp1.Fields("CP10")
         'Add By Sindy 2015/11/26 ex.FCT-38066 更正(BA4038541)
         Else
            m3_CP10 = m2_CP10
         '2015/11/26 END
         End If
      End If
      If Not IsNull(.Fields(5)) Then Label12(9) = TransDate(.Fields(5), 1) '本所期限
      If Not IsNull(.Fields(6)) Then Label12(10) = TransDate(.Fields(6), 1) '法定期限
   End If
   End With
   
   If m_CP10 = "302" And m2_CP10 = "1701" Then '更正註冊證
      Label12(0) = "更正-註冊證"
   ElseIf m2_CP10 = "1001" And m3_CP10 = "301" Then '註冊變更核准
      Label12(0) = "更正-註冊變更核准"
   ElseIf m2_CP10 = "1001" And m3_CP10 = "102" Then '延展核准
      Label12(0) = "更正-延展核准"
   ElseIf m2_CP10 = "1001" And m3_CP10 = "501" Then '移轉核准
      Label12(0) = "更正-移轉核准"
   ElseIf m2_CP10 = "1001" And m3_CP10 = "502" Then '授權核准
      Label12(0) = "更正-授權核准"
   'Add By Sindy 2012/6/1
   Else
      Label12(0) = "更正"
   '2012/6/1 End
   End If
End Sub

Private Sub Text5_Validate(Cancel As Boolean)
   Cancel = Not ChkLetterDate(Text5.Text)
   If Cancel = True Then TextInverse Text5
End Sub

Private Sub Text7_GotFocus()
  TextInverse Text7
End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 89 And KeyAscii <> 8 Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Function TxtValidate() As Boolean
   Dim Cancel As Boolean
   If lstNameAgent.Visible = True Then
      Cancel = False
      lstNameAgent_Validate Cancel
      If Cancel = True Then
         lstNameAgent.SetFocus
         Exit Function
      End If
   End If
   TxtValidate = True
End Function

Private Function FormSave() As Boolean

On Error GoTo ErrorHandler

   cnnConnection.BeginTrans
   
   If lstNameAgent.Visible = True Then
      strSql = " UPDATE CASEPROGRESS SET cp110=" & CNULL(m_CP110) & " WHERE CP09='" & strReceiveNo & "'"
      cnnConnection.Execute strSql
   End If
   'Added by Lydia 2020/12/31 預設為電子送件
   If m_CP118 = "Y" Then
        strSql = " UPDATE CASEPROGRESS SET CP118='Y' WHERE CP09='" & strReceiveNo & "' AND CP158=0 AND CP118 IS NULL"
        cnnConnection.Execute strSql
   End If
   'end 2020/12/31
   
   cnnConnection.CommitTrans
   FormSave = True
   
ErrorHandler:
   If Err.Number <> 0 Then
    cnnConnection.RollbackTrans
   End If
End Function

'檢查並設定cp110資料
Private Sub lstNameAgent_Validate(Cancel As Boolean)
   Dim ii As Integer
   Cancel = True
   m_CP110 = "": m_AgentName = ""
   For ii = 0 To lstNameAgent.ListCount - 1
      If lstNameAgent.Selected(ii) = True Then
         'modify by sonia 2016/10/7 員工編號已可非數字需做轉換
         'm_CP110 = m_CP110 & "," & lstNameAgent.ItemData(ii)
         'Modified by Lydia 2021/08/04 改模組
         'm_CP110 = m_CP110 & "," & PUB_Num2Id(lstNameAgent.ItemData(ii))
         m_CP110 = m_CP110 & "," & PUB_GetItemData(lstNameAgent.Tag, ii)
         m_AgentName = m_AgentName & "、" & lstNameAgent.List(ii)
         Cancel = False
      End If
   Next
   If Cancel = True Then
      MsgBox "出名代理人不可空白！", vbExclamation
   Else
      If Left(m_CP110, 1) = "," Then m_CP110 = Mid(m_CP110, 2)
      m_AgentName = Mid(m_AgentName, 2)
   End If
End Sub

'Added by Lydia 2020/12/31 各式申請書-電子送件申請書
Private Function StartLetter2(ByVal iET01 As String, ByVal iET03 As String, ByVal iCp09 As String) As Boolean
   Dim strTxt(1 To 30) As String, strTmp As String
   Dim StrSQLa As String
   Dim rsA As New ADODB.Recordset
   Dim ii As Integer, jj As Integer
   Dim strCP07 As String
   Dim tmpArr1 As Variant, tmpArr2 As Variant
   
   EndLetter iET01, iCp09, iET03, strUserNum
   
   ii = 1
   strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
      " VALUES ('" & iET01 & "','" & iCp09 & "','" & iET03 & "','" & strUserNum & "','本所案號','" & m_CaseNo & "')"
   
   '申請人資料
   'Modified by Lydia 2023/11/08 原本預設抓申請人基本檔之地址;現在改成預設抓案件申請人資料之地址
   'Call PUB_GetApplFCT_EData(iET01, iET03, iCp09, m_CP10, tm(), False)
   Call PUB_GetApplFCT_EData(iET01, iET03, iCp09, m_CP10, tm(), True)
   
   '出名代理人: 改成共用模組取得資料
   strExc(0) = PUB_GetAgentCP110(iCp09, m_CP110, "FCT", "4")
   If strExc(0) <> "" Then
       tmpArr1 = Split(strExc(0), "|")
       For jj = 0 To UBound(tmpArr1)
           If Trim(tmpArr1(jj)) <> "" Then
               tmpArr2 = Empty
               tmpArr2 = Split(tmpArr1(jj), ",")
                ii = ii + 1
                strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                   " VALUES ('" & iET01 & "','" & iCp09 & "','" & iET03 & "','" & strUserNum & "','代理人" & jj + 1 & "-證書字號','" & tmpArr2(0) & "')"
                ii = ii + 1
                strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                   " VALUES ('" & iET01 & "','" & iCp09 & "','" & iET03 & "','" & strUserNum & "','代理人" & jj + 1 & "-ID','" & tmpArr2(1) & "')"
                ii = ii + 1
                strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                   " VALUES ('" & iET01 & "','" & iCp09 & "','" & iET03 & "','" & strUserNum & "','代理人" & jj + 1 & "-中文姓名','" & PUB_ConvertNameFormat("" & tmpArr2(2)) & "')"
           End If
       Next jj
   End If
   
   If iET03 = "03" Then '基本資料表
        ii = ii + 1
        'FCT程序分機
        strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
              " VALUES ('" & iET01 & "','" & iCp09 & "','" & iET03 & "','" & strUserNum & "','FCT程序分機','" & m_F21st07 & "')"
   End If
   
   If iET03 = "10" Then '申請書
        ii = ii + 1
        '繳費金額
        strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
              " VALUES ('" & iET01 & "','" & iCp09 & "','" & iET03 & "','" & strUserNum & "','繳費金額','" & txtFee.Text & "')"
        
        'Add by Amy 2023/02/08 註冊證電子送件更正申請書-陳金蓮 ex:FCT-047789
        If m2_CP10 = "1701" And m_CP118 = "Y" Then
            strAppDetail = strAppDetail & "　　茲查電子證書內容有誤，謹請 鈞局更正如附件。" & vbCrLf
        ElseIf Trim(Label12(7).Caption) = "" Then  '無機關文號
            '相關總收文號-發文日期
            If m2_CP27 <> "" Then
                m2_CP27 = Val(m2_CP27) - 19110000
                If Len(m2_CP27) = 6 Then m2_CP27 = "0" & Trim(m2_CP27)
                strAppDetail = strAppDetail & "　　一、本案業於" & Val(Left(m2_CP27, 3)) & "年" & Mid(m2_CP27, 4, 2) & "月" & Right(m2_CP27, 2) & "日提出申請在案。" & vbCrLf
            Else
                strAppDetail = strAppDetail & "　　一、本案業於　年　月　日提出申請在案。" & vbCrLf
            End If
        Else
            If m2_CP10 = "1202" Then '核駁前先行通知
                strAppDetail = strAppDetail & "　　一、敬覆　鈞局" & Trim(Label12(7).Caption) & "核駁理由先行通知書。" & vbCrLf
             'Added by Lydia 2022/09/28 其對應之相關總收文號為「電話通知」時，申請書之申請內容第一點請帶：一、敬覆  鈞局XX年XX月XX日之電話通知。(日期為「電話通知」之收文日)
             ElseIf m2_CP10 = "1727" Then
                strAppDetail = strAppDetail & "　　一、敬覆　鈞局" & Val(Left(Label12(6), 3)) & "年" & Val(Mid(Label12(6), 4, 2)) & "月" & Val(Right(Label12(6), 2)) & "日之電話通知。"
             'end 2022/09/28
             'Add by Amy 2023/02/08  核淮 延展/變更 /移轉 申請書-陳金蓮 ex:FCT-030818
             ElseIf (m3_CP10 = "102" Or m3_CP10 = "301" Or m3_CP10 = "501") And m2_CP10 = "1001" Then
                strAppDetail = strAppDetail & "　　茲查核准函內容有誤，謹請 鈞局更正如附件。" & vbCrLf
             Else
                strAppDetail = strAppDetail & "　　一、敬覆　鈞局" & Trim(Label12(7).Caption) & "函。" & vbCrLf
             End If
        End If
        
        'Memo by Lydia 2020/12/31 阿蓮：畫面中”註冊證、註冊證變更...”等選項只適用在紙本，所以電子送件不列入內容
        '申請內容
        jj = 0
        strTmp = ""
        For intI = 1 To 4
            If chkAtt1(intI).Value = 1 Then
                 jj = jj + 1
                 strTmp = strTmp & vbCrLf & "　　　　" & jj & ". " & chkAtt1(intI).Caption
            End If
        Next intI
        ii = ii + 1
        strTmp = strAppDetail & IIf(strTmp <> "", "　　二、補正如下：" & strTmp & vbCrLf, "")
        strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
           " VALUES ('" & iET01 & "','" & iCp09 & "','" & iET03 & "','" & strUserNum & "','申請內容1', " & CNULL(ChgSQL(strTmp)) & ")"
        
        '附送書件
        'Modify by  Amy 2023/02/08 +註冊證、核淮函 原:For intI = 0 To 4
        For intI = 0 To 5
             If chkAtt1(intI).Value = 1 Then
                 ii = ii + 1
                 strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                    " VALUES ('" & iET01 & "','" & iCp09 & "','" & iET03 & "','" & strUserNum & "','附件-" & chkAtt1(intI).Caption & "', '" & m_CaseNo & chkAtt1(intI).Tag & "')"
             End If
        Next intI
        '若不勾選基本資料表，則附件名稱「未變更本案基本資料」並且不用產生.contact檔案
        If chkAtt1(0).Value = 0 Then
                ii = ii + 1
                strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                   " VALUES ('" & iET01 & "','" & iCp09 & "','" & iET03 & "','" & strUserNum & "','附件-" & chkAtt1(0).Caption & "', '未變更本案基本資料')"
        End If
   End If
   
   If Not ClsLawExecSQL(ii, strTxt) Then
      MsgBox "儲存例外欄位失敗，請洽系統管理員 !", vbCritical
   Else
      StartLetter2 = True
   End If
End Function


