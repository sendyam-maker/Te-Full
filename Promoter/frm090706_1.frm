VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm090706_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "繪圖人員工作進度資料查詢"
   ClientHeight    =   5730
   ClientLeft      =   -2955
   ClientTop       =   2220
   ClientWidth     =   9345
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5730
   ScaleWidth      =   9345
   Begin VB.CommandButton cmdok 
      Caption         =   "回前畫面(&U)"
      Height          =   350
      Index           =   1
      Left            =   8088
      TabIndex        =   1
      Top             =   20
      Width           =   1200
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "本月統計(&A)"
      Default         =   -1  'True
      Height          =   350
      Index           =   0
      Left            =   6864
      TabIndex        =   0
      Top             =   20
      Width           =   1200
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5328
      Left            =   0
      TabIndex        =   2
      Top             =   396
      Width           =   9300
      _ExtentX        =   16404
      _ExtentY        =   9393
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "瀏覽資料"
      TabPicture(0)   =   "frm090706_1.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label1(0)"
      Tab(0).Control(1)=   "Combo1"
      Tab(0).Control(2)=   "grd1"
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "詳細資料"
      TabPicture(1)   =   "frm090706_1.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "lbl1(29)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "lbl1(28)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "lbl1(13)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label1(1)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "lbl1(15)"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "lbl1(14)"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "lbl1(12)"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "lbl1(16)"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "lbl1(11)"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "lbl1(9)"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "lbl1(8)"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "lbl1(7)"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "lbl1(5)"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "lbl1(3)"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "lbl1(0)"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "lbl1(1)"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "lbl1(10)"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "Label1(33)"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).Control(18)=   "Label1(29)"
      Tab(1).Control(18).Enabled=   0   'False
      Tab(1).Control(19)=   "Label1(21)"
      Tab(1).Control(19).Enabled=   0   'False
      Tab(1).Control(20)=   "Label1(19)"
      Tab(1).Control(20).Enabled=   0   'False
      Tab(1).Control(21)=   "Label1(16)"
      Tab(1).Control(21).Enabled=   0   'False
      Tab(1).Control(22)=   "Label1(14)"
      Tab(1).Control(22).Enabled=   0   'False
      Tab(1).Control(23)=   "Label1(13)"
      Tab(1).Control(23).Enabled=   0   'False
      Tab(1).Control(24)=   "Label1(12)"
      Tab(1).Control(24).Enabled=   0   'False
      Tab(1).Control(25)=   "Label1(11)"
      Tab(1).Control(25).Enabled=   0   'False
      Tab(1).Control(26)=   "Label1(10)"
      Tab(1).Control(26).Enabled=   0   'False
      Tab(1).Control(27)=   "Label1(9)"
      Tab(1).Control(27).Enabled=   0   'False
      Tab(1).Control(28)=   "Label1(8)"
      Tab(1).Control(28).Enabled=   0   'False
      Tab(1).Control(29)=   "Label1(4)"
      Tab(1).Control(29).Enabled=   0   'False
      Tab(1).Control(30)=   "Label1(27)"
      Tab(1).Control(30).Enabled=   0   'False
      Tab(1).Control(31)=   "Label1(25)"
      Tab(1).Control(31).Enabled=   0   'False
      Tab(1).Control(32)=   "lbl1(26)"
      Tab(1).Control(32).Enabled=   0   'False
      Tab(1).Control(33)=   "Label1(24)"
      Tab(1).Control(33).Enabled=   0   'False
      Tab(1).Control(34)=   "Label1(5)"
      Tab(1).Control(34).Enabled=   0   'False
      Tab(1).Control(35)=   "lbl1(17)"
      Tab(1).Control(35).Enabled=   0   'False
      Tab(1).Control(36)=   "lbl1(18)"
      Tab(1).Control(36).Enabled=   0   'False
      Tab(1).Control(37)=   "lbl1(24)"
      Tab(1).Control(37).Enabled=   0   'False
      Tab(1).Control(38)=   "lbl1(22)"
      Tab(1).Control(38).Enabled=   0   'False
      Tab(1).Control(39)=   "lbl1(20)"
      Tab(1).Control(39).Enabled=   0   'False
      Tab(1).Control(40)=   "lbl1(25)"
      Tab(1).Control(40).Enabled=   0   'False
      Tab(1).Control(41)=   "lbl1(27)"
      Tab(1).Control(41).Enabled=   0   'False
      Tab(1).Control(42)=   "Label1(30)"
      Tab(1).Control(42).Enabled=   0   'False
      Tab(1).Control(43)=   "Label1(32)"
      Tab(1).Control(43).Enabled=   0   'False
      Tab(1).Control(44)=   "lbl1(19)"
      Tab(1).Control(44).Enabled=   0   'False
      Tab(1).Control(45)=   "Label1(20)"
      Tab(1).Control(45).Enabled=   0   'False
      Tab(1).Control(46)=   "lbl1(2)"
      Tab(1).Control(46).Enabled=   0   'False
      Tab(1).Control(47)=   "lbl1(21)"
      Tab(1).Control(47).Enabled=   0   'False
      Tab(1).Control(48)=   "Label1(18)"
      Tab(1).Control(48).Enabled=   0   'False
      Tab(1).Control(49)=   "lbl1(4)"
      Tab(1).Control(49).Enabled=   0   'False
      Tab(1).Control(50)=   "lbl1(23)"
      Tab(1).Control(50).Enabled=   0   'False
      Tab(1).Control(51)=   "Label1(15)"
      Tab(1).Control(51).Enabled=   0   'False
      Tab(1).Control(52)=   "lbl1(6)"
      Tab(1).Control(52).Enabled=   0   'False
      Tab(1).Control(53)=   "lbl1(30)"
      Tab(1).Control(53).Enabled=   0   'False
      Tab(1).Control(54)=   "Label1(34)"
      Tab(1).Control(54).Enabled=   0   'False
      Tab(1).Control(55)=   "Label1(2)"
      Tab(1).Control(55).Enabled=   0   'False
      Tab(1).Control(56)=   "Label1(26)"
      Tab(1).Control(56).Enabled=   0   'False
      Tab(1).Control(57)=   "Label1(31)"
      Tab(1).Control(57).Enabled=   0   'False
      Tab(1).Control(58)=   "Label1(22)"
      Tab(1).Control(58).Enabled=   0   'False
      Tab(1).Control(59)=   "Label1(6)"
      Tab(1).Control(59).Enabled=   0   'False
      Tab(1).Control(60)=   "Label1(3)"
      Tab(1).Control(60).Enabled=   0   'False
      Tab(1).Control(61)=   "Label1(23)"
      Tab(1).Control(61).Enabled=   0   'False
      Tab(1).Control(62)=   "Label1(17)"
      Tab(1).Control(62).Enabled=   0   'False
      Tab(1).Control(63)=   "Label1(7)"
      Tab(1).Control(63).Enabled=   0   'False
      Tab(1).Control(64)=   "Label1(28)"
      Tab(1).Control(64).Enabled=   0   'False
      Tab(1).Control(65)=   "lblFM(0)"
      Tab(1).Control(65).Enabled=   0   'False
      Tab(1).Control(66)=   "cmdPic"
      Tab(1).Control(66).Enabled=   0   'False
      Tab(1).ControlCount=   67
      Begin VB.CommandButton cmdPic 
         BackColor       =   &H00C0C0C0&
         Caption         =   "代表圖(&I)"
         Height          =   375
         Left            =   7530
         Style           =   1  '圖片外觀
         TabIndex        =   70
         Top             =   390
         Width           =   1665
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grd1 
         Height          =   4536
         Left            =   -74904
         TabIndex        =   3
         Top             =   684
         Width           =   9168
         _ExtentX        =   16166
         _ExtentY        =   7990
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
      Begin MSForms.ComboBox Combo1 
         Height          =   315
         Left            =   -73950
         TabIndex        =   72
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
      Begin MSForms.Label lblFM 
         Height          =   255
         Index           =   0
         Left            =   2130
         TabIndex        =   71
         Top             =   30
         Width           =   3405
         Caption         =   "lblFM"
         Size            =   "6011;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         Caption         =   "墨圖完稿日："
         Height          =   180
         Index           =   28
         Left            =   4524
         TabIndex        =   69
         Top             =   1836
         Width           =   1104
      End
      Begin VB.Label Label1 
         Caption         =   "承辦期限："
         Height          =   180
         Index           =   7
         Left            =   4524
         TabIndex        =   68
         Top             =   396
         Width           =   924
      End
      Begin VB.Label Label1 
         Caption         =   "是否算案件數："
         Height          =   180
         Index           =   17
         Left            =   4524
         TabIndex        =   67
         Top             =   2412
         Width           =   1272
      End
      Begin VB.Label Label1 
         Caption         =   "草圖齊備日："
         Height          =   180
         Index           =   23
         Left            =   4524
         TabIndex        =   66
         Top             =   684
         Width           =   1104
      End
      Begin VB.Label Label1 
         Caption         =   "草圖張數："
         Height          =   180
         Index           =   3
         Left            =   4524
         TabIndex        =   65
         Top             =   1260
         Width           =   924
      End
      Begin VB.Label Label1 
         Caption         =   "修改時數："
         Height          =   180
         Index           =   6
         Left            =   4524
         TabIndex        =   64
         Top             =   3276
         Width           =   924
      End
      Begin VB.Label Label1 
         Caption         =   "承辦時數："
         Height          =   180
         Index           =   22
         Left            =   4524
         TabIndex        =   63
         Top             =   2700
         Width           =   924
      End
      Begin VB.Label Label1 
         Caption         =   "備註："
         Height          =   180
         Index           =   31
         Left            =   4530
         TabIndex        =   62
         Top             =   4140
         Width           =   570
      End
      Begin VB.Label Label1 
         Caption         =   "草圖完稿日："
         Height          =   180
         Index           =   26
         Left            =   4524
         TabIndex        =   61
         Top             =   972
         Width           =   1104
      End
      Begin VB.Label Label1 
         Caption         =   "墨圖齊備日："
         Height          =   180
         Index           =   2
         Left            =   4524
         TabIndex        =   60
         Top             =   1548
         Width           =   1104
      End
      Begin VB.Label Label1 
         Caption         =   "墨圖張數："
         Height          =   180
         Index           =   34
         Left            =   4524
         TabIndex        =   59
         Top             =   2124
         Width           =   924
      End
      Begin MSForms.Label lbl1 
         Height          =   990
         Index           =   30
         Left            =   5115
         TabIndex        =   58
         Top             =   4140
         Width           =   4065
         BackColor       =   16777215
         Caption         =   "lblFM2"
         Size            =   "7170;1746"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   6
         Left            =   1005
         TabIndex        =   57
         Top             =   2130
         Width           =   1590
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "2805;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         Caption         =   "案件性質："
         Height          =   180
         Index           =   15
         Left            =   120
         TabIndex        =   56
         Top             =   2124
         Width           =   1176
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   23
         Left            =   5475
         TabIndex        =   55
         Top             =   2130
         Width           =   1590
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "2805;450"
         FontName        =   "新細明體"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   4
         Left            =   1050
         TabIndex        =   54
         Top             =   1545
         Width           =   3405
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "6006;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         Caption         =   "案件名稱："
         Height          =   180
         Index           =   18
         Left            =   120
         TabIndex        =   53
         Top             =   1548
         Width           =   924
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   21
         Left            =   5670
         TabIndex        =   52
         Top             =   1545
         Width           =   1590
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "2805;450"
         FontName        =   "新細明體"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   2
         Left            =   900
         TabIndex        =   51
         Top             =   975
         Width           =   1590
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "2805;450"
         FontName        =   "新細明體"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         Caption         =   "收文日："
         Height          =   180
         Index           =   20
         Left            =   120
         TabIndex        =   50
         Top             =   972
         Width           =   744
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   19
         Left            =   5670
         TabIndex        =   49
         Top             =   975
         Width           =   1590
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "2805;450"
         FontName        =   "新細明體"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         Caption         =   "(N：不算)"
         Height          =   180
         Index           =   32
         Left            =   6588
         TabIndex        =   48
         Top             =   2412
         Width           =   1068
      End
      Begin VB.Label Label1 
         Caption         =   "1."
         Height          =   180
         Index           =   30
         Left            =   5784
         TabIndex        =   47
         Top             =   3276
         Width           =   240
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   27
         Left            =   6060
         TabIndex        =   46
         Top             =   3270
         Width           =   1590
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "2805;450"
         FontName        =   "新細明體"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   25
         Left            =   6375
         TabIndex        =   45
         Top             =   2700
         Width           =   1590
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "2805;450"
         FontName        =   "新細明體"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   20
         Left            =   5550
         TabIndex        =   44
         Top             =   1260
         Width           =   1590
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "2805;450"
         FontName        =   "新細明體"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   22
         Left            =   5670
         TabIndex        =   43
         Top             =   1830
         Width           =   1590
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "2805;450"
         FontName        =   "新細明體"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   24
         Left            =   5940
         TabIndex        =   42
         Top             =   2415
         Width           =   420
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "741;450"
         FontName        =   "新細明體"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   18
         Left            =   5670
         TabIndex        =   41
         Top             =   690
         Width           =   1590
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "2805;450"
         FontName        =   "新細明體"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   17
         Left            =   5475
         TabIndex        =   40
         Top             =   390
         Width           =   1590
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "2805;450"
         FontName        =   "新細明體"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         Caption         =   "草圖："
         Height          =   180
         Index           =   5
         Left            =   5784
         TabIndex        =   39
         Top             =   2700
         Width           =   552
      End
      Begin VB.Label Label1 
         Caption         =   "墨圖："
         Height          =   180
         Index           =   24
         Left            =   5784
         TabIndex        =   38
         Top             =   2988
         Width           =   552
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   26
         Left            =   6375
         TabIndex        =   37
         Top             =   2985
         Width           =   1590
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "2805;450"
         FontName        =   "新細明體"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         Caption         =   "2."
         Height          =   180
         Index           =   25
         Left            =   5784
         TabIndex        =   36
         Top             =   3564
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "3."
         Height          =   180
         Index           =   27
         Left            =   5784
         TabIndex        =   35
         Top             =   3852
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "繪圖人員："
         Height          =   180
         Index           =   4
         Left            =   120
         TabIndex        =   34
         Top             =   396
         Width           =   924
      End
      Begin VB.Label Label1 
         Caption         =   "承辦人："
         Height          =   180
         Index           =   8
         Left            =   120
         TabIndex        =   33
         Top             =   3276
         Width           =   744
      End
      Begin VB.Label Label1 
         Caption         =   "國外案承辦人："
         Height          =   180
         Index           =   9
         Left            =   120
         TabIndex        =   32
         Top             =   4716
         Width           =   1284
      End
      Begin VB.Label Label1 
         Caption         =   "國外案本所案號："
         Height          =   180
         Index           =   10
         Left            =   120
         TabIndex        =   31
         Top             =   5004
         Width           =   1476
      End
      Begin VB.Label Label1 
         Caption         =   "點數："
         Height          =   180
         Index           =   11
         Left            =   120
         TabIndex        =   30
         Top             =   2412
         Width           =   672
      End
      Begin VB.Label Label1 
         Caption         =   "智權人員："
         Height          =   180
         Index           =   12
         Left            =   120
         TabIndex        =   29
         Top             =   3570
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "法定期限："
         Height          =   180
         Index           =   13
         Left            =   120
         TabIndex        =   28
         Top             =   2988
         Width           =   924
      End
      Begin VB.Label Label1 
         Caption         =   "本所期限："
         Height          =   180
         Index           =   14
         Left            =   120
         TabIndex        =   27
         Top             =   2700
         Width           =   924
      End
      Begin VB.Label Label1 
         Caption         =   "專利/商標種類："
         Height          =   180
         Index           =   16
         Left            =   120
         TabIndex        =   26
         Top             =   1836
         Width           =   1332
      End
      Begin VB.Label Label1 
         Caption         =   "本所案號："
         Height          =   180
         Index           =   19
         Left            =   120
         TabIndex        =   25
         Top             =   1260
         Width           =   924
      End
      Begin VB.Label Label1 
         Caption         =   "總收文號："
         Height          =   180
         Index           =   21
         Left            =   120
         TabIndex        =   24
         Top             =   684
         Width           =   924
      End
      Begin VB.Label Label1 
         Caption         =   "取消收文日："
         Height          =   180
         Index           =   29
         Left            =   120
         TabIndex        =   23
         Top             =   4428
         Width           =   1104
      End
      Begin VB.Label Label1 
         Caption         =   "草圖作業天數："
         Height          =   180
         Index           =   33
         Left            =   120
         TabIndex        =   22
         Top             =   3852
         Width           =   1284
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   10
         Left            =   900
         TabIndex        =   21
         Top             =   3270
         Width           =   1590
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "2805;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   1
         Left            =   1050
         TabIndex        =   20
         Top             =   690
         Width           =   1590
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "2805;450"
         FontName        =   "新細明體"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   0
         Left            =   1050
         TabIndex        =   19
         Top             =   390
         Width           =   1590
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "2805;450"
         FontName        =   "新細明體"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   3
         Left            =   1050
         TabIndex        =   18
         Top             =   1260
         Width           =   1665
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "2937;450"
         FontName        =   "新細明體"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   5
         Left            =   1470
         TabIndex        =   17
         Top             =   1830
         Width           =   1590
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "2805;450"
         FontName        =   "新細明體"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   7
         Left            =   795
         TabIndex        =   16
         Top             =   2415
         Width           =   1590
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "2805;450"
         FontName        =   "新細明體"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   8
         Left            =   1050
         TabIndex        =   15
         Top             =   2700
         Width           =   1590
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "2805;450"
         FontName        =   "新細明體"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   9
         Left            =   1050
         TabIndex        =   14
         Top             =   2985
         Width           =   1590
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "2805;450"
         FontName        =   "新細明體"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   11
         Left            =   1095
         TabIndex        =   13
         Top             =   3570
         Width           =   1590
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "2805;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   16
         Left            =   1590
         TabIndex        =   12
         Top             =   5010
         Width           =   1560
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "2752;450"
         FontName        =   "新細明體"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   12
         Left            =   1470
         TabIndex        =   11
         Top             =   3852
         Width           =   1590
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "2805;450"
         FontName        =   "新細明體"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   14
         Left            =   1260
         TabIndex        =   10
         Top             =   4428
         Width           =   1590
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "2805;450"
         FontName        =   "新細明體"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   15
         Left            =   1470
         TabIndex        =   9
         Top             =   4716
         Width           =   1590
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "2805;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         Caption         =   "墨圖作業天數："
         Height          =   180
         Index           =   1
         Left            =   120
         TabIndex        =   8
         Top             =   4140
         Width           =   1284
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   13
         Left            =   1470
         TabIndex        =   7
         Top             =   4140
         Width           =   1590
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "2805;450"
         FontName        =   "新細明體"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   28
         Left            =   6060
         TabIndex        =   6
         Top             =   3570
         Width           =   1590
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "2805;450"
         FontName        =   "新細明體"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   29
         Left            =   6060
         TabIndex        =   5
         Top             =   3855
         Width           =   1590
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "2805;450"
         FontName        =   "新細明體"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         Caption         =   "繪圖人員： "
         Height          =   180
         Index           =   0
         Left            =   -74844
         TabIndex        =   4
         Top             =   396
         Width           =   912
      End
   End
End
Attribute VB_Name = "frm090706_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2022/01/28 改成Form2.0 ; grd1改字型=新細明體-ExtB、Combo1、lb1(index)
'Memo By Morgan 2012/12/10 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/17 日期欄已修改
Option Explicit
Dim TextOk As Boolean, k As Integer
Dim strSql As String, strSQL1 As String, i As Integer, j As Integer, s As Integer, SavDay1 As String, SavDay2 As String, SavDay3 As String, StrSQL7 As String, StrSQL4 As String, strSQL5 As String
Dim strSQL2 As String, iPrint As Integer, Page As Integer, strTemp(0 To 21) As String, strTemp3 As String, TestOk As Boolean, StrTemp99(0 To 21) As String
Dim PLeft(0 To 21) As Integer, strTemp1 As Variant, strTemp2 As Variant, StrSQL6 As String, Str020401SysKind As String, Seekok As Integer, SeekTemp As Integer

Sub Process(strText As String)

'Added by Lydia 2022/01/28
Dim oLbl As Control
For Each oLbl In lbl1
   oLbl.Caption = ""
   oLbl.BackColor = &H8000000F
Next
'end 2022/01/28

'92.04.03 nick add left join
'strSQL = "SELECT S1.ST02,cp09," & SQLDate("CP05") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04,NVL(PA05,NVL(PA06,PA07)),DECODE(PA09,'000',PTM03,PTM04),decode(pa09,'000',cpm03,cpm04),cp18," & SQLDate("cp06") & "," & SQLDate("cp07") & ",S2.ST02,s3.st02,0,0," & SQLDate("CP57") & ",'',''," & SQLDate("cP48") & "," & SQLDate("eP14") & "," & SQLDate("eP15") & ",ep16," & SQLDate("ep17") & "," & SQLDate("ep18") & ",ep19,ep20,ep21,ep22,ep23,ep24,ep25,ep26 FROM ENGINEERPROGRESS,CASEPROGRESS,STAFF S1,STAFF S2,STAFF S3,CASEPROPERTYMAP,PATENTTRADEMARKMAP,PATENT WHERE EP02=CP09(+) AND  PA01=CP01 AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) and cp01=cpm01(+) and cp10=cpm02(+) AND '1'=PTM01(+)  AND PA08=PTM02(+) AND ep13=S1.ST01(+) AND eP05=S2.ST01(+) AND cP13=S3.ST01(+) AND EP02='" & StrText & "' "
strSql = "SELECT S1.ST02,cp09," & SQLDate("CP05") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04,NVL(PA05,NVL(PA06,PA07)),DECODE(PA09,'000',PTM03,PTM04),decode(pa09,'000',cpm03,cpm04),cp18," & SQLDate("cp06") & "," & SQLDate("cp07") & ",S2.ST02,s3.st02,0,0," & SQLDate("CP57") & ",'',''," & SQLDate("cP48") & "," & SQLDate("eP14") & "," & SQLDate("eP15") & ",ep16," & SQLDate("ep17") & "," & SQLDate("ep18") & ",ep19,ep20,ep21,ep22,ep23,ep24,ep25,ep26 FROM ENGINEERPROGRESS,CASEPROGRESS,STAFF S1,STAFF S2,STAFF S3,CASEPROPERTYMAP,PATENTTRADEMARKMAP,PATENT WHERE EP02=CP09(+) AND  cp01=pa01(+) and CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) and cp01=cpm01(+) and cp10=cpm02(+) AND '1'=PTM01(+)  AND PA08=PTM02(+) AND ep13=S1.ST01(+) AND eP05=S2.ST01(+) AND cP13=S3.ST01(+) AND EP02='" & strText & "' "
CheckOC
With adoRecordset
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        .MoveFirst
        For i = 0 To 30
            lbl1(i) = CheckStr(.Fields(i))
        Next i
        '92.04.03 nick add left join
        'strSQL = "SELECT CM01||'-'||CM02||'-'||CM03||'-'||CM04,ST02 FROM CASEMAP,CASEPROGRESS,STAFF WHERE CP01=CM01 AND CP02=CM02 AND CP03=CM03 AND CP04=CM04 AND CP14=ST01(+) AND CP31='Y' and cp09='" & StrText & "' order by CM01||'-'||CM02||'-'||CM03||'-'||CM04 "
        strSql = "SELECT CM01||'-'||CM02||'-'||CM03||'-'||CM04,ST02 FROM CASEMAP,CASEPROGRESS,STAFF WHERE CP01=CM01(+) AND CP02=CM02(+) AND CP03=CM03(+) AND CP04=CM04(+) AND CP14=ST01(+) AND CP31='Y' and cp09='" & strText & "' order by CM01||'-'||CM02||'-'||CM03||'-'||CM04 "
        CheckOC2
        adoRecordset1.CursorLocation = adUseClient
        adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
        If adoRecordset1.RecordCount <> 0 And adoRecordset1.RecordCount > 0 Then
            lbl1(16) = CheckStr(adoRecordset1.Fields(0))
            lbl1(15) = CheckStr(adoRecordset1.Fields(1))
        Else
            lbl1(16) = ""
            lbl1(15) = ""
        End If
        CheckOC2
        '計算草圖作業天數
        If Len(lbl1(18)) <> 0 And Len(lbl1(19)) <> 0 And Val(lbl1(18)) <> 0 And Val(lbl1(19)) <> 0 Then
            lbl1(12) = GetWorkDay(ChangeTStringToWString(ChangeTDateStringToTString(lbl1(19))), ChangeTStringToWString(ChangeTDateStringToTString(lbl1(18))))
        End If
        '計算墨圖作業天數
        If Len(lbl1(21)) <> 0 And Len(lbl1(22)) <> 0 And Val(lbl1(21)) <> 0 And Val(lbl1(22)) <> 0 Then
            lbl1(13) = GetWorkDay(ChangeTStringToWString(ChangeTDateStringToTString(lbl1(22))), ChangeTStringToWString(ChangeTDateStringToTString(lbl1(21))))
        End If
        'add by nickc 2007/08/03 檢查有無代表圖
        strSql = "SELECT ibf01,ibf02 FROM imgbytefile WHERE ibf01='" & SystemNumber(lbl1(3), 1) & "' and ibf02='" & SystemNumber(lbl1(3), 2) & "' and ibf03='" & SystemNumber(lbl1(3), 3) & "' and ibf04='" & SystemNumber(lbl1(3), 4) & "' and ibf05='1' "
        CheckOC2
        adoRecordset1.CursorLocation = adUseClient
        adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
        If adoRecordset1.RecordCount <> 0 And adoRecordset1.RecordCount > 0 Then
            cmdPic.Caption = "已設定代表圖(&I)"
            cmdPic.BackColor = &HC0FFC0
        Else
            cmdPic.Caption = "未設定代表圖(&I)"
            cmdPic.BackColor = &HC0C0FF
        End If
        CheckOC2
    End If
End With
CheckOC
End Sub

Private Sub cmdOK_Click(Index As Integer)
Select Case Index
Case 0
     Me.Hide
     frm090706_2.Show
Case 1
     Me.Hide
     frm090706.Show
     Unload Me
Case Else
End Select
End Sub

'Modified by Lydia 2022/01/28 Form2.0點選同一人不會觸發Click事件，改用DropButtonClick事件但要控制第2次才執行
'Private Sub Combo1_Click()
Private Sub Combo1_DropButtonClick()
   Static bClick As Boolean
   If bClick = False Then
      bClick = True
      Exit Sub
   End If
   bClick = False
'end 2022/01/28

StrMenu
SetGrd1
TextOk = True
Grd1_Click
TextOk = False
End Sub

Private Sub Form_Load()
MoveFormToCenter Me
TextOk = True
StrMenu1
StrMenu
SetGrd1
Grd1_Click
TextOk = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frm090706_1 = Nothing
End Sub

Private Sub Grd1_Click()
With grd1
    .Visible = False
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
    .col = 0
    If TextOk = True Then
        .row = 0
    Else
        .row = .MouseRow
    End If
    If .row = 0 Then
        .row = 1
    End If
    .col = 19
    Process (.Text)
    For i = 0 To .Cols - 1
        .col = i
        .CellBackColor = &HFFC0C0
    Next i
    .Visible = True
End With
End Sub

Sub StrMenu1()
strSql = "SELECT DISTINCT R110001 FROM R090706 WHERE ID='" & strUserNum & "' ORDER BY R110001 "
CheckOC
j = 0
With adoRecordset
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        .MoveFirst
        Do While .EOF = False
            Combo1.AddItem CheckStr(.Fields(0)), j
            j = j + 1
            .MoveNext
        Loop
    End If
End With
CheckOC
Combo1.Text = Combo1.List(0)
End Sub

Sub StrMenu()
strSql = "SELECT R110002,R110003,R110004,R110005,R110006,R110007,R110008,R110009,R110010,R110011,R110012,R110013,R110014,R110015,R110016,R110017,R110018,R110019,R110020,R110021 FROM R090706 WHERE ID='" & strUserNum & "' AND R110001='" & Combo1.Text & "' "
CheckOC
With adoRecordset
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        Set grd1.Recordset = adoRecordset
    End If
End With
CheckOC
End Sub

Private Sub SetGrd1()
With grd1
    .Visible = False
    .Cols = 20
    .row = 0
    .col = 0:   .Text = "收文類別"
    .ColWidth(0) = 800
    .CellAlignment = flexAlignCenterCenter
    .col = 1:   .Text = "收文日"
    .ColWidth(1) = 800
    .CellAlignment = flexAlignCenterCenter
    .col = 2:   .Text = "本所案號"
    .ColWidth(2) = 1500
    .CellAlignment = flexAlignCenterCenter
    .col = 3:   .Text = "案件名稱"
    .ColWidth(3) = 1500
    .CellAlignment = flexAlignCenterCenter
    .col = 4:   .Text = "是否算案件數"
    .ColWidth(4) = 1200
    .CellAlignment = flexAlignCenterCenter
    .col = 5:   .Text = "案件性質"
    .ColWidth(5) = 800
    .CellAlignment = flexAlignCenterCenter
    .col = 6:   .Text = "承辦人"
    .ColWidth(6) = 800
    .CellAlignment = flexAlignCenterCenter
    .col = 7:   .Text = "承辦期限"
    .ColWidth(7) = 800
    .CellAlignment = flexAlignCenterCenter
    .col = 8:   .Text = "點數"
    .ColWidth(8) = 800
    .CellAlignment = flexAlignCenterCenter
    .col = 9:   .Text = "草圖齊備日"
    .ColWidth(9) = 1000
    .CellAlignment = flexAlignCenterCenter
    .col = 10:  .Text = "草圖完稿日"
    .ColWidth(10) = 1000
    .CellAlignment = flexAlignCenterCenter
    .col = 11:  .Text = "草圖作業天數"
    .ColWidth(11) = 1200
    .CellAlignment = flexAlignCenterCenter
    .col = 12:  .Text = "墨圖齊備日"
    .ColWidth(12) = 1000
    .CellAlignment = flexAlignCenterCenter
    .col = 13:  .Text = "墨圖完稿日"
    .ColWidth(13) = 1000
    .CellAlignment = flexAlignCenterCenter
    .col = 14:  .Text = "墨圖作業天數"
    .ColWidth(14) = 1200
    .CellAlignment = flexAlignCenterCenter
    .col = 15:  .Text = "本所期限"
    .ColWidth(15) = 800
    .CellAlignment = flexAlignCenterCenter
    .col = 16:  .Text = "發文日"
    .ColWidth(16) = 800
    .CellAlignment = flexAlignCenterCenter
    .col = 17:  .Text = "備註"
    .ColWidth(17) = 800
    .CellAlignment = flexAlignCenterCenter
    .col = 18:  .Text = "智權人員"
    .ColWidth(18) = 800
    .CellAlignment = flexAlignCenterCenter
    .col = 19:  .Text = ""
    .ColWidth(19) = 0
    .CellAlignment = flexAlignCenterCenter
    .Visible = True
End With
End Sub

'add by nickc 2007/08/03
Private Sub CmdPic_Click()
frmPic001.oCP01 = SystemNumber(lbl1(3), 1)
frmPic001.oCP02 = SystemNumber(lbl1(3), 2)
frmPic001.oCP03 = SystemNumber(lbl1(3), 3)
frmPic001.oCP04 = SystemNumber(lbl1(3), 4)
frmPic001.StrMenu
frmPic001.Show vbModal

strSql = "SELECT ibf01,ibf02 FROM imgbytefile WHERE ibf01='" & SystemNumber(lbl1(3), 1) & "' and ibf02='" & SystemNumber(lbl1(3), 2) & "' and ibf03='" & SystemNumber(lbl1(3), 3) & "' and ibf04='" & SystemNumber(lbl1(3), 4) & "' and ibf05='1' "
CheckOC2
adoRecordset1.CursorLocation = adUseClient
adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If adoRecordset1.RecordCount <> 0 And adoRecordset1.RecordCount > 0 Then
    cmdPic.Caption = "已設定代表圖(&I)"
    cmdPic.BackColor = &HC0FFC0
Else
    cmdPic.Caption = "未設定代表圖(&I)"
    cmdPic.BackColor = &HC0C0FF
End If
CheckOC2
End Sub

