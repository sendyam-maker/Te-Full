VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm02050205 
   BorderStyle     =   1  '單線固定
   Caption         =   "服務業務基本資料維護 (其它業務)"
   ClientHeight    =   6396
   ClientLeft      =   96
   ClientTop       =   972
   ClientWidth     =   9132
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6396
   ScaleWidth      =   9132
   Begin VB.CommandButton Command3 
      Caption         =   "已設定代表圖"
      Enabled         =   0   'False
      Height          =   285
      Left            =   6300
      Style           =   1  '圖片外觀
      TabIndex        =   41
      Top             =   960
      Width           =   1395
   End
   Begin VB.CommandButton ButtonRelation 
      Caption         =   "相關卷號"
      Height          =   285
      Left            =   7680
      TabIndex        =   43
      Top             =   960
      Width           =   1395
   End
   Begin TabDlg.SSTab tabCtrl 
      Height          =   5385
      Left            =   90
      TabIndex        =   37
      Top             =   930
      Width           =   8955
      _ExtentX        =   15790
      _ExtentY        =   9504
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   5
      TabHeight       =   529
      TabMaxWidth     =   2646
      TabCaption(0)   =   "基本資料"
      TabPicture(0)   =   "frm02050205.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label11"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label36"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label19"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label18"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label16"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label14"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label13"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label10"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label9"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label8"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label6"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label5"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label4"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label3"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Label2"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Label1(0)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Label12"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Label17"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Label33"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Label1(160)"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Label1(172)"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "Label1(117)"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "Label1(112)"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "Label1(76)"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "textSP09_2"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "textSP08_2"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "cboContact"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "textSP17_2"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "textSP09"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "textCU79"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "textSP29"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "textSP17"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "textSP15"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "textSP28"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "textSP34"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "textSP16"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "textSP10"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "textSP08"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "textSP07"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).Control(39)=   "textSP06"
      Tab(0).Control(39).Enabled=   0   'False
      Tab(0).Control(40)=   "textSP05"
      Tab(0).Control(40).Enabled=   0   'False
      Tab(0).Control(41)=   "textSP05_1"
      Tab(0).Control(41).Enabled=   0   'False
      Tab(0).Control(42)=   "textSP73"
      Tab(0).Control(42).Enabled=   0   'False
      Tab(0).Control(43)=   "textSP74"
      Tab(0).Control(43).Enabled=   0   'False
      Tab(0).Control(44)=   "textSP85"
      Tab(0).Control(44).Enabled=   0   'False
      Tab(0).Control(45)=   "textSP83"
      Tab(0).Control(45).Enabled=   0   'False
      Tab(0).Control(46)=   "textSP80"
      Tab(0).Control(46).Enabled=   0   'False
      Tab(0).Control(47)=   "textSP04"
      Tab(0).Control(47).Enabled=   0   'False
      Tab(0).Control(48)=   "textSP03"
      Tab(0).Control(48).Enabled=   0   'False
      Tab(0).Control(49)=   "textSP02"
      Tab(0).Control(49).Enabled=   0   'False
      Tab(0).Control(50)=   "textSP01"
      Tab(0).Control(50).Enabled=   0   'False
      Tab(0).ControlCount=   51
      TabCaption(1)   =   "代理人相關資料"
      TabPicture(1)   =   "frm02050205.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "textSP75"
      Tab(1).Control(1)=   "textSP84"
      Tab(1).Control(2)=   "textSP71"
      Tab(1).Control(3)=   "textSP67"
      Tab(1).Control(4)=   "textSP36"
      Tab(1).Control(5)=   "textSP35"
      Tab(1).Control(6)=   "textSP26"
      Tab(1).Control(7)=   "textSP30"
      Tab(1).Control(8)=   "textSP37"
      Tab(1).Control(9)=   "textFA29"
      Tab(1).Control(10)=   "textSP27"
      Tab(1).Control(11)=   "textSP31"
      Tab(1).Control(12)=   "textSP33"
      Tab(1).Control(13)=   "textSP35_2"
      Tab(1).Control(14)=   "textSP37_2"
      Tab(1).Control(15)=   "textSP26_2"
      Tab(1).Control(16)=   "textSP67_2"
      Tab(1).Control(17)=   "Label32"
      Tab(1).Control(18)=   "Label31"
      Tab(1).Control(19)=   "Label15"
      Tab(1).Control(20)=   "Label7"
      Tab(1).Control(21)=   "Label20"
      Tab(1).Control(22)=   "Label21"
      Tab(1).Control(23)=   "Label22"
      Tab(1).Control(24)=   "Label23"
      Tab(1).Control(25)=   "Label24"
      Tab(1).Control(26)=   "Label25"
      Tab(1).Control(27)=   "Label26"
      Tab(1).Control(28)=   "Label27"
      Tab(1).Control(29)=   "Label28"
      Tab(1).Control(30)=   "Label29"
      Tab(1).Control(31)=   "Label30"
      Tab(1).ControlCount=   32
      TabCaption(2)   =   "銷卷資料"
      TabPicture(2)   =   "frm02050205.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lblSP61"
      Tab(2).Control(1)=   "lblSP68"
      Tab(2).Control(2)=   "lblSP69"
      Tab(2).Control(3)=   "lblSP70"
      Tab(2).Control(4)=   "Label78"
      Tab(2).Control(5)=   "Label79"
      Tab(2).Control(6)=   "Label80"
      Tab(2).Control(7)=   "Label81"
      Tab(2).ControlCount=   8
      TabCaption(3)   =   "參考備註"
      TabPicture(3)   =   "frm02050205.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "textSP18"
      Tab(3).Control(1)=   "cmdIns"
      Tab(3).ControlCount=   2
      Begin VB.CommandButton cmdIns 
         Caption         =   "各項指示"
         Height          =   285
         Left            =   -74880
         TabIndex        =   53
         Top             =   360
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox textSP01 
         Height          =   300
         Left            =   1320
         MaxLength       =   3
         TabIndex        =   0
         Top             =   300
         Width           =   732
      End
      Begin VB.TextBox textSP02 
         Height          =   300
         Left            =   2040
         MaxLength       =   6
         TabIndex        =   1
         Top             =   300
         Width           =   1092
      End
      Begin VB.TextBox textSP03 
         Height          =   300
         Left            =   3120
         MaxLength       =   1
         TabIndex        =   2
         Top             =   300
         Width           =   372
      End
      Begin VB.TextBox textSP04 
         Height          =   300
         Left            =   3480
         MaxLength       =   2
         TabIndex        =   3
         Top             =   300
         Width           =   732
      End
      Begin MSForms.TextBox textSP80 
         Height          =   300
         Left            =   4710
         TabIndex        =   13
         Top             =   2956
         Width           =   300
         VariousPropertyBits=   671105051
         BackColor       =   16777215
         MaxLength       =   1
         Size            =   "529;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textSP83 
         Height          =   300
         Left            =   7950
         TabIndex        =   14
         Top             =   2956
         Width           =   300
         VariousPropertyBits=   671105051
         BackColor       =   16777215
         MaxLength       =   1
         Size            =   "529;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textSP18 
         Height          =   4020
         Left            =   -74880
         TabIndex        =   55
         Top             =   720
         Width           =   8715
         VariousPropertyBits=   -1466941413
         BackColor       =   16777215
         MaxLength       =   2000
         ScrollBars      =   2
         Size            =   "15372;7091"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textSP75 
         Height          =   300
         Left            =   -73560
         TabIndex        =   31
         Top             =   2020
         Width           =   7275
         VariousPropertyBits=   671105051
         BackColor       =   16777215
         MaxLength       =   60
         Size            =   "12832;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textSP85 
         Height          =   300
         Left            =   1320
         TabIndex        =   23
         Top             =   5040
         Width           =   315
         VariousPropertyBits=   671105051
         BackColor       =   16777215
         MaxLength       =   1
         Size            =   "556;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textSP84 
         Height          =   300
         Left            =   -73020
         TabIndex        =   28
         Top             =   1060
         Width           =   2595
         VariousPropertyBits=   671105051
         BackColor       =   16777215
         MaxLength       =   20
         Size            =   "4577;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textSP74 
         Height          =   360
         Left            =   1320
         TabIndex        =   9
         Top             =   1945
         Width           =   7452
         VariousPropertyBits=   -1467989989
         BackColor       =   16777215
         MaxLength       =   300
         Size            =   "13144;635"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textSP73 
         Height          =   360
         Left            =   1320
         TabIndex        =   8
         Top             =   1568
         Width           =   7455
         VariousPropertyBits=   -1467989989
         BackColor       =   16777215
         MaxLength       =   395
         Size            =   "13150;635"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textSP71 
         Height          =   300
         Left            =   -73560
         TabIndex        =   30
         Top             =   1700
         Width           =   7275
         VariousPropertyBits=   671105051
         BackColor       =   16777215
         MaxLength       =   60
         Size            =   "12832;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textSP05_1 
         Height          =   960
         Left            =   1320
         TabIndex        =   4
         Top             =   600
         Width           =   7335
         VariousPropertyBits=   -1466941413
         BackColor       =   16777215
         MaxLength       =   140
         ScrollBars      =   2
         Size            =   "12933;1693"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textSP67 
         Height          =   300
         Left            =   -73350
         TabIndex        =   35
         Top             =   3300
         Width           =   1215
         VariousPropertyBits=   671105051
         BackColor       =   16777215
         MaxLength       =   8
         Size            =   "2143;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textSP36 
         Height          =   300
         Left            =   -73560
         TabIndex        =   34
         Top             =   2980
         Width           =   2295
         VariousPropertyBits=   671105051
         BackColor       =   16777215
         MaxLength       =   35
         Size            =   "4048;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textSP35 
         Height          =   300
         Left            =   -73560
         TabIndex        =   33
         Top             =   2660
         Width           =   1215
         VariousPropertyBits=   671105051
         BackColor       =   16777215
         MaxLength       =   8
         Size            =   "2143;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textSP05 
         Height          =   300
         Left            =   1320
         TabIndex        =   5
         Top             =   617
         Width           =   7332
         VariousPropertyBits=   -1467989985
         BackColor       =   16777215
         MaxLength       =   160
         ScrollBars      =   2
         Size            =   "12933;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textSP06 
         Height          =   300
         Left            =   1320
         TabIndex        =   6
         Top             =   934
         Width           =   7332
         VariousPropertyBits=   -1467989985
         BackColor       =   16777215
         MaxLength       =   180
         ScrollBars      =   2
         Size            =   "12933;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textSP07 
         Height          =   300
         Left            =   1320
         TabIndex        =   7
         Top             =   1251
         Width           =   7332
         VariousPropertyBits=   -1467989985
         BackColor       =   16777215
         MaxLength       =   160
         ScrollBars      =   2
         Size            =   "12933;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textSP08 
         Height          =   300
         Left            =   1320
         TabIndex        =   10
         Top             =   2322
         Width           =   1095
         VariousPropertyBits=   671105051
         BackColor       =   16777215
         MaxLength       =   9
         Size            =   "1931;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textSP10 
         Height          =   300
         Left            =   5640
         TabIndex        =   12
         Top             =   2639
         Width           =   1095
         VariousPropertyBits=   671105051
         BackColor       =   16777215
         MaxLength       =   7
         Size            =   "1931;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textSP16 
         Height          =   300
         Left            =   1320
         TabIndex        =   16
         Top             =   3273
         Width           =   1095
         VariousPropertyBits=   671105051
         BackColor       =   16777215
         MaxLength       =   7
         Size            =   "1931;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textSP34 
         Height          =   300
         Left            =   1320
         TabIndex        =   18
         Top             =   3590
         Width           =   375
         VariousPropertyBits=   671105051
         BackColor       =   16777215
         MaxLength       =   1
         Size            =   "661;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textSP28 
         Height          =   300
         Left            =   1320
         TabIndex        =   19
         Top             =   3907
         Width           =   2745
         VariousPropertyBits=   671105051
         BackColor       =   16777215
         MaxLength       =   50
         Size            =   "4842;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textSP15 
         Height          =   300
         Left            =   1320
         TabIndex        =   15
         Top             =   2956
         Width           =   375
         VariousPropertyBits=   671105051
         BackColor       =   16777215
         MaxLength       =   1
         Size            =   "661;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textSP17 
         Height          =   300
         Left            =   5640
         TabIndex        =   17
         Top             =   3273
         Width           =   495
         VariousPropertyBits=   671105051
         BackColor       =   16777215
         MaxLength       =   2
         Size            =   "873;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textSP29 
         Height          =   300
         Left            =   1320
         TabIndex        =   20
         Top             =   4224
         Width           =   2745
         VariousPropertyBits=   671105051
         BackColor       =   16777215
         MaxLength       =   50
         Size            =   "4842;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCU79 
         Height          =   480
         Left            =   1320
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   4541
         Width           =   7395
         VariousPropertyBits=   -1466941409
         ScrollBars      =   2
         Size            =   "13044;847"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textSP26 
         Height          =   300
         Left            =   -73560
         TabIndex        =   24
         Top             =   420
         Width           =   1215
         VariousPropertyBits=   671105051
         BackColor       =   16777215
         MaxLength       =   8
         Size            =   "2143;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textSP30 
         Height          =   300
         Left            =   -73560
         TabIndex        =   29
         Top             =   1380
         Width           =   7275
         VariousPropertyBits=   671105051
         BackColor       =   16777215
         MaxLength       =   60
         Size            =   "12832;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textSP37 
         Height          =   300
         Left            =   -73560
         TabIndex        =   32
         Top             =   2340
         Width           =   1215
         VariousPropertyBits=   671105051
         BackColor       =   16777215
         MaxLength       =   8
         Size            =   "2143;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textFA29 
         Height          =   1020
         Left            =   -74880
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   3885
         Width           =   8655
         VariousPropertyBits=   -1466941409
         ScrollBars      =   2
         Size            =   "15055;1693"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textSP27 
         Height          =   300
         Left            =   -73560
         TabIndex        =   25
         Top             =   740
         Width           =   2415
         VariousPropertyBits=   671105051
         BackColor       =   16777215
         MaxLength       =   50
         Size            =   "4260;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textSP31 
         Height          =   300
         Left            =   -70455
         TabIndex        =   26
         Top             =   740
         Width           =   615
         VariousPropertyBits=   671105051
         BackColor       =   16777215
         MaxLength       =   2
         Size            =   "1085;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textSP33 
         Height          =   300
         Left            =   -67710
         TabIndex        =   27
         Top             =   740
         Width           =   615
         VariousPropertyBits=   671105051
         BackColor       =   16777215
         MaxLength       =   1
         Size            =   "1085;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textSP09 
         Height          =   300
         Left            =   1320
         TabIndex        =   11
         Top             =   2639
         Width           =   495
         VariousPropertyBits=   671105051
         BackColor       =   16777215
         MaxLength       =   3
         Size            =   "873;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textSP17_2 
         Height          =   300
         Left            =   6180
         TabIndex        =   97
         TabStop         =   0   'False
         Top             =   3273
         Width           =   1455
         VariousPropertyBits=   671105055
         BackColor       =   16777215
         Size            =   "2566;529"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.ComboBox cboContact 
         Height          =   315
         Left            =   5640
         TabIndex        =   21
         Top             =   4217
         Width           =   1770
         VariousPropertyBits=   679495707
         DisplayStyle    =   7
         Size            =   "3122;556"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textSP35_2 
         Height          =   300
         Left            =   -72330
         TabIndex        =   96
         TabStop         =   0   'False
         Top             =   2660
         Width           =   5895
         VariousPropertyBits=   671105055
         BackColor       =   16777215
         Size            =   "10398;529"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textSP37_2 
         Height          =   300
         Left            =   -72330
         TabIndex        =   95
         TabStop         =   0   'False
         Top             =   2340
         Width           =   5895
         VariousPropertyBits=   671105055
         BackColor       =   16777215
         Size            =   "10398;529"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textSP26_2 
         Height          =   300
         Left            =   -72300
         TabIndex        =   94
         TabStop         =   0   'False
         Top             =   420
         Width           =   5895
         VariousPropertyBits=   671105055
         BackColor       =   16777215
         Size            =   "10398;529"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textSP67_2 
         Height          =   300
         Left            =   -72090
         TabIndex        =   93
         TabStop         =   0   'False
         Top             =   3300
         Width           =   5715
         VariousPropertyBits=   671105055
         BackColor       =   16777215
         Size            =   "10081;529"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textSP08_2 
         Height          =   300
         Left            =   2490
         TabIndex        =   92
         TabStop         =   0   'False
         Top             =   2322
         Width           =   6285
         VariousPropertyBits=   671105055
         BackColor       =   16777215
         Size            =   "11086;529"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textSP09_2 
         Height          =   300
         Left            =   1860
         TabIndex        =   91
         TabStop         =   0   'False
         Top             =   2639
         Width           =   1575
         VariousPropertyBits=   671105055
         BackColor       =   16777215
         Size            =   "2778;529"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lblSP61 
         Height          =   255
         Left            =   -73740
         TabIndex        =   90
         Top             =   450
         Width           =   1200
         BackColor       =   16777215
         VariousPropertyBits=   27
         Size            =   "2117;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lblSP68 
         Height          =   255
         Left            =   -73740
         TabIndex        =   89
         Top             =   750
         Width           =   1200
         BackColor       =   16777215
         VariousPropertyBits=   27
         Size            =   "2117;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lblSP69 
         Height          =   255
         Left            =   -73740
         TabIndex        =   88
         Top             =   1050
         Width           =   1200
         BackColor       =   16777215
         VariousPropertyBits=   27
         Size            =   "2117;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lblSP70 
         Height          =   255
         Left            =   -73590
         TabIndex        =   87
         Top             =   1350
         Width           =   5235
         BackColor       =   16777215
         VariousPropertyBits=   27
         Size            =   "9234;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label78 
         AutoSize        =   -1  'True
         Caption         =   "分所銷卷備註："
         Height          =   180
         Left            =   -74880
         TabIndex        =   86
         Top             =   1380
         Width           =   1260
      End
      Begin VB.Label Label79 
         AutoSize        =   -1  'True
         Caption         =   "分所銷卷員："
         Height          =   180
         Left            =   -74880
         TabIndex        =   85
         Top             =   1080
         Width           =   1080
      End
      Begin VB.Label Label80 
         AutoSize        =   -1  'True
         Caption         =   "分所銷卷日："
         Height          =   180
         Left            =   -74880
         TabIndex        =   84
         Top             =   780
         Width           =   1080
      End
      Begin VB.Label Label81 
         AutoSize        =   -1  'True
         Caption         =   "北所銷卷日："
         Height          =   180
         Left            =   -74880
         TabIndex        =   83
         Top             =   480
         Width           =   1080
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "以EMail通知:        (Y:是   D:僅D/N）"
         Height          =   180
         Index           =   76
         Left            =   3690
         TabIndex        =   81
         Top             =   3016
         Width           =   2715
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "EMail同時寄紙本 :         ( Y:是)"
         Height          =   180
         Index           =   112
         Left            =   6480
         TabIndex        =   80
         Top             =   3016
         Width           =   2325
      End
      Begin VB.Label Label32 
         Caption         =   "聯絡人2 :"
         Height          =   255
         Left            =   -74880
         TabIndex        =   79
         Top             =   2043
         Width           =   855
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "特殊出名公司 :         ( J:智權公司 空白:系統預設)"
         Height          =   180
         Index           =   117
         Left            =   120
         TabIndex        =   78
         Top             =   5100
         Width           =   3735
      End
      Begin VB.Label Label1 
         Caption         =   "與他案合併計算結餘，請於案件備註欄註明""與某案號合併計算結餘""！"
         ForeColor       =   &H000000FF&
         Height          =   360
         Index           =   172
         Left            =   5670
         TabIndex        =   77
         Top             =   3780
         Width           =   3000
      End
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         Caption         =   "CLIENT_MATTER_ID :"
         Height          =   180
         Left            =   -74880
         TabIndex        =   76
         Top             =   1120
         Width           =   1770
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "接洽人:"
         Height          =   180
         Index           =   160
         Left            =   4320
         TabIndex        =   75
         Top             =   4284
         Width           =   585
      End
      Begin VB.Label Label33 
         Caption         =   "商品組群 :"
         Height          =   255
         Left            =   120
         TabIndex        =   74
         Top             =   1945
         Width           =   1095
      End
      Begin VB.Label Label17 
         Caption         =   "商品類別 :"
         Height          =   255
         Left            =   120
         TabIndex        =   73
         Top             =   1568
         Width           =   1095
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "聯絡人部門 :"
         Height          =   180
         Left            =   -74880
         TabIndex        =   72
         Top             =   1760
         Width           =   990
      End
      Begin VB.Label Label12 
         Caption         =   "案件名稱 :"
         Height          =   255
         Left            =   120
         TabIndex        =   71
         Top             =   617
         Width           =   1215
      End
      Begin VB.Label Label7 
         Caption         =   "D/N固定列印對象 :"
         Height          =   255
         Left            =   -74880
         TabIndex        =   70
         Top             =   3323
         Width           =   1545
      End
      Begin VB.Label Label1 
         Caption         =   "本所案號 :"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   68
         Top             =   323
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "案件名稱(中) :"
         Height          =   255
         Left            =   120
         TabIndex        =   67
         Top             =   617
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "案件名稱(英) :"
         Height          =   255
         Left            =   120
         TabIndex        =   66
         Top             =   934
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "案件名稱(日) :"
         Height          =   255
         Left            =   120
         TabIndex        =   65
         Top             =   1251
         Width           =   1215
      End
      Begin VB.Label Label5 
         Caption         =   "申請人 :"
         Height          =   255
         Left            =   120
         TabIndex        =   64
         Top             =   2345
         Width           =   1215
      End
      Begin VB.Label Label6 
         Caption         =   "申請日 :"
         Height          =   255
         Left            =   4320
         TabIndex        =   63
         Top             =   2662
         Width           =   1215
      End
      Begin VB.Label Label8 
         Caption         =   "閉卷日期 :"
         Height          =   255
         Left            =   120
         TabIndex        =   62
         Top             =   3296
         Width           =   1215
      End
      Begin VB.Label Label9 
         Caption         =   "定稿語文 :"
         Height          =   255
         Left            =   120
         TabIndex        =   61
         Top             =   3613
         Width           =   1095
      End
      Begin VB.Label Label10 
         Caption         =   "分所案號 :"
         Height          =   255
         Left            =   120
         TabIndex        =   60
         Top             =   3930
         Width           =   1215
      End
      Begin VB.Label Label13 
         Caption         =   "是否閉卷 :"
         Height          =   255
         Left            =   120
         TabIndex        =   59
         Top             =   2979
         Width           =   975
      End
      Begin VB.Label Label14 
         Caption         =   "閉卷原因 :"
         Height          =   255
         Left            =   4320
         TabIndex        =   58
         Top             =   3296
         Width           =   1095
      End
      Begin VB.Label Label16 
         Caption         =   "客戶案件案號 :"
         Height          =   255
         Left            =   120
         TabIndex        =   57
         Top             =   4247
         Width           =   1215
      End
      Begin VB.Label Label18 
         Caption         =   "( 1:中 2:英 3:日 )"
         Height          =   255
         Left            =   1800
         TabIndex        =   56
         Top             =   3613
         Width           =   1455
      End
      Begin VB.Label Label19 
         Caption         =   "( Y:閉卷 )"
         Height          =   255
         Left            =   1800
         TabIndex        =   54
         Top             =   2979
         Width           =   1575
      End
      Begin VB.Label Label36 
         Caption         =   "客戶備註 :"
         Height          =   255
         Left            =   120
         TabIndex        =   52
         Top             =   4545
         Width           =   975
      End
      Begin VB.Label Label20 
         Caption         =   "FC代理人 :"
         Height          =   255
         Left            =   -74880
         TabIndex        =   51
         Top             =   443
         Width           =   855
      End
      Begin VB.Label Label21 
         Caption         =   "聯絡人 :"
         Height          =   255
         Left            =   -74880
         TabIndex        =   50
         Top             =   1403
         Width           =   855
      End
      Begin VB.Label Label22 
         Caption         =   "固定請款對象 :"
         Height          =   255
         Left            =   -74880
         TabIndex        =   49
         Top             =   2363
         Width           =   1215
      End
      Begin VB.Label Label23 
         Caption         =   "副本收受人 :"
         Height          =   255
         Left            =   -74880
         TabIndex        =   48
         Top             =   2683
         Width           =   1335
      End
      Begin VB.Label Label24 
         Caption         =   "代理人備註 :"
         Height          =   255
         Left            =   -74880
         TabIndex        =   47
         Top             =   3600
         Width           =   1095
      End
      Begin VB.Label Label25 
         Caption         =   "彼所案號 :"
         Height          =   255
         Left            =   -74880
         TabIndex        =   46
         Top             =   763
         Width           =   975
      End
      Begin VB.Label Label26 
         Caption         =   "折扣 :"
         Height          =   255
         Left            =   -70920
         TabIndex        =   45
         Top             =   763
         Width           =   735
      End
      Begin VB.Label Label27 
         Caption         =   "D/N是否列印申請人 :"
         Height          =   255
         Left            =   -69420
         TabIndex        =   44
         Top             =   763
         Width           =   1695
      End
      Begin VB.Label Label28 
         Caption         =   "副本聯絡人 :"
         Height          =   255
         Left            =   -74880
         TabIndex        =   42
         Top             =   3003
         Width           =   1215
      End
      Begin VB.Label Label29 
         Caption         =   "%"
         Height          =   255
         Left            =   -69825
         TabIndex        =   40
         Top             =   690
         Width           =   255
      End
      Begin VB.Label Label30 
         Caption         =   "( Y:印 )"
         Height          =   255
         Left            =   -67035
         TabIndex        =   39
         Top             =   763
         Width           =   735
      End
      Begin VB.Label Label11 
         Caption         =   "申請國家 :"
         Height          =   255
         Left            =   120
         TabIndex        =   38
         Top             =   2662
         Width           =   1095
      End
   End
   Begin MSComctlLib.ImageList ImgList 
      Left            =   8580
      Top             =   1200
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
            Picture         =   "frm02050205.frx":0070
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm02050205.frx":038C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm02050205.frx":06A8
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm02050205.frx":0884
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm02050205.frx":0BA0
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm02050205.frx":0EBC
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm02050205.frx":11D8
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm02050205.frx":14F4
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm02050205.frx":1810
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm02050205.frx":1B2C
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm02050205.frx":1E48
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbar 
      Align           =   1  '對齊表單上方
      Height          =   576
      Left            =   0
      TabIndex        =   69
      Top             =   0
      Width           =   9132
      _ExtentX        =   16108
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
   Begin MSForms.TextBox textCUID 
      Height          =   285
      Left            =   120
      TabIndex        =   82
      Top             =   630
      Width           =   8385
      VariousPropertyBits=   671105055
      Size            =   "14790;503"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
End
Attribute VB_Name = "frm02050205"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/12/20 Form2.0已修改 textSP05...(案件名)/textSP08_2/cboContact/textSP26_2/textSP30/textSP75/textSP37_2/textSP35_2/textSP36/textSP67_2/textFA29/lblSP69/lblSP70/textSP18/textCU79
'Memo By Sindy 2012/12/4 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/29 員工編號欄已修改
'Memo By Sindy 2010/8/10 日期欄已修改
Option Explicit

'Const MAX_FIELD = 61
'Const MAX_FIELD = 67
Dim MAX_FIELD As Integer
' 宣告欄位內容結構
Private Type FIELDITEM
   fiName As String
   fiOldData As String
   fiNewData As String
   fiType As Integer
End Type
'Dim m_FieldList(MAX_FIELD) As FIELDITEM
Dim m_FieldList() As FIELDITEM
' 變數宣告區
'Dim m_Recordset As New ADODB.Recordset
'Modify By Sindy 2012/2/20 改可以外部傳
'Dim m_EditMode As Integer
Public m_EditMode As Integer

' 辦識其為外商還是內商的程式
' 0 表內商
' 1 表外商
Dim m_SysKind As Integer
' 第一筆資料的本所案號
Dim m_FirstSP(4) As String
' 最後一筆資料的本所案號
Dim m_LastSP(4) As String
' 目前正在顯示的本所案號
Dim m_CurrSP(4) As String
'
Dim m_QuerySystem As String
' 90.07.13 modify by louis (執行各項功能的權限)
Dim m_bInsert As Boolean
Dim m_bUpdate As Boolean
Dim m_bDelete As Boolean
Dim m_bQuery As Boolean
Dim strChkCuAreaMail As String, strChkCuAreaMailTo As String 'Added by Lydia 2017/06/19 檢查收文智權人員和客戶智權人員不同業務區要發mail的內文和通知人員
Dim m_MeTrackMode  As String 'Added by Lydia 2021/12/20 Form2.0 記錄鍵盤傳入順序

'Add By Sindy 2012/2/20
' 設定顯示的本所案號
Public Sub SetCurrKey(Optional ByVal strKEY01 As String = Empty, Optional ByVal strKEY02 As String = Empty, Optional ByVal strKEY03 As String = Empty, Optional ByVal strKEY04 As String = Empty)
   If IsEmptyText(strKEY01) Or IsEmptyText(strKEY02) Then
      m_CurrSP(0) = Empty
      m_CurrSP(1) = Empty
      m_CurrSP(2) = Empty
      m_CurrSP(3) = Empty
      Exit Sub
   End If
   m_CurrSP(0) = strKEY01
   m_CurrSP(1) = strKEY02
   m_CurrSP(2) = strKEY03
   If IsEmptyText(m_CurrSP(2)) Then
      m_CurrSP(2) = "0"
   End If
   m_CurrSP(3) = strKEY04
   If IsEmptyText(m_CurrSP(3)) Then
      m_CurrSP(3) = "00"
   End If
End Sub

Private Sub RefreshRange()
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   
   'Select Case m_SysKind
   '   Case 0:
   '      strSQL = "SELECT SP01,SP02,SP03,SP04 FROM SERVICEPRACTICE " & _
   '               "WHERE SP01 IN " & m_QuerySystem & " AND " & _
   '                     "SP02 = (SELECT MIN(SP02) FROM SERVICEPRACTICE WHERE SP01 IN " & m_QuerySystem & ") AND " & _
   '                     "SP03 = (SELECT MIN(SP03) FROM SERVICEPRACTICE WHERE SP01 IN " & m_QuerySystem & " AND SP02 = (SELECT MIN(SP02) FROM SERVICEPRACTICE WHERE SP01 IN " & m_QuerySystem & ")) AND " & _
   '                     "SP04 = (SELECT MIN(SP04) FROM SERVICEPRACTICE WHERE SP01 IN " & m_QuerySystem & " AND SP02 = (SELECT MIN(SP02) FROM SERVICEPRACTICE WHERE SP01 IN " & m_QuerySystem & ") AND SP03 = (SELECT MIN(SP03) FROM SERVICEPRACTICE WHERE SP01 IN " & m_QuerySystem & " AND SP02 = (SELECT MIN(SP02) FROM SERVICEPRACTICE WHERE SP01 IN " & m_QuerySystem & "))) "
   '   Case 1:
   '      strSQL = "SELECT SP01,SP02,SP03,SP04 FROM SERVICEPRACTICE " & _
   '               "WHERE SP01 = '" & "S" & "' AND " & _
   '                     "SP02 = (SELECT MIN(SP02) FROM SERVICEPRACTICE WHERE SP01 = '" & "S" & "') AND " & _
   '                     "SP03 = (SELECT MIN(SP03) FROM SERVICEPRACTICE WHERE SP01 = '" & "S" & "' AND SP02 = (SELECT MIN(SP02) FROM SERVICEPRACTICE WHERE SP01 = '" & "S" & "' )) AND " & _
   '                     "SP04 = (SELECT MIN(SP04) FROM SERVICEPRACTICE WHERE SP01 = '" & "S" & "' AND SP02 = (SELECT MIN(SP02) FROM SERVICEPRACTICE WHERE SP01 = '" & "S" & "' ) AND SP03 = (SELECT MIN(SP03) FROM SERVICEPRACTICE WHERE SP01 = '" & "S" & "' AND SP02 = (SELECT MIN(SP02) FROM SERVICEPRACTICE WHERE SP01 = '" & "S" & "' ))) "
   'End Select
   
   strSql = "SELECT SP01,SP02,SP03,SP04 FROM SERVICEPRACTICE " & _
            "WHERE SP01 = (SELECT MIN(SP01) FROM SERVICEPRACTICE " & _
                          "WHERE SP01 IN " & m_QuerySystem & ") AND " & _
                  "SP02 = (SELECT MIN(SP02) FROM SERVICEPRACTICE " & _
                          "WHERE SP01 = (SELECT MIN(SP01) FROM SERVICEPRACTICE " & _
                                        "WHERE SP01 IN " & m_QuerySystem & ")) AND " & _
                  "SP03 = (SELECT MIN(SP03) FROM SERVICEPRACTICE " & _
                          "WHERE SP01 = (SELECT MIN(SP01) FROM SERVICEPRACTICE " & _
                                        "WHERE SP01 IN " & m_QuerySystem & ") AND " & _
                                "SP02 = (SELECT MIN(SP02) FROM SERVICEPRACTICE " & _
                                        "WHERE SP01 = (SELECT MIN(SP01) FROM SERVICEPRACTICE " & _
                                                      "WHERE SP01 IN " & m_QuerySystem & "))) AND " & _
                  "SP04 = (SELECT MIN(SP04) FROM SERVICEPRACTICE " & _
                          "WHERE SP01 = (SELECT MIN(SP01) FROM SERVICEPRACTICE " & _
                                        "WHERE SP01 IN " & m_QuerySystem & ") AND " & _
                                "SP02 = (SELECT MIN(SP02) FROM SERVICEPRACTICE " & _
                                        "WHERE SP01 = (SELECT MIN(SP01) FROM SERVICEPRACTICE " & _
                                                      "WHERE SP01 IN " & m_QuerySystem & ")) AND " & _
                                "SP03 = (SELECT MIN(SP03) FROM SERVICEPRACTICE " & _
                                        "WHERE SP01 = (SELECT MIN(SP01) FROM SERVICEPRACTICE " & _
                                                      "WHERE SP01 IN " & m_QuerySystem & ") AND " & _
                                              "SP02 = (SELECT MIN(SP02) FROM SERVICEPRACTICE " & _
                                                      "WHERE SP01 = (SELECT MIN(SP01) FROM SERVICEPRACTICE " & _
                                                                    "WHERE SP01 IN " & m_QuerySystem & ")))) "
   
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("SP01")) = False Then: m_FirstSP(0) = rsTmp.Fields("SP01")
      If IsNull(rsTmp.Fields("SP02")) = False Then: m_FirstSP(1) = rsTmp.Fields("SP02")
      If IsNull(rsTmp.Fields("SP03")) = False Then: m_FirstSP(2) = rsTmp.Fields("SP03")
      If IsNull(rsTmp.Fields("SP04")) = False Then: m_FirstSP(3) = rsTmp.Fields("SP04")
   End If
   rsTmp.Close

   'Select Case m_SysKind
   '   Case 0:
   '      strSQL = "SELECT SP01,SP02,SP03,SP04 FROM SERVICEPRACTICE " & _
   '               "WHERE SP01 IN " & m_QuerySystem & " AND " & _
   '                     "SP02 = (SELECT MAX(SP02) FROM SERVICEPRACTICE WHERE SP01 IN " & m_QuerySystem & ") AND " & _
   '                     "SP03 = (SELECT MAX(SP03) FROM SERVICEPRACTICE WHERE SP01 IN " & m_QuerySystem & " AND SP02 = (SELECT MAX(SP02) FROM SERVICEPRACTICE WHERE SP01 IN " & m_QuerySystem & ")) AND " & _
   '                     "SP04 = (SELECT MAX(SP04) FROM SERVICEPRACTICE WHERE SP01 IN " & m_QuerySystem & " AND SP02 = (SELECT MAX(SP02) FROM SERVICEPRACTICE WHERE SP01 IN " & m_QuerySystem & ") AND SP03 = (SELECT MAX(SP03) FROM SERVICEPRACTICE WHERE SP01 IN " & m_QuerySystem & " AND SP02 = (SELECT MAX(SP02) FROM SERVICEPRACTICE WHERE SP01 IN " & m_QuerySystem & "))) "
   '   Case 1:
   '      strSQL = "SELECT SP01,SP02,SP03,SP04 FROM SERVICEPRACTICE " & _
   '               "WHERE SP01 = '" & "S" & "' AND " & _
   '                     "SP02 = (SELECT MAX(SP02) FROM SERVICEPRACTICE WHERE SP01 = '" & "S" & "') AND " & _
   '                     "SP03 = (SELECT MAX(SP03) FROM SERVICEPRACTICE WHERE SP01 = '" & "S" & "' AND SP02 = (SELECT MAX(SP02) FROM SERVICEPRACTICE WHERE SP01 = '" & "S" & "' )) AND " & _
   '                     "SP04 = (SELECT MAX(SP04) FROM SERVICEPRACTICE WHERE SP01 = '" & "S" & "' AND SP02 = (SELECT MAX(SP02) FROM SERVICEPRACTICE WHERE SP01 = '" & "S" & "' ) AND SP03 = (SELECT MAX(SP03) FROM SERVICEPRACTICE WHERE SP01 = '" & "S" & "' AND SP02 = (SELECT MAX(SP02) FROM SERVICEPRACTICE WHERE SP01 = '" & "S" & "' ))) "
   'End Select
   strSql = "SELECT SP01,SP02,SP03,SP04 FROM SERVICEPRACTICE " & _
            "WHERE SP01 = (SELECT MAX(SP01) FROM SERVICEPRACTICE " & _
                          "WHERE SP01 IN " & m_QuerySystem & ") AND " & _
                  "SP02 = (SELECT MAX(SP02) FROM SERVICEPRACTICE " & _
                          "WHERE SP01 = (SELECT MAX(SP01) FROM SERVICEPRACTICE " & _
                                        "WHERE SP01 IN " & m_QuerySystem & ")) AND " & _
                  "SP03 = (SELECT MAX(SP03) FROM SERVICEPRACTICE " & _
                          "WHERE SP01 = (SELECT MAX(SP01) FROM SERVICEPRACTICE " & _
                                        "WHERE SP01 IN " & m_QuerySystem & ") AND " & _
                                "SP02 = (SELECT MAX(SP02) FROM SERVICEPRACTICE " & _
                                        "WHERE SP01 = (SELECT MAX(SP01) FROM SERVICEPRACTICE " & _
                                                      "WHERE SP01 IN " & m_QuerySystem & "))) AND " & _
                  "SP04 = (SELECT MAX(SP04) FROM SERVICEPRACTICE " & _
                          "WHERE SP01 = (SELECT MAX(SP01) FROM SERVICEPRACTICE " & _
                                        "WHERE SP01 IN " & m_QuerySystem & ") AND " & _
                                "SP02 = (SELECT MAX(SP02) FROM SERVICEPRACTICE " & _
                                        "WHERE SP01 = (SELECT MAX(SP01) FROM SERVICEPRACTICE " & _
                                                      "WHERE SP01 IN " & m_QuerySystem & ")) AND " & _
                                "SP03 = (SELECT MAX(SP03) FROM SERVICEPRACTICE " & _
                                        "WHERE SP01 = (SELECT MAX(SP01) FROM SERVICEPRACTICE " & _
                                                      "WHERE SP01 IN " & m_QuerySystem & ") AND " & _
                                              "SP02 = (SELECT MAX(SP02) FROM SERVICEPRACTICE " & _
                                                      "WHERE SP01 = (SELECT MAX(SP01) FROM SERVICEPRACTICE " & _
                                                                    "WHERE SP01 IN " & m_QuerySystem & ")))) "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("SP01")) = False Then: m_LastSP(0) = rsTmp.Fields("SP01")
      If IsNull(rsTmp.Fields("SP02")) = False Then: m_LastSP(1) = rsTmp.Fields("SP02")
      If IsNull(rsTmp.Fields("SP03")) = False Then: m_LastSP(2) = rsTmp.Fields("SP03")
      If IsNull(rsTmp.Fields("SP04")) = False Then: m_LastSP(3) = rsTmp.Fields("SP04")
   End If
   rsTmp.Close
   
   Set rsTmp = Nothing
End Sub

Private Sub FilterSystem()
   Dim nIndex As Integer
   Dim nCount As Integer
   Dim strSys As String
   Dim strTemp As String
   m_QuerySystem = Empty
   
   strSys = GetUserSystemKind
   nCount = GetSubStringCount(strSys)
   For nIndex = 1 To nCount
      strTemp = GetSubString(strSys, nIndex)
      Select Case strTemp
         Case "T", "TF", "CFT", "FCT", "TB", "TC", "TD", "TM", "P", "CFP", "FCP", "L", "CFL", "FCL", "LA":
            GoTo NextRecord
         Case Else:
            If IsEmptyText(m_QuerySystem) = False Then m_QuerySystem = m_QuerySystem & ","
            m_QuerySystem = m_QuerySystem & "'" & strTemp & "'"
      End Select
NextRecord:
   Next nIndex
   
   m_QuerySystem = "(" & m_QuerySystem & ")"

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

Private Sub ButtonRelation_Click()
   Dim strSP01 As String
   Dim strSP02 As String
   Dim strSP03 As String
   Dim strSP04 As String
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   strSP01 = textSP01
   strSP02 = textSP02
   strSP03 = textSP03
   If IsEmptyText(strSP03) = True Then: strSP03 = "0"
   strSP04 = textSP04
   If IsEmptyText(strSP04) = True Then: strSP04 = "00"
   
   If IsEmptyText(strSP01) = True Or IsEmptyText(strSP02) = True Then
      strTit = "檢核資料"
      strMsg = "請先輸入本所案號"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
   Else
      Where1103ComeFrom Me, strSP01, strSP02, strSP03, strSP04
   End If
End Sub

'Added by Lydia 2016/11/24 各項指示
Private Sub cmdIns_Click()
   If Me.textSP01.Text = "" Or Me.textSP02.Text = "" Then
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
   
   frm12040159.SetParent "E", Trim(Me.textSP01.Text & Me.textSP02.Text & Me.textSP03.Text & Me.textSP04.Text), Me
   frm12040159.Show
End Sub

'Add by Amy 2016/08/26
Private Sub Command3_Click()
    frmPic001.oCP01 = textSP01
    frmPic001.oCP02 = textSP02
    frmPic001.oCP03 = textSP03
    frmPic001.oCP04 = textSP04
    frmPic001.StrMenu
    frmPic001.CanScan
    frmPic001.SetSeekCmdok 'Add by Amy 2018/07/16
    frmPic001.Show vbModal
    '檢查有無代表圖
    Call ReadPic
End Sub

Private Sub Form_Initialize()
   MAX_FIELD = tf_SP
   ReDim m_FieldList(MAX_FIELD) As FIELDITEM
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   Call PUB_SaveMeTrackMode(m_MeTrackMode, 1, KeyCode)  'Added by Lydia 2021/12/20 Form2.0 記錄鍵盤傳入順序
   
'Memo by Lydia 2021/12/20 從Form_KeyDown搬來
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
      Case vbKeyF9, vbKeyF10:
         If m_EditMode <> 0 Then
            'Added by Lydia 2021/12/20 Form2.0 控制Function鍵：記錄鍵盤傳入順序，判斷是否可執行
            If KeyCode = vbKeyF9 Then
                If PUB_ChkMeTrackMode(m_MeTrackMode) = False Then
                    Exit Sub
                End If
            End If
            'end 2021/12/20
            OnAction KeyCode
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

' Load Form
Private Sub Form_Load()

   tabCtrl.Tab = 0
   
   ' 90.07.13 modify by louis (取得使用者執行各項功能的權限)
   m_bInsert = IsUserHasRightOfFunction("frm02050205", strAdd, False)
   m_bUpdate = IsUserHasRightOfFunction("frm02050205", strEdit, False)
   m_bDelete = IsUserHasRightOfFunction("frm02050205", strDel, False)
   m_bQuery = IsUserHasRightOfFunction("frm02050205", strFind, False)
   
   m_EditMode = 0
   MoveFormToCenter Me
   
   textCUID.BackColor = &H8000000F
   textSP08_2.BackColor = &H8000000F
   textSP09_2.BackColor = &H8000000F
   textSP17_2.BackColor = &H8000000F
   textSP26_2.BackColor = &H8000000F
   textSP35_2.BackColor = &H8000000F
   textSP37_2.BackColor = &H8000000F
   textCU79.BackColor = &H8000000F
   textFA29.BackColor = &H8000000F
   textSP67_2.BackColor = &H8000000F
   textSP80.BackColor = &H8000000F 'Add By Sindy 2017/11/9
   textSP83.BackColor = &H8000000F 'Add By Sindy 2017/11/9
   'Added by Lydia 2021/12/20
   lblSP61.BackColor = &H8000000F
   lblSP68.BackColor = &H8000000F
   lblSP69.BackColor = &H8000000F
   lblSP70.BackColor = &H8000000F
   
   ' 設定可以使用的系統別
   FilterSystem
   
   InitialField
   'QueryDB
   'Modify By Sindy 2012/2/20 Mark
'   RefreshRange
'   ShowFirstRecord
'   SetCtrlReadOnly True
'   UpdateToolbarState
   
   'Modify By Sindy 2012/2/20
   If Not IsEmptyText(m_CurrSP(0)) And Not IsEmptyText(m_CurrSP(1)) And Not IsEmptyText(m_CurrSP(2)) And Not IsEmptyText(m_CurrSP(3)) Then
      ShowCurrRecord m_CurrSP(0), m_CurrSP(1), m_CurrSP(2), m_CurrSP(3)
      UpdateToolbarState
      SetCtrlReadOnly True
   Else
   '2012/2/20 End
      'Add By Cheng 2002/01/04
      SetQueryStatus
   End If
   
   'Added by Lydia 2020/05/05 各項指示：顯示按鈕
   If strSrvDate(1) >= 各項指示啟用日 Then
      cmdIns.Visible = True
   Else
      cmdIns.Visible = False
      textSP18.Top = 360
      textSP18.Height = 4215
   End If
   
End Sub

' 初始化欄位陣列
Private Sub InitialField()
   Dim nIndex As Integer
   Dim strTmp As String
   ' 初始化欄位陣列
   For nIndex = 1 To MAX_FIELD
      strTmp = Format(nIndex, "00")
      m_FieldList(nIndex - 1).fiName = "SP" & strTmp
      m_FieldList(nIndex - 1).fiOldData = Empty
      m_FieldList(nIndex - 1).fiNewData = Empty
      m_FieldList(nIndex - 1).fiType = 0
      Select Case nIndex
         Case 10, 12, 16, 20, 21, 31, 39, 40, 53, 54, 56, 57:
            m_FieldList(nIndex - 1).fiType = 1
      End Select
   Next nIndex
   lblSP61 = ""
   lblSP68 = ""
   lblSP69 = ""
   lblSP70 = ""
End Sub

' 設定欄位的內容
Private Sub SetFieldNewData(ByVal strName As String, ByVal strData As String)
   Dim nIndex As Integer
   For nIndex = 0 To MAX_FIELD - 1
      If strName = m_FieldList(nIndex).fiName Then
         m_FieldList(nIndex).fiNewData = strData
         Exit For
      End If
   Next nIndex
End Sub

' 更新欄位的內容
Private Sub UpdateFieldNewData()
   Dim nIndex As Integer
   Dim strData1 As String
   Dim strData2 As String
   
   SetFieldNewData "SP01", textSP01
   SetFieldNewData "SP02", textSP02
   If IsEmptyText(textSP03) = True Then: textSP03 = "0"
   SetFieldNewData "SP03", textSP03
   If IsEmptyText(textSP04) = True Then: textSP04 = "00"
   SetFieldNewData "SP04", textSP04
    Select Case Me.textSP01.Text
    Case "TS", "S"
        SetFieldNewData "SP05", textSP05_1
    Case Else
        SetFieldNewData "SP05", textSP05
    End Select
   SetFieldNewData "SP06", textSP06: SetFieldNewData "SP07", textSP07
   ' 申請人
   If IsEmptyText(textSP08) = False Then
      SetFieldNewData "SP08", textSP08 & String(9 - Len(textSP08), "0")
   Else
      SetFieldNewData "SP08", textSP08
   End If
   SetFieldNewData "SP09", textSP09
   ' 申請日
   If IsEmptyText(textSP10) = False Then
      SetFieldNewData "SP10", DBDATE(textSP10)
   Else
      SetFieldNewData "SP10", textSP10
   End If
   SetFieldNewData "SP15", textSP15
   ' 閉卷日期
   If IsEmptyText(textSP16) = False Then
      SetFieldNewData "SP16", DBDATE(textSP16)
   Else
      SetFieldNewData "SP16", textSP16
   End If
   SetFieldNewData "SP17", textSP17
   SetFieldNewData "SP18", textSP18
   ' FC代理人
   If IsEmptyText(textSP26) = False Then
      SetFieldNewData "SP26", textSP26 & String(9 - Len(textSP26), "0")
   Else
      SetFieldNewData "SP26", textSP26
   End If
   SetFieldNewData "SP27", textSP27: SetFieldNewData "SP28", textSP28: SetFieldNewData "SP29", textSP29: SetFieldNewData "SP30", textSP30 '聯絡人
   SetFieldNewData "SP75", textSP75 'Add By Sindy 2016/11/4 聯絡人2
   SetFieldNewData "SP31", textSP31: SetFieldNewData "SP33", textSP33: SetFieldNewData "SP34", textSP34: SetFieldNewData "SP71", textSP71
   ' 副本收受人
   If IsEmptyText(textSP35) = False Then
      SetFieldNewData "SP35", textSP35 & String(9 - Len(textSP35), "0")
   Else
      SetFieldNewData "SP35", textSP35
   End If
   SetFieldNewData "SP36", textSP36
   ' 固定請款對象
   If IsEmptyText(textSP37) = False Then
      SetFieldNewData "SP37", textSP37 & String(9 - Len(textSP37), "0")
   Else
      SetFieldNewData "SP37", textSP37
   End If
'   SetFieldNewData "SP61", textSP61
   ' D/N固定列印對象
   If IsEmptyText(textSP67) = False Then
      SetFieldNewData "SP67", textSP67 & String(9 - Len(textSP67), "0")
   Else
      SetFieldNewData "SP67", textSP67
   End If
   'add by nickc 2006/12/07
   SetFieldNewData "SP73", textSP73
   SetFieldNewData "SP74", textSP74
   SetFieldNewData "SP84", textSP84 'Add by Morgan 2010/11/9
   SetFieldNewData "SP85", textSP85 'Add By Sindy 2014/2/10
   SetFieldNewData "SP80", textSP80 'Add By Sindy 2017/11/9
   SetFieldNewData "SP83", textSP83 'Add By Sindy 2017/11/9
End Sub

' 從記錄中更新欄位內容
Private Sub UpdateFieldOldData(ByRef rsSrcTmp As ADODB.Recordset)
   Dim nIndex As Integer
   Dim strTmp As String
   
   For nIndex = 0 To MAX_FIELD - 1
      If m_FieldList(nIndex).fiName <> Empty Then
         If IsNull(rsSrcTmp.Fields(m_FieldList(nIndex).fiName)) = False Then
            m_FieldList(nIndex).fiOldData = rsSrcTmp.Fields(m_FieldList(nIndex).fiName)
            'add by nickc 2006/06/08 因為很多欄位並不顯示在畫面上，所以舊值會跟 null 比而被清掉
            m_FieldList(nIndex).fiNewData = rsSrcTmp.Fields(m_FieldList(nIndex).fiName)
         Else
            m_FieldList(nIndex).fiOldData = Empty
            'add by nickc 2006/06/08 因為很多欄位並不顯示在畫面上，所以舊值會跟 null 比而被清掉
            m_FieldList(nIndex).fiNewData = Empty
         End If
      End If
   Next nIndex
EXITSUB:
End Sub

' 讀取資料庫所有的資料
Private Sub QueryDB()
   'Dim strSQL As String
   
   ' 檢查RecordSet的狀態
   'If m_Recordset.State <> adStateClosed Then
   '   m_Recordset.Close
   'End If
   ' 設定 Query 的命令
   'Select Case m_SysKind
   '   Case 0:
   '      strSQL = "SELECT * FROM ServicePractice " & _
   '               "WHERE SP01 = 'TT' " & _
   '               "ORDER BY SP01, SP02, SP03, SP04"
   '   Case 1:
   '      strSQL = "SELECT * FROM ServicePractice " & _
   '               "WHERE SP01 = 'S' " & _
   '               "ORDER BY SP01, SP02, SP03, SP04"
   'End Select
      
   ' 讀取資料庫
   'm_Recordset.CursorLocation = adUseClient
   'm_Recordset.Open strSQL, cnnConnection, adOpenDynamic
   RefreshRange
End Sub

' 清除欄位內的資料內容
Private Sub ClearField()
   Dim nIndex As Integer
   ' 依照是內商還是外商來判斷
   Select Case m_SysKind
      Case 0:
         textSP01 = Empty
      Case 1:
         textSP01 = "S"
   End Select
   textSP02 = Empty: textSP03 = Empty: textSP04 = Empty: textSP05 = Empty: textSP05_1 = Empty
   textSP06 = Empty: textSP07 = Empty: textSP08 = Empty: textSP09 = Empty: textSP10 = Empty
   textSP15 = Empty: textSP16 = Empty: textSP17 = Empty: textSP18 = Empty
   textSP26 = Empty: textSP27 = Empty: textSP28 = Empty: textSP29 = Empty: textSP30 = Empty: textSP75 = Empty
   textSP31 = Empty: textSP33 = Empty: textSP34 = Empty: textSP35 = Empty: textSP71 = Empty
   textSP36 = Empty: textSP37 = Empty:  textSP67 = Empty: textSP84 = Empty
   textSP80 = Empty: textSP83 = Empty 'Add By Sindy 2017/11/9
   
   textSP08_2 = Empty: textSP09_2 = Empty: textSP17_2 = Empty: textSP26_2 = Empty: textSP35_2 = Empty: textSP37_2 = Empty
   textCU79 = Empty: textFA29 = Empty: textSP67_2 = Empty
   textCUID = Empty
   
   'add by nickc 2006/12/07
   textSP73 = Empty
   textSP74 = Empty
   
   For nIndex = 0 To MAX_FIELD - 1
      m_FieldList(nIndex).fiOldData = Empty
      m_FieldList(nIndex).fiNewData = Empty
   Next nIndex
   lblSP61 = ""
   lblSP68 = ""
   lblSP69 = ""
   lblSP70 = ""
   cboContact.Clear 'Add by Morgan 2008/8/4
   textSP85 = Empty 'Add By Sindy 2014/2/10
End Sub

' 更新各控制項的狀態
Private Sub SetCtrlReadOnly(ByVal bEnable As Boolean)
   textSP01.Locked = bEnable: textSP02.Locked = bEnable: textSP03.Locked = bEnable: textSP04.Locked = bEnable: textSP05.Locked = bEnable: textSP05_1.Locked = bEnable
   textSP06.Locked = bEnable: textSP07.Locked = bEnable: textSP08.Locked = bEnable: textSP09.Locked = bEnable: textSP10.Locked = bEnable
   textSP15.Locked = bEnable: textSP16.Locked = bEnable: textSP17.Locked = bEnable: textSP18.Locked = bEnable
   textSP26.Locked = bEnable: textSP27.Locked = bEnable: textSP28.Locked = bEnable: textSP29.Locked = bEnable: textSP30.Locked = bEnable: textSP75.Locked = bEnable
   textSP31.Locked = bEnable: textSP33.Locked = bEnable: textSP34.Locked = bEnable: textSP35.Locked = bEnable: textSP71.Locked = bEnable
   textSP36.Locked = bEnable: textSP37.Locked = bEnable:  textSP67.Locked = bEnable
   'add by nickc 2006/12/07
   textSP73.Locked = bEnable: textSP74.Locked = bEnable
   textSP84.Locked = bEnable 'Add by Morgan 2010/11/9
   'Modify by Amy 2018/07/03 只有電腦中心才可改 特殊出名公司
   textSP85.Locked = True
   If Pub_StrUserSt03 = "M51" Then
      textSP85.Locked = bEnable 'Add By Sindy 2014/2/10
   End If
   textSP80.Locked = bEnable 'Add By Sindy 2017/11/9
   textSP83.Locked = bEnable 'Add By Sindy 2017/11/9
End Sub
' 更新Key的狀態
Private Sub SetKeyReadOnly(ByVal bEnable As Boolean)
   textSP01.Locked = bEnable: textSP02.Locked = bEnable: textSP03.Locked = bEnable: textSP04.Locked = bEnable
End Sub

' 將資料庫中的資料更新到所有欄位中
Private Sub UpdateCtrlData()
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   
   Command3.Enabled = False 'Add by Amy 2016/08/26
   strSql = "SELECT * FROM ServicePractice " & _
            "WHERE SP01 = '" & m_CurrSP(0) & "' AND " & _
                  "SP02 = '" & m_CurrSP(1) & "' AND " & _
                  "SP03 = '" & m_CurrSP(2) & "' AND " & _
                  "SP04 = '" & m_CurrSP(3) & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
   'modify by sonia 2020/3/23 +S
   'If m_CurrSP(0) = "TS" Then Command3.Enabled = True 'Add by Amy 2016/08/26
   Select Case m_CurrSP(0)
      Case "TS", "S"
         Command3.Enabled = True
      Case Else
   End Select
   'end 2020/3/23
   ClearField
   textSP01 = rsTmp.Fields("SP01")
   textSP02 = rsTmp.Fields("SP02")
   textSP03 = rsTmp.Fields("SP03")
   textSP04 = rsTmp.Fields("SP04")
    Select Case Me.textSP01.Text
    Case "TS", "S"
        If Not IsNull(rsTmp.Fields("SP05")) Then: textSP05_1 = rsTmp.Fields("SP05"): 'End If
    Case Else
        If Not IsNull(rsTmp.Fields("SP05")) Then: textSP05 = rsTmp.Fields("SP05"): 'End If
    End Select
   If Not IsNull(rsTmp.Fields("SP06")) Then: textSP06 = rsTmp.Fields("SP06"): 'End If
   If Not IsNull(rsTmp.Fields("SP07")) Then: textSP07 = rsTmp.Fields("SP07"): 'End If
   If Not IsNull(rsTmp.Fields("SP08")) Then: textSP08 = rsTmp.Fields("SP08"): 'End If
   textSP08.Tag = "" & rsTmp.Fields("SP08") 'Added by Lydia 2024/06/13
   If Not IsNull(rsTmp.Fields("SP09")) Then: textSP09 = rsTmp.Fields("SP09"): 'End If
   If Not IsNull(rsTmp.Fields("SP10")) Then
      textSP10 = ChangeWStringToTString(rsTmp.Fields("SP10"))
   End If
   If Not IsNull(rsTmp.Fields("SP15")) Then: textSP15 = rsTmp.Fields("SP15"): 'End If
   If Not IsNull(rsTmp.Fields("SP16")) Then
      textSP16 = ChangeWStringToTString(rsTmp.Fields("SP16"))
   End If
   If Not IsNull(rsTmp.Fields("SP17")) Then: textSP17 = rsTmp.Fields("SP17"): 'End If
   If Not IsNull(rsTmp.Fields("SP18")) Then: textSP18 = rsTmp.Fields("SP18"): 'End If
   If Not IsNull(rsTmp.Fields("SP26")) Then: textSP26 = Mid(rsTmp.Fields("SP26"), 1, 8): 'End If
   textSP26.Tag = "" & rsTmp.Fields("SP26") 'Added by Lydia 2024/06/13
   If Not IsNull(rsTmp.Fields("SP27")) Then: textSP27 = rsTmp.Fields("SP27"): 'End If
   If Not IsNull(rsTmp.Fields("SP28")) Then: textSP28 = rsTmp.Fields("SP28"): 'End If
   If Not IsNull(rsTmp.Fields("SP29")) Then: textSP29 = rsTmp.Fields("SP29"): 'End If
   If Not IsNull(rsTmp.Fields("SP30")) Then: textSP30 = rsTmp.Fields("SP30"): 'End If
   If Not IsNull(rsTmp.Fields("SP75")) Then: textSP75 = rsTmp.Fields("SP75"): 'Add By Sindy 2016/11/4 聯絡人2
   If Not IsNull(rsTmp.Fields("SP31")) Then: textSP31 = rsTmp.Fields("SP31"): 'End If
   If Not IsNull(rsTmp.Fields("SP33")) Then: textSP33 = rsTmp.Fields("SP33"): 'End If
   If Not IsNull(rsTmp.Fields("SP34")) Then: textSP34 = rsTmp.Fields("SP34"): 'End If
   If Not IsNull(rsTmp.Fields("SP35")) Then: textSP35 = Mid(rsTmp.Fields("SP35"), 1, 8): 'End If
   If Not IsNull(rsTmp.Fields("SP36")) Then: textSP36 = rsTmp.Fields("SP36"): 'End If
   If Not IsNull(rsTmp.Fields("SP37")) Then: textSP37 = Mid(rsTmp.Fields("SP37"), 1, 8): 'End If
'   If Not IsNull(rsTmp.Fields("SP61")) Then: textSP61 = rsTmp.Fields("SP61"): 'End If
   If Not IsNull(rsTmp.Fields("SP67")) Then: textSP67 = Mid(rsTmp.Fields("SP67"), 1, 8): 'End If
   If Not IsNull(rsTmp.Fields("SP71")) Then: textSP71 = rsTmp.Fields("SP71") 'Add by Morgan 2006/10/18
   'add by nickc 2007/03/03
   If Not IsNull(rsTmp.Fields("SP73")) Then: textSP73 = rsTmp.Fields("SP73")
   If Not IsNull(rsTmp.Fields("SP74")) Then: textSP74 = rsTmp.Fields("SP74")
   textSP84 = "" & rsTmp.Fields("SP84") 'Add by Morgan 2010/11/9
   If Not IsNull(rsTmp.Fields("SP85")) Then: textSP85 = rsTmp.Fields("SP85") 'Add By Sindy 2014/2/10
   If Not IsNull(rsTmp.Fields("SP80")) Then: textSP80 = rsTmp.Fields("SP80") 'Add By Sindy 2017/11/9
   If Not IsNull(rsTmp.Fields("SP83")) Then: textSP83 = rsTmp.Fields("SP83") 'Add By Sindy 2017/11/9
   
   'Modified by Lydia 2021/12/17 改Form 2.0
   'PUB_AddContact "" & rsTmp("sp08"), cboContact, "" & rsTmp("sp78") 'Add by Morgan 2008/8/4
   PUB_AddContact "" & rsTmp("sp08"), cboContact, "" & rsTmp("sp78"), , True
   UpdateFieldOldData rsTmp
   
   'add by nickc 2006/07/13
   Dim strTemp As String
   If IsNull(rsTmp.Fields("SP61")) = False Then
      If IsEmptyText(rsTmp.Fields("SP61")) = False Then
         strTemp = TAIWANDATE(rsTmp.Fields("SP61"))
         lblSP61 = Format(strTemp, "###/##/##")
      End If
   End If
   If IsNull(rsTmp.Fields("SP68")) = False Then
      If IsEmptyText(rsTmp.Fields("SP68")) = False Then
         strTemp = TAIWANDATE(rsTmp.Fields("SP68"))
         lblSP68 = Format(strTemp, "###/##/##")
      End If
   End If
   If IsNull(rsTmp.Fields("SP69")) = False Then
      If IsEmptyText(rsTmp.Fields("SP69")) = False Then
         lblSP69 = GetStaffName(rsTmp.Fields("SP69"), True)
      End If
   End If
   If Not IsNull(rsTmp.Fields("SP70")) Then: lblSP70 = rsTmp.Fields("SP70")
   
   '檢查有無代表圖
   'modify by sonia 2020/3/23 +S
   'If m_CurrSP(0) = "TS" Then Call ReadPic 'Add by Amy 2016/08/26
   Select Case m_CurrSP(0)
      Case "TS", "S"
         Call ReadPic
      Case Else
   End Select
   'end 2020/3/23
   
   ' 更新顯示 Create 及 Update 的人
   UpdateCUID rsTmp
   
   ' 更新控制項中需帶出的資料
   textSP08_Validate False
   textSP09_Validate False
   textSP17_Validate False
   textSP26_Validate False
   textSP35_Validate False
   textSP37_Validate False
   textSP67_Validate False
   End If
EXITSUB:
   Set rsTmp = Nothing
End Sub

' 更新 Create 及 Update 的人
Private Sub UpdateCUID(ByRef rsSrcTmp As ADODB.Recordset)
   Dim strTemp As String
   Dim strCName As String
   Dim strCDate As String
   Dim strCTime As String
   Dim strUName As String
   Dim strUDate As String
   Dim strUTime As String
   If IsNull(rsSrcTmp.Fields("SP52")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("SP52")) = False Then
         strCName = GetStaffName(rsSrcTmp.Fields("SP52"), True)
      End If
   End If
   If IsNull(rsSrcTmp.Fields("SP53")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("SP53")) = False Then
         strTemp = TAIWANDATE(rsSrcTmp.Fields("SP53"))
         strCDate = Format(strTemp, "###/##/##")
      End If
   End If
   If IsNull(rsSrcTmp.Fields("SP54")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("SP54")) = False Then
         strTemp = rsSrcTmp.Fields("SP54")
         strCTime = Format(strTemp, "##:##")
      End If
   End If
   If IsNull(rsSrcTmp.Fields("SP55")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("SP55")) = False Then
         strUName = GetStaffName(rsSrcTmp.Fields("SP55"), True)
      End If
   End If
   If IsNull(rsSrcTmp.Fields("SP56")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("SP56")) = False Then
         strTemp = TAIWANDATE(rsSrcTmp.Fields("SP56"))
         strUDate = Format(strTemp, "###/##/##")
      End If
   End If
   If IsNull(rsSrcTmp.Fields("SP57")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("SP57")) = False Then
         strTemp = rsSrcTmp.Fields("SP57")
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

' 顯示資料
Private Sub ShowCurrRecord(ByVal strSP01 As String, ByVal strSP02 As String, ByVal strSP03 As String, ByVal strSP04 As String)
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   
   If IsRecordExist(strSP01, strSP02, strSP03, strSP04) = True Then
      m_CurrSP(0) = strSP01
      m_CurrSP(1) = strSP02
      m_CurrSP(2) = strSP03
      m_CurrSP(3) = strSP04
   Else
      strSql = "SELECT SP01,SP02,SP03,SP04 FROM SERVICEPRACTICE " & _
               "WHERE SP01 = '" & strSP01 & "' AND " & _
                     "SP02 = '" & strSP02 & "' AND " & _
                     "SP03 = '" & strSP03 & "' AND " & _
                     "SP04 = (SELECT MIN(SP04) FROM SERVICEPRACTICE " & _
                             "WHERE SP01 = '" & strSP01 & "' AND " & _
                                   "SP02 = '" & strSP02 & "' AND " & _
                                   "SP03 = '" & strSP03 & "' AND " & _
                                   "SP04 > '" & strSP04 & "' )"
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         If IsNull(rsTmp.Fields("SP01")) = False Then: m_CurrSP(0) = rsTmp.Fields("SP01")
         If IsNull(rsTmp.Fields("SP02")) = False Then: m_CurrSP(1) = rsTmp.Fields("SP02")
         If IsNull(rsTmp.Fields("SP03")) = False Then: m_CurrSP(2) = rsTmp.Fields("SP03")
         If IsNull(rsTmp.Fields("SP04")) = False Then: m_CurrSP(3) = rsTmp.Fields("SP04")
         rsTmp.Close
         UpdateCtrlData
         GoTo EXITSUB
      End If
      rsTmp.Close
      
      strSql = "SELECT SP01,SP02,SP03,SP04 FROM SERVICEPRACTICE " & _
               "WHERE SP01 = '" & strSP01 & "' AND " & _
                     "SP02 = '" & strSP02 & "' AND " & _
                     "SP03 = (SELECT MIN(SP03) FROM SERVICEPRACTICE " & _
                             "WHERE SP01 = '" & strSP01 & "' AND " & _
                                   "SP02 = '" & strSP02 & "' AND " & _
                                   "SP03 > '" & strSP03 & "') AND " & _
                     "SP04 = (SELECT MIN(SP04) FROM SERVICEPRACTICE " & _
                             "WHERE SP01 = '" & strSP01 & "' AND " & _
                                   "SP02 = '" & strSP02 & "' AND " & _
                                   "SP03 = (SELECT MIN(SP03) FROM SERVICEPRACTICE " & _
                                           "WHERE SP01 = '" & strSP01 & "' AND " & _
                                                 "SP02 = '" & strSP02 & "' AND " & _
                                                 "SP03 > '" & strSP03 & "')) "
                                                       
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         If IsNull(rsTmp.Fields("SP01")) = False Then: m_CurrSP(0) = rsTmp.Fields("SP01")
         If IsNull(rsTmp.Fields("SP02")) = False Then: m_CurrSP(1) = rsTmp.Fields("SP02")
         If IsNull(rsTmp.Fields("SP03")) = False Then: m_CurrSP(2) = rsTmp.Fields("SP03")
         If IsNull(rsTmp.Fields("SP04")) = False Then: m_CurrSP(3) = rsTmp.Fields("SP04")
         rsTmp.Close
         UpdateCtrlData
         GoTo EXITSUB
      End If
      rsTmp.Close
                                   
      strSql = "SELECT SP01,SP02,SP03,SP04 FROM SERVICEPRACTICE " & _
               "WHERE SP01 = '" & strSP01 & "' AND " & _
                     "SP02 = (SELECT MIN(SP02) FROM SERVICEPRACTICE " & _
                             "WHERE SP01 = '" & strSP01 & "' AND " & _
                                   "SP02 > '" & strSP02 & "') AND " & _
                     "SP03 = (SELECT MIN(SP03) FROM SERVICEPRACTICE " & _
                             "WHERE SP01 = '" & strSP01 & "' AND " & _
                                   "SP02 = (SELECT MIN(SP02) FROM SERVICEPRACTICE " & _
                                           "WHERE SP01 = '" & strSP01 & "' AND " & _
                                                 "SP02 > '" & strSP02 & "')) AND " & _
                     "SP04 = (SELECT MIN(SP04) FROM SERVICEPRACTICE " & _
                             "WHERE SP01 = '" & strSP01 & "' AND " & _
                                   "SP02 = (SELECT MIN(SP02) FROM SERVICEPRACTICE " & _
                                           "WHERE SP01 = '" & strSP01 & "' AND " & _
                                                 "SP02 > '" & strSP02 & "') AND " & _
                                   "SP03 = (SELECT MIN(SP03) FROM SERVICEPRACTICE " & _
                                           "WHERE SP01 = '" & strSP01 & "' AND " & _
                                                 "SP02 = (SELECT MIN (SP02) FROM SERVICEPRACTICE " & _
                                                         "WHERE SP01 = '" & strSP01 & "' AND " & _
                                                               "SP02 > '" & strSP02 & "'))) "
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         If IsNull(rsTmp.Fields("SP01")) = False Then: m_CurrSP(0) = rsTmp.Fields("SP01")
         If IsNull(rsTmp.Fields("SP02")) = False Then: m_CurrSP(1) = rsTmp.Fields("SP02")
         If IsNull(rsTmp.Fields("SP03")) = False Then: m_CurrSP(2) = rsTmp.Fields("SP03")
         If IsNull(rsTmp.Fields("SP04")) = False Then: m_CurrSP(3) = rsTmp.Fields("SP04")
         rsTmp.Close
         UpdateCtrlData
         GoTo EXITSUB
      End If
      rsTmp.Close
      
      If m_SysKind = 0 Then
         'strSQL = "SELECT SP01,SP02,SP03,SP04 FROM SERVICEPRACTICE " & _
         '         "WHERE SP01 = (SELECT MIN(SP01) FROM SERVICEPRACTICE " & _
         '                       "WHERE SP01 > '" & strSP01 & "' AND " & "(SP01 <> 'TB' AND SP01 <> 'TM' AND SP01 <> 'TD' AND SP01 <> 'TC')) AND " & _
         '               "SP02 = (SELECT MIN(SP02) FROM SERVICEPRACTICE " & _
         '                       "WHERE SP01 = (SELECT MIN(SP01) FROM SERVICEPRACTICE " & _
         '                                     "WHERE SP01 > '" & strSP01 & "' AND " & "(SP01 <> 'TB' AND SP01 <> 'TM' AND SP01 <> 'TD' AND SP01 <> 'TC'))) AND " & _
         '               "SP03 = (SELECT MIN(SP03) FROM SERVICEPRACTICE " & _
         '                       "WHERE SP01 = (SELECT MIN(SP01) FROM SERVICEPRACTICE " & _
         '                                     "WHERE SP01 > '" & strSP01 & "' AND " & "(SP01 <> 'TB' AND SP01 <> 'TM' AND SP01 <> 'TD' AND SP01 <> 'TC')) AND " & _
         '                                           "SP02 = (SELECT MIN(SP02) FROM SERVICEPRACTICE " & _
         '                                                   "WHERE SP01 = (SELECT MIN(SP01) FROM SERVICEPRACTICE " & _
         '                                                                 "WHERE SP01 > '" & strSP01 & "' AND " & "(SP01 <> 'TB' AND SP01 <> 'TM' AND SP01 <> 'TD' AND SP01 <> 'TC')))) AND " & _
         '               "SP04 = (SELECT MIN(SP04) FROM SERVICEPRACTICE " & _
         '                       "WHERE SP01 = (SELECT MIN(SP01) FROM SERVICEPRACTICE " & _
         '                                     "WHERE SP01 > '" & strSP01 & "' AND " & "(SP01 <> 'TB' AND SP01 <> 'TM' AND SP01 <> 'TD' AND SP01 <> 'TC')) AND " & _
         '                                           "SP02 = (SELECT MIN(SP02) FROM SERVICEPRACTICE " & _
         '                                                   "WHERE SP01 = (SELECT MIN(SP01) FROM SERVICEPRACTICE " & _
         '                                                                 "WHERE SP01 > '" & strSP01 & "' AND " & "(SP01 <> 'TB' AND SP01 <> 'TM' AND SP01 <> 'TD' AND SP01 <> 'TC'))) AND " & _
         '                                                                       "SP03 = (SELECT MIN(SP03) FROM SERVICEPRACTICE " & _
         '                                                                               "WHERE SP01 = (SELECT MIN(SP01) FROM SERVICEPRACTICE " & _
         '                                                                                             "WHERE SP01 > '" & strSP01 & "' AND " & "(SP01 <> 'TB' AND SP01 <> 'TM' AND SP01 <> 'TD' AND SP01 <> 'TC')) AND " & _
         '                                                                                                   "SP02 = (SELECT MIN(SP02) FROM SERVICEPRACTICE " & _
         '                                                                                                           "WHERE SP01 = (SELECT MIN(SP01) FROM SERVICEPRACTICE " & _
         '                                                                                                                         "WHERE SP01 > '" & strSP01 & "' AND " & "(SP01 <> 'TB' AND SP01 <> 'TM' AND SP01 <> 'TD' AND SP01 <> 'TC'))))) "
         '
         strSql = "SELECT SP01,SP02,SP03,SP04 FROM SERVICEPRACTICE " & _
                  "WHERE SP01 = (SELECT MIN(SP01) FROM SERVICEPRACTICE " & _
                                "WHERE SP01 > '" & strSP01 & "' AND " & "SP01 IN " & m_QuerySystem & ") AND " & _
                        "SP02 = (SELECT MIN(SP02) FROM SERVICEPRACTICE " & _
                                "WHERE SP01 = (SELECT MIN(SP01) FROM SERVICEPRACTICE " & _
                                              "WHERE SP01 > '" & strSP01 & "' AND " & "SP01 IN " & m_QuerySystem & ")) AND " & _
                        "SP03 = (SELECT MIN(SP03) FROM SERVICEPRACTICE " & _
                                "WHERE SP01 = (SELECT MIN(SP01) FROM SERVICEPRACTICE " & _
                                              "WHERE SP01 > '" & strSP01 & "' AND " & "SP01 IN " & m_QuerySystem & ") AND " & _
                                                    "SP02 = (SELECT MIN(SP02) FROM SERVICEPRACTICE " & _
                                                            "WHERE SP01 = (SELECT MIN(SP01) FROM SERVICEPRACTICE " & _
                                                                          "WHERE SP01 > '" & strSP01 & "' AND " & "SP01 IN " & m_QuerySystem & "))) AND " & _
                        "SP04 = (SELECT MIN(SP04) FROM SERVICEPRACTICE " & _
                                "WHERE SP01 = (SELECT MIN(SP01) FROM SERVICEPRACTICE " & _
                                              "WHERE SP01 > '" & strSP01 & "' AND " & "SP01 IN " & m_QuerySystem & ") AND " & _
                                                    "SP02 = (SELECT MIN(SP02) FROM SERVICEPRACTICE " & _
                                                            "WHERE SP01 = (SELECT MIN(SP01) FROM SERVICEPRACTICE " & _
                                                                          "WHERE SP01 > '" & strSP01 & "' AND " & "SP01 IN " & m_QuerySystem & ")) AND " & _
                                                                                "SP03 = (SELECT MIN(SP03) FROM SERVICEPRACTICE " & _
                                                                                        "WHERE SP01 = (SELECT MIN(SP01) FROM SERVICEPRACTICE " & _
                                                                                                      "WHERE SP01 > '" & strSP01 & "' AND " & "SP01 IN " & m_QuerySystem & ") AND " & _
                                                                                                            "SP02 = (SELECT MIN(SP02) FROM SERVICEPRACTICE " & _
                                                                                                                    "WHERE SP01 = (SELECT MIN(SP01) FROM SERVICEPRACTICE " & _
                                                                                                                                  "WHERE SP01 > '" & strSP01 & "' AND " & "SP01 IN " & m_QuerySystem & ")))) "
         rsTmp.CursorLocation = adUseClient
         rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
         If rsTmp.RecordCount > 0 Then
            If IsNull(rsTmp.Fields("SP01")) = False Then: m_CurrSP(0) = rsTmp.Fields("SP01")
            If IsNull(rsTmp.Fields("SP02")) = False Then: m_CurrSP(1) = rsTmp.Fields("SP02")
            If IsNull(rsTmp.Fields("SP03")) = False Then: m_CurrSP(2) = rsTmp.Fields("SP03")
            If IsNull(rsTmp.Fields("SP04")) = False Then: m_CurrSP(3) = rsTmp.Fields("SP04")
            rsTmp.Close
            UpdateCtrlData
            GoTo EXITSUB
         Else
            rsTmp.Close
            ShowLastRecord
            GoTo EXITSUB
         End If
      End If
      rsTmp.Close
   End If
   UpdateCtrlData
EXITSUB:
End Sub

' 顯示第一筆資料
Private Sub ShowFirstRecord()
   m_CurrSP(0) = m_FirstSP(0)
   m_CurrSP(1) = m_FirstSP(1)
   m_CurrSP(2) = m_FirstSP(2)
   m_CurrSP(3) = m_FirstSP(3)
   
   UpdateCtrlData
End Sub

' 顯示上一筆資料
Private Sub ShowPrevRecord()
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   
   If m_CurrSP(0) = m_FirstSP(0) And m_CurrSP(1) = m_FirstSP(1) And m_CurrSP(2) = m_FirstSP(2) And m_CurrSP(3) = m_FirstSP(3) Then
      ShowMsg MsgText(9008)
      GoTo EXITSUB
   End If
  
   
   'Select Case m_SysKind
   '   Case 0:
   '      strSQL = "SELECT SP01,SP02,SP03,SP04 FROM SERVICEPRACTICE " & _
   '               "WHERE SP01 = '" & m_CurrSP(0) & "' AND " & _
   '                     "SP02 = '" & m_CurrSP(1) & "' AND " & _
   '                     "SP03 = '" & m_CurrSP(2) & "' AND " & _
   '                     "SP04 = (SELECT MAX(SP04) FROM SERVICEPRACTICE " & _
   '                             "WHERE SP01 = '" & m_CurrSP(0) & "' AND " & _
   '                                   "SP02 = '" & m_CurrSP(1) & "' AND " & _
   '                                   "SP03 = '" & m_CurrSP(2) & "' AND " & _
   '                                   "SP04 < '" & m_CurrSP(3) & "' )"
   '   Case 1:
   '      strSQL = "SELECT SP01,SP02,SP03,SP04 FROM SERVICEPRACTICE " & _
   '               "WHERE SP01 = '" & m_CurrSP(0) & "' AND " & _
   '                     "SP02 = '" & m_CurrSP(1) & "' AND " & _
   '                     "SP03 = '" & m_CurrSP(2) & "' AND " & _
   '                     "SP04 = (SELECT MAX(SP04) FROM SERVICEPRACTICE " & _
   '                             "WHERE SP01 = '" & m_CurrSP(0) & "' AND " & _
   '                                   "SP02 = '" & m_CurrSP(1) & "' AND " & _
   '                                   "SP03 = '" & m_CurrSP(2) & "' AND " & _
   '                                   "SP04 < '" & m_CurrSP(3) & "' )"
   'End Select
   strSql = "SELECT SP01,SP02,SP03,SP04 FROM SERVICEPRACTICE " & _
            "WHERE SP01 = '" & m_CurrSP(0) & "' AND " & _
                  "SP02 = '" & m_CurrSP(1) & "' AND " & _
                  "SP03 = '" & m_CurrSP(2) & "' AND " & _
                  "SP04 = (SELECT MAX(SP04) FROM SERVICEPRACTICE " & _
                          "WHERE SP01 = '" & m_CurrSP(0) & "' AND " & _
                                "SP02 = '" & m_CurrSP(1) & "' AND " & _
                                "SP03 = '" & m_CurrSP(2) & "' AND " & _
                                "SP04 < '" & m_CurrSP(3) & "' )"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("SP01")) = False Then: m_CurrSP(0) = rsTmp.Fields("SP01")
      If IsNull(rsTmp.Fields("SP02")) = False Then: m_CurrSP(1) = rsTmp.Fields("SP02")
      If IsNull(rsTmp.Fields("SP03")) = False Then: m_CurrSP(2) = rsTmp.Fields("SP03")
      If IsNull(rsTmp.Fields("SP04")) = False Then: m_CurrSP(3) = rsTmp.Fields("SP04")
      rsTmp.Close
      UpdateCtrlData
      GoTo EXITSUB
   End If
   rsTmp.Close
   
   'Select Case m_SysKind
   '   Case 0:
   '      strSQL = "SELECT SP01,SP02,SP03,SP04 FROM SERVICEPRACTICE " & _
   '               "WHERE SP01 = '" & m_CurrSP(0) & "' AND " & _
   '                     "SP02 = '" & m_CurrSP(1) & "' AND " & _
   '                     "SP03 = (SELECT MAX(SP03) FROM SERVICEPRACTICE " & _
   '                             "WHERE SP01 = '" & m_CurrSP(0) & "' AND " & _
   '                                   "SP02 = '" & m_CurrSP(1) & "' AND " & _
   '                                   "SP03 < '" & m_CurrSP(2) & "') AND " & _
   '                     "SP04 = (SELECT MAX(SP04) FROM SERVICEPRACTICE " & _
   '                             "WHERE SP01 = '" & m_CurrSP(0) & "' AND " & _
   '                                   "SP02 = '" & m_CurrSP(1) & "' AND " & _
   '                                   "SP03 = (SELECT MAX(SP03) FROM SERVICEPRACTICE " & _
   '                                           "WHERE SP01 = '" & m_CurrSP(0) & "' AND " & _
   '                                                 "SP02 = '" & m_CurrSP(1) & "' AND " & _
   '                                                 "SP03 < '" & m_CurrSP(2) & "')) "
   '   Case 1:
   '      strSQL = "SELECT SP01,SP02,SP03,SP04 FROM SERVICEPRACTICE " & _
   '               "WHERE SP01 = '" & m_CurrSP(0) & "' AND " & _
   '                     "SP02 = '" & m_CurrSP(1) & "' AND " & _
   '                     "SP03 = (SELECT MAX(SP03) FROM SERVICEPRACTICE " & _
   '                             "WHERE SP01 = '" & m_CurrSP(0) & "' AND " & _
   '                                   "SP02 = '" & m_CurrSP(1) & "' AND " & _
   '                                   "SP03 < '" & m_CurrSP(2) & "') AND " & _
   '                     "SP04 = (SELECT MAX(SP04) FROM SERVICEPRACTICE " & _
   '                             "WHERE SP01 = '" & m_CurrSP(0) & "' AND " & _
   '                                   "SP02 = '" & m_CurrSP(1) & "' AND " & _
   '                                   "SP03 = (SELECT MAX(SP03) FROM SERVICEPRACTICE " & _
   '                                           "WHERE SP01 = '" & m_CurrSP(0) & "' AND " & _
   '                                                 "SP02 = '" & m_CurrSP(1) & "' AND " & _
   '                                                 "SP03 < '" & m_CurrSP(2) & "'))"
   'End Select
   strSql = "SELECT SP01,SP02,SP03,SP04 FROM SERVICEPRACTICE " & _
            "WHERE SP01 = '" & m_CurrSP(0) & "' AND " & _
                  "SP02 = '" & m_CurrSP(1) & "' AND " & _
                  "SP03 = (SELECT MAX(SP03) FROM SERVICEPRACTICE " & _
                          "WHERE SP01 = '" & m_CurrSP(0) & "' AND " & _
                                "SP02 = '" & m_CurrSP(1) & "' AND " & _
                                "SP03 < '" & m_CurrSP(2) & "') AND " & _
                  "SP04 = (SELECT MAX(SP04) FROM SERVICEPRACTICE " & _
                          "WHERE SP01 = '" & m_CurrSP(0) & "' AND " & _
                                "SP02 = '" & m_CurrSP(1) & "' AND " & _
                                "SP03 = (SELECT MAX(SP03) FROM SERVICEPRACTICE " & _
                                        "WHERE SP01 = '" & m_CurrSP(0) & "' AND " & _
                                              "SP02 = '" & m_CurrSP(1) & "' AND " & _
                                              "SP03 < '" & m_CurrSP(2) & "')) "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("SP01")) = False Then: m_CurrSP(0) = rsTmp.Fields("SP01")
      If IsNull(rsTmp.Fields("SP02")) = False Then: m_CurrSP(1) = rsTmp.Fields("SP02")
      If IsNull(rsTmp.Fields("SP03")) = False Then: m_CurrSP(2) = rsTmp.Fields("SP03")
      If IsNull(rsTmp.Fields("SP04")) = False Then: m_CurrSP(3) = rsTmp.Fields("SP04")
      rsTmp.Close
      UpdateCtrlData
      GoTo EXITSUB
   End If
   rsTmp.Close

   'Select Case m_SysKind
   '   Case 0:
   '      strSQL = "SELECT SP01,SP02,SP03,SP04 FROM SERVICEPRACTICE " & _
   '               "WHERE SP01 = '" & m_CurrSP(0) & "' AND " & _
   '                     "SP02 = (SELECT MAX(SP02) FROM SERVICEPRACTICE " & _
   '                             "WHERE SP01 = '" & m_CurrSP(0) & "' AND " & _
   '                                   "SP02 < '" & m_CurrSP(1) & "') AND " & _
   '                     "SP03 = (SELECT MAX(SP03) FROM SERVICEPRACTICE " & _
   '                             "WHERE SP01 = '" & m_CurrSP(0) & "' AND " & _
   '                                   "SP02 = (SELECT MAX(SP02) FROM SERVICEPRACTICE " & _
   '                                           "WHERE SP01 = '" & m_CurrSP(0) & "' AND " & _
   '                                                 "SP02 < '" & m_CurrSP(1) & "')) AND " & _
   '                     "SP04 = (SELECT MAX(SP04) FROM SERVICEPRACTICE " & _
   '                             "WHERE SP01 = '" & m_CurrSP(0) & "' AND " & _
   '                                   "SP02 = (SELECT MAX(SP02) FROM SERVICEPRACTICE " & _
   '                                           "WHERE SP01 = '" & m_CurrSP(0) & "' AND " & _
   '                                                 "SP02 < '" & m_CurrSP(1) & "') AND " & _
   '                                   "SP03 = (SELECT MAX(SP03) FROM SERVICEPRACTICE " & _
   '                                           "WHERE SP01 = '" & m_CurrSP(0) & "' AND " & _
   '                                                 "SP02 = (SELECT MAX(SP02) FROM SERVICEPRACTICE " & _
   '                                                         "WHERE SP01 = '" & m_CurrSP(0) & "' AND " & _
   '                                                               "SP02 < '" & m_CurrSP(1) & "'))) "
   '   Case 1:
   '      strSQL = "SELECT SP01,SP02,SP03,SP04 FROM SERVICEPRACTICE " & _
   '               "WHERE SP01 = '" & m_CurrSP(0) & "' AND " & _
   '                     "SP02 = (SELECT MAX(SP02) FROM SERVICEPRACTICE " & _
   '                             "WHERE SP01 = '" & m_CurrSP(0) & "' AND " & _
   '                                   "SP02 < '" & m_CurrSP(1) & "') AND " & _
   '                     "SP03 = (SELECT MAX(SP03) FROM SERVICEPRACTICE " & _
   '                             "WHERE SP01 = '" & m_CurrSP(0) & "' AND " & _
   '                                   "SP02 = (SELECT MAX(SP02) FROM SERVICEPRACTICE " & _
   '                                           "WHERE SP01 = '" & m_CurrSP(0) & "' AND " & _
   '                                                 "SP02 < '" & m_CurrSP(1) & "')) AND " & _
   '                     "SP04 = (SELECT MAX(SP04) FROM SERVICEPRACTICE " & _
   '                             "WHERE SP01 = '" & m_CurrSP(0) & "' AND " & _
   '                                   "SP02 = (SELECT MAX(SP02) FROM SERVICEPRACTICE " & _
   '                                           "WHERE SP01 = '" & m_CurrSP(0) & "' AND " & _
   '                                                 "SP02 < '" & m_CurrSP(1) & "') AND " & _
   '                                                 "SP03 = (SELECT MAX(SP03) FROM SERVICEPRACTICE " & _
   '                                                         "WHERE SP01 = '" & m_CurrSP(0) & "' AND " & _
   '                                                               "SP02 = (SELECT MAX(SP02) FROM SERVICEPRACTICE " & _
   '                                                                       "WHERE SP01 = '" & m_CurrSP(0) & "' AND " & _
   '                                                                             "SP02 < '" & m_CurrSP(1) & "'))) "
   'End Select
   strSql = "SELECT SP01,SP02,SP03,SP04 FROM SERVICEPRACTICE " & _
            "WHERE SP01 = '" & m_CurrSP(0) & "' AND " & _
                  "SP02 = (SELECT MAX(SP02) FROM SERVICEPRACTICE " & _
                          "WHERE SP01 = '" & m_CurrSP(0) & "' AND " & _
                                "SP02 < '" & m_CurrSP(1) & "') AND " & _
                  "SP03 = (SELECT MAX(SP03) FROM SERVICEPRACTICE " & _
                          "WHERE SP01 = '" & m_CurrSP(0) & "' AND " & _
                                "SP02 = (SELECT MAX(SP02) FROM SERVICEPRACTICE " & _
                                        "WHERE SP01 = '" & m_CurrSP(0) & "' AND " & _
                                              "SP02 < '" & m_CurrSP(1) & "')) AND " & _
                  "SP04 = (SELECT MAX(SP04) FROM SERVICEPRACTICE " & _
                          "WHERE SP01 = '" & m_CurrSP(0) & "' AND " & _
                                "SP02 = (SELECT MAX(SP02) FROM SERVICEPRACTICE " & _
                                        "WHERE SP01 = '" & m_CurrSP(0) & "' AND " & _
                                              "SP02 < '" & m_CurrSP(1) & "') AND " & _
                                "SP03 = (SELECT MAX(SP03) FROM SERVICEPRACTICE " & _
                                        "WHERE SP01 = '" & m_CurrSP(0) & "' AND " & _
                                              "SP02 = (SELECT MAX(SP02) FROM SERVICEPRACTICE " & _
                                                      "WHERE SP01 = '" & m_CurrSP(0) & "' AND " & _
                                                            "SP02 < '" & m_CurrSP(1) & "'))) "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("SP01")) = False Then: m_CurrSP(0) = rsTmp.Fields("SP01")
      If IsNull(rsTmp.Fields("SP02")) = False Then: m_CurrSP(1) = rsTmp.Fields("SP02")
      If IsNull(rsTmp.Fields("SP03")) = False Then: m_CurrSP(2) = rsTmp.Fields("SP03")
      If IsNull(rsTmp.Fields("SP04")) = False Then: m_CurrSP(3) = rsTmp.Fields("SP04")
      rsTmp.Close
      UpdateCtrlData
      GoTo EXITSUB
   End If
   rsTmp.Close
   
   'If m_SysKind = 0 Then
      strSql = "SELECT SP01,SP02,SP03,SP04 FROM SERVICEPRACTICE " & _
               "WHERE SP01 = (SELECT MAX(SP01) FROM SERVICEPRACTICE " & _
                             "WHERE SP01 < '" & m_CurrSP(0) & "' AND " & "SP01 IN " & m_QuerySystem & ") AND " & _
                     "SP02 = (SELECT MAX(SP02) FROM SERVICEPRACTICE " & _
                             "WHERE SP01 = (SELECT MAX(SP01) FROM SERVICEPRACTICE " & _
                                           "WHERE SP01 < '" & m_CurrSP(0) & "' AND " & "SP01 IN " & m_QuerySystem & ")) AND " & _
                     "SP03 = (SELECT MAX(SP03) FROM SERVICEPRACTICE " & _
                             "WHERE SP01 = (SELECT MAX(SP01) FROM SERVICEPRACTICE " & _
                                           "WHERE SP01 < '" & m_CurrSP(0) & "' AND " & "SP01 IN " & m_QuerySystem & ") AND " & _
                                                 "SP02 = (SELECT MAX(SP02) FROM SERVICEPRACTICE " & _
                                                         "WHERE SP01 = (SELECT MAX(SP01) FROM SERVICEPRACTICE " & _
                                                                       "WHERE SP01 < '" & m_CurrSP(0) & "' AND " & "SP01 IN " & m_QuerySystem & "))) AND " & _
                     "SP04 = (SELECT MAX(SP04) FROM SERVICEPRACTICE " & _
                             "WHERE SP01 = (SELECT MAX(SP01) FROM SERVICEPRACTICE " & _
                                           "WHERE SP01 < '" & m_CurrSP(0) & "' AND " & "SP01 IN " & m_QuerySystem & ") AND " & _
                                                 "SP02 = (SELECT MAX(SP02) FROM SERVICEPRACTICE " & _
                                                         "WHERE SP01 = (SELECT MAX(SP01) FROM SERVICEPRACTICE " & _
                                                                       "WHERE SP01 < '" & m_CurrSP(0) & "' AND " & "SP01 IN " & m_QuerySystem & ")) AND " & _
                                                                             "SP03 = (SELECT MAX(SP03) FROM SERVICEPRACTICE " & _
                                                                                     "WHERE SP01 = (SELECT MAX(SP01) FROM SERVICEPRACTICE " & _
                                                                                                   "WHERE SP01 < '" & m_CurrSP(0) & "' AND " & "SP01 IN " & m_QuerySystem & ") AND " & _
                                                                                                         "SP02 = (SELECT MAX(SP02) FROM SERVICEPRACTICE " & _
                                                                                                                 "WHERE SP01 = (SELECT MAX(SP01) FROM SERVICEPRACTICE " & _
                                                                                                                               "WHERE SP01 < '" & m_CurrSP(0) & "' AND " & "SP01 IN " & m_QuerySystem & ")))) "
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         If IsNull(rsTmp.Fields("SP01")) = False Then: m_CurrSP(0) = rsTmp.Fields("SP01")
         If IsNull(rsTmp.Fields("SP02")) = False Then: m_CurrSP(1) = rsTmp.Fields("SP02")
         If IsNull(rsTmp.Fields("SP03")) = False Then: m_CurrSP(2) = rsTmp.Fields("SP03")
         If IsNull(rsTmp.Fields("SP04")) = False Then: m_CurrSP(3) = rsTmp.Fields("SP04")
         rsTmp.Close
         UpdateCtrlData
         GoTo EXITSUB
      End If
      rsTmp.Close
      UpdateCtrlData
   'End If
   'rsTmp.Close
   
   
EXITSUB:
   Set rsTmp = Nothing
End Sub

' 顯示下一筆資料
Private Sub ShowNextRecord()
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   
   If m_CurrSP(0) = m_LastSP(0) And m_CurrSP(1) = m_LastSP(1) And m_CurrSP(2) = m_LastSP(2) And m_CurrSP(3) = m_LastSP(3) Then
      ShowMsg MsgText(9009)
      GoTo EXITSUB
   End If
   
   'Select Case m_SysKind
   '   Case 0:
   '      strSQL = "SELECT SP01,SP02,SP03,SP04 FROM SERVICEPRACTICE " & _
   '               "WHERE SP01 = '" & m_CurrSP(0) & "' AND " & _
   '                     "SP02 = '" & m_CurrSP(1) & "' AND " & _
   '                     "SP03 = '" & m_CurrSP(2) & "' AND " & _
   '                     "SP04 = (SELECT MIN(SP04) FROM SERVICEPRACTICE " & _
   '                             "WHERE SP01 = '" & m_CurrSP(0) & "' AND " & _
   '                                   "SP02 = '" & m_CurrSP(1) & "' AND " & _
   '                                   "SP03 = '" & m_CurrSP(2) & "' AND " & _
   '                                   "SP04 > '" & m_CurrSP(3) & "' )"
   '   Case 1:
   '      strSQL = "SELECT SP01,SP02,SP03,SP04 FROM SERVICEPRACTICE " & _
   '               "WHERE SP01 = '" & m_CurrSP(0) & "' AND " & _
   '                     "SP02 = '" & m_CurrSP(1) & "' AND " & _
   '                     "SP03 = '" & m_CurrSP(2) & "' AND " & _
   '                     "SP04 = (SELECT MIN(SP04) FROM SERVICEPRACTICE " & _
   '                             "WHERE SP01 = '" & m_CurrSP(0) & "' AND " & _
   '                                   "SP02 = '" & m_CurrSP(1) & "' AND " & _
   '                                   "SP03 = '" & m_CurrSP(2) & "' AND " & _
   '                                   "SP04 > '" & m_CurrSP(3) & "' )"
   'End Select
   strSql = "SELECT SP01,SP02,SP03,SP04 FROM SERVICEPRACTICE " & _
            "WHERE SP01 = '" & m_CurrSP(0) & "' AND " & _
                  "SP02 = '" & m_CurrSP(1) & "' AND " & _
                  "SP03 = '" & m_CurrSP(2) & "' AND " & _
                  "SP04 = (SELECT MIN(SP04) FROM SERVICEPRACTICE " & _
                          "WHERE SP01 = '" & m_CurrSP(0) & "' AND " & _
                                "SP02 = '" & m_CurrSP(1) & "' AND " & _
                                "SP03 = '" & m_CurrSP(2) & "' AND " & _
                                "SP04 > '" & m_CurrSP(3) & "' )"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("SP01")) = False Then: m_CurrSP(0) = rsTmp.Fields("SP01")
      If IsNull(rsTmp.Fields("SP02")) = False Then: m_CurrSP(1) = rsTmp.Fields("SP02")
      If IsNull(rsTmp.Fields("SP03")) = False Then: m_CurrSP(2) = rsTmp.Fields("SP03")
      If IsNull(rsTmp.Fields("SP04")) = False Then: m_CurrSP(3) = rsTmp.Fields("SP04")
      rsTmp.Close
      UpdateCtrlData
      GoTo EXITSUB
   End If
   rsTmp.Close
   
   'Select Case m_SysKind
   '   Case 0:
   '      strSQL = "SELECT SP01,SP02,SP03,SP04 FROM SERVICEPRACTICE " & _
   '               "WHERE SP01 = '" & m_CurrSP(0) & "' AND " & _
   '                     "SP02 = '" & m_CurrSP(1) & "' AND " & _
   '                     "SP03 = (SELECT MIN(SP03) FROM SERVICEPRACTICE " & _
   '                             "WHERE SP01 = '" & m_CurrSP(0) & "' AND " & _
   '                                   "SP02 = '" & m_CurrSP(1) & "' AND " & _
   '                                   "SP03 > '" & m_CurrSP(2) & "') AND " & _
   '                     "SP04 = (SELECT MIN(SP04) FROM SERVICEPRACTICE " & _
   '                             "WHERE SP01 = '" & m_CurrSP(0) & "' AND " & _
   '                                   "SP02 = '" & m_CurrSP(1) & "' AND " & _
   '                                   "SP03 = (SELECT MIN(SP03) FROM SERVICEPRACTICE " & _
   '                                           "WHERE SP01 = '" & m_CurrSP(0) & "' AND " & _
   '                                                 "SP02 = '" & m_CurrSP(1) & "' AND " & _
   '                                                 "SP03 > '" & m_CurrSP(2) & "')) "
   '   Case 1:
   '      strSQL = "SELECT SP01,SP02,SP03,SP04 FROM SERVICEPRACTICE " & _
   '               "WHERE SP01 = '" & m_CurrSP(0) & "' AND " & _
   '                     "SP02 = '" & m_CurrSP(1) & "' AND " & _
   '                     "SP03 = (SELECT MIN(SP03) FROM SERVICEPRACTICE " & _
   '                             "WHERE SP01 = '" & m_CurrSP(0) & "' AND " & _
   '                                   "SP02 = '" & m_CurrSP(1) & "' AND " & _
   '                                   "SP03 > '" & m_CurrSP(2) & "') AND " & _
   '                     "SP04 = (SELECT MIN(SP04) FROM SERVICEPRACTICE " & _
   '                             "WHERE SP01 = '" & m_CurrSP(0) & "' AND " & _
   '                                   "SP02 = '" & m_CurrSP(1) & "' AND " & _
   '                                   "SP03 = (SELECT MIN(SP03) FROM SERVICEPRACTICE " & _
   '                                           "WHERE SP01 = '" & m_CurrSP(0) & "' AND " & _
   '                                                 "SP02 = '" & m_CurrSP(1) & "' AND " & _
   '                                                 "SP03 > '" & m_CurrSP(2) & "'))"
   'End Select
   strSql = "SELECT SP01,SP02,SP03,SP04 FROM SERVICEPRACTICE " & _
            "WHERE SP01 = '" & m_CurrSP(0) & "' AND " & _
                  "SP02 = '" & m_CurrSP(1) & "' AND " & _
                  "SP03 = (SELECT MIN(SP03) FROM SERVICEPRACTICE " & _
                          "WHERE SP01 = '" & m_CurrSP(0) & "' AND " & _
                                "SP02 = '" & m_CurrSP(1) & "' AND " & _
                                "SP03 > '" & m_CurrSP(2) & "') AND " & _
                  "SP04 = (SELECT MIN(SP04) FROM SERVICEPRACTICE " & _
                          "WHERE SP01 = '" & m_CurrSP(0) & "' AND " & _
                                "SP02 = '" & m_CurrSP(1) & "' AND " & _
                                "SP03 = (SELECT MIN(SP03) FROM SERVICEPRACTICE " & _
                                        "WHERE SP01 = '" & m_CurrSP(0) & "' AND " & _
                                              "SP02 = '" & m_CurrSP(1) & "' AND " & _
                                              "SP03 > '" & m_CurrSP(2) & "')) "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("SP01")) = False Then: m_CurrSP(0) = rsTmp.Fields("SP01")
      If IsNull(rsTmp.Fields("SP02")) = False Then: m_CurrSP(1) = rsTmp.Fields("SP02")
      If IsNull(rsTmp.Fields("SP03")) = False Then: m_CurrSP(2) = rsTmp.Fields("SP03")
      If IsNull(rsTmp.Fields("SP04")) = False Then: m_CurrSP(3) = rsTmp.Fields("SP04")
      rsTmp.Close
      UpdateCtrlData
      GoTo EXITSUB
   End If
   rsTmp.Close
                                
   'Select Case m_SysKind
   '   Case 0:
   '      strSQL = "SELECT SP01,SP02,SP03,SP04 FROM SERVICEPRACTICE " & _
   '               "WHERE SP01 = '" & m_CurrSP(0) & "' AND " & _
   '                     "SP02 = (SELECT MIN(SP02) FROM SERVICEPRACTICE " & _
   '                             "WHERE SP01 = '" & m_CurrSP(0) & "' AND " & _
   '                                   "SP02 > '" & m_CurrSP(1) & "') AND " & _
   '                     "SP03 = (SELECT MIN(SP03) FROM SERVICEPRACTICE " & _
   '                             "WHERE SP01 = '" & m_CurrSP(0) & "' AND " & _
   '                                   "SP02 = (SELECT MIN(SP02) FROM SERVICEPRACTICE " & _
   '                                           "WHERE SP01 = '" & m_CurrSP(0) & "' AND " & _
   '                                                 "SP02 > '" & m_CurrSP(1) & "')) AND " & _
   '                     "SP04 = (SELECT MIN(SP04) FROM SERVICEPRACTICE " & _
   '                             "WHERE SP01 = '" & m_CurrSP(0) & "' AND " & _
   '                                   "SP02 = (SELECT MIN(SP02) FROM SERVICEPRACTICE " & _
   '                                           "WHERE SP01 = '" & m_CurrSP(0) & "' AND " & _
   '                                                 "SP02 > '" & m_CurrSP(1) & "') AND " & _
   '                                   "SP03 = (SELECT MIN(SP03) FROM SERVICEPRACTICE " & _
   '                                           "WHERE SP01 = '" & m_CurrSP(0) & "' AND " & _
   '                                                 "SP02 = (SELECT MIN (SP02) FROM SERVICEPRACTICE " & _
   '                                                         "WHERE SP01 = '" & m_CurrSP(0) & "' AND " & _
   '                                                              "SP02 > '" & m_CurrSP(1) & "'))) "
   '   Case 1:
   '      strSQL = "SELECT SP01,SP02,SP03,SP04 FROM SERVICEPRACTICE " & _
   '               "WHERE SP01 = '" & m_CurrSP(0) & "' AND " & _
   '                     "SP02 = (SELECT MIN(SP02) FROM SERVICEPRACTICE " & _
   '                             "WHERE SP01 = '" & m_CurrSP(0) & "' AND " & _
   '                                   "SP02 > '" & m_CurrSP(1) & "') AND " & _
   '                     "SP03 = (SELECT MIN(SP03) FROM SERVICEPRACTICE " & _
   '                             "WHERE SP01 = '" & m_CurrSP(0) & "' AND " & _
   '                                   "SP02 = (SELECT MIN(SP02) FROM SERVICEPRACTICE " & _
   '                                           "WHERE SP01 = '" & m_CurrSP(0) & "' AND " & _
   '                                                 "SP02 > '" & m_CurrSP(1) & "')) AND " & _
   '                     "SP04 = (SELECT MIN(SP04) FROM SERVICEPRACTICE " & _
   '                             "WHERE SP01 = '" & m_CurrSP(0) & "' AND " & _
   '                                   "SP02 = (SELECT MIN(SP02) FROM SERVICEPRACTICE " & _
   '                                           "WHERE SP01 = '" & m_CurrSP(0) & "' AND " & _
   '                                                 "SP02 > '" & m_CurrSP(1) & "') AND " & _
   '                                                 "SP03 = (SELECT MIN(SP03) FROM SERVICEPRACTICE " & _
   '                                                         "WHERE SP01 = '" & m_CurrSP(0) & "' AND " & _
   '                                                               "SP02 = (SELECT MIN(SP02) FROM SERVICEPRACTICE " & _
   '                                                                       "WHERE SP01 = '" & m_CurrSP(0) & "' AND " & _
   '                                                                             "SP02 > '" & m_CurrSP(1) & "'))) "
   'End Select
   strSql = "SELECT SP01,SP02,SP03,SP04 FROM SERVICEPRACTICE " & _
            "WHERE SP01 = '" & m_CurrSP(0) & "' AND " & _
                  "SP02 = (SELECT MIN(SP02) FROM SERVICEPRACTICE " & _
                          "WHERE SP01 = '" & m_CurrSP(0) & "' AND " & _
                                "SP02 > '" & m_CurrSP(1) & "') AND " & _
                  "SP03 = (SELECT MIN(SP03) FROM SERVICEPRACTICE " & _
                          "WHERE SP01 = '" & m_CurrSP(0) & "' AND " & _
                                "SP02 = (SELECT MIN(SP02) FROM SERVICEPRACTICE " & _
                                        "WHERE SP01 = '" & m_CurrSP(0) & "' AND " & _
                                              "SP02 > '" & m_CurrSP(1) & "')) AND " & _
                  "SP04 = (SELECT MIN(SP04) FROM SERVICEPRACTICE " & _
                          "WHERE SP01 = '" & m_CurrSP(0) & "' AND " & _
                                "SP02 = (SELECT MIN(SP02) FROM SERVICEPRACTICE " & _
                                        "WHERE SP01 = '" & m_CurrSP(0) & "' AND " & _
                                              "SP02 > '" & m_CurrSP(1) & "') AND " & _
                                "SP03 = (SELECT MIN(SP03) FROM SERVICEPRACTICE " & _
                                        "WHERE SP01 = '" & m_CurrSP(0) & "' AND " & _
                                              "SP02 = (SELECT MIN (SP02) FROM SERVICEPRACTICE " & _
                                                      "WHERE SP01 = '" & m_CurrSP(0) & "' AND " & _
                                                            "SP02 > '" & m_CurrSP(1) & "'))) "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("SP01")) = False Then: m_CurrSP(0) = rsTmp.Fields("SP01")
      If IsNull(rsTmp.Fields("SP02")) = False Then: m_CurrSP(1) = rsTmp.Fields("SP02")
      If IsNull(rsTmp.Fields("SP03")) = False Then: m_CurrSP(2) = rsTmp.Fields("SP03")
      If IsNull(rsTmp.Fields("SP04")) = False Then: m_CurrSP(3) = rsTmp.Fields("SP04")
      rsTmp.Close
      UpdateCtrlData
      GoTo EXITSUB
   End If
   rsTmp.Close
   
   'If m_SysKind = 0 Then
      strSql = "SELECT SP01,SP02,SP03,SP04 FROM SERVICEPRACTICE " & _
               "WHERE SP01 = (SELECT MIN(SP01) FROM SERVICEPRACTICE " & _
                             "WHERE SP01 > '" & m_CurrSP(0) & "' AND " & "SP01 IN " & m_QuerySystem & ") AND " & _
                     "SP02 = (SELECT MIN(SP02) FROM SERVICEPRACTICE " & _
                             "WHERE SP01 = (SELECT MIN(SP01) FROM SERVICEPRACTICE " & _
                                           "WHERE SP01 > '" & m_CurrSP(0) & "' AND " & "SP01 IN " & m_QuerySystem & ")) AND " & _
                     "SP03 = (SELECT MIN(SP03) FROM SERVICEPRACTICE " & _
                             "WHERE SP01 = (SELECT MIN(SP01) FROM SERVICEPRACTICE " & _
                                           "WHERE SP01 > '" & m_CurrSP(0) & "' AND " & "SP01 IN " & m_QuerySystem & ") AND " & _
                                                 "SP02 = (SELECT MIN(SP02) FROM SERVICEPRACTICE " & _
                                                         "WHERE SP01 = (SELECT MIN(SP01) FROM SERVICEPRACTICE " & _
                                                                       "WHERE SP01 > '" & m_CurrSP(0) & "' AND " & "SP01 IN " & m_QuerySystem & "))) AND " & _
                     "SP04 = (SELECT MIN(SP04) FROM SERVICEPRACTICE " & _
                             "WHERE SP01 = (SELECT MIN(SP01) FROM SERVICEPRACTICE " & _
                                           "WHERE SP01 > '" & m_CurrSP(0) & "' AND " & "SP01 IN " & m_QuerySystem & ") AND " & _
                                                 "SP02 = (SELECT MIN(SP02) FROM SERVICEPRACTICE " & _
                                                         "WHERE SP01 = (SELECT MIN(SP01) FROM SERVICEPRACTICE " & _
                                                                       "WHERE SP01 > '" & m_CurrSP(0) & "' AND " & "SP01 IN " & m_QuerySystem & ")) AND " & _
                                                                             "SP03 = (SELECT MIN(SP03) FROM SERVICEPRACTICE " & _
                                                                                     "WHERE SP01 = (SELECT MIN(SP01) FROM SERVICEPRACTICE " & _
                                                                                                   "WHERE SP01 > '" & m_CurrSP(0) & "' AND " & "SP01 IN " & m_QuerySystem & ") AND " & _
                                                                                                         "SP02 = (SELECT MIN(SP02) FROM SERVICEPRACTICE " & _
                                                                                                                 "WHERE SP01 = (SELECT MIN(SP01) FROM SERVICEPRACTICE " & _
                                                                                                                               "WHERE SP01 > '" & m_CurrSP(0) & "' AND " & "SP01 IN " & m_QuerySystem & ")))) "
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         If IsNull(rsTmp.Fields("SP01")) = False Then: m_CurrSP(0) = rsTmp.Fields("SP01")
         If IsNull(rsTmp.Fields("SP02")) = False Then: m_CurrSP(1) = rsTmp.Fields("SP02")
         If IsNull(rsTmp.Fields("SP03")) = False Then: m_CurrSP(2) = rsTmp.Fields("SP03")
         If IsNull(rsTmp.Fields("SP04")) = False Then: m_CurrSP(3) = rsTmp.Fields("SP04")
         rsTmp.Close
         UpdateCtrlData
         GoTo EXITSUB
      End If
      rsTmp.Close
      UpdateCtrlData
   'End If
   'rsTmp.Close
   
EXITSUB:
   Set rsTmp = Nothing
End Sub
' 顯示最後一筆資料
Private Sub ShowLastRecord()
   m_CurrSP(0) = m_LastSP(0)
   m_CurrSP(1) = m_LastSP(1)
   m_CurrSP(2) = m_LastSP(2)
   m_CurrSP(3) = m_LastSP(3)
   
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
            tlbar.Buttons(6).Enabled = True
            tlbar.Buttons(7).Enabled = True
            tlbar.Buttons(8).Enabled = True
            tlbar.Buttons(9).Enabled = True
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
         strTit = "資料輸入有誤"
         strMsg = "請輸入 Y 或 空白"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
   End Select
End Function

' 按下按鍵
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   Call PUB_SaveMeTrackMode(m_MeTrackMode, 0, KeyCode)  'Added by Lydia 2021/12/20 Form2.0 記錄鍵盤傳入順序
   
'Memo by Lydia 2021/12/20 原程式搬到Form_KeyUp

End Sub
'add by nickc 2006/11/13 Enter 事件，等於存檔，做完取消，不然 form 內其他物件有寫 keycode 或是 keyascii 事件的話，也會做到Private Sub Form_KeyPress(KeyAscii As Integer)
Private Sub Form_KeyPress(KeyAscii As Integer)
'Remove by Lydia 2021/12/20 取消以ENTER控制為換行的功能 (Form2.0修改之維護資料功能Toolbar之修改統一)
'    Select Case KeyAscii
'      Case vbKeyReturn:
'         If m_EditMode <> 0 Then
'            KeyAscii = 0
'            OnAction vbKeyF9
'         End If
'    End Select
'end 2021/12/20
End Sub

' 執行指令
Private Sub OnAction(ByVal KeyCode As Integer)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
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
         UpdateToolbarState
         SetInputEntry
      ' 刪除
      Case vbKeyF5:
         If IsCaseProgressExist(textSP01, textSP02, textSP03, textSP04) = True Then
            strTit = "檢核資料"
            strMsg = "此本所案號在案件進度檔中仍有資料, 不可刪除!"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         Else
            strTit = "詢問"
            strMsg = "是否要刪除此筆資料?"
            nResponse = MsgBox(strMsg, vbYesNo + vbCritical + vbDefaultButton2, strTit)
            If nResponse = vbYes Then
               m_EditMode = 3
               'edit by nickc 2006/06/08
               'OnWork
               If OnWork = False Then Exit Sub
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
         UpdateFieldNewData
        'edit by nickc 2006/06/08
        'OnWork
        If OnWork = False Then Exit Sub
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

Private Sub Form_Unload(Cancel As Integer)
   'Add By Cheng 2002/07/19
   Set frm02050205 = Nothing
End Sub

Private Sub textSP01_Change()
    Select Case Me.textSP01
    Case "TS", "S"
        Me.Label12.Visible = True
        Me.textSP05_1.Visible = True
        Me.textSP05_1.Enabled = True
        Me.Label2.Visible = False
        Me.textSP05.Visible = False
        Me.textSP05.Enabled = False
        Me.Label3.Visible = False
        Me.textSP06.Visible = False
        Me.textSP06.Enabled = False
        Me.Label4.Visible = False
        Me.textSP07.Visible = False
        Me.textSP07.Enabled = False
    Case Else
        Me.Label12.Visible = False
        Me.textSP05_1.Visible = False
        Me.textSP05_1.Enabled = False
        Me.Label2.Visible = True
        Me.textSP05.Visible = True
        Me.textSP05.Enabled = True
        Me.Label3.Visible = True
        Me.textSP06.Visible = True
        Me.textSP06.Enabled = True
        Me.Label4.Visible = True
        Me.textSP07.Visible = True
        Me.textSP07.Enabled = True
    End Select
End Sub

Private Sub textSP01_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

' 系統類別
Private Sub textSP01_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   Cancel = False
   If IsEmptyText(textSP01) = False Then
      Select Case m_EditMode
         Case 1, 4:
            If IsUserHasRightOfSystem(strUserNum, textSP01) = False Then
               Cancel = True
               strTit = "資料檢核"
               strMsg = "您沒有使有此系統別的權限"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textSP01_GotFocus
               GoTo EXITSUB
            End If

            Select Case textSP01
               Case "TB", "TM", "TD", "TC":
                  Cancel = True
                  strTit = "資料檢核"
                  strMsg = "本所案號中的系統別不正確"
                  nResponse = MsgBox(strMsg, vbOKOnly, strTit)
                  textSP01_GotFocus
               Case Else:
            End Select
            
            'Select Case m_SysKind
            '   Case 0:
            '      Select Case textSP01
            '         Case "TB", "TM", "TD", "TC":
            '            Cancel = True
            '            strTit = "資料檢核"
            '            strMsg = "本所案號中的系統別不正確"
            '            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            '            textSP01_GotFocus
            '         Case Else:
            '      End Select
            '   Case 1:
            '      Select Case textSP01
            '         Case "S":
            '         Case Else:
            '            Cancel = True
            '            strTit = "資料檢核"
            '            strMsg = "本所案號中的系統別不正確"
            '            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            '            textSP01_GotFocus
            '      End Select
            'End Select
         Case Else:
      End Select
   End If
EXITSUB:
End Sub

' 本所案號輸入完後
'Private Sub textSP04_Validate(Cancel As Boolean)
Private Sub textSP04_LostFocus()
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Dim strSP01 As String
   Dim strSP02 As String
   Dim strSP03 As String
   Dim strSP04 As String
   
   strSP01 = textSP01
   strSP02 = textSP02
   strSP03 = textSP03 & String(1 - Len(textSP03), "0")
   strSP04 = textSP04 & String(2 - Len(textSP04), "0")

   If m_EditMode = 1 Then
      If IsRecordExist(strSP01, strSP02, strSP03, strSP04) = True Then
         strTit = "檢核資料"
         strMsg = "此筆資料已存在資料庫中"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textSP02.SetFocus
         GoTo EXITSUB
      End If
      ' 檢查是否超過自動編號
      'If IsOverAutoNumber(strSP01, DBYEAR(SystemDate()), strSP02) = True Then
      If IsOverAutoNumber(strSP01, Empty, strSP02) = True Then
         strTit = "檢核資料"
         strMsg = "本所案號中的流水號超過自動編號"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textSP02.SetFocus
         GoTo EXITSUB
      End If
   End If
EXITSUB:
End Sub

Private Sub textSP05_1_GotFocus()
    TextInverse Me.textSP05_1
    'edit by nickc 2007/06/06 切換輸入法改用API
    OpenIme
End Sub

Private Sub textSP05_1_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   
   If CheckLengthIsOK(textSP05_1, 140) = False Then
      Cancel = True
      strTit = "檢核資料"
      strMsg = "案件名稱內容太長"
      textSP05_1_GotFocus
   End If
   'edit by nickc 2007/06/06 切換輸入法改用API
   If Cancel = False Then CloseIme
End Sub

' 案件名稱(中)
Private Sub textSP05_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   
   'modify by sonia 2018/8/14 欄位已從60改為140,但此處沒改
   'Modified by Lydia 2021/02/19 改欄位寬度
   'If CheckLengthIsOK(textSP05, 140) = False Then
   If CheckLengthIsOK(textSP05, textSP05.MaxLength) = False Then
      Cancel = True
      strTit = "檢核資料"
      strMsg = "案件名稱(中)內容太長"
      textSP05_GotFocus
   End If
   'edit by nickc 2007/06/06 切換輸入法改用API
   'If Cancel = False Then: textSP05.IMEMode = 2
   If Cancel = False Then CloseIme
End Sub

'Added by Lydia 2021/02/19
Private Sub textSP06_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False

   If CheckLengthIsOK(textSP06, textSP06.MaxLength) = False Then
      Cancel = True
      strTit = "檢核資料"
      strMsg = "案件名稱(英)內容太長"
      textSP06_GotFocus
   End If
   If Cancel = False Then CloseIme
End Sub

' 案件名稱(日)
Private Sub textSP07_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   
   'Modified by Lydia 2021/02/19 改欄位寬度
   'If CheckLengthIsOK(textSP07, 60) = False Then
   If CheckLengthIsOK(textSP07, textSP07.MaxLength) = False Then
      Cancel = True
      strTit = "檢核資料"
      strMsg = "案件名稱(日)內容太長"
      textSP07_GotFocus
   End If
   'edit by nickc 2007/06/06 切換輸入法改用API
   'If Cancel = False Then: textSP07.IMEMode = 2
   If Cancel = False Then CloseIme
End Sub

'Modified by Lydia 2021/12/20 改成Form 2.0; KeyAscii As Integer=>MSForms.ReturnInteger
Private Sub textSP08_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = UpperCase(KeyAscii)
End Sub

'Modified by Lydia 2021/12/20 改成Form 2.0; KeyAscii As Integer=>MSForms.ReturnInteger
Private Sub textSP15_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = UpperCase(KeyAscii)
End Sub

' 案件備註
Private Sub textSP18_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   
   If CheckLengthIsOK(textSP18, 2000) = False Then
      Cancel = True
      strTit = "檢核資料"
      strMsg = "案件備註內容太長"
      textSP18_GotFocus
   End If
   'edit by nickc 2007/06/06 切換輸入法改用API
   'If Cancel = False Then: textSP18.IMEMode = 2
   If Cancel = False Then CloseIme
End Sub
'Modified by Lydia 2021/12/20 改成Form 2.0; KeyAscii As Integer=>MSForms.ReturnInteger
Private Sub textSP26_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = UpperCase(KeyAscii)
End Sub

' 折扣
Private Sub textSP31_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textSP31) = False Then
      If IsNumeric(textSP31) = False Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "只可輸入數值資料"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textSP31_GotFocus
      End If
   End If
End Sub
'Modified by Lydia 2021/12/20 改成Form 2.0; KeyAscii As Integer=>MSForms.ReturnInteger
Private Sub textSP33_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = UpperCase(KeyAscii)
End Sub
'Modified by Lydia 2021/12/20 改成Form 2.0; KeyAscii As Integer=>MSForms.ReturnInteger
Private Sub textSP35_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = UpperCase(KeyAscii)
End Sub
'Modified by Lydia 2021/12/20 改成Form 2.0; KeyAscii As Integer=>MSForms.ReturnInteger
Private Sub textSP36_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = UpperCase(KeyAscii)
End Sub
'Modified by Lydia 2021/12/20 改成Form 2.0; KeyAscii As Integer=>MSForms.ReturnInteger
Private Sub textSP37_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textSP67_GotFocus()
    TextInverse Me.textSP67
End Sub
'Modified by Lydia 2021/12/20 改成Form 2.0; KeyAscii As Integer=>MSForms.ReturnInteger
Private Sub textSP67_KeyPress(KeyAscii As MSForms.ReturnInteger)
    KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textSP67_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   'Add By Cheng 2002/07/09
   Dim strTempName As String
   
   Cancel = False
   textSP67_2 = Empty
   If IsEmptyText(textSP67) = False Then
      'Modify By Cheng 2002/07/09
'      textSP67_2 = GetAgentOrCustName(textSP67)
      If Left(Me.textSP67.Text, 1) = "X" Then
         textSP67_2 = GetAgentOrCustName(Me.textSP67.Text)
      Else
         If PUB_GetAgentName(Me.textSP01.Text, Me.textSP67.Text, strTempName) Then
            textSP67_2 = strTempName
         Else
            textSP67_2 = ""
         End If
      End If
      
      If IsEmptyText(textSP67_2) = True Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "D/N固定列印對象不存在"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textSP67_GotFocus
      End If
   End If
End Sub


'add by nickc 2006/12/07
Private Sub textSP73_GotFocus()
InverseTextBox textSP73
End Sub
Private Sub textSP73_Validate(Cancel As Boolean)
   Dim nCount As Integer
   Dim nIndex As Integer
   Dim strTit As String
   Dim strMsg As String
   Dim strTemp As String
   Dim nResponse
   If m_EditMode = 4 Then Exit Sub
   nCount = GetSubStringCount(textSP73)
   For nIndex = 1 To nCount
      strTemp = GetSubString(textSP73, nIndex)
      For nCount = 1 To nCount
         If nIndex <> nCount Then
            If strTemp = GetSubString(textSP73, nCount) Then
               Cancel = True
               strTit = "檢核資料"
               strMsg = "商品類別<" & strTemp & ">不可重覆"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textSP73_GotFocus
               GoTo EXITSUB
            End If
         End If
      Next nCount
   Next nIndex
textSP73 = Replace(textSP73, " ", "")
EXITSUB:
End Sub
Private Sub textSP74_GotFocus()
InverseTextBox textSP74
End Sub
Private Sub textSP74_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Dim nCount As Integer
   Dim nIndex As Integer
   Dim strSql As String
   Dim strTemp As String
   
   Cancel = False
   ' 無資料時不做任何檢查
   If IsEmptyText(textSP74) = True And m_EditMode = 4 Then
      GoTo EXITSUB
   End If
   
   'Modify By Sindy 2024/4/18 商品組群欄人員貼上資料後將全形或半形的「；」分號，轉為半形的逗號存入TM32。
   textSP74 = Replace(Replace(textSP74, ";", ","), "；", ",")
   '2024/4/18 END
   nCount = GetSubStringCount(textSP74)
   For nIndex = 1 To nCount
      strTemp = GetSubString(textSP74, nIndex)
      If Len(strTemp) > 6 Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "商品組群<" & strTemp & ">不正確"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textSP74_GotFocus
         GoTo EXITSUB
      End If
   Next nIndex
   
   For nIndex = 1 To nCount
      strTemp = GetSubString(textSP74, nIndex)
      For nCount = 1 To nCount
         If nIndex <> nCount Then
            If strTemp = GetSubString(textSP74, nCount) Then
               Cancel = True
               strTit = "檢核資料"
               strMsg = "商品組群<" & strTemp & ">不可重覆"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textSP74_GotFocus
               GoTo EXITSUB
            End If
         End If
      Next nCount
   Next nIndex
   
EXITSUB:
End Sub

'Add By Sindy 2017/11/9
Private Sub textSP80_GotFocus()
   CloseIme
   TextInverse textSP80
End Sub

'Modified by Lydia 2021/12/20 改成Form 2.0; KeyAscii As Integer=>MSForms.ReturnInteger
Private Sub textSP80_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 8 And KeyAscii <> Asc("Y") And KeyAscii <> Asc("D") Then
      KeyAscii = 0
      Beep
   End If
End Sub
Private Sub textSP80_Validate(Cancel As Boolean)
   If (textSP80 = "" And textSP83 = "Y") Then
      MsgBox "【EMail 同時寄紙本】為 Y 時，【以EMail 通知】欄位也必須為 Y！"
      Cancel = True
   End If
End Sub
Private Sub textSP83_GotFocus()
   InverseTextBox textSP83
End Sub
'Modified by Lydia 2021/12/20 改成Form 2.0; KeyAscii As Integer=>MSForms.ReturnInteger
Private Sub textSP83_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = UpperCase(KeyAscii)
End Sub
'EMail同時寄紙本
Private Sub textSP83_Validate(Cancel As Boolean)
Dim strTit As String
Dim strMsg As String
Dim nResponse
   
   Cancel = False
   If IsEmptyText(textSP83) = False Then
      If IsYesOrSpace(textSP83) = False Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "EMail同時寄紙本請輸入Y或空白"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textSP83_GotFocus
      End If
   End If
   If Cancel = False Then
      If (textSP80 = "" And textSP83 = "Y") Then
         MsgBox "【EMail 同時寄紙本】為 Y 時，【以EMail 通知】欄位也必須為 Y！"
         Cancel = True
      End If
   End If
End Sub
'2017/11/9 END

Private Sub textSP84_GotFocus()
   TextInverse textSP84
   CloseIme
End Sub

'Add By Sindy 2014/2/10
Private Sub textSP85_GotFocus()
   InverseTextBox textSP85
End Sub
'Modified by Lydia 2021/12/20 改成Form 2.0; KeyAscii As Integer=>MSForms.ReturnInteger
Private Sub textSP85_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = UpperCase(KeyAscii)
End Sub
'特殊出名公司
Private Sub textSP85_Validate(Cancel As Boolean)
Dim strTit As String
Dim strMsg As String
Dim nResponse
   
   Cancel = False
   textSP85 = Trim(textSP85)
   If textSP85 <> "" And textSP85 <> "J" Then
      Cancel = True
      strTit = "檢核資料"
      strMsg = "請輸入J或空白"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textSP85_GotFocus
   End If
End Sub

' 按下 ToolBar 的 Button
Private Sub tlbar_ButtonClick(ByVal Button As MSComctlLib.Button)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
  
   Call Pub_SaveMeToolBar(m_MeTrackMode, Me.tlbar, Button.Index) 'Added by Lydia 2021/12/20 若有交錯使用Function鍵和Toolbar鍵會失去記錄造成無法判斷，所以ToolBar鍵另外記錄
   
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
Private Function IsRecordExist(ByVal strSP01 As String, ByVal strSP02 As String, ByVal strSP03 As String, ByVal strSP04 As String) As Boolean
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   
   IsRecordExist = False
   strSql = "SELECT * FROM ServicePractice " & _
            "WHERE SP01 = '" & strSP01 & "' AND " & _
                  "SP02 = '" & strSP02 & "' AND " & _
                  "SP03 = '" & strSP03 & "' AND " & _
                  "SP04 = '" & strSP04 & "'"
                  
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
   Dim strSP01, strSP02, strSP03, strSP04 As String
   
   'add by nickc 2006/06/08
   AddRecord = False
   
   strSP01 = textSP01
   strSP02 = textSP02
   strSP03 = textSP03
   strSP04 = textSP04
   
   ' 檢查記錄是否已存在
   If IsRecordExist(strSP01, strSP02, strSP03, strSP04) = True Then
      strTit = "新增資料"
      strMsg = "該筆記錄已存在"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      UpdateCtrlData
      GoTo EXITSUB
   End If
   
   bFirst = True
   bDifference = False
   strSql = "INSERT INTO ServicePractice ("
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
   
   'add by nickc 2006/06/08
   On Error GoTo oErr
   cnnConnection.BeginTrans
   Pub_SeekTbLog strSql
   
   cnnConnection.Execute strSql
   
   'add by nickc 2006/06/08
   cnnConnection.CommitTrans
   
   If ((strSP01 & strSP02 & strSP03 & strSP04) < (m_FirstSP(0) & m_FirstSP(1) & m_FirstSP(2) & m_FirstSP(3))) Or ((strSP01 & strSP02 & strSP03 & strSP04) > (m_LastSP(0) & m_LastSP(1) & m_LastSP(2) & m_LastSP(3))) Then
      RefreshRange
   End If

   ShowCurrRecord strSP01, strSP02, strSP03, strSP04
    AddRecord = True
EXITSUB:
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
   Dim strSP01, strSP02, strSP03, strSP04 As String
   Dim bolData As Boolean, strMCTF(0) As String 'Add by Amy 2017/03/22
   
   'add by nickc 2006/06/08
   ModRecord = False
   
   strSP01 = textSP01
   strSP02 = textSP02
   strSP03 = textSP03
   strSP04 = textSP04
   
   '910910  nick tigger
   '***** start
   'strsql = "UPDATE ServicePractice SET "
   'edit by nickc 2006/06/08
   'strSQL = "begin user_data.user_enabled:=1; UPDATE ServicePractice SET "
   strSql = "UPDATE ServicePractice SET "
   '***** end
   bFirst = True
   bDifference = False
   For nIndex = 0 To MAX_FIELD - 1
      strTmp = Empty
      '92.05.22 nick 跳過create & update 相關項目
      If (nIndex < 51 Or nIndex > 56) And nIndex <> 60 And nIndex <> 67 And nIndex <> 68 And nIndex <> 69 Then
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
                  "WHERE SP01 = '" & strSP01 & "' AND " & _
                     "SP02 = '" & strSP02 & "' AND " & _
                     "SP03 = '" & strSP03 & "' AND " & _
                     "SP04 = '" & strSP04 & "'"
    'edit by nickc 2006/06/08
    'strSQL = strSQL & " " & _
                  "WHERE SP01 = '" & strSP01 & "' AND " & _
                     "SP02 = '" & strSP02 & "' AND " & _
                     "SP03 = '" & strSP03 & "' AND " & _
                     "SP04 = '" & strSP04 & "'; end ;"
   strSql = strSql & " " & _
                  "WHERE SP01 = '" & strSP01 & "' AND " & _
                     "SP02 = '" & strSP02 & "' AND " & _
                     "SP03 = '" & strSP03 & "' AND " & _
                     "SP04 = '" & strSP04 & "'"
    '***** end
'910910 nick tigger
'***** start
On Error GoTo ErrHand
'***** end
   
   If bDifference = True Then
      '910910 nick tigger
      '**** start
      cnnConnection.BeginTrans
     'add by nickc 2006/06/08 紀錄分析語法
      Pub_SeekTbLog strSql
      '***** end
      'edit by nickc 2006/06/08
      'cnnConnection.Execute strSQL
      cnnConnection.Execute "begin user_data.user_enabled:=1;" & strSql & "; end;"
      
      'Add by Amy 2017/03/22 FC代理人修改為MCTF時更新客戶檔及下一程序
      If Left(textSP01, 1) = "T" And Trim(Len(textSP26)) > 0 And m_FieldList(25).fiOldData <> ChangeCustomerL(textSP26) Then
        bolData = GetCusORFagentData(ChangeCustomerL(textSP26), "FA120", strMCTF())
        If Left(strMCTF(0), 4) = "MCTF" Then
            If UpdMCTF_NP(ChangeCustomerL(textSP26), strMCTF(0), textSP01 & textSP02 & textSP03 & textSP04) = False Then GoTo ErrHand
        End If
      End If
        
      '910910 nick tigger
      '***** start
      cnnConnection.CommitTrans
      '***** end
      QueryDB
      ShowCurrRecord strSP01, strSP02, strSP03, strSP04
      'add by nickc 2005/08/23 紀錄修改案號
      pub_ModifyCaseNum = strSP01 & "-" & strSP02 & "-" & strSP03 & "-" & strSP04
   End If
'910910 nick tigger
'***** start
    'add by nickc 2006/06/08
    ModRecord = True
   Exit Function
ErrHand:
    cnnConnection.RollbackTrans
'******* end
    'add by nickc 2006/06/08
    MsgBox Err.Description
End Function

' 刪除記錄
'edit by nickc 2006/06/08
'Private Sub DelRecord()
Private Function DelRecord() As Boolean
   Dim strSql As String
   Dim strSP01 As String
   Dim strSP02 As String
   Dim strSP03 As String
   Dim strSP04 As String
   
   
   'add by nickc 2006/06/08
   DelRecord = False
   
   strSP01 = textSP01
   strSP02 = textSP02
   strSP03 = textSP03
   strSP04 = textSP04
   
   'Add By Sindy 2010/7/1
   If ChkCaseCode("CP", strSP01, strSP02, strSP03, strSP04) = False Then Exit Function
   If ChkCaseCode("NP", strSP01, strSP02, strSP03, strSP04) = False Then Exit Function
   '2010/7/1 End
   
   If OnDataDeleteRecord(0, strSP01 & strSP02 & strSP03 & strSP04) <> 0 Then
      GoTo EXITSUB
   End If
   
   strSql = "DELETE FROM ServicePractice " & _
            "WHERE SP01 = '" & textSP01 & "' AND " & _
                  "SP02 = '" & textSP02 & "' AND " & _
                  "SP03 = '" & textSP03 & "' AND " & _
                  "SP04 = '" & textSP04 & "'"
   'add by nickc 2006/06/08
   On Error GoTo oErr
   cnnConnection.BeginTrans
   Pub_SeekTbLog strSql
   cnnConnection.Execute strSql
   
    'Added by Lydia 2016/11/24 一併刪除各項指示
    strSql = "DELETE FROM INSTRUCTIONS WHERE ITS01=" & CNULL(Pub_GetITS01Type(textSP01)) & " AND ITS02=" & CNULL(textSP01 & textSP02 & textSP03 & textSP04)
    Pub_SeekTbLog strSql
    cnnConnection.Execute strSql
    'end 2016/11/24
    
   'add by nickc 2006/06/08
   cnnConnection.CommitTrans
   
   ' 只有刪除的是第一筆或是最後一筆才須重新取的第一筆及最後一筆的本所案號
   If (strSP01 = m_FirstSP(0) And strSP02 = m_FirstSP(1) And strSP03 = m_FirstSP(2) And strSP04 = m_FirstSP(3)) Or (strSP01 = m_LastSP(0) And strSP02 = m_LastSP(1) And strSP03 = m_LastSP(2) And strSP04 = m_LastSP(3)) Then
      RefreshRange
   End If
   
   ShowCurrRecord strSP01, strSP02, strSP03, strSP04
   'add by nickc 2006/06/08
   DelRecord = True
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
   
   If IsEmptyText(textSP03) = True Then: textSP03 = "0"
   If IsEmptyText(textSP04) = True Then: textSP04 = "00"
   
   If IsRecordExist(textSP01, textSP02, textSP03, textSP04) = True Then
      m_CurrSP(0) = textSP01
      m_CurrSP(1) = textSP02
      m_CurrSP(2) = textSP03
      m_CurrSP(3) = textSP04
      QueryRecord = True
      UpdateCtrlData
   Else
      QueryRecord = False
   End If
   
   UpdateToolbarState
End Function

' 使用者按下確定的按紐
'edit by nickc 2006/06/08
'Private Sub OnWork()
Private Function OnWork() As Boolean
   Dim strMsg As String
   Dim strTit As String
   Dim nResponse
   'add by nickc 2006/06/08
   OnWork = False
   
   Select Case m_EditMode
      Case 1:
         If CheckDataValid() = True Then
            'Add By Cheng 2002/06/24
            If TxtValidate = False Then GoTo EXITSUB
            
            'add by nickc 2006/06/08
            'AddRecord
            If AddRecord = False Then GoTo EXITSUB
            
            RefreshRange
         Else
            GoTo EXITSUB
         End If
      Case 2:
         If CheckDataValid() = True Then
            'Add By Cheng 2002/06/24
            If TxtValidate = False Then GoTo EXITSUB
            'Added by Lydia 2017/06/19 (存檔前)檢查收文智權人員和客戶智權人員不同業務區要發mail的內文和通知人員
            strChkCuAreaMail = PUB_ChkSameCustSales(Trim(textSP01), Trim(textSP02), Trim(textSP03), Trim(textSP04), "", Trim(textSP08), "", "", "", "", strChkCuAreaMailTo)
            
            'edit by nickc 2006/06/08
            'ModRecord
            If ModRecord = False Then GoTo EXITSUB
            
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
         If DelRecord = False Then GoTo EXITSUB
         
         RefreshRange
      Case 4:
         If QueryRecord = False Then
            strMsg = "無此資料"
            strTit = "查詢資料"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            UpdateCtrlData
         End If
   End Select
   m_EditMode = 0
   SetCtrlReadOnly True
   'add by nickc 2006/06/08
   OnWork = True
EXITSUB:
End Function

' 檢查台灣日期
Private Function IsValidTDate(ByVal strDate As String) As Boolean
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   IsValidTDate = True
   If strDate <> Empty Then
      If CheckIsTaiwanDate(strDate, False) = False Then
         IsValidTDate = False
         strMsg = "日期不正確, 請重新輸入"
         strTit = "資料檢核"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      End If
   End If
End Function

' 檢查西元日期
Private Function IsValidDate(ByVal strDate As String) As Boolean
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   IsValidDate = True
   If strDate <> Empty Then
      If CheckIsDate(strDate, False) = False Then
         IsValidDate = False
         strMsg = "日期不正確, 請重新輸入"
         strTit = "資料檢核"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      End If
   End If
End Function

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

' 申請人欄位
Private Sub textSP08_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Dim rsTmp As ADODB.Recordset
   Dim strKey As String
   Dim strSql As String
   Dim strData As String
   
   Cancel = False
   textSP08_2 = Empty
   textCU79 = Empty
   If IsEmptyText(textSP08) = False Then
      strData = textSP08 & String(9 - Len(textSP08), "0")
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
               textSP08_2 = rsTmp.Fields("CU04")
            ElseIf IsNull(rsTmp.Fields("CU05")) = False Then
               textSP08_2 = rsTmp.Fields("CU05")
            ElseIf IsNull(rsTmp.Fields("CU06")) = False Then
               textSP08_2 = rsTmp.Fields("CU06")
            End If
            If IsNull(rsTmp.Fields("CU79")) = False Then
               textCU79 = rsTmp.Fields("CU79")
            End If
         Else
            Cancel = True
            strTit = "資料檢核"
            strMsg = "申請人代號不存在"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textSP08_GotFocus
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
         ' 檢查讀取的資料筆數
         If rsTmp.RecordCount > 0 Then
            rsTmp.MoveFirst
            If IsNull(rsTmp.Fields("FA03")) = False Then
               strKey = rsTmp.Fields("FA03")
               textSP08 = rsTmp.Fields("FA03")
               rsTmp.Close
               If Len(strKey) > 8 Then
                  strSql = "SELECT * FROM Customer " & _
                        "WHERE CU01 = '" & Mid(strKey, 1, 8) & "' AND " & _
                              "CU02 = '" & Mid(strKey, 9, 1) & "'"
               Else
                  strSql = "SELECT * FROM Customer " & _
                        "WHERE CU01 = '" & Mid(strKey, 1, 8) & "'"
               End If
               rsTmp.CursorLocation = adUseClient
               rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly   'edit by nickc 2005/08/04 原先是  動態開啟
               If rsTmp.RecordCount > 0 Then
                  rsTmp.MoveFirst
                  If IsNull(rsTmp.Fields("CU04")) = False Then
                     textSP08_2 = rsTmp.Fields("CU04")
                  ElseIf IsNull(rsTmp.Fields("CU05")) = False Then
                     textSP08_2 = rsTmp.Fields("CU05")
                  ElseIf IsNull(rsTmp.Fields("CU06")) = False Then
                     textSP08_2 = rsTmp.Fields("CU06")
                  End If
                  If IsNull(rsTmp.Fields("CU79")) = False Then
                     textCU79 = rsTmp.Fields("CU79")
                  End If
               Else
                  Cancel = True
                  strTit = "資料檢核"
                  strMsg = "申請人代號不存在"
                  nResponse = MsgBox(strMsg, vbOKOnly, strTit)
                  textSP08_GotFocus
               End If
            Else
               Cancel = True
               strTit = "資料檢核"
               strMsg = "申請人代號不存在"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textSP08_GotFocus
            End If
         Else
            Cancel = True
            strTit = "資料檢核"
            strMsg = "申請人代號不存在"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textSP08_GotFocus
         End If
         rsTmp.Close
      Case Else:
         Cancel = True
         strTit = "檢核資料"
         strMsg = "申請人代號不正確"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textSP08_GotFocus
      End Select
   End If
   Set rsTmp = Nothing
End Sub
' 申請國家
Private Sub textSP09_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   Cancel = False
   textSP09_2 = Empty
   If IsEmptyText(textSP09) = False Then
      ' 申請國家不可輸入 001 - 008
      Select Case textSP09
         Case "001", "002", "003", "004", "005", "006", "007", "008":
            Cancel = True
            strTit = "檢核資料"
            strMsg = "申請國家不可輸入001-008"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textSP09_GotFocus
            GoTo EXITSUB
         Case Else:
      End Select
      textSP09_2 = GetNationName(textSP09, 0)
      If IsEmptyText(textSP09_2) = True Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "申請國家代碼不存在"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textSP09_GotFocus
         GoTo EXITSUB
      End If
   End If
EXITSUB:
End Sub

' 申請日
Private Sub textSP10_Validate(Cancel As Boolean)
   Cancel = False
   If IsValidTDate(textSP10) = False Then
      Cancel = True
      textSP10_GotFocus
   End If
End Sub

' 是否閉卷
Private Sub textSP15_Validate(Cancel As Boolean)
   Cancel = False
   If IsYesOrSpace(textSP15) = False Then
      Cancel = True
      textSP15_GotFocus
   End If
End Sub

' 閉卷日期
Private Sub textSP16_Validate(Cancel As Boolean)
   Cancel = False
   If IsValidTDate(textSP16) = False Then
      Cancel = True
      textSP16_GotFocus
   End If
End Sub

' 閉卷原因
Private Sub textSP17_Validate(Cancel As Boolean)
   Dim rsTmp As ADODB.Recordset
   Dim strSql As String
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   textSP17_2 = Empty
   If IsEmptyText(textSP17) = False Then
      Set rsTmp = New ADODB.Recordset
      strSql = "SELECT * FROM ReasonOfRelief " & _
               "WHERE ROR01 = '" & textSP17 & "'"
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly   'edit by nickc 2005/08/04 原先是  動態開啟
      ' 檢查讀取的資料筆數
      If rsTmp.RecordCount > 0 Then
         rsTmp.MoveFirst
         If IsNull(rsTmp.Fields("ROR02")) = False Then
            textSP17_2 = rsTmp.Fields("ROR02")
         End If
      Else
         Cancel = True
         strTit = "檢核資料"
         strMsg = "閉卷原因不存在"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textSP17_GotFocus
      End If
      rsTmp.Close
   End If
   Set rsTmp = Nothing
End Sub

' FC代理人
Private Sub textSP26_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Dim rsTmp As ADODB.Recordset
   Dim strKey As String
   Dim strSql As String
   Dim strData As String
   Cancel = False
   textSP26_2 = Empty
   textFA29 = Empty
   If IsEmptyText(textSP26) = False Then
      strData = textSP26 & String(9 - Len(textSP26), "0")
      ' 不滿8碼補滿8碼
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
               textSP26 = rsTmp.Fields("CU03")
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
                     textSP26_2 = rsTmp.Fields("FA05")
                  ElseIf IsNull(rsTmp.Fields("FA04")) = False Then
                     textSP26_2 = rsTmp.Fields("FA04")
                  ElseIf IsNull(rsTmp.Fields("FA06")) = False Then
                     textSP26_2 = rsTmp.Fields("FA06")
                  End If
                  If IsNull(rsTmp.Fields("FA29")) = False Then
                     textFA29 = rsTmp.Fields("FA29")
                  End If
               Else
                  Cancel = True
                  strTit = "資料檢核"
                  strMsg = "FC代理人代號不存在"
                  nResponse = MsgBox(strMsg, vbOKOnly, strTit)
                  textSP26_GotFocus
               End If
            Else
               Cancel = True
               strTit = "資料檢核"
               strMsg = "FC代理人代號不存在"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textSP26_GotFocus
            End If
         Else
            Cancel = True
            strTit = "資料檢核"
            strMsg = "FC代理人代號不存在"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textSP26_GotFocus
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
               textSP26_2 = rsTmp.Fields("FA05")
            ElseIf IsNull(rsTmp.Fields("FA04")) = False Then
               textSP26_2 = rsTmp.Fields("FA04")
            ElseIf IsNull(rsTmp.Fields("FA06")) = False Then
               textSP26_2 = rsTmp.Fields("FA06")
            End If
            If IsNull(rsTmp.Fields("FA29")) = False Then
               textFA29 = rsTmp.Fields("FA29")
            End If
         Else
            Cancel = True
            strTit = "資料檢核"
            strMsg = "FC代理人代號不存在"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textSP26_GotFocus
         End If
         rsTmp.Close
      Case Else:
         Cancel = True
         strTit = "檢核資料"
         strMsg = "代理人代號不正確"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textSP26_GotFocus
      End Select
   End If
   Set rsTmp = Nothing
End Sub

' D/N是否列印申請人
Private Sub textSP33_Validate(Cancel As Boolean)
   Cancel = False
   If IsYesOrSpace(textSP33) = False Then
      Cancel = True
   End If
End Sub

' 定稿語文
Private Sub textSP34_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   Select Case textSP34
      Case "", "1", "2", "3"
      Case Else
         Cancel = True
         strTit = "資料輸入有誤"
         strMsg = "請輸入 1 , 2 或 3"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textSP34_GotFocus
   End Select
End Sub

' 副本收受人
Private Sub textSP35_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   'Add By Cheng 2002/07/09
   Dim strTempName As String
   
   Cancel = False
   textSP35_2 = Empty
   If IsEmptyText(textSP35) = False Then
      'Modify By Cheng 2002/07/09
'      textSP35_2 = GetAgentOrCustName(textSP35)
      If Left(Me.textSP35.Text, 1) = "X" Then
         textSP35_2 = GetAgentOrCustName(Me.textSP35.Text)
      Else
         If PUB_GetAgentName(Me.textSP01.Text, Me.textSP35.Text, strTempName) Then
            textSP35_2 = strTempName
         Else
            textSP35_2 = ""
         End If
      End If

      If IsEmptyText(textSP35_2) = True Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "副本收受人不存在"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textSP35_GotFocus
      End If
   End If
End Sub

' 固定請款對象
Private Sub textSP37_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   'Add By Cheng 2002/07/09
   Dim strTempName As String
   
   Cancel = False
   textSP37_2 = Empty
   If IsEmptyText(textSP37) = False Then
      'Modify By Cheng 2002/07/09
'      textSP37_2 = GetAgentOrCustName(textSP37)
      If Left(Me.textSP37.Text, 1) = "X" Then
         textSP37_2 = GetAgentOrCustName(Me.textSP37.Text)
      Else
         If PUB_GetAgentName(Me.textSP01.Text, Me.textSP37.Text, strTempName) Then
            textSP37_2 = strTempName
         Else
            textSP37_2 = ""
         End If
      End If
      
      If IsEmptyText(textSP37_2) = True Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "固定請款對象不存在"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textSP37_GotFocus
      End If
   End If
End Sub

' 是否銷卷
'Private Sub textSP61_Validate(Cancel As Boolean)
'   Cancel = False
'   If IsYesOrSpace(textSP61) = False Then
'      Cancel = True
'      textSP61_GotFocus
'   End If
'End Sub

Private Sub textSP01_GotFocus()
   InverseTextBox textSP01
End Sub

Private Sub textSP02_GotFocus()
   InverseTextBox textSP02
End Sub

Private Sub textSP03_GotFocus()
   InverseTextBox textSP03
End Sub

Private Sub textSP04_GotFocus()
   InverseTextBox textSP04
End Sub

Private Sub textSP05_GotFocus()
   InverseTextBox textSP05
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textSP05.IMEMode = 1
   OpenIme
End Sub

Private Sub textSP06_GotFocus()
   InverseTextBox textSP06
End Sub

Private Sub textSP07_GotFocus()
   InverseTextBox textSP07
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textSP07.IMEMode = 1
   OpenIme
End Sub

Private Sub textSP08_GotFocus()
   InverseTextBox textSP08
End Sub

Private Sub textSP09_GotFocus()
   InverseTextBox textSP09
End Sub

Private Sub textSP10_GotFocus()
   InverseTextBox textSP10
End Sub

Private Sub textSP15_GotFocus()
   InverseTextBox textSP15
End Sub

Private Sub textSP16_GotFocus()
   InverseTextBox textSP16
End Sub

Private Sub textSP17_GotFocus()
   InverseTextBox textSP17
End Sub

Private Sub textSP18_GotFocus()
   InverseTextBox textSP18
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textSP18.IMEMode = 1
   OpenIme
End Sub

Private Sub textSP26_GotFocus()
   InverseTextBox textSP26
End Sub

Private Sub textSP27_GotFocus()
   InverseTextBox textSP27
End Sub

Private Sub textSP28_GotFocus()
   InverseTextBox textSP28
End Sub

Private Sub textSP29_GotFocus()
   InverseTextBox textSP29
End Sub

Private Sub textSP30_GotFocus()
   InverseTextBox textSP30
End Sub

Private Sub textSP30_Validate(Cancel As Boolean)
   If CheckLengthIsOK(textSP30, textSP30.MaxLength) = False Then
      Cancel = True
      textSP30_GotFocus
   End If
End Sub

'Add By Sindy 2016/11/4 聯絡人2
Private Sub textSP75_GotFocus()
   InverseTextBox textSP75
End Sub
Private Sub textSP75_Validate(Cancel As Boolean)
   If CheckLengthIsOK(textSP75, textSP75.MaxLength) = False Then
      Cancel = True
      textSP75_GotFocus
   End If
End Sub
'2016/11/4 END

Private Sub textSP71_GotFocus()
   InverseTextBox textSP71
End Sub

Private Sub textSP71_Validate(Cancel As Boolean)
   If CheckLengthIsOK(textSP71, textSP71.MaxLength) = False Then
      Cancel = True
      textSP71_GotFocus
   End If
End Sub

Private Sub textSP31_GotFocus()
   InverseTextBox textSP31
End Sub

Private Sub textSP33_GotFocus()
   InverseTextBox textSP33
End Sub

Private Sub textSP34_GotFocus()
   InverseTextBox textSP34
End Sub

Private Sub textSP35_GotFocus()
   InverseTextBox textSP35
End Sub

Private Sub textSP36_GotFocus()
   InverseTextBox textSP36
End Sub

Private Sub textSP37_GotFocus()
   InverseTextBox textSP37
End Sub

'Private Sub textSP61_GotFocus()
'   InverseTextBox textSP61
'End Sub

Private Sub textCU79_GotFocus()
   InverseTextBox textCU79
End Sub

Private Sub textFA29_GotFocus()
   InverseTextBox textFA29
End Sub

' 案件進度檔
Private Function IsCaseProgressExist(ByVal strSP01 As String, ByVal strSP02 As String, ByVal strSP03 As String, ByVal strSP04 As String) As Boolean
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   
   IsCaseProgressExist = False
   strSql = "SELECT * from CaseProgress " & _
            "WHERE CP01 = '" & strSP01 & "' AND " & _
                  "CP02 = '" & strSP02 & "' AND " & _
                  "CP03 = '" & strSP03 & "' AND " & _
                  "CP04 = '" & strSP04 & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly   'edit by nickc 2005/08/04 原先是  動態開啟
   If rsTmp.RecordCount > 0 Then
      IsCaseProgressExist = True
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Function

' 開始輸入資料
Private Sub SetInputEntry()
    Select Case m_EditMode
    Case 1: textSP01.SetFocus
    Case 2
        Select Case Me.textSP01
        Case "TS", "S"
            If Me.textSP05_1.Enabled = True Then textSP05_1.SetFocus
        Case Else
            If Me.textSP05.Enabled = True Then textSP05.SetFocus
        End Select
    Case 4: textSP01.SetFocus
    End Select
End Sub

' 檢查輸入是否完整
Private Function CheckDataValid() As Boolean
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   CheckDataValid = False
   
   Select Case m_EditMode
      Case 1, 2, 4:
         ' 本所案號第一欄
         If IsEmptyText(textSP01) = True Then
            strTit = "檢核資料"
            strMsg = "請輸入本所案號"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textSP01.SetFocus
            GoTo EXITSUB
         End If

         ' 本所案號第二欄
         If IsEmptyText(textSP02) = True Then
            strTit = "檢核資料"
            strMsg = "請輸入本所案號"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textSP02.SetFocus
            GoTo EXITSUB
         End If
      Case Else:
   End Select
   
      'add by nickc 2007/02/07 新增時，編號不可大於自動編號
   If m_EditMode = 1 Then
        If ClsPDChkCaseNum(textSP01, textSP02) Then
                GoTo EXITSUB
        End If
   End If
   
   Select Case m_EditMode
      Case 1, 2:
        Select Case Me.textSP01.Text
        Case "TS", "S"
            ' 案件名稱不可同時空白
            If IsEmptyText(textSP05_1) = True Then
               strTit = "檢核資料"
               strMsg = "案件名稱不可空白"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textSP05_1.SetFocus
               GoTo EXITSUB
            End If
        Case Else
            ' 案件名稱(中)(英)(日)不可同時空白
            If IsEmptyText(textSP05) = True And IsEmptyText(textSP06) = True And IsEmptyText(textSP07) = True Then
               strTit = "檢核資料"
               strMsg = "案件名稱(中)(英)(日)不可同時空白"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textSP05.SetFocus
               GoTo EXITSUB
            End If
        End Select
         ' 申請國家不可空白
         If IsEmptyText(textSP09) = True Then
            strTit = "檢核資料"
            strMsg = "申請國家不可空白"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textSP09.SetFocus
            GoTo EXITSUB
         End If
         ' 申請人編號不可空白
         If IsEmptyText(textSP08) = True And IsEmptyText(textSP26) = True Then
            strTit = "檢核資料"
            strMsg = "申請人編號或代理人編號不可同時空白"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textSP08.SetFocus
            GoTo EXITSUB
         End If
      Case Else:
   End Select
   
   CheckDataValid = True
EXITSUB:
End Function
'Add By Cheng 2002/01/04
Private Sub SetQueryStatus()
m_EditMode = 4
SetCtrlReadOnly True
SetKeyReadOnly False
ClearField
UpdateToolbarState
'SetInputEntry
End Sub

'Add By Cheng 2002/06/24
Private Function TxtValidate() As Boolean
Dim objTxt As Object
Dim ii As Integer
Dim Cancel As Boolean
'Add by Amy 2017/03/22
Dim bolData As Boolean, strMsg As String
Dim strMCTFNew(0) As String, strTmp(0) As String

TxtValidate = False

If Me.textSP01.Enabled = True Then
   Cancel = False
   textSP01_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textSP05_1.Enabled = True Then
   Cancel = False
   textSP05_1_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textSP05.Enabled = True Then
   Cancel = False
   textSP05_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textSP07.Enabled = True Then
   Cancel = False
   textSP07_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textSP08.Enabled = True Then
   Cancel = False
   textSP08_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textSP09.Enabled = True Then
   Cancel = False
   textSP09_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textSP10.Enabled = True Then
   Cancel = False
   textSP10_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textSP15.Enabled = True Then
   Cancel = False
   textSP15_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textSP16.Enabled = True Then
   Cancel = False
   textSP16_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textSP17.Enabled = True Then
   Cancel = False
   textSP17_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textSP18.Enabled = True Then
   Cancel = False
   textSP18_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textSP26.Enabled = True Then
   Cancel = False
   textSP26_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textSP31.Enabled = True Then
   Cancel = False
   textSP31_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textSP33.Enabled = True Then
   Cancel = False
   textSP33_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textSP34.Enabled = True Then
   Cancel = False
   textSP34_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textSP35.Enabled = True Then
   Cancel = False
   textSP35_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textSP37.Enabled = True Then
   Cancel = False
   textSP37_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

'If Me.textSP61.Enabled = True Then
'   Cancel = False
'   textSP61_Validate Cancel
'   If Cancel = True Then
'      Exit Function
'   End If
'End If

If Me.textSP67.Enabled = True Then
   Cancel = False
   textSP67_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

'Add By Sindy 2017/11/9
If Me.textSP80.Enabled = True Then
   Cancel = False
   textSP80_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If
If Me.textSP83.Enabled = True Then
   Cancel = False
   textSP83_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If
'2017/11/9 END

'Add By Sindy 2014/2/10
If Me.textSP85.Enabled = True Then
   Cancel = False
   textSP85_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

'Add by Morgan 2007/5/10
If Not ((textSP15.Text = "" And textSP16.Text = "" And textSP17.Text = "") Or (textSP15.Text <> "" And textSP16.Text <> "" And textSP17.Text <> "")) Then
   MsgBox "是否閉卷、閉卷日期、閉卷原因三個欄位須同時空白或有值！", vbExclamation
   Exit Function
End If
'end 2007/5/10

'Add by Amy 2017/03/22 T字頭若FC代理人之管控智權人員為MCTF時,修改成不同組別不可存檔
If m_EditMode = 2 And Left(textSP01, 1) = "T" And Trim(Len(textSP26)) > 0 And m_FieldList(25).fiOldData <> ChangeCustomerL(textSP26) Then
    strMsg = ""
    bolData = GetCusORFagentData(ChangeCustomerL(textSP26), "FA120", strMCTFNew())
    If Left(strMCTFNew(0), 4) = "MCTF" And Len(Trim(textSP08)) > 0 Then
        bolData = GetCusORFagentData(ChangeCustomerL(textSP08), "CU13", strTmp())
        If strMCTFNew(0) <> strTmp(0) And Left(strTmp(0), 4) = "MCTF" Then
            MsgBox "申請人：" & textSP26 & " (" & strTmp(0) & ")" & vbCrLf & "與代理人" & textSP26 & _
                "商標管控智權人員(" & strMCTFNew(0) & ")不同，不可存檔！"
            Exit Function
        End If
    End If
End If

'Added by Lydia 2024/06/13 檢查更新代理人／申請人狀態排除「不得代理」
strExc(1) = ChangeCustomerL(textSP08)
strExc(2) = ChangeCustomerL(textSP08.Tag)
If strExc(1) <> "" And strExc(1) <> strExc(2) Then
   If GetCustomerAndState(strExc(1), strExc(3), , , , textSP01, strExc(8), False, Me.Name, textSP02, textSP03, textSP04) = False Then
      Me.tabCtrl.Tab = 0
      textSP08.SetFocus
      textSP08_GotFocus
      Exit Function
   End If
End If
strExc(1) = ChangeCustomerL(textSP26)
strExc(2) = ChangeCustomerL(textSP26.Tag)
If strExc(1) <> "" And strExc(1) <> strExc(2) Then
   If GetAgentAndState(strExc(1), strExc(3), , , , textSP01, strExc(8), False) = False Then
      Me.tabCtrl.Tab = 1
      textSP26.SetFocus
      textSP26_GotFocus
      Exit Function
   End If
End If
'end 2024/06/13

'Added by Lydia 2021/12/20 檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
If PUB_ChkUniText(Me, , True, "TextBox") = False Then
    Exit Function
End If

TxtValidate = True
End Function

'Add by Amy 2016/08/26
Private Sub ReadPic()
    '檢查有無代表圖
    'Modify by Amy 2018/07/16  改寫至function
'    strSql = "SELECT ibf01,ibf02 FROM imgbytefile WHERE ibf01='" & textSP01 & "' and ibf02='" & textSP02 & "' and ibf03='" & textSP03 & "' and ibf04='" & textSP04 & "' and ibf05='1'"
'    CheckOC2
'    adoRecordset1.CursorLocation = adUseClient
'    adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
'    If adoRecordset1.RecordCount <> 0 And adoRecordset1.RecordCount > 0 Then
    If ChkImgByteFile(textSP01, textSP02, textSP03, textSP04) = True Then
        'Modofied by Lydia 2021/12/20 拿掉快速鍵
        Command3.Caption = "已設定代表圖"
        Command3.BackColor = &HC0FFC0
    Else
        'Modofied by Lydia 2021/12/20 拿掉快速鍵
        Command3.Caption = "未設定代表圖"
        Command3.BackColor = &HC0C0FF
    End If
'    CheckOC2
    'end 2018/07/16
End Sub
