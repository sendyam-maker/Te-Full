VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm030101_05 
   BorderStyle     =   1  '單線固定
   Caption         =   "發文(變更事項)"
   ClientHeight    =   5595
   ClientLeft      =   5505
   ClientTop       =   1770
   ClientWidth     =   9150
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5595
   ScaleWidth      =   9150
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Left            =   7200
      TabIndex        =   3
      Top             =   60
      Width           =   912
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   8160
      TabIndex        =   2
      Top             =   60
      Width           =   912
   End
   Begin TabDlg.SSTab tabCtrl 
      Height          =   4632
      Left            =   120
      TabIndex        =   4
      Top             =   900
      Width           =   8952
      _ExtentX        =   15796
      _ExtentY        =   8176
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   5
      TabHeight       =   420
      TabCaption(0)   =   "第一頁"
      TabPicture(0)   =   "frm030101_05.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "checkCE65"
      Tab(0).Control(1)=   "checkCE22"
      Tab(0).Control(2)=   "checkCE52"
      Tab(0).Control(3)=   "checkCE54"
      Tab(0).Control(4)=   "checkCE56"
      Tab(0).Control(5)=   "checkCE16"
      Tab(0).Control(6)=   "checkCE03"
      Tab(0).Control(7)=   "checkCE09"
      Tab(0).Control(8)=   "textCE04_2"
      Tab(0).Control(9)=   "textCE04"
      Tab(0).Control(10)=   "textCE02"
      Tab(0).Control(11)=   "textCE10"
      Tab(0).Control(12)=   "textCE11"
      Tab(0).Control(13)=   "textCE12"
      Tab(0).Control(14)=   "textCE13"
      Tab(0).Control(15)=   "textCE14"
      Tab(0).Control(16)=   "textCE15"
      Tab(0).Control(17)=   "textCE17"
      Tab(0).Control(18)=   "textCE63"
      Tab(0).Control(19)=   "textCE64"
      Tab(0).Control(20)=   "Label2"
      Tab(0).Control(21)=   "Label4"
      Tab(0).Control(22)=   "Label6"
      Tab(0).Control(23)=   "Label7"
      Tab(0).Control(24)=   "Label8"
      Tab(0).Control(25)=   "Label9"
      Tab(0).Control(26)=   "Label10"
      Tab(0).Control(27)=   "Label11"
      Tab(0).Control(28)=   "Label13"
      Tab(0).Control(29)=   "Label15"
      Tab(0).Control(30)=   "Label17"
      Tab(0).Control(31)=   "Label19"
      Tab(0).Control(32)=   "Label21"
      Tab(0).Control(33)=   "Label22"
      Tab(0).ControlCount=   34
      TabCaption(1)   =   "第二頁"
      TabPicture(1)   =   "frm030101_05.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label44"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label42"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label40"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label38"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Label36"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Label35"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Label34"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Label32"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "Label30"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "Label28"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "Label26"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "Label25"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "Label24"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "Label3"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "textCE61"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "textCE49"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "textCE47"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "textCE45"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).Control(18)=   "textCE43"
      Tab(1).Control(18).Enabled=   0   'False
      Tab(1).Control(19)=   "textCE42"
      Tab(1).Control(19).Enabled=   0   'False
      Tab(1).Control(20)=   "textCE41"
      Tab(1).Control(20).Enabled=   0   'False
      Tab(1).Control(21)=   "textCE39"
      Tab(1).Control(21).Enabled=   0   'False
      Tab(1).Control(22)=   "textCE57"
      Tab(1).Control(22).Enabled=   0   'False
      Tab(1).Control(23)=   "textCE25"
      Tab(1).Control(23).Enabled=   0   'False
      Tab(1).Control(24)=   "textCE24"
      Tab(1).Control(24).Enabled=   0   'False
      Tab(1).Control(25)=   "textCE23"
      Tab(1).Control(25).Enabled=   0   'False
      Tab(1).Control(26)=   "textCE39_2"
      Tab(1).Control(26).Enabled=   0   'False
      Tab(1).Control(27)=   "textCE41_1"
      Tab(1).Control(27).Enabled=   0   'False
      Tab(1).Control(28)=   "checkCE38"
      Tab(1).Control(28).Enabled=   0   'False
      Tab(1).Control(29)=   "checkCE58"
      Tab(1).Control(29).Enabled=   0   'False
      Tab(1).Control(30)=   "checkCE40"
      Tab(1).Control(30).Enabled=   0   'False
      Tab(1).Control(31)=   "checkCE60"
      Tab(1).Control(31).Enabled=   0   'False
      Tab(1).Control(32)=   "checkCE44"
      Tab(1).Control(32).Enabled=   0   'False
      Tab(1).Control(33)=   "checkCE46"
      Tab(1).Control(33).Enabled=   0   'False
      Tab(1).Control(34)=   "checkCE48"
      Tab(1).Control(34).Enabled=   0   'False
      Tab(1).Control(35)=   "checkCE50"
      Tab(1).Control(35).Enabled=   0   'False
      Tab(1).Control(36)=   "checkCE62"
      Tab(1).Control(36).Enabled=   0   'False
      Tab(1).Control(37)=   "cmdGoods"
      Tab(1).Control(37).Enabled=   0   'False
      Tab(1).ControlCount=   38
      Begin VB.CommandButton cmdGoods 
         Caption         =   "商品名稱"
         Height          =   315
         Left            =   390
         TabIndex        =   9
         Top             =   3300
         Width           =   1005
      End
      Begin VB.CheckBox checkCE62 
         Height          =   180
         Left            =   120
         TabIndex        =   74
         Top             =   4260
         Width           =   252
      End
      Begin VB.CheckBox checkCE50 
         Height          =   180
         Left            =   120
         TabIndex        =   73
         Top             =   3960
         Width           =   252
      End
      Begin VB.CheckBox checkCE48 
         Height          =   180
         Left            =   120
         TabIndex        =   72
         Top             =   3660
         Width           =   252
      End
      Begin VB.CheckBox checkCE46 
         Height          =   180
         Left            =   120
         TabIndex        =   71
         Top             =   3060
         Width           =   252
      End
      Begin VB.CheckBox checkCE44 
         Height          =   180
         Left            =   120
         TabIndex        =   70
         Top             =   2160
         Width           =   252
      End
      Begin VB.CheckBox checkCE60 
         Height          =   180
         Left            =   120
         TabIndex        =   69
         Top             =   1860
         Width           =   252
      End
      Begin VB.CheckBox checkCE40 
         Height          =   180
         Left            =   120
         TabIndex        =   68
         Top             =   1560
         Width           =   252
      End
      Begin VB.CheckBox checkCE58 
         Height          =   180
         Left            =   120
         TabIndex        =   67
         Top             =   1260
         Width           =   252
      End
      Begin VB.CheckBox checkCE38 
         Height          =   180
         Left            =   120
         TabIndex        =   66
         Top             =   360
         Width           =   252
      End
      Begin VB.CheckBox checkCE65 
         Height          =   180
         Left            =   -74880
         TabIndex        =   65
         Top             =   3960
         Width           =   252
      End
      Begin VB.CheckBox checkCE22 
         Height          =   180
         Left            =   -74880
         TabIndex        =   64
         Top             =   3660
         Width           =   252
      End
      Begin VB.CheckBox checkCE52 
         Height          =   180
         Left            =   -74880
         TabIndex        =   63
         Top             =   3360
         Width           =   252
      End
      Begin VB.CheckBox checkCE54 
         Height          =   180
         Left            =   -74880
         TabIndex        =   62
         Top             =   3060
         Width           =   252
      End
      Begin VB.CheckBox checkCE56 
         Height          =   180
         Left            =   -74880
         TabIndex        =   61
         Top             =   2760
         Width           =   252
      End
      Begin VB.CheckBox checkCE16 
         Height          =   180
         Left            =   -74880
         TabIndex        =   60
         Top             =   960
         Width           =   252
      End
      Begin VB.CheckBox checkCE03 
         Height          =   180
         Left            =   -74880
         TabIndex        =   59
         Top             =   660
         Width           =   252
      End
      Begin VB.CheckBox checkCE09 
         Height          =   180
         Left            =   -74880
         TabIndex        =   58
         Top             =   360
         Width           =   252
      End
      Begin MSForms.TextBox textCE41_1 
         Height          =   915
         Left            =   1680
         TabIndex        =   77
         Top             =   2160
         Width           =   7035
         VariousPropertyBits=   -1475330021
         ScrollBars      =   2
         Size            =   "12409;1623"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE39_2 
         Height          =   285
         Left            =   3000
         TabIndex        =   76
         TabStop         =   0   'False
         Top             =   1560
         Width           =   5712
         VariousPropertyBits=   671105051
         Size            =   "10075;503"
         BorderColor     =   16777215
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE04_2 
         Height          =   285
         Left            =   -72000
         TabIndex        =   75
         TabStop         =   0   'False
         Top             =   360
         Width           =   5772
         VariousPropertyBits=   671105055
         Size            =   "10181;503"
         BorderColor     =   16777215
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE04 
         Height          =   285
         Left            =   -73320
         TabIndex        =   28
         Top             =   360
         Width           =   1212
         VariousPropertyBits=   671105051
         MaxLength       =   9
         Size            =   "2138;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE02 
         Height          =   285
         Left            =   -73320
         TabIndex        =   27
         Top             =   660
         Width           =   1212
         VariousPropertyBits=   671105051
         MaxLength       =   8
         Size            =   "2138;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE10 
         Height          =   285
         Left            =   -73320
         TabIndex        =   26
         Top             =   960
         Width           =   7092
         VariousPropertyBits=   671105051
         MaxLength       =   50
         Size            =   "12509;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE11 
         Height          =   285
         Left            =   -73320
         TabIndex        =   25
         Top             =   1260
         Width           =   7092
         VariousPropertyBits=   671105051
         MaxLength       =   80
         Size            =   "12509;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE12 
         Height          =   285
         Left            =   -73320
         TabIndex        =   24
         Top             =   1560
         Width           =   7092
         VariousPropertyBits=   671105051
         MaxLength       =   40
         Size            =   "12509;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE13 
         Height          =   285
         Left            =   -73320
         TabIndex        =   23
         Top             =   1860
         Width           =   7092
         VariousPropertyBits=   671105051
         MaxLength       =   50
         Size            =   "12509;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE14 
         Height          =   285
         Left            =   -73320
         TabIndex        =   22
         Top             =   2160
         Width           =   7092
         VariousPropertyBits=   671105051
         MaxLength       =   80
         Size            =   "12509;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE15 
         Height          =   285
         Left            =   -73320
         TabIndex        =   21
         Top             =   2460
         Width           =   7092
         VariousPropertyBits=   671105051
         MaxLength       =   40
         Size            =   "12509;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE17 
         Height          =   285
         Left            =   -73320
         TabIndex        =   20
         Top             =   3660
         Width           =   7092
         VariousPropertyBits=   671105051
         MaxLength       =   60
         Size            =   "12509;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE63 
         Height          =   285
         Left            =   -73320
         TabIndex        =   19
         Top             =   3960
         Width           =   7092
         VariousPropertyBits=   671105051
         MaxLength       =   60
         Size            =   "12509;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE64 
         Height          =   285
         Left            =   -73320
         TabIndex        =   18
         Top             =   4260
         Width           =   7092
         VariousPropertyBits=   671105051
         MaxLength       =   60
         Size            =   "12509;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE23 
         Height          =   285
         Left            =   1680
         TabIndex        =   17
         Top             =   360
         Width           =   7032
         VariousPropertyBits=   671105051
         MaxLength       =   70
         Size            =   "12404;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE24 
         Height          =   285
         Left            =   1680
         TabIndex        =   16
         Top             =   660
         Width           =   7032
         VariousPropertyBits=   671105051
         MaxLength       =   150
         Size            =   "12404;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE25 
         Height          =   285
         Left            =   1680
         TabIndex        =   15
         Top             =   960
         Width           =   7032
         VariousPropertyBits=   671105051
         MaxLength       =   70
         Size            =   "12404;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE57 
         Height          =   285
         Left            =   1680
         TabIndex        =   14
         Top             =   1260
         Width           =   1212
         VariousPropertyBits=   671105051
         MaxLength       =   20
         Size            =   "2138;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE39 
         Height          =   285
         Left            =   1680
         TabIndex        =   13
         Top             =   1560
         Width           =   1212
         VariousPropertyBits=   671105051
         MaxLength       =   1
         Size            =   "2138;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE41 
         Height          =   285
         Left            =   1680
         TabIndex        =   12
         Top             =   2160
         Width           =   7032
         VariousPropertyBits=   671105051
         MaxLength       =   160
         Size            =   "12404;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE42 
         Height          =   285
         Left            =   1680
         TabIndex        =   11
         Top             =   2460
         Width           =   7032
         VariousPropertyBits=   671105051
         MaxLength       =   180
         Size            =   "12404;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE43 
         Height          =   285
         Left            =   1680
         TabIndex        =   10
         Top             =   2760
         Width           =   7032
         VariousPropertyBits=   671105051
         MaxLength       =   160
         Size            =   "12404;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE45 
         Height          =   285
         Left            =   1680
         TabIndex        =   8
         Top             =   3060
         Width           =   7032
         VariousPropertyBits=   671105051
         MaxLength       =   200
         Size            =   "12404;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE47 
         Height          =   285
         Left            =   1680
         TabIndex        =   7
         Top             =   3660
         Width           =   7032
         VariousPropertyBits=   671105051
         MaxLength       =   395
         Size            =   "12404;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE49 
         Height          =   285
         Left            =   1680
         TabIndex        =   6
         Top             =   3960
         Width           =   7032
         VariousPropertyBits=   671105051
         MaxLength       =   699
         Size            =   "12404;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE61 
         Height          =   285
         Left            =   1680
         TabIndex        =   5
         Top             =   4260
         Width           =   7032
         VariousPropertyBits=   671105051
         MaxLength       =   2000
         Size            =   "12404;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label3 
         Caption         =   "案件名稱 :"
         Height          =   252
         Left            =   360
         TabIndex        =   78
         Top             =   2160
         Width           =   1212
      End
      Begin VB.Label Label2 
         Caption         =   "申請人 :"
         Height          =   252
         Left            =   -74640
         TabIndex        =   55
         Top             =   360
         Width           =   732
      End
      Begin VB.Label Label4 
         Caption         =   "申請日 :"
         Height          =   252
         Left            =   -74640
         TabIndex        =   54
         Top             =   660
         Width           =   732
      End
      Begin VB.Label Label6 
         Caption         =   "代表人1(中) :"
         Height          =   252
         Left            =   -74640
         TabIndex        =   53
         Top             =   960
         Width           =   1092
      End
      Begin VB.Label Label7 
         Caption         =   "代表人1(英) :"
         Height          =   252
         Left            =   -74640
         TabIndex        =   52
         Top             =   1260
         Width           =   1092
      End
      Begin VB.Label Label8 
         Caption         =   "代表人1(日) :"
         Height          =   252
         Left            =   -74640
         TabIndex        =   51
         Top             =   1560
         Width           =   1092
      End
      Begin VB.Label Label9 
         Caption         =   "代表人2(中) :"
         Height          =   252
         Left            =   -74640
         TabIndex        =   50
         Top             =   1860
         Width           =   1092
      End
      Begin VB.Label Label10 
         Caption         =   "代表人2(英) :"
         Height          =   252
         Left            =   -74640
         TabIndex        =   49
         Top             =   2160
         Width           =   1092
      End
      Begin VB.Label Label11 
         Caption         =   "代表人2(日) :"
         Height          =   252
         Left            =   -74640
         TabIndex        =   48
         Top             =   2460
         Width           =   1092
      End
      Begin VB.Label Label13 
         Caption         =   "代理人 :"
         Height          =   252
         Left            =   -74640
         TabIndex        =   47
         Top             =   2760
         Width           =   732
      End
      Begin VB.Label Label15 
         Caption         =   "代表人印鑑 :"
         Height          =   252
         Left            =   -74640
         TabIndex        =   46
         Top             =   3060
         Width           =   1092
      End
      Begin VB.Label Label17 
         Caption         =   "申請人印鑑 :"
         Height          =   252
         Left            =   -74640
         TabIndex        =   45
         Top             =   3360
         Width           =   1092
      End
      Begin VB.Label Label19 
         Caption         =   "申請人中譯文 :"
         Height          =   252
         Left            =   -74640
         TabIndex        =   44
         Top             =   3660
         Width           =   1332
      End
      Begin VB.Label Label21 
         Caption         =   "代表人1中譯文 :"
         Height          =   252
         Left            =   -74640
         TabIndex        =   43
         Top             =   3960
         Width           =   1332
      End
      Begin VB.Label Label22 
         Caption         =   "代表人2中譯文 :"
         Height          =   252
         Left            =   -74640
         TabIndex        =   42
         Top             =   4260
         Width           =   1332
      End
      Begin VB.Label Label24 
         Caption         =   "申請地址(中) :"
         Height          =   252
         Left            =   360
         TabIndex        =   41
         Top             =   360
         Width           =   1212
      End
      Begin VB.Label Label25 
         Caption         =   "申請地址(英) :"
         Height          =   252
         Left            =   360
         TabIndex        =   40
         Top             =   660
         Width           =   1212
      End
      Begin VB.Label Label26 
         Caption         =   "申請地址(日) :"
         Height          =   252
         Left            =   360
         TabIndex        =   39
         Top             =   960
         Width           =   1212
      End
      Begin VB.Label Label28 
         Caption         =   "正商標號數 :"
         Height          =   252
         Left            =   360
         TabIndex        =   38
         Top             =   1260
         Width           =   1212
      End
      Begin VB.Label Label30 
         Caption         =   "商標種類 :"
         Height          =   252
         Left            =   360
         TabIndex        =   37
         Top             =   1560
         Width           =   1212
      End
      Begin VB.Label Label32 
         Caption         =   "圖樣 :"
         Height          =   252
         Left            =   360
         TabIndex        =   36
         Top             =   1860
         Width           =   612
      End
      Begin VB.Label Label34 
         Caption         =   "案件名稱(中) :"
         Height          =   252
         Left            =   360
         TabIndex        =   35
         Top             =   2160
         Width           =   1212
      End
      Begin VB.Label Label35 
         Caption         =   "案件名稱(英) :"
         Height          =   252
         Left            =   360
         TabIndex        =   34
         Top             =   2460
         Width           =   1212
      End
      Begin VB.Label Label36 
         Caption         =   "案件名稱(日) :"
         Height          =   252
         Left            =   360
         TabIndex        =   33
         Top             =   2760
         Width           =   1212
      End
      Begin VB.Label Label38 
         Caption         =   "縮減商品 :"
         Height          =   252
         Left            =   360
         TabIndex        =   32
         Top             =   3060
         Width           =   1212
      End
      Begin VB.Label Label40 
         Caption         =   "商品類別 :"
         Height          =   255
         Left            =   360
         TabIndex        =   31
         Top             =   3660
         Width           =   1215
      End
      Begin VB.Label Label42 
         Caption         =   "商品群組 :"
         Height          =   255
         Left            =   360
         TabIndex        =   30
         Top             =   3960
         Width           =   1215
      End
      Begin VB.Label Label44 
         Caption         =   "其它 :"
         Height          =   255
         Left            =   360
         TabIndex        =   29
         Top             =   4260
         Width           =   615
      End
   End
   Begin MSForms.TextBox textTMKey 
      Height          =   285
      Left            =   1080
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   540
      Width           =   2532
      VariousPropertyBits=   671105051
      Size            =   "4466;503"
      BorderColor     =   16777215
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textCE01 
      Height          =   285
      Left            =   5280
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   540
      Width           =   2532
      VariousPropertyBits=   671105051
      Size            =   "4466;503"
      BorderColor     =   16777215
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      Caption         =   "本所案號 :"
      Height          =   252
      Index           =   0
      Left            =   120
      TabIndex        =   57
      Top             =   540
      Width           =   852
   End
   Begin VB.Label Label1 
      Caption         =   "收文號 :"
      Height          =   252
      Index           =   2
      Left            =   4440
      TabIndex        =   56
      Top             =   540
      Width           =   732
   End
End
Attribute VB_Name = "frm030101_05"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/08/11 改成Form2.0 ; 全部TextBox
'Memo By Sindy 2012/12/4 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/29 員工編號欄已修改
'Memo By Sindy 2010/8/11 日期欄已修改
Option Explicit

' 本所案號
Dim m_TM01 As String
Dim m_TM02 As String
Dim m_TM03 As String
Dim m_TM04 As String
' 收文號
Dim m_CE01 As String

' 宣告欄位內容結構
Private Type FIELDITEM
   fiName As String
   fiNewData As String
   fiType As Integer
End Type
Dim m_CEList() As FIELDITEM
Dim m_CECount As Integer

' 更新商標基本檔時所使用的變更事項檔欄位的暫存資料
' 申請日
Dim m_CE02 As String
' 申請人
Dim m_CE04 As String
' 商品種類代碼
Dim m_CE39 As String
' 前畫面
Dim m_Parent As String
'Add By Cheng 2003/04/10
Private Type TMFIELDITEM
   tiName As String
   tiData As String
   tiType As String
End Type
Dim m_TMList() As TMFIELDITEM
Dim m_TMListCount As Integer
Private Type SRFIELDITEM
   siName As String
   siData As String
   siType As String
End Type
Dim m_SRList() As SRFIELDITEM
Dim m_SRListCount As Integer
Dim tmpOldCE04 As String
Dim tmpOldCE02 As String
Dim tmpOldCE10 As String
Dim tmpOldCE11 As String
Dim tmpOldCE12 As String
Dim tmpOldCE13 As String
Dim tmpOldCE14 As String
Dim tmpOldCE15 As String
Dim tmpOldCE23 As String
Dim tmpOldCE24 As String
Dim tmpOldCE25 As String
Dim tmpOldCE39 As String
Dim tmpOldCE41 As String
Dim tmpOldCE42 As String
Dim tmpOldCE43 As String
Dim tmpOldCE47 As String
Dim tmpOldCE49 As String
Dim tmpOldCE57 As String
'add by nickc 2007/04/03
Dim m_TM09 As String
Public ChkTG As Boolean
'Add By Sindy 2011/8/3
Dim m_TM23 As String
Dim m_CP31 As String 'Add By Sindy 2011/8/23 是否新案件


' 檢查該欄位是否存在
Private Function IsCEFieldExist(ByVal strField As String) As Boolean
   Dim nIndex As Integer
   IsCEFieldExist = False
   For nIndex = 0 To m_CECount - 1
      If m_CEList(nIndex).fiName = strField Then
         IsCEFieldExist = True
         Exit For
      End If
   Next nIndex
End Function
' 設定欄位新值
Private Sub SetCEFieldData(ByVal strField As String, ByVal strNewData As String, ByVal nType As Integer)
   Dim bFind As Boolean
   Dim nIndex As Integer
   bFind = False
   For nIndex = 0 To m_CECount - 1
      If m_CEList(nIndex).fiName = strField Then
         bFind = True
         m_CEList(nIndex).fiNewData = strNewData
         Exit For
      End If
   Next nIndex
   If bFind = False Then
      ReDim Preserve m_CEList(m_CECount + 1)
      m_CEList(m_CECount).fiName = strField
      m_CEList(m_CECount).fiNewData = strNewData
      m_CEList(m_CECount).fiType = nType
      m_CECount = m_CECount + 1
   End If
End Sub
' 清除欄位串列
Private Sub ClearCEFields()
   Erase m_CEList
   m_CECount = 0
End Sub

' 更新欄位內容
Private Sub UpdateFieldNewData()
   SetCEFieldData "CE01", m_CE01, 0
   If checkCE03.Value = 1 Then
      SetCEFieldData "CE02", DBDATE(textCE02), 1
   End If
   If checkCE09.Value = 1 Then
      'Modify by Sindy 2013/1/22
      'SetCEFieldData "CE04", textCE04 & String(9 - Len(textCE04), "0"), 0
      textCE04 = textCE04 & String(9 - Len(textCE04), "0")
      SetCEFieldData "CE04", textCE04, 0
      '2013/1/22 End
   End If
   If checkCE16.Value = 1 Then
      SetCEFieldData "CE10", textCE10, 0
      SetCEFieldData "CE11", textCE11, 0
      SetCEFieldData "CE12", textCE12, 0
      SetCEFieldData "CE13", textCE13, 0
      SetCEFieldData "CE14", textCE14, 0
      SetCEFieldData "CE15", textCE15, 0
   End If
   If checkCE56.Value = 1 Then
      SetCEFieldData "CE55", "V", 0
   End If
   If checkCE54.Value = 1 Then
      SetCEFieldData "CE53", "V", 0
   End If
   If checkCE52.Value = 1 Then
      SetCEFieldData "CE51", "V", 0
   End If
   If checkCE22.Value = 1 Then
      SetCEFieldData "CE17", textCE17, 0
   End If
   If checkCE65.Value = 1 Then
      SetCEFieldData "CE63", textCE63, 0
      SetCEFieldData "CE64", textCE64, 0
   End If
   If checkCE38.Value = 1 Then
      SetCEFieldData "CE23", textCE23, 0
      SetCEFieldData "CE24", textCE24, 0
      SetCEFieldData "CE25", textCE25, 0
   End If
   If checkCE58.Value = 1 Then
      SetCEFieldData "CE57", textCE57, 0
   End If
   If checkCE40.Value = 1 Then
      SetCEFieldData "CE39", textCE39, 0
   End If
   If checkCE60.Value = 1 Then
      SetCEFieldData "CE59", "V", 0
   End If
   If checkCE44.Value = 1 Then
        Select Case m_TM01
        Case "T", "FCT", "CFT", "TF", "S"
            SetCEFieldData "CE41", textCE41_1, 0
        Case Else
            SetCEFieldData "CE41", textCE41, 0
            SetCEFieldData "CE42", textCE42, 0
            SetCEFieldData "CE43", textCE43, 0
        End Select
   End If
   If checkCE46.Value = 1 Then
      SetCEFieldData "CE45", textCE45, 0
   End If
   If checkCE48.Value = 1 Then
      SetCEFieldData "CE47", textCE47, 0
   End If
   If checkCE50.Value = 1 Then
      SetCEFieldData "CE49", textCE49, 0
   End If
   If checkCE62.Value = 1 Then
      SetCEFieldData "CE61", textCE61, 0
   End If
End Sub

Private Sub cmdCancel_Click()
   Unload Me
   frm02010401_4.Show
End Sub

Private Sub checkCE38_Click()
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
    
    'Add By Cheng 2003/04/17
    '若有勾選變更申請地址, 且未勾選變更申請人
    If Me.checkCE38.Value = vbChecked And Me.checkCE09.Value = vbUnchecked Then
        StrSQLa = "Select TM23 From Trademark Where " & ChgTradeMark(Replace(textTMKey.Text, "-", ""))
        StrSQLa = StrSQLa & " union Select SP08 From Servicepractice Where " & ChgService(Replace(textTMKey.Text, "-", ""))
        rsA.CursorLocation = adUseClient
        rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
        If rsA.RecordCount > 0 Then
            '顯示申請人地址
            textCE23.Text = PUB_GetCustEachAdd("" & rsA.Fields(0).Value, "1")
            textCE24.Text = PUB_GetCustEachAdd("" & rsA.Fields(0).Value, "2")
            textCE25.Text = PUB_GetCustEachAdd("" & rsA.Fields(0).Value, "3")
        End If
        If rsA.State <> adStateClosed Then rsA.Close
        Set rsA = Nothing
    '若有勾選變更申請地址, 同時勾選變更申請人
    ElseIf Me.checkCE38.Value = vbChecked And Me.checkCE09.Value = vbChecked Then
        '顯示申請人地址
        textCE23.Text = PUB_GetCustEachAdd(ChangeCustomerL(Me.textCE04.Text), "1")
        textCE24.Text = PUB_GetCustEachAdd(ChangeCustomerL(Me.textCE04.Text), "2")
        textCE25.Text = PUB_GetCustEachAdd(ChangeCustomerL(Me.textCE04.Text), "3")
    End If
End Sub

Private Sub cmdExit_Click()
    Me.Enabled = False
    'Add By Cheng 2004/04/08
'move to unload by nick 2004/10/22
'    '列印接洽接案單
'    PUB_PrintCaseCloseSheet strUserNum
'    '刪除暫存資料
'    PUB_DeleteCaseCloseSheet strUserNum
   Select Case m_Parent
      Case "frm030101_03":
         Unload frm030101_03
      Case "frm030101_06":
         Unload frm030101_06
      Case "frm030101_07"
         Unload frm030101_07
      Case "frm030101_08"
         Unload frm030101_08
      Case "frm030101_09"
         Unload frm030101_09
      Case "frm030101_10"
         Unload frm030101_10
      Case Else
   End Select
   Unload frm030101_01
   Unload Me
End Sub

Private Sub cmdGoods_Click()
frm03010303_04.Hide
Set frm03010303_04.UpForm = Me
frm03010303_04.TGKey = m_TM01 & "-" & m_TM02 & "-" & m_TM03 & "-" & m_TM04
frm03010303_04.AllClass = m_TM09
frm03010303_04.cmdOK(2).Visible = True
Me.Hide
frm03010303_04.QueryData
frm03010303_04.Show vbModal 'Modify By Sindy 2009/09/17 改為強制回應表單
End Sub

Private Sub cmdOK_Click()
   If CheckDataValid = True Then
      'Add By Cheng 2002/05/23
      '重新檢查欄位有效性
      If TxtValidate = False Then Exit Sub
      ' 設定滑鼠游標為等待狀態
      Screen.MousePointer = vbHourglass
      ' 儲存資料
      UpdateFieldNewData
      'edit by nick 2004/11/03
      'OnSaveData
      If OnSaveData = False Then MsgBox "存檔失敗，請洽系統管理員 !", vbCritical: Screen.MousePointer = vbDefault: Exit Sub
      ' 設定滑鼠游標為預設
      Screen.MousePointer = vbDefault
      Unload Me
      'Add By Sindy 2012/4/17 有變更事項資料時,必須按確定鍵
      Select Case m_Parent
         Case "frm030101_03":
            frm030101_03.m_blnClkChgButton = True
            frm030101_03.Show
            frm030101_03.QueryData
         Case "frm030101_06":
            frm030101_06.m_blnClkChgButton = True
            frm030101_06.Show
            frm030101_06.QueryData
         Case "frm030101_07"
            frm030101_07.m_blnClkChgButton = True
            frm030101_07.Show
            frm030101_07.QueryData
         Case "frm030101_08"
            frm030101_08.m_blnClkChgButton = True
            frm030101_08.Show
            frm030101_08.QueryData
         Case "frm030101_09"
            frm030101_09.m_blnClkChgButton = True
            frm030101_09.Show
            frm030101_09.QueryData
         Case "frm030101_10"
            frm030101_10.m_blnClkChgButton = True
            frm030101_10.Show
            frm030101_10.QueryData
         Case Else
      End Select
   End If
End Sub

Private Sub Form_Load()
   textTMKey.BackColor = &H8000000F
   textCE01.BackColor = &H8000000F
   textCE04_2.BackColor = &H8000000F
   textCE39_2.BackColor = &H8000000F
   
    'Added by Lydia 2016/09/10 設定代表人中文名稱和英文名稱長度
    textCE10.MaxLength = Pub_MaxCEL10
    textCE11.MaxLength = Pub_MaxCEL11
    textCE13.MaxLength = Pub_MaxCEL10
    textCE14.MaxLength = Pub_MaxCEL11
    'end 2016/09/10
    
   tabCtrl.Tab = 0
   
   MoveFormToCenter Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ClearCEFields
   'Add By Cheng 2002/07/19
   Set frm030101_05 = Nothing
End Sub

' 由客戶代碼取得客戶名稱
Private Function GetCustomer(ByVal strData As String) As String
   Dim rsTmp As ADODB.Recordset
   Dim strKey As String
   Dim strSql As String
   
   GetCustomer = Empty
   If IsEmptyText(strData) = False Then
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
            GetCustomer = rsTmp.Fields("CU04")
         ElseIf IsNull(rsTmp.Fields("CU05")) = False Then
            GetCustomer = rsTmp.Fields("CU05")
         ElseIf IsNull(rsTmp.Fields("CU06")) = False Then
            GetCustomer = rsTmp.Fields("CU06")
         End If
      End If
      rsTmp.Close
   End If
   Set rsTmp = Nothing
End Function

' 由客戶代號取得地址
' Input : strData ==> 客戶代號
'         nType ==> 種類
'                   0 : 表要取得的是中文地址
'                   1 : 表要取得的是英文地址
'                   2 : 表要取得的是日文地址
Private Function GetAddress(ByVal strData As String, ByVal nType As Integer) As String
   Dim rsTmp As New ADODB.Recordset
   Dim strKey As String
   Dim strSql As String
   
   GetAddress = Empty
   If IsEmptyText(strData) = False Then
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
         Select Case nType
            Case 0:
               If IsNull(rsTmp.Fields("CU23")) = False Then
                  GetAddress = rsTmp.Fields("CU23")
               End If
            Case 1:
               If IsNull(rsTmp.Fields("CU24")) = False Then
                  GetAddress = rsTmp.Fields("CU24")
               End If
            Case 2:
               If IsNull(rsTmp.Fields("CU29")) = False Then
                  GetAddress = rsTmp.Fields("CU29")
               End If
         End Select
      End If
      rsTmp.Close
   End If
   Set rsTmp = Nothing
End Function

Public Sub SetData(ByVal nType As Integer, ByVal strData As String, Optional ByVal bClear As Boolean = False)
   ' 清除搜尋的Key
   If bClear = True Then
      m_TM01 = Empty
      m_TM02 = Empty
      m_TM03 = Empty
      m_TM04 = Empty
      m_CE01 = Empty
   End If
   
   Select Case nType
      ' 本所案號 欄位1
      Case 0: m_TM01 = strData
      ' 本所案號 欄位2
      Case 1: m_TM02 = strData
      ' 本所案號 欄位3
      Case 2: m_TM03 = strData
      ' 本所案號 欄位4
      Case 3: m_TM04 = strData
      ' 收文號
      Case 4: m_CE01 = strData
   End Select
End Sub

Public Sub SetParent(ByVal strParent As String)
   m_Parent = strParent
End Sub

Private Sub QueryCaseProgress()
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   strSql = "SELECT * FROM CaseProgress " & _
            "WHERE CP09 = '" & m_CE01 & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      'Add By Sindy 2011/8/23 是否新案件
      m_CP31 = ""
      If IsNull(rsTmp.Fields("CP31")) = False Then
         m_CP31 = rsTmp.Fields("CP31")
      End If
   End If
End Sub

Public Sub QueryData()
Dim rsTmp As New ADODB.Recordset
Dim strSql As String
'Add By Cheng 2003/09/03
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
   
   ' 清除欄位串列
   'ClearFields
   
   ' 清除暫存變數
   m_CE02 = Empty
   m_CE04 = Empty
   
   tmpOldCE04 = Empty
   tmpOldCE02 = Empty
   tmpOldCE10 = Empty
   tmpOldCE11 = Empty
   tmpOldCE12 = Empty
   tmpOldCE13 = Empty
   tmpOldCE14 = Empty
   tmpOldCE15 = Empty
   tmpOldCE23 = Empty
   tmpOldCE24 = Empty
   tmpOldCE25 = Empty
   tmpOldCE39 = Empty
   tmpOldCE41 = Empty
   tmpOldCE42 = Empty
   tmpOldCE43 = Empty
   tmpOldCE47 = Empty
   tmpOldCE49 = Empty
   tmpOldCE57 = Empty
    'Add By Cheng 2003/11/11
    Select Case m_TM01
    Case "T", "FCT", "CFT", "TF", "S"
        Me.Label34.Visible = False
        Me.textCE41.Visible = False
        Me.textCE41.Enabled = False
        Me.Label35.Visible = False
        Me.textCE42.Visible = False
        Me.textCE42.Enabled = False
        Me.Label36.Visible = False
        Me.textCE43.Visible = False
        Me.textCE43.Enabled = False
    Case Else
        Me.Label3.Visible = False
        Me.textCE41_1.Visible = False
        Me.textCE41_1.Enabled = False
    End Select
    'End
   textTMKey = m_TM01 & m_TM02 & m_TM03 & m_TM04
   textCE01 = m_CE01
    'Add By Cheng 2003/09/03
    'Begin
    'edit by nickc  2007/04/03
    m_TM09 = Empty
    'StrSQLa = "Select TM11, TM23, TM47, TM48, TM49, TM50, TM51, TM52, TM08, TM05, TM06, TM07, TM09, TM32, TM27 From Trademark Where " & ChgTradeMark(textTMKey)
    'StrSQLa = StrSQLa & " Union Select SP10, SP08, SP42, '', '', '', '', '', '', SP05, SP06, SP07, '', '', '' From Servicepractice Where " & ChgService(textTMKey)
    StrSQLa = "Select TM11, TM23, TM47, TM48, TM49, TM50, TM51, TM52, TM08, TM05, TM06, TM07, TM09, TM32, TM27,tm09,TM24,TM25,TM26 From Trademark Where " & ChgTradeMark(textTMKey)
    StrSQLa = StrSQLa & " Union Select SP10, SP08, SP42, '', '', '', '', '', '', SP05, SP06, SP07, '', '', '',sp73,'','','' From Servicepractice Where " & ChgService(textTMKey)
    rsA.CursorLocation = adUseClient
    rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
    If rsA.RecordCount > 0 Then
      ' 申請日
        tmpOldCE02 = "" & rsA.Fields(0).Value
      ' 申請人
        m_TM23 = "" & rsA.Fields(1).Value 'Add By Sindy 2011/8/3
        tmpOldCE04 = "" & rsA.Fields(1).Value
      '申請人地址
'        tmpOldCE23 = PUB_GetCustEachAdd(tmpOldCE04, "1")
'        tmpOldCE24 = PUB_GetCustEachAdd(tmpOldCE04, "2")
'        tmpOldCE25 = PUB_GetCustEachAdd(tmpOldCE04, "3")
        'Modify By Sindy 2011/2/1
        tmpOldCE23 = Trim("" & rsA.Fields(16).Value)
        tmpOldCE24 = Trim("" & rsA.Fields(17).Value)
        tmpOldCE25 = Trim("" & rsA.Fields(18).Value)
        '2011/2/1 End
      ' 代表人
        tmpOldCE10 = "" & rsA.Fields(2).Value
        tmpOldCE11 = "" & rsA.Fields(3).Value
        tmpOldCE12 = "" & rsA.Fields(4).Value
        tmpOldCE13 = "" & rsA.Fields(5).Value
        tmpOldCE14 = "" & rsA.Fields(6).Value
        tmpOldCE15 = "" & rsA.Fields(7).Value
      ' 專利商標種類代號
         tmpOldCE39 = "" & rsA.Fields(8).Value
      ' 案件名稱
        tmpOldCE41 = "" & rsA.Fields(9).Value
        tmpOldCE42 = "" & rsA.Fields(10).Value
        tmpOldCE43 = "" & rsA.Fields(11).Value
      ' 商品類別
        tmpOldCE47 = "" & rsA.Fields(12).Value
      ' 商品群組
        tmpOldCE49 = "" & rsA.Fields(13).Value
      ' 正商標號數
        tmpOldCE57 = "" & rsA.Fields(14).Value
        'add by nickc 2007/04/03
        m_TM09 = "" & rsA.Fields(15).Value
    End If
    If rsA.State <> adStateClosed Then rsA.Close
    Set rsA = Nothing
    'End
   
   'Modify By Sindy 2012/5/18 Mark暫存變數，因暫存變數是為比對基本檔和畫面上的欄位值是否相同,所以暫存變數不須再存取變更檔資料
   strSql = "SELECT * FROM ChangeEvent " & _
            "WHERE CE01 = '" & m_CE01 & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      ' 申請日
      If IsNull(rsTmp.Fields("CE02")) = False Then
         checkCE03.Value = 1 'Add By Sindy 2012/3/7
         m_CE02 = rsTmp.Fields("CE02")
'         tmpOldCE02 = CheckStr(rsTmp.Fields("CE02"))
         textCE02 = ChangeWStringToTString(rsTmp.Fields("CE02"))
      End If
      ' 申請人
      If IsNull(rsTmp.Fields("CE04")) = False Then
         checkCE09.Value = 1 'Add By Sindy 2012/3/7
         m_CE04 = rsTmp.Fields("CE04")
         textCE04 = rsTmp.Fields("CE04")
'         tmpOldCE04 = textCE04.Text
         textCE04_2 = GetCustomer(rsTmp.Fields("CE04"))
         '顯示申請人地址
         textCE23.Text = PUB_GetCustEachAdd(Me.textCE04.Text, "1")
         textCE24.Text = PUB_GetCustEachAdd(Me.textCE04.Text, "2")
         textCE25.Text = PUB_GetCustEachAdd(Me.textCE04.Text, "3")
'         tmpOldCE23 = textCE23.Text
'         tmpOldCE24 = textCE24.Text
'         tmpOldCE25 = textCE25.Text
      End If
      'If IsNull(rsTmp.Fields("CE09")) = False Then
      '   textCE09 = rsTmp.Fields("CE09")
      'End If
      ' 代表人
      If IsNull(rsTmp.Fields("CE10")) = False Then
         checkCE16.Value = 1 'Add By Sindy 2012/3/7
         textCE10 = rsTmp.Fields("CE10")
'         tmpOldCE10 = textCE10.Text
      End If
      If IsNull(rsTmp.Fields("CE11")) = False Then
         checkCE16.Value = 1 'Add By Sindy 2012/3/7
         textCE11 = rsTmp.Fields("CE11")
'         tmpOldCE11 = textCE11.Text
      End If
      If IsNull(rsTmp.Fields("CE12")) = False Then
         checkCE16.Value = 1 'Add By Sindy 2012/3/7
         textCE12 = rsTmp.Fields("CE12")
'         tmpOldCE12 = textCE12.Text
      End If
      If IsNull(rsTmp.Fields("CE13")) = False Then
         checkCE16.Value = 1 'Add By Sindy 2012/3/7
         textCE13 = rsTmp.Fields("CE13")
'         tmpOldCE13 = textCE13.Text
      End If
      If IsNull(rsTmp.Fields("CE14")) = False Then
         checkCE16.Value = 1 'Add By Sindy 2012/3/7
         textCE14 = rsTmp.Fields("CE14")
'         tmpOldCE14 = textCE14.Text
      End If
      If IsNull(rsTmp.Fields("CE15")) = False Then
         checkCE16.Value = 1 'Add By Sindy 2012/3/7
         textCE15 = rsTmp.Fields("CE15")
'         tmpOldCE15 = textCE15.Text
      End If
      ' 申請地址
      If IsNull(rsTmp.Fields("CE23")) = False Then
         checkCE38.Value = 1 'Add By Sindy 2012/3/7
         textCE23 = rsTmp.Fields("CE23")
'         tmpOldCE23 = textCE23.Text
      End If
      If IsNull(rsTmp.Fields("CE24")) = False Then
         checkCE38.Value = 1 'Add By Sindy 2012/3/7
         textCE24 = rsTmp.Fields("CE24")
'         tmpOldCE24 = textCE24.Text
      End If
      If IsNull(rsTmp.Fields("CE25")) = False Then
         checkCE38.Value = 1 'Add By Sindy 2012/3/7
         textCE25 = rsTmp.Fields("CE25")
'         tmpOldCE25 = textCE25.Text
      End If
      ' 專利商標種類代號
      'Add By Cheng 2002/07/18
      m_CE39 = Empty
      If IsNull(rsTmp.Fields("CE39")) = False Then
         checkCE40.Value = 1 'Add By Sindy 2012/3/7
         m_CE39 = rsTmp.Fields("CE39")
         textCE39 = rsTmp.Fields("CE39")
         If IsEmptyText(textCE39) = False Then: textCE39_Validate (False)
'         tmpOldCE39 = textCE39.Text
      End If
        Select Case m_TM01
        Case "T", "FCT", "CFT", "TF", "S"
            ' 案件名稱
            If IsNull(rsTmp.Fields("CE41")) = False Then
               checkCE44.Value = 1 'Add By Sindy 2012/3/7
               textCE41_1 = rsTmp.Fields("CE41")
'               tmpOldCE41 = textCE41_1.Text
            End If
        Case Else
            ' 案件名稱
            If IsNull(rsTmp.Fields("CE41")) = False Then
               checkCE44.Value = 1 'Add By Sindy 2012/3/7
               textCE41 = rsTmp.Fields("CE41")
'               tmpOldCE41 = textCE41.Text
            End If
            If IsNull(rsTmp.Fields("CE42")) = False Then
               checkCE44.Value = 1 'Add By Sindy 2012/3/7
               textCE42 = rsTmp.Fields("CE42")
'               tmpOldCE42 = textCE42.Text
            End If
            If IsNull(rsTmp.Fields("CE43")) = False Then
               checkCE44.Value = 1 'Add By Sindy 2012/3/7
               textCE43 = rsTmp.Fields("CE43")
'               tmpOldCE43 = textCE43.Text
            End If
        End Select
      ' 縮減商品
      If IsNull(rsTmp.Fields("CE45")) = False Then
         checkCE46.Value = 1 'Add By Sindy 2012/3/7
         textCE45 = rsTmp.Fields("CE45")
      End If
      ' 商品類別
      If IsNull(rsTmp.Fields("CE47")) = False Then
         checkCE48.Value = 1 'Add By Sindy 2012/3/7
         textCE47 = rsTmp.Fields("CE47")
'         tmpOldCE47 = textCE47.Text
      End If
      ' 商品群組
      If IsNull(rsTmp.Fields("CE49")) = False Then
         checkCE50.Value = 1 'Add By Sindy 2012/3/7
         textCE49 = rsTmp.Fields("CE49")
'         tmpOldCE49 = textCE49.Text
      End If
      ' 申請人印鑑
      If IsNull(rsTmp.Fields("CE51")) = False Then
         checkCE52.Value = 1 'Add By Sindy 2012/3/7
         'textCE51 = rsTmp.Fields("CE51")
      End If
      ' 代表人印鑑
      If IsNull(rsTmp.Fields("CE53")) = False Then
         checkCE54.Value = 1 'Add By Sindy 2012/3/7
         'textCE53 = rsTmp.Fields("CE53")
      End If
      ' 代理人
      If IsNull(rsTmp.Fields("CE55")) = False Then
         checkCE56.Value = 1 'Add By Sindy 2012/3/7
         'textCE55 = rsTmp.Fields("CE55")
      End If
      ' 正商標號數
      If IsNull(rsTmp.Fields("CE57")) = False Then
         checkCE58.Value = 1 'Add By Sindy 2012/3/7
         textCE57 = rsTmp.Fields("CE57")
'         tmpOldCE57 = textCE57.Text
      End If
      ' 圖樣
      If IsNull(rsTmp.Fields("CE59")) = False Then
         checkCE60.Value = 1 'Add By Sindy 2012/3/7
         'textCE59 = rsTmp.Fields("CE59")
      End If
      ' 其它
      If IsNull(rsTmp.Fields("CE61")) = False Then
         checkCE62.Value = 1 'Add By Sindy 2012/3/7
         textCE61 = rsTmp.Fields("CE61")
      End If
      ' 代表人譯文
      If IsNull(rsTmp.Fields("CE63")) = False Then
         checkCE65.Value = 1 'Add By Sindy 2012/3/7
         textCE63 = rsTmp.Fields("CE63")
      End If
      If IsNull(rsTmp.Fields("CE64")) = False Then
         checkCE65.Value = 1 'Add By Sindy 2012/3/7
         textCE64 = rsTmp.Fields("CE64")
      End If
   End If
   rsTmp.Close
   Set rsTmp = Nothing
   
   Call QueryCaseProgress 'Add By Sindy 2011/8/23
End Sub

'edit b nick 2004/11/03
'Private sub OnSaveData()
Public Function OnSaveData() As Boolean
OnSaveData = True
   Dim strSql As String
   Dim strTmp As String
   Dim nIndex As Integer
   Dim bDifference As Boolean
   Dim bFirst As Boolean
   Dim rsTmp As New ADODB.Recordset
   
   
'911106 nick transation
On Error GoTo CheckingErr
cnnConnection.BeginTrans

   ' 先刪除掉已存在的資料
   strSql = "SELECT * FROM ChangeEvent " & _
            "WHERE CE01 = '" & m_CE01 & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      rsTmp.Close
      strSql = "DELETE FROM ChangeEvent " & _
               "WHERE CE01 = '" & m_CE01 & "' "
      cnnConnection.Execute strSql
   Else
      rsTmp.Close
   End If

   ' 新增一筆資料到變更事項檔
   bFirst = True
   bDifference = False
   strSql = "INSERT INTO ChangeEvent ("
   For nIndex = 0 To m_CECount - 1
      strTmp = m_CEList(nIndex).fiName
      If IsEmptyText(strTmp) = False And IsEmptyText(m_CEList(nIndex).fiNewData) = False Then
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
   For nIndex = 0 To m_CECount - 1
      strTmp = Empty
      If m_CEList(nIndex).fiType = 0 Then
         'Modify by Morgan 2005/1/25 避免單引號錯誤
         'If IsEmptyText(m_CEList(nIndex).fiNewData) = False Then: strTmp = "'" & m_CEList(nIndex).fiNewData & "'"
         If IsEmptyText(m_CEList(nIndex).fiNewData) = False Then: strTmp = "'" & ChgSQL(m_CEList(nIndex).fiNewData) & "'"
      Else
         strTmp = m_CEList(nIndex).fiNewData
      End If
      If IsEmptyText(m_CEList(nIndex).fiName) = False And IsEmptyText(m_CEList(nIndex).fiNewData) = False Then
         If bFirst = True Then
            'Modify by Morgan 2005/1/25 有錯，還原，移到上面控制
            ''edit by nick 2004/12/01  解單引號錯誤
            ''StrSql = StrSql & strTmp
            'StrSql = StrSql & ChgSQL(strTmp)
            strSql = strSql & strTmp
            bFirst = False
         Else
            'Modify by Morgan 2005/1/25 有錯，還原，移到上面控制
            ''edit by nick 2004/12/01  解單引號錯誤
            ''StrSql = StrSql & "," & strTmp
            'StrSql = StrSql & "," & ChgSQL(strTmp)
            strSql = strSql & "," & strTmp
         End If
      End If
   Next nIndex
   strSql = strSql & ")"
   
   If bDifference = True Then
      cnnConnection.Execute strSql
   End If
    'Add By Cheng 2003/04/10
    Select Case m_TM01
    Case "CFT", "FCT", "T", "TF"
        If OnSaveTrademark = False Then GoTo CheckingErr
    Case Else
        If OnSaveServicePractice = False Then GoTo CheckingErr
    End Select
       
    Set rsTmp = Nothing
'911106 nick transation
    cnnConnection.CommitTrans
    Exit Function
CheckingErr:
    MsgBox (Err.Description)
     cnnConnection.RollbackTrans
   OnSaveData = False
End Function

' 申請日
Private Sub textCE02_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   If IsEmptyText(textCE02) = False Then
      If CheckIsTaiwanDate(textCE02, False) = False Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "請輸入正確的申請日"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCE02_GotFocus
      End If
   End If
End Sub

'Add By Sindy 2010/5/13
'Modified by Lydia 2021/08/11改成Form2.0 ;
'Private Sub textCE04_KeyPress(KeyAscii As Integer)
Private Sub textCE04_KeyPress(KeyAscii As MSForms.ReturnInteger)
 KeyAscii = UpperCase(KeyAscii)
End Sub

' 申請人
Private Sub textCE04_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   Cancel = False
   textCE04_2 = Empty
   If IsEmptyText(textCE04) = False Then
        'Add By Cheng 2003/04/14
        '申請人編號補滿9碼
        Me.textCE04.Text = Left(Me.textCE04.Text & "000000000", 9)
      'edit by 2004/07/22 nick  檢查該申請人或代理人狀態，若為不再使用則停在原地
      Dim oState As Boolean
      oState = True
      'textCE04_2 = GetCustomerName(textCE04)
      textCE04_2 = GetCustomerNameAndState(textCE04, "0", oState)
      If oState = False Then
            Cancel = True
            Exit Sub
      End If
      If textCE04_2 = Empty Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "申請人代碼<" & textCE04 & ">不正確"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCE04_GotFocus
      End If
   End If
End Sub

Private Sub textCE39_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   textCE39_2 = Empty
   Cancel = False
   If IsEmptyText(textCE39) = False Then
      textCE39_2 = GetTradeMarkName(textCE39, 0)
      If IsEmptyText(textCE39_2) = True Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "商標種類不存在"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCE39_GotFocus
      End If
   End If
End Sub

Private Sub textCE02_GotFocus()
   InverseTextBox textCE02
End Sub

Private Sub textCE04_GotFocus()
   InverseTextBox textCE04
End Sub

Private Sub textCE10_GotFocus()
   InverseTextBox textCE10
End Sub

Private Sub textCE11_GotFocus()
   InverseTextBox textCE11
End Sub

Private Sub textCE12_GotFocus()
   InverseTextBox textCE12
End Sub

Private Sub textCE13_GotFocus()
   InverseTextBox textCE13
End Sub

Private Sub textCE14_GotFocus()
   InverseTextBox textCE14
End Sub

Private Sub textCE15_GotFocus()
   InverseTextBox textCE15
End Sub

Private Sub textCE17_GotFocus()
   InverseTextBox textCE17
End Sub

Private Sub textCE23_GotFocus()
   InverseTextBox textCE23
End Sub

Private Sub textCE24_GotFocus()
   InverseTextBox textCE24
End Sub

Private Sub textCE25_GotFocus()
   InverseTextBox textCE25
End Sub

Private Sub textCE39_GotFocus()
   InverseTextBox textCE39
End Sub

Private Sub textCE41_1_GotFocus()
    TextInverse Me.textCE41_1
End Sub

Private Sub textCE41_GotFocus()
   InverseTextBox textCE41
End Sub

Private Sub textCE42_GotFocus()
   InverseTextBox textCE42
End Sub

Private Sub textCE43_GotFocus()
   InverseTextBox textCE43
End Sub

Private Sub textCE45_GotFocus()
   InverseTextBox textCE45
End Sub

Private Sub textCE47_GotFocus()
   InverseTextBox textCE47
End Sub

Private Sub textCE47_Validate(Cancel As Boolean)
'add by nickc 2005/06/03
textCE47 = Replace(textCE47, " ", "")
End Sub

Private Sub textCE49_GotFocus()
   InverseTextBox textCE49
End Sub

Private Sub textCE57_GotFocus()
   InverseTextBox textCE57
End Sub

Private Sub textCE61_GotFocus()
   InverseTextBox textCE61
End Sub

Private Sub textCE63_GotFocus()
   InverseTextBox textCE63
End Sub

Private Sub textCE64_GotFocus()
   InverseTextBox textCE64
End Sub

Private Function CheckDataValid() As Boolean
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Dim bUpdate As Boolean
   
   CheckDataValid = False
   bUpdate = False
    '第一頁
    'Modify By Cheng 2003/03/07
'   If checkCE09.Value = True Then bUpdate = True
'   If checkCE03.Value = True Then bUpdate = True
'   If checkCE16.Value = True Then bUpdate = True
'   If checkCE56.Value = True Then bUpdate = True
'   If checkCE54.Value = True Then bUpdate = True
'   If checkCE52.Value = True Then bUpdate = True
'   If checkCE22.Value = True Then bUpdate = True
'   If checkCE65.Value = True Then bUpdate = True
   If checkCE09.Value = vbChecked Then bUpdate = True
   If checkCE03.Value = vbChecked Then bUpdate = True
   If checkCE16.Value = vbChecked Then bUpdate = True
   If checkCE56.Value = vbChecked Then bUpdate = True
   If checkCE54.Value = vbChecked Then bUpdate = True
   If checkCE52.Value = vbChecked Then bUpdate = True
   If checkCE22.Value = vbChecked Then bUpdate = True
   If checkCE65.Value = vbChecked Then bUpdate = True
    '第二頁
'   If checkCE38.Value = True Then bUpdate = True
'   If checkCE58.Value = True Then bUpdate = True
'   If checkCE40.Value = True Then bUpdate = True
'   If checkCE60.Value = True Then bUpdate = True
'   If checkCE44.Value = True Then bUpdate = True
'   If checkCE46.Value = True Then bUpdate = True
'   If checkCE48.Value = True Then bUpdate = True
'   If checkCE50.Value = True Then bUpdate = True
'   If checkCE52.Value = True Then bUpdate = True
   If checkCE38.Value = vbChecked Then bUpdate = True
   If checkCE58.Value = vbChecked Then bUpdate = True
   If checkCE40.Value = vbChecked Then bUpdate = True
   If checkCE60.Value = vbChecked Then bUpdate = True
   If checkCE44.Value = vbChecked Then bUpdate = True
   If checkCE46.Value = vbChecked Then bUpdate = True
   If checkCE48.Value = vbChecked Then bUpdate = True
   If checkCE50.Value = vbChecked Then bUpdate = True
   If checkCE52.Value = vbChecked Then bUpdate = True
   If checkCE62.Value = vbChecked Then bUpdate = True
   If bUpdate = False Then
      strTit = "檢核資料"
      strMsg = "請勾選變更項目 !"
      nResponse = MsgBox(strMsg, vbOKOnly + vbCritical, strTit)
      GoTo EXITSUB
   End If
    'Add By Cheng 2003/04/18
    If Me.checkCE09.Value = vbChecked Then
        If Me.textCE04.Text = "" Then
            MsgBox "請輸入申請人代號!!!", vbExclamation + vbOKOnly
            Me.textCE04.SetFocus
            textCE04_GotFocus
            Exit Function
        End If
        'Add By Sindy 2011/8/3
        If m_CP31 <> "Y" Then 'Add By Sindy 2011/8/23 新案時不檢查
            If ChangeCustomerL(textCE04) = m_TM23 Then
                MsgBox "新申請人編號與目前相同 !", vbCritical
                Me.textCE04.SetFocus
                textCE04_GotFocus
                Exit Function
            End If
        End If
        '2011/8/3 End
    End If
    If Me.checkCE03.Value = vbChecked Then
        If Me.textCE02.Text = "" Then
            MsgBox "請輸入申請日!!!", vbExclamation + vbOKOnly
            Me.textCE02.SetFocus
            textCE02_GotFocus
            Exit Function
        End If
    End If
    If Me.checkCE16.Value = vbChecked Then
        If Me.textCE10.Text = "" And Me.textCE11.Text = "" And Me.textCE12.Text = "" And Me.textCE13.Text = "" And Me.textCE14.Text = "" And Me.textCE15.Text = "" Then
            MsgBox "請輸入代表人名稱!!!", vbExclamation + vbOKOnly
            Me.textCE10.SetFocus
            textCE10_GotFocus
            Exit Function
        End If
    End If
    If Me.checkCE22.Value = vbChecked Then
        If Me.textCE17.Text = "" Then
            MsgBox "請輸入申請人中譯文!!!", vbExclamation + vbOKOnly
            Me.textCE17.SetFocus
            textCE17_GotFocus
            Exit Function
        End If
    End If
    If Me.checkCE65.Value = vbChecked Then
        If Me.textCE63.Text = "" And Me.textCE64.Text = "" Then
            MsgBox "請輸入代表人中譯文!!!", vbExclamation + vbOKOnly
            Me.textCE63.SetFocus
            textCE63_GotFocus
            Exit Function
        End If
    End If
    If Me.checkCE38.Value = vbChecked Then
        If Me.textCE23.Text = "" And Me.textCE24.Text = "" And Me.textCE25.Text = "" Then
            MsgBox "請輸入申請地址!!!", vbExclamation + vbOKOnly
            Me.textCE23.SetFocus
            textCE23_GotFocus
            Exit Function
        End If
    End If
    If Me.checkCE58.Value = vbChecked Then
        If Me.textCE57.Text = "" Then
            MsgBox "請輸入正商標號數!!!", vbExclamation + vbOKOnly
            Me.textCE57.SetFocus
            textCE57_GotFocus
            Exit Function
        End If
    End If
    If Me.checkCE40.Value = vbChecked Then
        If Me.textCE39.Text = "" Then
            MsgBox "請輸入正商標種類!!!", vbExclamation + vbOKOnly
            Me.textCE39.SetFocus
            textCE39_GotFocus
            Exit Function
        End If
    End If
    If Me.checkCE44.Value = vbChecked Then
        Select Case m_TM01
        Case "T", "FCT", "CFT", "TF", "S"
            If Me.textCE41_1.Text = "" Then
                MsgBox "請輸入案件名稱!!!", vbExclamation + vbOKOnly
                Me.textCE41_1.SetFocus
                textCE41_1_GotFocus
                Exit Function
            End If
        Case Else
            If Me.textCE41.Text = "" And Me.textCE42.Text = "" And Me.textCE43.Text = "" Then
                MsgBox "請輸入案件名稱!!!", vbExclamation + vbOKOnly
                Me.textCE41.SetFocus
                textCE41_GotFocus
                Exit Function
            End If
        End Select
    End If
    If Me.checkCE46.Value = vbChecked Then
        If Me.textCE45.Text = "" Then
            MsgBox "請輸入縮減商品!!!", vbExclamation + vbOKOnly
            Me.textCE45.SetFocus
            textCE45_GotFocus
            Exit Function
        End If
    End If
    If Me.checkCE48.Value = vbChecked Then
        If Me.textCE47.Text = "" Then
            MsgBox "請輸入商品類別!!!", vbExclamation + vbOKOnly
            Me.textCE47.SetFocus
            textCE47_GotFocus
            Exit Function
        End If
    End If
    If Me.checkCE50.Value = vbChecked Then
        If Me.textCE49.Text = "" Then
            MsgBox "請輸入商品群組!!!", vbExclamation + vbOKOnly
            Me.textCE49.SetFocus
            textCE49_GotFocus
            Exit Function
        End If
    End If
    If Me.checkCE62.Value = vbChecked Then
        If Me.textCE61.Text = "" Then
            MsgBox "請輸入其他!!!", vbExclamation + vbOKOnly
            Me.textCE61.SetFocus
            textCE61_GotFocus
            Exit Function
        End If
    End If
   CheckDataValid = True
EXITSUB:
End Function

'Add By Cheng 2002/05/23
Private Function TxtValidate() As Boolean
Dim objTxt As Object
Dim ii As Integer
Dim Cancel As Boolean

TxtValidate = False
If Me.textCE02.Enabled = True Then
   Cancel = False
   textCE02_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textCE04.Enabled = True Then
   Cancel = False
   textCE04_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textCE39.Enabled = True Then
   Cancel = False
   textCE39_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

   'Added by Lydia 2021/08/20 檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
   If PUB_ChkUniText(Me, , True, "TextBox") = False Then
        Exit Function
   End If
   
TxtValidate = True
End Function


'Add By Cheng 2003/04/10
Private Function OnSaveTrademark() As Boolean
Dim strSql As String
Dim bFirst As Boolean
Dim bDifference As Boolean
Dim strTmp As String
Dim nIndex As Integer
Dim tmpCp64 As String
Dim rsnick911204 As New ADODB.Recordset
   
On Error GoTo ErrorHandler
    
    OnSaveTrademark = True
    tmpCp64 = ""
    tmpCp64 = " select cp64 from caseprogress where cP09= '" & m_CE01 & "'"
    Set rsnick911204 = New ADODB.Recordset
    rsnick911204.CursorLocation = adUseClient
    rsnick911204.Open tmpCp64, cnnConnection, adOpenStatic, adLockReadOnly
    tmpCp64 = ""
    If rsnick911204.RecordCount > 0 Then
         tmpCp64 = CheckStr(rsnick911204.Fields(0).Value) & " "
    End If
    ' 申請人
'    If checkCE09.Value = True Then
    If checkCE09.Value = vbChecked Then
       If tmpOldCE04 <> textCE04 Then
             SetTMFieldData "TM23", textCE04, 0
'             tmpCp64 = tmpCp64 & "原申請人:" & tmpOldCE04 & " "
             If tmpOldCE04 <> "" Then tmpCp64 = tmpCp64 & "原申請人:" & tmpOldCE04 & " "
            '申請地址
             SetTMFieldData "TM24", PUB_GetCustEachAdd(textCE04, 1), 0
             SetTMFieldData "TM25", PUB_GetCustEachAdd(textCE04, 2), 0
             SetTMFieldData "TM26", PUB_GetCustEachAdd(textCE04, 3), 0
       End If
    End If
    ' 申請日
'    If checkCE03.Value = True Then
    If checkCE03.Value = vbChecked Then
       If tmpOldCE02 <> DBDATE(textCE02) Then
             SetTMFieldData "TM11", DBDATE(textCE02), 1
'             tmpCp64 = tmpCp64 & "原申請日:" & tmpOldCE02 & " "
             If tmpOldCE02 <> "" Then tmpCp64 = tmpCp64 & "原申請日:" & tmpOldCE02 & " "
       End If
    End If
    '911204 nick 新增代表人 只判斷是否有變更
    If checkCE16.Value = 1 Then
       If textCE10 <> tmpOldCE10 Then
             SetTMFieldData "TM47", textCE10, 0
'             tmpCp64 = tmpCp64 & "原代表人:" & tmpOldCE10 & " "
             If tmpOldCE10 <> "" Then tmpCp64 = tmpCp64 & "原代表人1(中):" & tmpOldCE10 & " "
       End If
       If textCE11 <> tmpOldCE11 Then
             SetTMFieldData "TM48", textCE11, 0
'             tmpCp64 = tmpCp64 & "原代表人:" & tmpOldCE11 & " "
             If tmpOldCE11 <> "" Then tmpCp64 = tmpCp64 & "原代表人1(英):" & tmpOldCE11 & " "
       End If
       If textCE12 <> tmpOldCE12 Then
             SetTMFieldData "TM49", textCE12, 0
'             tmpCp64 = tmpCp64 & "原代表人:" & tmpOldCE12 & " "
             If tmpOldCE12 <> "" Then tmpCp64 = tmpCp64 & "原代表人1(日):" & tmpOldCE12 & " "
       End If
       If textCE13 <> tmpOldCE13 Then
             SetTMFieldData "TM50", textCE13, 0
'             tmpCp64 = tmpCp64 & "原代表人:" & tmpOldCE13 & " "
             If tmpOldCE13 <> "" Then tmpCp64 = tmpCp64 & "原代表人2(中):" & tmpOldCE13 & " "
       End If
       If textCE14 <> tmpOldCE14 Then
             SetTMFieldData "TM51", textCE14, 0
'             tmpCp64 = tmpCp64 & "原代表人:" & tmpOldCE14 & " "
             If tmpOldCE14 <> "" Then tmpCp64 = tmpCp64 & "原代表人2(英):" & tmpOldCE14 & " "
       End If
       If textCE15 <> tmpOldCE15 Then
             SetTMFieldData "TM52", textCE15, 0
'             tmpCp64 = tmpCp64 & "原代表人:" & tmpOldCE15 & " "
             If tmpOldCE15 <> "" Then tmpCp64 = tmpCp64 & "原代表人2(日):" & tmpOldCE15 & " "
       End If
    End If
    ' 申請地址
'    If checkCE38.Value = True Then
    If checkCE38.Value = vbChecked Then
         If tmpOldCE23 <> textCE23 Then
             SetTMFieldData "TM24", textCE23, 0
'             tmpCp64 = tmpCp64 & "原申請中文地址:" & tmpOldCE23 & " "
             If tmpOldCE23 <> "" Then tmpCp64 = tmpCp64 & "原申請中文地址:" & tmpOldCE23 & " "
         End If
         If tmpOldCE24 <> textCE24 Then
             SetTMFieldData "TM25", textCE24, 0
'             tmpCp64 = tmpCp64 & "原申請英文地址:" & tmpOldCE24 & " "
             If tmpOldCE24 <> "" Then tmpCp64 = tmpCp64 & "原申請英文地址:" & tmpOldCE24 & " "
         End If
         If tmpOldCE25 <> textCE25 Then
             SetTMFieldData "TM26", textCE25, 0
'             tmpCp64 = tmpCp64 & "原申請日文地址:" & tmpOldCE25 & " "
             If tmpOldCE25 <> "" Then tmpCp64 = tmpCp64 & "原申請日文地址:" & tmpOldCE25 & " "
         End If
    End If
    '正商標號數
'    If checkCE58.Value = True Then
    If checkCE58.Value = vbChecked Then
         If tmpOldCE57 <> textCE57 Then
             SetTMFieldData "TM27", textCE57, 0
'             tmpCp64 = tmpCp64 & "原正商標號數:" & tmpOldCE57 & " "
             If tmpOldCE57 <> "" Then tmpCp64 = tmpCp64 & "原正商標號數:" & tmpOldCE57 & " "
         End If
    End If
    '商標種類
'    If checkCE40.Value = True Then
    If checkCE40.Value = vbChecked Then
         If tmpOldCE39 <> textCE39 Then
             SetTMFieldData "TM08", textCE39, 0
'             tmpCp64 = tmpCp64 & "原商標種類:" & tmpOldCE39 & " "
             If tmpOldCE39 <> "" Then tmpCp64 = tmpCp64 & "原商標種類:" & tmpOldCE39 & " "
            'Add By Cheng 2003/09/09
            '聯合商標變更為正商標, 清除基本檔的正商標號數
            If (tmpOldCE39 = "2" And Me.textCE39.Text = "1") Or (tmpOldCE39 = "5" And Me.textCE39.Text = "4") Then
                SetTMFieldData "TM27", "", 0
            End If
         End If
    End If
    ' 案件名稱
'    If checkCE44.Value = True Then
    If checkCE44.Value = vbChecked Then
        If tmpOldCE41 <> textCE41_1 Then
            SetTMFieldData "TM05", textCE41_1, 0
            If tmpOldCE41 <> "" Then tmpCp64 = tmpCp64 & "原案件名稱:" & tmpOldCE41 & " "
        End If
'         '911204 nick
'         If tmpOldCE41 <> textCE41 Then
'             SetTMFieldData "TM05", textCE41, 0
'             '911204 nick
''             tmpCp64 = tmpCp64 & "原案件中文名稱:" & tmpOldCE41 & " "
'             If tmpOldCE41 <> "" Then tmpCp64 = tmpCp64 & "原案件中文名稱:" & tmpOldCE41 & " "
'         End If
'         '911204 nick
'         If tmpOldCE42 <> textCE42 Then
'             SetTMFieldData "TM06", textCE42, 0
'             '911204 nick
''             tmpCp64 = tmpCp64 & "原案件英文名稱:" & tmpOldCE42 & " "
'             If tmpOldCE42 <> "" Then tmpCp64 = tmpCp64 & "原案件英文名稱:" & tmpOldCE42 & " "
'         End If
'         '911204 nick
'         If tmpOldCE43 <> textCE43 Then
'             SetTMFieldData "TM07", textCE43, 0
'             '911204 nick
''             tmpCp64 = tmpCp64 & "原案件日文名稱:" & tmpOldCE43 & " "
'             If tmpOldCE43 <> "" Then tmpCp64 = tmpCp64 & "原案件日文名稱:" & tmpOldCE43 & " "
'         End If
    End If
    '商品類別
'    If checkCE48.Value = True Then
    If checkCE48.Value = vbChecked Then
         If tmpOldCE47 <> textCE47 Then
             SetTMFieldData "TM09", textCE47, 0
'             tmpCp64 = tmpCp64 & "原商標類別:" & tmpOldCE47 & " "
             If tmpOldCE47 <> "" Then tmpCp64 = tmpCp64 & "原商標類別:" & tmpOldCE47 & " "
         End If
    End If
    '商品群組
'    If checkCE50.Value = True Then
    If checkCE50.Value = vbChecked Then
         If tmpOldCE49 <> textCE49 Then
             SetTMFieldData "TM32", textCE49, 0
'             tmpCp64 = tmpCp64 & "原商標群組:" & tmpOldCE49 & " "
             If tmpOldCE49 <> "" Then tmpCp64 = tmpCp64 & "原商標群組:" & tmpOldCE49 & " "
         End If
    End If
    ' 更新商標基本檔
    strSql = "UPDATE Trademark SET "
    bFirst = True
    bDifference = False
    For nIndex = 0 To m_TMListCount - 1
       strTmp = Empty
       If m_TMList(nIndex).tiType = 0 Then
          'edit by nick 2004/12/01  解單引號錯誤
          'strTmp = m_TMList(nIndex).tiName & " = '" & m_TMList(nIndex).tiData & "'"
          strTmp = m_TMList(nIndex).tiName & " = '" & ChgSQL(m_TMList(nIndex).tiData) & "'"
       Else
          If m_TMList(nIndex).tiData = Empty Then
             strTmp = m_TMList(nIndex).tiName & " = " & 0
          Else
             'edit by nick 2004/12/01  解單引號錯誤
             'strTmp = m_TMList(nIndex).tiName & " = " & m_TMList(nIndex).tiData
             strTmp = m_TMList(nIndex).tiName & " = " & ChgSQL(m_TMList(nIndex).tiData)
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
    Next nIndex
    ' 組成SQL語法
    strSql = strSql & " " & _
                   "WHERE TM01 = '" & m_TM01 & "' AND " & _
                         "TM02 = '" & m_TM02 & "' AND " & _
                         "TM03 = '" & m_TM03 & "' AND " & _
                         "TM04 = '" & m_TM04 & "'"
    ' 執行SQL指令
    If bDifference = True Then
       cnnConnection.Execute strSql
       '911226 nick 更新回原本收文號的備註
       'edit by nick 2004/12/01  解單引號錯誤
       'StrSql = "update caseprogress set cp64='" & tmpCp64 & "' where cp09='" & m_CE01 & "' "
       strSql = "update caseprogress set cp64='" & ChgSQL(tmpCp64) & "' where cp09='" & m_CE01 & "' "
       cnnConnection.Execute strSql
    End If
    
    ' 清除所佔用的記憶體
    If m_TMListCount > 0 Then
       Erase m_TMList
       m_TMListCount = 0
    End If
'Add By Cheng 2002/11/06
Exit Function
ErrorHandler:
    OnSaveTrademark = False
End Function

' 設定欄位新值
Private Sub SetTMFieldData(ByVal strField As String, ByVal strNewData As String, ByVal nType As Integer)
   Dim bFind As Boolean
   Dim nIndex As Integer
   ' 搜尋是否存在該欄位
   bFind = False
   For nIndex = 0 To m_TMListCount - 1
      If m_TMList(nIndex).tiName = strField Then
         bFind = True
         m_TMList(nIndex).tiData = strNewData
         Exit For
      End If
   Next nIndex
   ' 不存在則新增該欄位
   If bFind = False Then
      ReDim Preserve m_TMList(m_TMListCount + 1)
      m_TMList(m_TMListCount).tiName = strField
      m_TMList(m_TMListCount).tiData = strNewData
      m_TMList(m_TMListCount).tiType = nType
      m_TMListCount = m_TMListCount + 1
   End If
End Sub


' 91.09.02 modify by louis
'Modify By Cheng 2002/11/06
'Private Sub OnSaveServicePractice()
Private Function OnSaveServicePractice() As Boolean
   Dim strSql As String
   Dim bFirst As Boolean
   Dim bDifference As Boolean
   Dim strTmp As String
   Dim nIndex As Integer
   '911204 nick
   Dim tmpCp64 As String
   Dim rsnick911204 As New ADODB.Recordset
   
   
'Add By Cheng 2002/11/06
On Error GoTo ErrorHandler
OnSaveServicePractice = True

   '911204 nick
   tmpCp64 = ""
    'Modify By Cheng 2003/04/10
    '取消限制
'   ' 只有系統類別為TC及案件性質為變更301才更新基本檔
'   If m_TM01 <> "TC" Or m_CP10 <> "301" Then
'      Exit Function
'   End If
   
   '911204 nick
   tmpCp64 = " select cp64 from caseprogress where cP09= '" & m_CE01 & "'"
   Set rsnick911204 = New ADODB.Recordset
   rsnick911204.CursorLocation = adUseClient
   rsnick911204.Open tmpCp64, cnnConnection, adOpenStatic, adLockReadOnly
   tmpCp64 = ""
   If rsnick911204.RecordCount > 0 Then
        tmpCp64 = CheckStr(rsnick911204.Fields(0).Value) & " "
   End If
   
   ' 申請人
'   If checkCE09.Value = True Then
   If checkCE09.Value = vbChecked Then
      '911204 nick
      If tmpOldCE04 <> textCE04 Then
            SetSRFieldData "SP08", textCE04, 0
            '911204 nick
'            tmpCp64 = tmpCp64 & "原申請人:" & tmpOldCE04 & " "
            If tmpOldCE04 <> "" Then tmpCp64 = tmpCp64 & "原申請人:" & tmpOldCE04 & " "
      End If
   End If
   ' 申請日
'   If checkCE03.Value = True Then
   If checkCE03.Value = vbChecked Then
      '911204 nick
      If tmpOldCE02 <> DBDATE(textCE02) Then
            SetSRFieldData "SP10", DBDATE(textCE02), 1
            '911204 nick
'            tmpCp64 = tmpCp64 & "原申請日:" & tmpOldCE02 & " "
            If tmpOldCE02 <> "" Then tmpCp64 = tmpCp64 & "原申請日:" & tmpOldCE02 & " "
      End If
   End If
   '911204 nick 新增代表人 只判斷是否有變更
   If checkCE16.Value = 1 Then
      If textCE10 <> tmpOldCE10 Then
            SetSRFieldData "SP42", textCE10, 0
'            tmpCp64 = tmpCp64 & "原代表人:" & tmpOldCE10 & " "
            If tmpOldCE10 <> "" Then tmpCp64 = tmpCp64 & "原代表人:" & tmpOldCE10 & " "
      End If
   End If
   ' 案件名稱
'   If checkCE44.Value = True Then
   If checkCE44.Value = vbChecked Then
        Select Case m_TM01
        Case "S"
            If tmpOldCE41 <> textCE41_1 Then
                SetSRFieldData "SP05", textCE41_1, 0
                If tmpOldCE41 <> "" Then tmpCp64 = tmpCp64 & "原案件中文名稱:" & tmpOldCE41 & " "
            End If
        Case Else
            '911204 nick
            If tmpOldCE41 <> textCE41 Then
                SetSRFieldData "SP05", textCE41, 0
                '911204 nick
    '            tmpCp64 = tmpCp64 & "原案件中文名稱:" & tmpOldCE41 & " "
                If tmpOldCE41 <> "" Then tmpCp64 = tmpCp64 & "原案件中文名稱:" & tmpOldCE41 & " "
            End If
        End Select
        '911204 nick
        If tmpOldCE42 <> textCE42 Then
            SetSRFieldData "SP06", textCE42, 0
            '911204 nick
'            tmpCp64 = tmpCp64 & "原案件英文名稱:" & tmpOldCE42 & " "
            If tmpOldCE42 <> "" Then tmpCp64 = tmpCp64 & "原案件英文名稱:" & tmpOldCE42 & " "
        End If
        '911204 nick
        If tmpOldCE43 <> textCE43 Then
            SetSRFieldData "SP07", textCE43, 0
            '911204 nick
'            tmpCp64 = tmpCp64 & "原案件日文名稱:" & tmpOldCE43 & " "
            If tmpOldCE43 <> "" Then tmpCp64 = tmpCp64 & "原案件日文名稱:" & tmpOldCE43 & " "
        End If
   End If
   
   ' 更新服務業務基本檔
   strSql = "UPDATE ServicePractice SET "
   bFirst = True
   bDifference = False
   For nIndex = 0 To m_SRListCount - 1
      strTmp = Empty
      If m_SRList(nIndex).siType = 0 Then
         'edit by nick 2004/12/01  解單引號錯誤
         'strTmp = m_SRList(nIndex).siName & " = '" & m_SRList(nIndex).siData & "'"
         strTmp = m_SRList(nIndex).siName & " = '" & ChgSQL(m_SRList(nIndex).siData) & "'"
      Else
         If m_SRList(nIndex).siData = Empty Then
            strTmp = m_SRList(nIndex).siName & " = " & 0
         Else
            'edit by nick 2004/12/01  解單引號錯誤
            'strTmp = m_SRList(nIndex).siName & " = " & m_SRList(nIndex).siData
            strTmp = m_SRList(nIndex).siName & " = " & ChgSQL(m_SRList(nIndex).siData)
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
   Next nIndex
   ' 組成SQL語法
   strSql = strSql & " " & _
                  "WHERE SP01 = '" & m_TM01 & "' AND " & _
                        "SP02 = '" & m_TM02 & "' AND " & _
                        "SP03 = '" & m_TM03 & "' AND " & _
                        "SP04 = '" & m_TM04 & "'"
   ' 執行SQL指令
   If bDifference = True Then
      cnnConnection.Execute strSql
      '911226 nick 更新回原本收文號的備註
      'edit by nick 2004/12/01  解單引號錯誤
      'StrSql = "update caseprogress set cp64='" & tmpCp64 & "' where cp09='" & m_CE01 & "' "
      strSql = "update caseprogress set cp64='" & ChgSQL(tmpCp64) & "' where cp09='" & m_CE01 & "' "
      cnnConnection.Execute strSql
   End If
   
   ' 清除所佔用的記憶體
   If m_SRListCount > 0 Then
      Erase m_SRList
      m_SRListCount = 0
   End If
'Add By Cheng 2002/11/06
Exit Function
ErrorHandler:
    OnSaveServicePractice = False
End Function


' 設定欄位新值
Private Sub SetSRFieldData(ByVal strField As String, ByVal strNewData As String, ByVal nType As Integer)
   Dim bFind As Boolean
   Dim nIndex As Integer
   ' 搜尋是否存在該欄位
   bFind = False
   For nIndex = 0 To m_SRListCount - 1
      If m_SRList(nIndex).siName = strField Then
         bFind = True
         m_SRList(nIndex).siData = strNewData
         Exit For
      End If
   Next nIndex
   ' 不存在則新增該欄位
   If bFind = False Then
      ReDim Preserve m_SRList(m_SRListCount + 1)
      m_SRList(m_SRListCount).siName = strField
      m_SRList(m_SRListCount).siData = strNewData
      m_SRList(m_SRListCount).siType = nType
      m_SRListCount = m_SRListCount + 1
   End If
End Sub

