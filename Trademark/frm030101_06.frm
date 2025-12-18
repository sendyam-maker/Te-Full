VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm030101_06 
   BorderStyle     =   1  '單線固定
   Caption         =   "發文(CFC申請)"
   ClientHeight    =   5790
   ClientLeft      =   5590
   ClientTop       =   1720
   ClientWidth     =   9150
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5790
   ScaleWidth      =   9150
   Begin VB.TextBox textTM15 
      BorderStyle     =   0  '沒有框線
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   57
      TabStop         =   0   'False
      Top             =   1140
      Width           =   2532
   End
   Begin VB.TextBox textTM45 
      BorderStyle     =   0  '沒有框線
      Height          =   285
      Left            =   5700
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   56
      TabStop         =   0   'False
      Top             =   1440
      Width           =   2532
   End
   Begin VB.TextBox textCP10 
      BorderStyle     =   0  '沒有框線
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   55
      TabStop         =   0   'False
      Top             =   1440
      Width           =   2532
   End
   Begin VB.TextBox textCP09 
      BorderStyle     =   0  '沒有框線
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   54
      TabStop         =   0   'False
      Top             =   540
      Width           =   2532
   End
   Begin VB.TextBox textTMKey 
      BorderStyle     =   0  '沒有框線
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   53
      TabStop         =   0   'False
      Top             =   840
      Width           =   2532
   End
   Begin VB.TextBox textTM12 
      BorderStyle     =   0  '沒有框線
      Height          =   285
      Left            =   5700
      Locked          =   -1  'True
      TabIndex        =   52
      TabStop         =   0   'False
      Top             =   1140
      Width           =   2532
   End
   Begin VB.TextBox textTM20 
      BorderStyle     =   0  '沒有框線
      Height          =   285
      Left            =   5700
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   51
      TabStop         =   0   'False
      Top             =   840
      Width           =   2532
   End
   Begin VB.TextBox textCP12 
      BorderStyle     =   0  '沒有框線
      Height          =   285
      Left            =   5700
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   50
      TabStop         =   0   'False
      Top             =   540
      Width           =   2532
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   8160
      TabIndex        =   29
      Top             =   60
      Width           =   912
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   6060
      TabIndex        =   24
      Top             =   60
      Width           =   912
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "回前畫面(&U)"
      Height          =   400
      Left            =   7020
      TabIndex        =   28
      Top             =   60
      Width           =   1092
   End
   Begin VB.CommandButton cmdMod 
      BackColor       =   &H008080FF&
      Caption         =   "變更事項(&R)"
      Height          =   400
      Left            =   4824
      Style           =   1  '圖片外觀
      TabIndex        =   27
      Top             =   60
      Width           =   1215
   End
   Begin VB.CommandButton cmdRelate 
      Caption         =   "相關卷號(&F)"
      Height          =   400
      Left            =   3660
      TabIndex        =   26
      Top             =   60
      Width           =   1092
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "同時發文(&N)"
      Height          =   400
      Index           =   1
      Left            =   2520
      TabIndex        =   25
      Top             =   60
      Width           =   1092
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3672
      Left            =   120
      TabIndex        =   30
      Top             =   2100
      Width           =   8892
      _ExtentX        =   15681
      _ExtentY        =   6473
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   5
      TabHeight       =   420
      TabCaption(0)   =   "基本資料1"
      TabPicture(0)   =   "frm030101_06.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label11"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1(12)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label16"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label15"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label1(10)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label9"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label8"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label7"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label23"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label22"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label4"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label14"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label25"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label6"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Label5"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Label10"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Label12"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Label13"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "textSP05"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "textSP06"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "textSP07"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "textCP44_2"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "textSP08_2"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "textSP58_2"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "textSP59_2"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "textCP44"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "textCF09"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "textCP26"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "textCP18"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "textPrint"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "textUargeDate"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "textCP27"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "textSP08"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "textSP58"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "textSP59"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "checkAttach(0)"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "checkAttach(1)"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "checkAttach(2)"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "checkAttach(3)"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).Control(39)=   "textSP48"
      Tab(0).Control(39).Enabled=   0   'False
      Tab(0).ControlCount=   40
      TabCaption(1)   =   "基本資料2"
      TabPicture(1)   =   "frm030101_06.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label29"
      Tab(1).Control(1)=   "Label28"
      Tab(1).Control(2)=   "Label18"
      Tab(1).Control(3)=   "Label17"
      Tab(1).Control(4)=   "Label19"
      Tab(1).Control(5)=   "Label20"
      Tab(1).Control(6)=   "Label21"
      Tab(1).Control(7)=   "textSP18"
      Tab(1).Control(8)=   "textCP64"
      Tab(1).Control(9)=   "textSP46"
      Tab(1).Control(10)=   "textAttach"
      Tab(1).Control(11)=   "textSP42"
      Tab(1).Control(12)=   "textPrintLetter"
      Tab(1).ControlCount=   13
      Begin VB.TextBox textSP48 
         Height          =   264
         Left            =   7536
         MaxLength       =   1
         TabIndex        =   5
         Top             =   960
         Width           =   372
      End
      Begin VB.CheckBox checkAttach 
         Caption         =   "書"
         Height          =   252
         Index           =   3
         Left            =   4200
         TabIndex        =   11
         Top             =   1560
         Width           =   852
      End
      Begin VB.CheckBox checkAttach 
         Caption         =   "圖示"
         Height          =   252
         Index           =   2
         Left            =   3240
         TabIndex        =   10
         Top             =   1560
         Width           =   852
      End
      Begin VB.CheckBox checkAttach 
         Caption         =   "照片"
         Height          =   252
         Index           =   1
         Left            =   2280
         TabIndex        =   9
         Top             =   1560
         Width           =   852
      End
      Begin VB.CheckBox checkAttach 
         Caption         =   "申請書"
         Height          =   252
         Index           =   0
         Left            =   1200
         TabIndex        =   8
         Top             =   1560
         Width           =   852
      End
      Begin VB.TextBox textSP59 
         Height          =   264
         Left            =   1200
         MaxLength       =   9
         TabIndex        =   17
         Top             =   3300
         Width           =   1092
      End
      Begin VB.TextBox textSP58 
         Height          =   264
         Left            =   1200
         MaxLength       =   9
         TabIndex        =   16
         Top             =   3000
         Width           =   1092
      End
      Begin VB.TextBox textSP08 
         Height          =   264
         Left            =   1200
         MaxLength       =   9
         TabIndex        =   15
         Top             =   2700
         Width           =   1092
      End
      Begin VB.TextBox textCP27 
         Height          =   264
         Left            =   1200
         MaxLength       =   8
         TabIndex        =   0
         Top             =   360
         Width           =   1092
      End
      Begin VB.TextBox textUargeDate 
         Height          =   264
         Left            =   1200
         TabIndex        =   2
         Top             =   660
         Width           =   1092
      End
      Begin VB.TextBox textPrint 
         Height          =   264
         Left            =   1200
         MaxLength       =   1
         TabIndex        =   6
         Top             =   1260
         Width           =   372
      End
      Begin VB.TextBox textCP18 
         Height          =   264
         Left            =   5520
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   1
         Top             =   360
         Width           =   2532
      End
      Begin VB.TextBox textCP26 
         Height          =   264
         Left            =   5760
         MaxLength       =   1
         TabIndex        =   3
         Top             =   660
         Width           =   372
      End
      Begin VB.TextBox textCF09 
         Height          =   264
         Left            =   4920
         MaxLength       =   3
         TabIndex        =   7
         Top             =   1260
         Width           =   612
      End
      Begin VB.TextBox textPrintLetter 
         Height          =   264
         Left            =   -73560
         MaxLength       =   1
         TabIndex        =   20
         Top             =   960
         Width           =   372
      End
      Begin VB.ComboBox textCP44 
         Height          =   300
         Left            =   1200
         TabIndex        =   4
         Top             =   960
         Width           =   1524
      End
      Begin MSForms.TextBox textSP42 
         Height          =   285
         Left            =   -73560
         TabIndex        =   21
         Top             =   1260
         Width           =   7275
         VariousPropertyBits=   671105051
         MaxLength       =   120
         Size            =   "12827;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textAttach 
         Height          =   285
         Left            =   -73560
         TabIndex        =   19
         Top             =   660
         Width           =   7272
         VariousPropertyBits=   671105051
         MaxLength       =   100
         Size            =   "12827;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textSP46 
         Height          =   285
         Left            =   -73560
         TabIndex        =   18
         Top             =   360
         Width           =   7272
         VariousPropertyBits=   671105051
         MaxLength       =   20
         Size            =   "12827;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textSP59_2 
         Height          =   285
         Left            =   2400
         TabIndex        =   72
         TabStop         =   0   'False
         Top             =   3300
         Width           =   6312
         VariousPropertyBits=   671105055
         ForeColor       =   -2147483641
         Size            =   "11134;503"
         BorderColor     =   16777215
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textSP58_2 
         Height          =   285
         Left            =   2400
         TabIndex        =   70
         TabStop         =   0   'False
         Top             =   3000
         Width           =   6312
         VariousPropertyBits=   671105055
         ForeColor       =   -2147483641
         Size            =   "11134;503"
         BorderColor     =   16777215
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textSP08_2 
         Height          =   285
         Left            =   2400
         TabIndex        =   68
         TabStop         =   0   'False
         Top             =   2700
         Width           =   6312
         VariousPropertyBits=   671105055
         ForeColor       =   -2147483641
         Size            =   "11134;503"
         BorderColor     =   16777215
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCP44_2 
         Height          =   285
         Left            =   2784
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   960
         Width           =   3660
         VariousPropertyBits=   671105055
         MaxLength       =   20
         Size            =   "6456;503"
         BorderColor     =   16777215
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textSP07 
         Height          =   285
         Left            =   1440
         TabIndex        =   14
         Top             =   2400
         Width           =   7272
         VariousPropertyBits=   671105051
         MaxLength       =   60
         Size            =   "12827;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textSP06 
         Height          =   285
         Left            =   1440
         TabIndex        =   13
         Top             =   2100
         Width           =   7272
         VariousPropertyBits=   671105051
         MaxLength       =   60
         Size            =   "12827;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textSP05 
         Height          =   285
         Left            =   1440
         TabIndex        =   12
         Top             =   1800
         Width           =   7272
         VariousPropertyBits=   671105051
         MaxLength       =   60
         Size            =   "12827;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCP64 
         Height          =   950
         Left            =   -73560
         TabIndex        =   22
         Top             =   1560
         Width           =   7272
         VariousPropertyBits=   -1466941413
         MaxLength       =   2000
         ScrollBars      =   2
         Size            =   "12827;1676"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textSP18 
         Height          =   950
         Left            =   -73560
         TabIndex        =   23
         Top             =   2580
         Width           =   7272
         VariousPropertyBits=   -1466941413
         MaxLength       =   2000
         ScrollBars      =   2
         Size            =   "12827;1676"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label21 
         Caption         =   "代表人 :"
         Height          =   252
         Left            =   -74880
         TabIndex        =   78
         Top             =   1260
         Width           =   1212
      End
      Begin VB.Label Label20 
         Caption         =   "附件種類 :"
         Height          =   252
         Left            =   -74880
         TabIndex        =   77
         Top             =   660
         Width           =   1212
      End
      Begin VB.Label Label19 
         Caption         =   "作品種類 :"
         Height          =   252
         Left            =   -74880
         TabIndex        =   76
         Top             =   360
         Width           =   1212
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "是否發行 :              (Y:發行)"
         Height          =   180
         Left            =   6552
         TabIndex        =   75
         Top             =   960
         Width           =   2136
      End
      Begin VB.Label Label12 
         Caption         =   "函知附件:"
         Height          =   252
         Left            =   120
         TabIndex        =   74
         Top             =   1560
         Width           =   972
      End
      Begin VB.Label Label10 
         Caption         =   "申請人3 :"
         Height          =   252
         Left            =   120
         TabIndex        =   73
         Top             =   3300
         Width           =   852
      End
      Begin VB.Label Label5 
         Caption         =   "申請人2 :"
         Height          =   252
         Left            =   120
         TabIndex        =   71
         Top             =   3000
         Width           =   852
      End
      Begin VB.Label Label6 
         Caption         =   "申請人1 :"
         Height          =   252
         Left            =   120
         TabIndex        =   69
         Top             =   2700
         Width           =   852
      End
      Begin VB.Label Label25 
         Caption         =   "發文日 :"
         Height          =   252
         Left            =   120
         TabIndex        =   48
         Top             =   360
         Width           =   852
      End
      Begin VB.Label Label14 
         Caption         =   "催審期限 :"
         Height          =   252
         Left            =   120
         TabIndex        =   47
         Top             =   660
         Width           =   972
      End
      Begin VB.Label Label4 
         Caption         =   "代理人 :"
         Height          =   252
         Left            =   120
         TabIndex        =   46
         Top             =   960
         Width           =   972
      End
      Begin VB.Label Label22 
         Caption         =   "列印定稿 :"
         Height          =   252
         Left            =   120
         TabIndex        =   45
         Top             =   1260
         Width           =   972
      End
      Begin VB.Label Label23 
         Caption         =   "(N:不印)"
         Height          =   252
         Left            =   1680
         TabIndex        =   44
         Top             =   1260
         Width           =   1092
      End
      Begin VB.Label Label7 
         Caption         =   "案件日文名稱 :"
         Height          =   252
         Left            =   120
         TabIndex        =   43
         Top             =   2400
         Width           =   1452
      End
      Begin VB.Label Label8 
         Caption         =   "案件英文名稱 :"
         Height          =   252
         Left            =   120
         TabIndex        =   42
         Top             =   2100
         Width           =   1212
      End
      Begin VB.Label Label9 
         Caption         =   "案件中文名稱 :"
         Height          =   252
         Left            =   120
         TabIndex        =   41
         Top             =   1800
         Width           =   1332
      End
      Begin VB.Label Label1 
         Caption         =   "點數 :"
         Height          =   252
         Index           =   10
         Left            =   4440
         TabIndex        =   40
         Top             =   360
         Width           =   972
      End
      Begin VB.Label Label15 
         Caption         =   "(N:不算)"
         Height          =   252
         Left            =   6240
         TabIndex        =   39
         Top             =   660
         Width           =   972
      End
      Begin VB.Label Label16 
         Caption         =   "是否算案件數 :"
         Height          =   252
         Left            =   4440
         TabIndex        =   38
         Top             =   660
         Width           =   1212
      End
      Begin VB.Label Label1 
         Caption         =   "大約"
         Height          =   252
         Index           =   12
         Left            =   4440
         TabIndex        =   37
         Top             =   1260
         Width           =   492
      End
      Begin VB.Label Label11 
         Caption         =   "可接獲回音"
         Height          =   252
         Left            =   5640
         TabIndex        =   36
         Top             =   1260
         Width           =   1212
      End
      Begin VB.Label Label17 
         Caption         =   "(N:不印)"
         Height          =   252
         Left            =   -73080
         TabIndex        =   35
         Top             =   960
         Width           =   972
      End
      Begin VB.Label Label18 
         Caption         =   "是列印指示信 :"
         Height          =   252
         Left            =   -74880
         TabIndex        =   34
         Top             =   960
         Width           =   1212
      End
      Begin VB.Label Label28 
         Caption         =   "進度備註 :"
         Height          =   252
         Left            =   -74880
         TabIndex        =   33
         Top             =   1560
         Width           =   972
      End
      Begin VB.Label Label29 
         Caption         =   "案件備註 :"
         Height          =   252
         Left            =   -74880
         TabIndex        =   32
         Top             =   3000
         Width           =   972
      End
   End
   Begin MSForms.TextBox textCP14 
      Height          =   285
      Left            =   1320
      TabIndex        =   49
      TabStop         =   0   'False
      Top             =   1740
      Width           =   2532
      VariousPropertyBits=   671105055
      MaxLength       =   20
      Size            =   "4466;503"
      BorderColor     =   16777215
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textCP13 
      Height          =   285
      Left            =   5700
      TabIndex        =   79
      Top             =   1740
      Width           =   2535
      VariousPropertyBits=   671105055
      Size            =   "4466;503"
      BorderColor     =   16777215
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label2 
      Caption         =   "審定號數 :"
      Height          =   255
      Left            =   240
      TabIndex        =   67
      Top             =   1140
      Width           =   915
   End
   Begin VB.Label Label1 
      Caption         =   "智權人員 :"
      Height          =   252
      Index           =   11
      Left            =   4740
      TabIndex        =   66
      Top             =   1740
      Width           =   972
   End
   Begin VB.Label Label1 
      Caption         =   "彼所案號 :"
      Height          =   252
      Index           =   9
      Left            =   4740
      TabIndex        =   65
      Top             =   1440
      Width           =   972
   End
   Begin VB.Label Label1 
      Caption         =   "案件性質 :"
      Height          =   252
      Index           =   6
      Left            =   240
      TabIndex        =   64
      Top             =   1440
      Width           =   972
   End
   Begin VB.Label Label1 
      Caption         =   "收文號 :"
      Height          =   252
      Index           =   1
      Left            =   240
      TabIndex        =   63
      Top             =   540
      Width           =   732
   End
   Begin VB.Label Label1 
      Caption         =   "本所案號 :"
      Height          =   252
      Index           =   0
      Left            =   240
      TabIndex        =   62
      Top             =   840
      Width           =   852
   End
   Begin VB.Label Label3 
      Caption         =   "申請案號 :"
      Height          =   252
      Left            =   4740
      TabIndex        =   61
      Top             =   1140
      Width           =   852
   End
   Begin VB.Label Label1 
      Caption         =   "發證日 :"
      Height          =   252
      Index           =   3
      Left            =   4740
      TabIndex        =   60
      Top             =   840
      Width           =   732
   End
   Begin VB.Label Label1 
      Caption         =   "業務區別 :"
      Height          =   252
      Index           =   2
      Left            =   4740
      TabIndex        =   59
      Top             =   540
      Width           =   972
   End
   Begin VB.Label Label24 
      Caption         =   "承辦人 :"
      Height          =   252
      Left            =   240
      TabIndex        =   58
      Top             =   1740
      Width           =   852
   End
End
Attribute VB_Name = "frm030101_06"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/08/11 改成Form2.0 ; textCP13、textCP14、textSP05、textSP06、textSP07、textCP44_2、textSP08_2、textSP58_2、textSP59_2、textSP46、textAttach、textCP64、textSP18、textSP42
'Memo By Sindy 2012/12/4 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/29 員工編號欄已修改
'Memo By Sindy 2010/8/11 日期欄已修改
Option Explicit

Const CB_SHOWDROPDOWN = &H14F
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
' 本所案號
Dim m_TM01 As String
Dim m_TM02 As String
Dim m_TM03 As String
Dim m_TM04 As String
' 收文號
Dim m_CP09 As String
' 申請國家
Dim m_TM10 As String
' 案件性質代號
Dim m_CP10 As String
' 承辦人代號
Dim m_CP14 As String
' 宣告欄位內容結構
Private Type FIELDITEM
   fiName As String
   fiOldData As String
   fiNewData As String
   fiType As Integer
End Type
' 儲存商標基本檔或服務業務基本檔檔案欄位的串列
Dim m_TMSPList() As FIELDITEM
Dim m_TMSPCount As Integer
' 儲存案件進度檔檔案欄位的串列
Dim m_CPList() As FIELDITEM
Dim m_CPCount As Integer
' 宣告代理人內容結構
Private Type AGENTITEM
   aiCode As String
   aiName As String
End Type
Dim m_AgentList() As AGENTITEM
Dim m_AgentCount As Integer
'Add By Cheng 2002/08/23
Dim m_strCust1 As String '申請人1
Dim m_strCust2 As String '申請人2
Dim m_strCust3 As String '申請人3
'Add By Cheng 2003/10/06
Public m_blnClkChgButton As Boolean '是否按下變更事項按鈕
'add by nickc 2008/02/22
Dim m_CP44New As String
Dim m_CP116 As String
Dim m_TM44 As String
Dim m_TM119 As String
Dim m_TM120 As String
Dim m_990CP09 As String 'Add By Sindy 2016/12/20
Dim m_CP07 As String 'Add By Sindy 2019/6/11


Private Sub cmdCancel_Click()
   frm030101_01.Show
   Unload Me
End Sub

Private Sub cmdExit_Click()
   Unload frm030101_01
   Unload Me
End Sub

Private Sub cmdMod_Click()
   frm030101_05.SetData 0, m_TM01, True
   frm030101_05.SetData 1, m_TM02, False
   frm030101_05.SetData 2, m_TM03, False
   frm030101_05.SetData 3, m_TM04, False
   frm030101_05.SetData 4, m_CP09, False
   frm030101_05.SetParent "frm030101_06"
   'Me.Hide
   frm030101_05.Show
   frm030101_05.QueryData
'    m_blnClkChgButton = True
End Sub

Private Sub cmdOK_Click(Index As Integer)
   'Modify By Sindy 2010/11/19 把「確定」及「同時發文」按鈕程式碼合併
   Select Case Index
      Case 0, 1
         If CheckDataValid = True Then
            'Add By Cheng 2002/05/23
            '重新檢查欄位有效性
            If TxtValidate = False Then Exit Sub
            ' 設定滑鼠游標為等待狀態
            Screen.MousePointer = vbHourglass
            ' 更新欄位輸入的內容
            OnUpdateField
            ' 存檔
            'edit by nick 2004/11/03
            'OnSaveData
            If OnSaveData = False Then MsgBox "存檔失敗，請洽系統管理員 !", vbCritical: Screen.MousePointer = vbDefault: Exit Sub
      
            ' 設定滑鼠游標為預設
            Screen.MousePointer = vbDefault
            
            'Add By Sindy 2012/4/5 CFT,FCT所有案件性質發文時,檢查代表圖是否存在
            'Modify by Amy 2018/07/31 ChkIsExistImg不使用
            'Call ChkIsExistImg(m_TM01, m_TM02, m_TM03, m_TM04)
            If ChkImgByteFile(m_TM01, m_TM02, m_TM03, m_TM04) = False Then MsgBox "本案尚未放代表圖至系統！"
            
            'Add By Sindy 2024/8/19
            If frm030101_01.bolIsEMPFlow = True Then
               frm090202_4.QueryData
            End If
            '2024/8/19 End
            If Index = 0 Then '確定鍵
               '*********** 90.11.23   nick  清畫面
               'frm030101_01.radio(0).Value = True
               'frm030101_01.textCP09.Enabled = True
               'frm030101_01.textCP09.Text = ""
               'frm030101_01.textTM01.Enabled = False
               'frm030101_01.textTM01.Text = ""
               'frm030101_01.textTM02.Enabled = False
               'frm030101_01.textTM02.Text = ""
               'frm030101_01.textTM02_2.Enabled = False
               'frm030101_01.textTM02_2.Text = ""
               'frm030101_01.textTM03.Enabled = False
               'frm030101_01.textTM03.Text = "'"
               'frm030101_01.textTM04.Enabled = False
               'frm030101_01.textTM04.Text = ""
               'frm030101_01.grdList.Clear
               'frm030101_01.grdList.Rows = 2
               'frm030101_01.RefreshData
               '***********************************
               'Add By Cheng 2002/04/30
               '若有未發文資料顯示警告
               If PUB_GetCPunIssueDatas("" & Me.textTMKey.Text) = False Then
                  'Add By Sindy 2024/8/19
                  If frm030101_01.bolIsEMPFlow = True Then
                     Unload frm030101_01
                     frm090202_4.Show
                     Unload Me
                     Exit Sub
                  End If
                  '2024/8/19 End
               End If
               frm030101_01.Show
               ' 90.12.07 modify by louis
         '      frm030101_01.Clear
               'Add By Cheng 2002/01/10
               frm030101_01.Clear1
               Unload Me
            ElseIf Index = 1 Then '同時發文鍵
               ' 呼叫第一個畫面
               frm030101_01.SetData 0, m_TM01, True
               frm030101_01.SetData 1, m_TM02, False
               frm030101_01.SetData 2, m_TM03, False
               frm030101_01.SetData 3, m_TM04, False
               frm030101_01.SetQueryFromTM
               Unload Me
               frm030101_01.Show
               frm030101_01.radio(1).Value = True
               frm030101_01.radio_Click 1
               frm030101_01.QueryData
            End If
         End If
      Case Else
   End Select
End Sub

Private Sub cmdPriority_Click()
   'frm880002.strPriority1 = strPriority1
   'frm880002.strPriority2 = strPriority2
   'frm880002.strPriority3 = strPriority3
   'frm880002.Show vbModal
   'strPriority1 = frm880002.strPriority1
   'strPriority2 = frm880002.strPriority2
   'strPriority3 = frm880002.strPriority3
End Sub

Private Sub cmdRelate_Click()
   Where1103ComeFrom Me, m_TM01, m_TM02, m_TM03, m_TM04
End Sub

'Private Sub cmdTogether_Click()
'   If CheckDataValid = True Then
'      'Add By Cheng 2002/07/15
'      '重新檢查欄位有效性
'      If TxtValidate = False Then Exit Sub
'      ' 設定滑鼠游標為等待狀態
'      Screen.MousePointer = vbHourglass
'      ' 更新欄位輸入的內容
'      OnUpdateField
'      ' 存檔
'      'edit by nick 2004/11/03
'      'OnSaveData
'      If OnSaveData = False Then MsgBox "存檔失敗，請洽系統管理員 !", vbCritical: Screen.MousePointer = vbDefault: Exit Sub
'
'      ' 設定滑鼠游標為預設
'      Screen.MousePointer = vbDefault
'
'      ' 呼叫第一個畫面
'      frm030101_01.SetData 0, m_TM01, True
'      frm030101_01.SetData 1, m_TM02, False
'      frm030101_01.SetData 2, m_TM03, False
'      frm030101_01.SetData 3, m_TM04, False
'      frm030101_01.SetQueryFromTM
'      Unload Me
'      frm030101_01.Show
'      frm030101_01.radio(1).Value = True
'      frm030101_01.radio_Click 1
'      frm030101_01.QueryData
'   End If
'End Sub

'Private Sub Form_Activate()
    'Add By Cheng 2003/10/06
    '若有按下變更事項按鈕, 則重新讀取資料
    'edit by nickc 2005/08/23
    'If m_blnClkChgButton = True Then
'Modify By Sindy 2012/10/1 下列程式無意義Mark
'    If m_blnClkChgButton = True Or (pub_ModifyCaseNum = m_TM01 & "-" & m_TM02 & "-" & m_TM03 & "-" & m_TM04 And pub_ModifyCaseNum <> "") Then
'        pub_ModifyCaseNum = ""
'        QueryData
''        m_blnClkChgButton = False
'    End If
'End Sub

Private Sub Form_Load()
   ' 設定控制項的背景顏色
   textTMKey.BackColor = &H8000000F
   textTM12.BackColor = &H8000000F
   textTM15.BackColor = &H8000000F
   textTM20.BackColor = &H8000000F
   textTM45.BackColor = &H8000000F
   textSP08_2.BackColor = &H8000000F
   textSP58_2.BackColor = &H8000000F
   textSP59_2.BackColor = &H8000000F
   
   textCP09.BackColor = &H8000000F
   textCP10.BackColor = &H8000000F
   textCP12.BackColor = &H8000000F
   textCP13.BackColor = &H8000000F
   textCP14.BackColor = &H8000000F
   textCP44_2.BackColor = &H8000000F
   
   SSTab1.Tab = 0
   
   MoveFormToCenter Me
'    m_blnClkChgButton = False
End Sub

Public Sub SetData(ByVal nType As Integer, ByVal strData As String, Optional ByVal bClear As Boolean = False)
Dim rsTmp As New ADODB.Recordset 'Add By Sindy 2012/4/17
   
   ' 清除搜尋的Key
   If bClear = True Then
      m_TM01 = Empty
      m_TM02 = Empty
      m_TM03 = Empty
      m_TM04 = Empty
      m_CP09 = Empty
   End If
   
   Select Case nType
      ' 收文號
      Case 0: m_CP09 = strData
         'Add By Sindy 2012/4/17
         strSql = "SELECT * FROM ChangeEvent " & _
                  "WHERE CE01 = '" & m_CP09 & "' "
         rsTmp.CursorLocation = adUseClient
         rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
         If rsTmp.RecordCount <= 0 Then
            m_blnClkChgButton = True
         Else
            m_blnClkChgButton = False
         End If
         rsTmp.Close
   End Select
End Sub

Private Sub ClearAgentList()
   If m_AgentCount > 0 Then
      Erase m_AgentList
   End If
   m_AgentCount = 0
End Sub

Private Sub AddAgent(ByVal strAgentCode As String, ByVal strAgentName As String)
   Dim nIndex As Integer
   Dim bFind As Boolean
   bFind = False
   For nIndex = 0 To m_AgentCount - 1
      If m_AgentList(nIndex).aiCode = strAgentCode Then
         bFind = True
         Exit For
      End If
   Next nIndex
   If bFind = False Then
      ReDim Preserve m_AgentList(m_AgentCount + 1)
      m_AgentList(m_AgentCount).aiCode = strAgentCode
      m_AgentList(m_AgentCount).aiName = strAgentName
      m_AgentCount = m_AgentCount + 1
   End If
End Sub

' 清除商標基本檔檔案欄位串列
Private Sub ClearTMSPFieldList()
   If m_TMSPCount > 0 Then
      Erase m_TMSPList
   End If
   m_TMSPCount = 0
End Sub

' 設定商標基本檔或服務業務基本檔欄位串列中的欄位內容
Private Sub SetTMSPFieldOldData(ByVal strFieldName As String, ByVal strFieldData As String, ByVal nFieldType As Integer)
   Dim nPos As Integer
   Dim bFind As Boolean
   bFind = False
   For nPos = 0 To m_TMSPCount - 1
      If m_TMSPList(nPos).fiName = strFieldName Then
         bFind = True
         m_TMSPList(nPos).fiOldData = strFieldData
         m_TMSPList(nPos).fiNewData = strFieldData
         m_TMSPList(nPos).fiType = nFieldType
         Exit For
      End If
   Next nPos
   If bFind = False Then
      ReDim Preserve m_TMSPList(m_TMSPCount + 1)
      m_TMSPList(m_TMSPCount).fiName = strFieldName
      m_TMSPList(m_TMSPCount).fiOldData = strFieldData
      m_TMSPList(m_TMSPCount).fiNewData = strFieldData
      m_TMSPList(m_TMSPCount).fiType = nFieldType
      m_TMSPCount = m_TMSPCount + 1
   End If
End Sub

' 設定商標基本檔或服務業務基本檔欄位串列中的欄位內容
Private Sub SetTMSPFieldNewData(ByVal strFieldName As String, ByVal strFieldData As String)
   Dim nPos As Integer
   Dim bFind As Boolean
   bFind = False
   For nPos = 0 To m_TMSPCount - 1
      If m_TMSPList(nPos).fiName = strFieldName Then
         bFind = True
         m_TMSPList(nPos).fiNewData = strFieldData
         Exit For
      End If
   Next nPos
End Sub

' 清除案件進度檔檔案欄位串列
Private Sub ClearCPFieldList()
   If m_CPCount > 0 Then
      Erase m_CPList
   End If
   m_CPCount = 0
End Sub

' 設定案件進度檔欄位串列中的欄位內容
Private Sub SetCPFieldOldData(ByVal strFieldName As String, ByVal strFieldData As String, ByVal nFieldType As Integer)
   Dim nPos As Integer
   Dim bFind As Boolean
   bFind = False
   For nPos = 0 To m_CPCount - 1
      If m_CPList(nPos).fiName = strFieldName Then
         bFind = True
         m_CPList(nPos).fiOldData = strFieldData
         m_CPList(nPos).fiNewData = strFieldData
         m_CPList(nPos).fiType = nFieldType
         Exit For
      End If
   Next nPos
   If bFind = False Then
      ReDim Preserve m_CPList(m_CPCount + 1)
      m_CPList(m_CPCount).fiName = strFieldName
      m_CPList(m_CPCount).fiOldData = strFieldData
      m_CPList(m_CPCount).fiNewData = strFieldData
      m_CPList(m_CPCount).fiType = nFieldType
      m_CPCount = m_CPCount + 1
   End If
End Sub

' 設定案件進度檔欄位串列中的欄位內容
Private Sub SetCPFieldNewData(ByVal strFieldName As String, ByVal strFieldData As String)
   Dim nPos As Integer
   Dim bFind As Boolean
   bFind = False
   For nPos = 0 To m_CPCount - 1
      If m_CPList(nPos).fiName = strFieldName Then
         bFind = True
         m_CPList(nPos).fiNewData = strFieldData
         Exit For
      End If
   Next nPos
End Sub

' 取得服務業務基本檔的欄位內容
Private Sub QueryServicePractice()
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   
   strSql = "SELECT * FROM ServicePractice " & _
            "WHERE SP01 = '" & m_TM01 & "' AND " & _
                  "SP02 = '" & m_TM02 & "' AND " & _
                  "SP03 = '" & m_TM03 & "' AND " & _
                  "SP04 = '" & m_TM04 & "' "
                  
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then

        'add by nickc 2008/02/22
        m_TM44 = CheckStr(rsTmp.Fields("SP26"))
        
      ' 案件中文名稱
      textSP05 = Empty
      If IsNull(rsTmp.Fields("SP05")) = False Then
         textSP05 = rsTmp.Fields("SP05")
      End If
      SetTMSPFieldOldData "SP05", textSP05, 0
      ' 案件英文名稱
      textSP06 = Empty
      If IsNull(rsTmp.Fields("SP06")) = False Then
         textSP06 = rsTmp.Fields("SP06")
      End If
      SetTMSPFieldOldData "SP06", textSP06, 0
      ' 案件日文名稱
      textSP07 = Empty
      If IsNull(rsTmp.Fields("SP07")) = False Then
         textSP07 = rsTmp.Fields("SP07")
      End If
      SetTMSPFieldOldData "SP07", textSP07, 0
      ' 申請人1
      textSP08 = Empty
      If IsNull(rsTmp.Fields("SP08")) = False Then
         textSP08 = rsTmp.Fields("SP08")
         textSP08_2 = GetCustomerName(rsTmp.Fields("SP08"), 0)
      End If
      SetTMSPFieldOldData "SP08", textSP08, 0
      ' 申請人2
      textSP58 = Empty
      If IsNull(rsTmp.Fields("SP58")) = False Then
         textSP58 = rsTmp.Fields("SP58")
         textSP58_2 = GetCustomerName(rsTmp.Fields("SP58"), 0)
      End If
      SetTMSPFieldOldData "SP58", textSP58, 0
      ' 申請人3
      textSP59 = Empty
      If IsNull(rsTmp.Fields("SP59")) = False Then
         textSP59 = rsTmp.Fields("SP59")
         textSP59_2 = GetCustomerName(rsTmp.Fields("SP59"), 0)
      End If
      SetTMSPFieldOldData "SP59", textSP59, 0
      'Add By Cheng 2002/08/23
      m_strCust1 = "" & Me.textSP08.Text
      m_strCust2 = "" & Me.textSP58.Text
      m_strCust3 = "" & Me.textSP59.Text
      
      ' 申請國家
      If IsNull(rsTmp.Fields("SP09")) = False Then
         m_TM10 = rsTmp.Fields("SP09")
      End If
      ' 申請案號
      If IsNull(rsTmp.Fields("SP11")) = False Then
         textTM12 = rsTmp.Fields("SP11")
      End If
      ' 發證日
      If IsNull(rsTmp.Fields("SP12")) = False Then
         textTM20 = DBDATE(rsTmp.Fields("SP12"))
      End If
      ' 案件備註
      textSP18 = Empty
      If IsNull(rsTmp.Fields("SP18")) = False Then
         textSP18 = rsTmp.Fields("SP18")
      End If
      SetTMSPFieldOldData "SP18", textSP18, 0
      ' 代表人
      textSP42 = Empty
      If IsNull(rsTmp.Fields("SP42")) = False Then
         textSP42 = rsTmp.Fields("SP42")
      End If
      SetTMSPFieldOldData "SP42", textSP42, 0
      ' 作品種類
      textSP46 = Empty
      If IsNull(rsTmp.Fields("SP46")) = False Then
         textSP46 = rsTmp.Fields("SP46")
      End If
      SetTMSPFieldOldData "SP46", textSP46, 0
      ' 是否發行
      textSP48 = Empty
      If IsNull(rsTmp.Fields("SP48")) = False Then
         textSP48 = rsTmp.Fields("SP48")
      End If
      SetTMSPFieldOldData "SP48", textSP48, 0
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Sub

' 取得案件進度檔的欄位內容
Private Sub QueryCaseProgress()
   Dim strSql As String
   Dim strSubSQL As String
   Dim rsTmp As New ADODB.Recordset
   Dim rsSubTmp As New ADODB.Recordset
   Dim strDate As String
   Dim strCP27 As String
   Dim strCP44 As String
   Dim strCP45 As String
   Dim nIndex As Integer
   Dim bFind As Boolean
   
   strSql = "SELECT * FROM CaseProgress " & _
            "WHERE CP09 = '" & m_CP09 & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      'add by nickc 2008/02/22
      m_CP116 = CheckStr(rsTmp.Fields("CP116"))
      
      'Add By Sindy 2019/6/11
      '法定期限
      m_CP07 = Empty
      If IsNull(rsTmp.Fields("CP07")) = False Then: m_CP07 = rsTmp.Fields("CP07")
      '2019/6/11 End
      
      ' 案件性質
      'Add By Cheng 2002/07/18
      m_CP10 = Empty: m_CP14 = Empty
      If IsNull(rsTmp.Fields("CP10")) = False Then
         m_CP10 = rsTmp.Fields("CP10")
         If m_TM10 < "010" Then
            textCP10 = GetCaseTypeName(m_TM01, rsTmp.Fields("CP10"), 0)
         Else
            textCP10 = GetCaseTypeName(m_TM01, rsTmp.Fields("CP10"), 1)
         End If
      End If
      ' 業務區別
      If IsNull(rsTmp.Fields("CP12")) = False Then
         textCP12 = GetDepartmentName(rsTmp.Fields("CP12"))
      End If
      ' 智權人員
      If IsNull(rsTmp.Fields("CP13")) = False Then
         textCP13 = GetStaffName(rsTmp.Fields("CP13"))
      End If
      ' 承辦人員
      If IsNull(rsTmp.Fields("CP14")) = False Then
         '92.10.6 ADD BY SONIA
         m_CP14 = rsTmp.Fields("CP14")
         '92'10'6 END
         textCP14 = GetStaffName(rsTmp.Fields("CP14"))
      End If
      ' 發文日(預設為系統日)
      strCP27 = Empty
      'edit by nickc 2006/03/17
      'textCP27 = DBDATE(Date)
      textCP27 = strSrvDate(1)
      If IsNull(rsTmp.Fields("CP27")) = False Then
         strCP27 = rsTmp.Fields("CP27")
      End If
      SetCPFieldOldData "CP27", strCP27, 1
      ' 代理人
      textCP44 = Empty
      If IsNull(rsTmp.Fields("CP44")) = False Then
         textCP44 = rsTmp.Fields("CP44")
      End If
      SetCPFieldOldData "CP44", textCP44, 0
      ' 點數
      textCP18 = Empty
      If IsNull(rsTmp.Fields("CP18")) = False Then
         textCP18 = rsTmp.Fields("CP18")
      End If
      SetCPFieldOldData "CP18", textCP18, 0
      ' 是否算案件數
      textCP26 = Empty
      If IsNull(rsTmp.Fields("CP26")) = False Then
         textCP26 = rsTmp.Fields("CP26")
      End If
      SetCPFieldOldData "CP26", textCP26, 0
      ' 彼所案號
      strCP45 = Empty
      If IsNull(rsTmp.Fields("CP45")) = False Then
         strCP45 = rsTmp.Fields("CP45")
      End If
      SetCPFieldOldData "CP45", strCP45, 0
      ' 進度備註
      textCP64 = Empty
      If IsNull(rsTmp.Fields("CP64")) = False Then
         textCP64 = rsTmp.Fields("CP64")
      End If
      SetCPFieldOldData "CP64", textCP64, 0
      
      ' 代理人
      ClearAgentList
      'Add By Sindy 2013/5/23 若是原先有，也要加入
      If textCP44.Text <> "" Then
'         If InStr(textCP44, "-") > 0 Then
'            If ClsPDGetContact(textCP44, strCP44) Then
'               AddAgent textCP44, strCP44
'            End If
'         Else
            strCP44 = GetFAgentName(textCP44)
            AddAgent textCP44, strCP44
'         End If
      End If
      '2013/5/23 End
      '2010/9/7 Modify by Sindy 文件簽證711及申請英文證明304不要列入
      strSubSQL = "SELECT CP44, MAX(CP27) AS CP27 FROM CASEPROGRESS " & _
                  "WHERE CP01 = '" & m_TM01 & "' AND " & _
                        "CP02 = '" & m_TM02 & "' AND " & _
                        "CP03 = '" & m_TM03 & "' AND " & _
                        "CP04 = '" & m_TM04 & "' AND " & _
                        "CP09 <> '" & m_CP09 & "' And CP09<'C' And CP44 Is Not Null " & _
                        "AND CP10 NOT IN ('711','304') " & _
                  "GROUP BY CP44 " & _
                  "ORDER BY CP27 DESC "
      rsSubTmp.CursorLocation = adUseClient
      rsSubTmp.Open strSubSQL, cnnConnection, adOpenStatic, adLockReadOnly
      If rsSubTmp.RecordCount > 0 Then
         rsSubTmp.MoveFirst
         ' 依序將代理人加入到系統串列中
         Do While rsSubTmp.EOF = False
            If IsNull(rsSubTmp.Fields("CP44")) = False Then
               strCP44 = GetFAgentName(rsSubTmp.Fields("CP44"))
               AddAgent rsSubTmp.Fields("CP44"), GetFAgentName(rsSubTmp.Fields("CP44"))
            End If
            rsSubTmp.MoveNext
         Loop
      End If
      rsSubTmp.Close
      ' 從系統串列中取得所有代理人並放入Combo Box中
      For nIndex = 0 To m_AgentCount - 1
         textCP44.AddItem m_AgentList(nIndex).aiCode
      Next nIndex
      ' 設定顯示為第一筆
      If textCP44.ListCount > 0 Then
         textCP44.ListIndex = 0
         textCP44_Validate False
      End If
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Sub

' 讀取資料庫
Public Sub QueryData()
   Dim strDay As String
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   
   ' 先清除商標基本檔或服務業務基本檔欄位串列
   ClearTMSPFieldList
   ' 先清除案件進度檔欄位串列
   ClearCPFieldList
   
   ' 先取得本所案號
   strSql = "SELECT * FROM CaseProgress " & _
            "WHERE CP09 = '" & m_CP09 & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      ' 本所案號
      If IsNull(rsTmp.Fields("CP01")) = False Then: m_TM01 = rsTmp.Fields("CP01")
      If IsNull(rsTmp.Fields("CP02")) = False Then: m_TM02 = rsTmp.Fields("CP02")
      If IsNull(rsTmp.Fields("CP03")) = False Then: m_TM03 = rsTmp.Fields("CP03")
      If IsNull(rsTmp.Fields("CP04")) = False Then: m_TM04 = rsTmp.Fields("CP04")
   End If
   rsTmp.Close
   
   ' 取得國家代碼
   m_TM10 = GetNationNo(m_TM01, m_TM02, m_TM03, m_TM04)
   
   ' 本所案號
'   textTMKey = m_TM01 & m_TM02 & m_TM03 & m_TM04
   textTMKey.Text = m_TM01 & "-" & m_TM02 & "-" & IIf(Len("" & m_TM03) <= 0, "0", m_TM03) & "-" & IIf(Len("" & m_TM04) <= 0, "00", m_TM04)

   ' 收文號
   textCP09 = m_CP09
   
   ' 取得案件進度檔的欄位
   QueryCaseProgress
   ' 讀取服務業務基本檔
   QueryServicePractice
   
   ' 大約?可接獲回音(欄位)
   textCF09 = Empty
   strSql = "SELECT * FROM CaseFee " & _
            "WHERE CF01 = '" & m_TM01 & "' AND " & _
                  "CF02 = '" & m_TM10 & "' AND " & _
                  "CF03 = '" & m_CP10 & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("CF09")) = False Then
         textCF09 = rsTmp.Fields("CF09")
      End If
   End If
   rsTmp.Close
   
   ' 計算催審期限
   strDay = GetUrgeDate(m_TM01, m_TM10, m_CP10, textCP27)
   If IsEmptyText(strDay) = False Then
      textUargeDate = strDay
   End If
   textCP27.Tag = textCP27.Text 'Added by Lydia 2019/11/08 記錄發文日，有修改發文日再重新計算催審期限
   
   Set rsTmp = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
   'Add By Cheng 2002/07/19
   Set frm030101_06 = Nothing
End Sub

' 點數
Private Sub textCP18_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textCP18) = False Then
      If IsNumeric(textCP18) = False Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "點數只可輸入數值資料"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP18_GotFocus
      End If
   End If
End Sub

Private Sub textCP26_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

' 是否算案件數
Private Sub textCP26_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textCP26) = False Then
      Select Case textCP26
         Case " ", "N":
         Case Else
            Cancel = True
            strTit = "資料檢核"
            strMsg = "只可輸入空白或N"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textCP26_GotFocus
      End Select
   End If
End Sub

' 發文日
Private Sub textCP27_Validate(Cancel As Boolean)
   Dim strDay As String
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Dim strDate As String
   Cancel = False
   
   If IsEmptyText(textCP27) = False Then
      ' 發文日日期不正確
      If CheckIsDate(textCP27, False) = False Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "請輸入正確的發文日"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP27_GotFocus
         GoTo EXITSUB
      End If
      
      ' 發文日日期不可超過系統日
      'edit by nickc 2006/03/17
      'If Val(DBDATE(textCP27)) > Val(DBDATE(Date)) Then
      If Val(DBDATE(textCP27)) > Val(strSrvDate(1)) Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "發文日不可超過系統日"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP27_GotFocus
         GoTo EXITSUB
      End If
      
      ' 計算催審期限
      If Me.textCP27.Tag <> Me.textCP27.Text Then 'Added by Lydia 2019/11/08 記錄發文日，有修改發文日再重新計算催審期限
            strDay = GetUrgeDate(m_TM01, m_TM10, m_CP10, textCP27)
            If IsEmptyText(strDay) = False Then
               textUargeDate = strDay
            End If
      'Added by Lydia 2019/11/08
      End If
      Me.textCP27.Tag = Me.textCP27.Text
      'end 2019/11/08
   End If
EXITSUB:
End Sub

Private Sub textCP44_Click()
   textCP44_2 = m_AgentList(textCP44.ListIndex).aiName
End Sub

' 當使用者按向下鍵時, 將ComboBox顯示成下拉式的樣子
Private Sub textCP44_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyDown Then
      SendMessage textCP44.hWnd, CB_SHOWDROPDOWN, True, ByVal 0&
   End If
End Sub

Private Sub textCP44_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

' 代理人
Private Sub textCP44_Validate(Cancel As Boolean)
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Dim strTempName As String   '2010/11/24 add by sonia
   
   Cancel = False
   'Add By Cheng 2002/03/08
   If m_TM10 <> 台灣國家代號 Then
      If Len(Me.textCP44.Text) <= 0 Then
         MsgBox "當申請國家非台灣時, 代理人欄不可為空白!!!", vbExclamation
         Cancel = True
         Exit Sub
      End If
   End If
   
   If textCP44.ListIndex >= 0 Then
      textCP44 = m_AgentList(textCP44.ListIndex).aiCode
   End If
   textCP44_2 = Empty
   If IsEmptyText(textCP44) = False Then
      'edit by 2004/07/22 nick  檢查該申請人或代理人狀態，若為不再使用則停在原地
      '2010/11/24 modify by sonia 取消basQuery的GetFAgentNameAndState
      'Dim oState As Boolean
      'oState = True
      ''textCP44_2 = GetFAgentName(textCP44)
      'textCP44_2 = GetFAgentNameAndState(textCP44, oState)
      'If oState = False Then
      '       Cancel = True
      '        Exit Sub
      'End If
      If PUB_GetAgentNameAndState(m_TM01, textCP44.Text, strTempName) Then
         textCP44_2 = strTempName
      Else
         textCP44_2.Text = ""
         If strTempName <> "" Then
            Cancel = True
            Exit Sub
         End If
      End If
      '2010/11/24 end
      If IsEmptyText(textCP44_2) = True Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "代理人不存在"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP44_GotFocus
      Else
         ' 依所選擇的代理人找出案件進度檔中其收文日最大的一筆其彼所案號更新到畫面上的彼所案號欄位
         strSql = "SELECT CP45 FROM CaseProgress " & _
                  "WHERE CP01 = '" & m_TM01 & "' AND " & _
                        "CP02 = '" & m_TM02 & "' AND " & _
                        "CP03 = '" & m_TM03 & "' AND " & _
                        "CP04 = '" & m_TM04 & "' AND " & _
                        "CP44 = '" & textCP44 & "' AND " & _
                        "CP05 IN (SELECT MAX(CP05) FROM CASEPROGRESS " & _
                                 "WHERE CP01 = '" & m_TM01 & "' AND " & _
                                       "CP02 = '" & m_TM02 & "' AND " & _
                                       "CP03 = '" & m_TM03 & "' AND " & _
                                       "CP04 = '" & m_TM04 & "' AND " & _
                                       "CP44 = '" & textCP44 & "')"
         rsTmp.CursorLocation = adUseClient
         rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
         If rsTmp.RecordCount > 0 Then
            rsTmp.MoveFirst
            If IsNull(rsTmp.Fields("CP45")) = False Then
               textTM45 = rsTmp.Fields("CP45")
            End If
         End If
         rsTmp.Close
      End If
   End If
   Set rsTmp = Nothing
End Sub

' 進度備註
Private Sub textCP64_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If CheckLengthIsOK(textCP64, 2000) = False Then
      Cancel = True
      strTit = "檢核資料"
      strMsg = "進度備註內容太長"
      'nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textCP64_GotFocus
   End If
End Sub

Private Sub textSP08_KeyPress(KeyAscii As Integer)
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
      'nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textSP18_GotFocus
   End If
End Sub

Private Sub textPrint_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

' 列印定稿
Private Sub textPrint_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
      
   If IsEmptyText(textPrint) = False Then
      Select Case textPrint
         Case " ", "N":
         Case Else:
            Cancel = True
            strTit = "資料檢核"
            strMsg = "只可輸入空白或N"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textPrint_GotFocus
      End Select
   End If
End Sub

Private Sub textPrintLetter_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

' 是否列印指示信
Private Sub textPrintLetter_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textPrintLetter) = False Then
      Select Case textPrintLetter
         Case " ", "N":
         Case Else
            Cancel = True
            strTit = "資料檢核"
            strMsg = "只可輸入空白或N"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textPrintLetter_GotFocus
      End Select
   End If
End Sub

' 更新欄位的內容
Private Sub OnUpdateField()
   SetCPFieldNewData "CP18", textCP18
   SetCPFieldNewData "CP26", textCP26
   SetCPFieldNewData "CP27", DBDATE(textCP27)
   ' 代理人
   If IsEmptyText(textCP44) = False Then
      SetCPFieldNewData "CP44", textCP44 & String(9 - Len(textCP44), "0")
      'add by nickc 2008/02/22
      m_CP44New = textCP44 & String(9 - Len(textCP44), "0")
   Else
      SetCPFieldNewData "CP44", textCP44
      'add by nickc 2008/02/22
      m_CP44New = textCP44
   End If
   SetCPFieldNewData "CP45", textTM45
   SetCPFieldNewData "CP64", textCP64
   
   SetTMSPFieldNewData "SP05", textSP05
   SetTMSPFieldNewData "SP06", textSP06
   SetTMSPFieldNewData "SP07", textSP07
   ' 申請人
   If IsEmptyText(textSP08) = False Then
      SetTMSPFieldNewData "SP08", textSP08 & String(9 - Len(textSP08), "0")
   Else
      SetTMSPFieldNewData "SP08", textSP08
   End If
   SetTMSPFieldNewData "SP18", textSP18
   SetTMSPFieldNewData "SP42", textSP42
   SetTMSPFieldNewData "SP46", textSP46
   SetTMSPFieldNewData "SP48", textSP48
   If IsEmptyText(textSP58) = False Then
      SetTMSPFieldNewData "SP58", textSP58 & String(9 - Len(textSP58), "0")
   Else
      SetTMSPFieldNewData "SP58", textSP58
   End If
   If IsEmptyText(textSP59) = False Then
      SetTMSPFieldNewData "SP59", textSP59 & String(9 - Len(textSP59), "0")
   Else
      SetTMSPFieldNewData "SP59", textSP59
   End If
      
End Sub

' 更新服務業務基本檔的相關欄位
Private Sub OnUpdateServicePractice()
   Dim strTmp As String
   Dim strSql As String
   Dim nIndex As Integer
   Dim bFirst As Boolean
   Dim bDifference As Boolean
      
   ' 更新案件進度檔
   strSql = "UPDATE ServicePractice SET "
   bFirst = True
   bDifference = False
   For nIndex = 0 To m_TMSPCount - 1
      strTmp = Empty
      If m_TMSPList(nIndex).fiOldData <> m_TMSPList(nIndex).fiNewData Then
         If m_TMSPList(nIndex).fiType = 0 Then
            ' 91.03.25 modify by louis (單引號)
            'strTmp = m_TMSPList(nIndex).fiName & " = '" & m_TMSPList(nIndex).fiNewData & "'"
            strTmp = m_TMSPList(nIndex).fiName & " = '" & ChgSQL(m_TMSPList(nIndex).fiNewData) & "'"
         Else
            If m_TMSPList(nIndex).fiNewData = Empty Then
               strTmp = m_TMSPList(nIndex).fiName & " = " & 0
            Else
               strTmp = m_TMSPList(nIndex).fiName & " = " & m_TMSPList(nIndex).fiNewData
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
   Next nIndex
   ' 設定SQL語法更新的條件
   strSql = strSql & " " & _
                  "WHERE SP01 = '" & m_TM01 & "' AND " & _
                        "SP02 = '" & m_TM02 & "' AND " & _
                        "SP03 = '" & m_TM03 & "' AND " & _
                        "SP04 = '" & m_TM04 & "' "
   ' 執行SQL指令
   If bDifference = True Then: cnnConnection.Execute strSql
End Sub

'edit b nick 2004/11/03
'Private sub OnSaveData()
Public Function OnSaveData() As Boolean
OnSaveData = True
   Dim strTmp As String
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   Dim nIndex As Integer
   Dim bFirst As Boolean
   Dim bDifference As Boolean
   Dim strNP08 As String
   Dim strNP07 As String
   Dim strNP22 As String
      
'911106 nick transation
On Error GoTo CheckingErr
cnnConnection.BeginTrans

   ' 更新案件進度檔
   strSql = "UPDATE CaseProgress SET "
   bFirst = True
   bDifference = False
   For nIndex = 0 To m_CPCount - 1
      strTmp = Empty
      If m_CPList(nIndex).fiOldData <> m_CPList(nIndex).fiNewData Then
         If m_CPList(nIndex).fiType = 0 Then
            ' 91.03.25 modify by louis (單引號)
            'strTmp = m_CPList(nIndex).fiName & " = '" & m_CPList(nIndex).fiNewData & "'"
            strTmp = m_CPList(nIndex).fiName & " = '" & ChgSQL(m_CPList(nIndex).fiNewData) & "'"
         Else
            If m_CPList(nIndex).fiNewData = Empty Then
               strTmp = m_CPList(nIndex).fiName & " = " & 0
            Else
               strTmp = m_CPList(nIndex).fiName & " = " & m_CPList(nIndex).fiNewData
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
   Next nIndex
   ' 設定SQL語法更新的條件
   strSql = strSql & " " & _
                  "WHERE CP09 = '" & m_CP09 & "' "
   ' 執行SQL指令
   If bDifference = True Then: cnnConnection.Execute strSql
   
   ' 更新服務業務基本檔
   OnUpdateServicePractice
   
   ' 若有輸入催審期限時, 新增一筆催審的記錄到下一程序檔
   If IsEmptyText(textUargeDate) = False Then
      strNP07 = "305"
      strNP22 = GetNextProgressNo()
      '92.10.6 MODIFY BY SONIA
      'strSQL = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
      '         "VALUES ('" & m_CP09 & "','" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & strNP07 & "," & _
      '                    DBDATE(textUargeDate) & "," & DBDATE(textUargeDate) & ",'" & strUserNum & "'," & strNP22 & ")"
      'Modified by Lydia 2020/07/09 本所期限檢查：若本所期限非工作天則直接調整至最近的工作天
      'strSql = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
               "VALUES ('" & m_CP09 & "','" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & strNP07 & "," & _
                          DBDATE(textUargeDate) & "," & DBDATE(textUargeDate) & ",'" & m_CP14 & "'," & strNP22 & ")"
      strSql = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
               "VALUES ('" & m_CP09 & "','" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & strNP07 & "," & _
                          PUB_GetWorkDay1(textUargeDate, True) & "," & DBDATE(textUargeDate) & ",'" & m_CP14 & "'," & strNP22 & ")"
      '92.10.6 END
      cnnConnection.Execute strSql
      ' 延展, 使用宣誓, 刊登廣告, 繳年費, 催審, 提申, 收達不印接洽結案單
      Select Case strNP07
         Case "102", "105", "702", "708", "305", "998", "997":
         Case Else:
            ' 列印國內案件接洽及結案記錄單
'            g_PrtForm001.PrintForm strNP22, m_TM01, m_TM02, m_TM03, m_TM04
            'Add By Cheng 2004/04/08
            '新增列印接洽結案單資料
            pub_AddressListSN = pub_AddressListSN + 1
            PUB_AddNewCaseCloseSheet strUserNum, "" & pub_AddressListSN, "" & strNP22, "" & m_TM01, "" & m_TM02, "" & m_TM03, "" & m_TM04
      End Select
   End If
   
   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   ' 若有定義代理人收達天數時, 新增一筆收達的記錄到下一程序檔
   strSql = "SELECT * FROM CaseFee " & _
            "WHERE CF01 = '" & m_TM01 & "' AND " & _
                  "CF02 = '" & m_TM10 & "' AND " & _
                  "CF03 = '" & m_CP10 & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      If IsNull(rsTmp.Fields("CF23")) = False Then
         strNP07 = "997"
         strNP08 = DBDATE(textCP27)
        'Modify By Cheng 2003/09/02
'         strNP08 = DBDATE(Format(DateSerial(Val(DBYEAR(strNP08)), Val(DBMONTH(strNP08)), Val(DBDAY(strNP08)) + Val(rsTmp.Fields("CF23")))))
         strNP08 = DBDATE(DateAdd("d", Val(rsTmp.Fields("CF23")), ChangeWStringToWDateString(DBDATE(strNP08))))
         'Add By Sindy 2019/6/11 檢查期限是否正確
         strNP08 = PUB_T997998LimitDate(strNP08, m_CP07, 1)
         '2019/6/11 END
         strNP22 = GetNextProgressNo()
         '92.10.6 MODIFY BY SONIA
         'strSQL = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
         '         "VALUES ('" & m_CP09 & "','" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & strNP07 & "," & _
         '                   strNP08 & "," & strNP08 & ",'" & strUserNum & "'," & strNP22 & ")"
         'Modified by Lydia 2020/07/09 本所期限檢查：若本所期限非工作天則直接調整至最近的工作天
         'strSql = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
                  "VALUES ('" & m_CP09 & "','" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & strNP07 & "," & _
                            strNP08 & "," & strNP08 & ",'" & m_CP14 & "'," & strNP22 & ")"
         strSql = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
                  "VALUES ('" & m_CP09 & "','" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & strNP07 & "," & _
                            PUB_GetWorkDay1(strNP08, True) & "," & strNP08 & ",'" & m_CP14 & "'," & strNP22 & ")"
         '92.10.6 END
         cnnConnection.Execute strSql
         ' 延展, 使用宣誓, 刊登廣告, 繳年費, 催審, 提申, 收達不印接洽結案單
         Select Case strNP07
            Case "102", "105", "702", "708", "305", "998", "997":
            Case Else:
               ' 列印國內案件接洽及結案記錄單
'               g_PrtForm001.PrintForm strNP22, m_TM01, m_TM02, m_TM03, m_TM04
                'Add By Cheng 2004/04/08
                '新增列印接洽結案單資料
                pub_AddressListSN = pub_AddressListSN + 1
                PUB_AddNewCaseCloseSheet strUserNum, "" & pub_AddressListSN, "" & strNP22, "" & m_TM01, "" & m_TM02, "" & m_TM03, "" & m_TM04
         End Select
      End If
   End If
   rsTmp.Close
   
   'Added by Lydia 2024/07/09 判斷案件國家收費表內有設定提申期限(天)CF11，要加掛提申(998)期限；
   Call Pub_GetCF11to998(m_TM10, m_TM01, m_TM02, m_TM03, m_TM04, m_CP07, m_CP09, m_CP10, m_CP14, textCP27)
      
   'Add By Sindy 2016/12/20
   If m_990CP09 <> "" Then
      strSql = "update caseprogress set cp27=" & strSrvDate(1) & " where cp09='" & m_990CP09 & "' and cp27 is null"
      cnnConnection.Execute strSql
   End If
   '2016/12/20 END
   
   Set rsTmp = Nothing
'911106 nick transation
    cnnConnection.CommitTrans
   
     'Add by nickc 2008/02/22 檢查代理人Email(需考慮可能為FF案件)
    PUB_CheckEMail m_CP44New, m_CP116
    PUB_CheckEMail m_TM44, m_TM119
    If m_TM120 <> "" Then
       PUB_CheckEMail m_TM44, m_TM120
    End If
    'end 2008/02/22
    
   ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   ' 列印定稿
   If textPrint <> "N" Then
      PrintLetter
   End If
   
   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   ' 列印指示信
   If textPrintLetter <> "N" Then
      PrintLetter_2
   End If
   
    Exit Function
CheckingErr:
    MsgBox (Err.Description)
     cnnConnection.RollbackTrans
   OnSaveData = False
End Function

' 案件中文名稱
Private Sub textSP05_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If CheckLengthIsOK(textSP05, 60) = False Then
      Cancel = True
      strTit = "檢核資料"
      strMsg = "案件中文名稱內容太長"
      'nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textSP05_GotFocus
   End If
End Sub

' 案件英文名稱
Private Sub textSP06_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If CheckLengthIsOK(textSP06, 60) = False Then
      Cancel = True
      strTit = "檢核資料"
      strMsg = "案件英文名稱內容太長"
      'nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textSP06_GotFocus
   End If
End Sub

' 案件日文名稱
Private Sub textSP07_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If CheckLengthIsOK(textSP07, 60) = False Then
      Cancel = True
      strTit = "檢核資料"
      strMsg = "案件日文名稱內容太長"
      'nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textSP07_GotFocus
   End If
End Sub

' 申請人1
Private Sub textSP08_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   Cancel = False
   textSP08_2 = Empty
   If IsEmptyText(textSP08) = False Then
        Me.textSP08.Text = ChangeCustomerL(Me.textSP08.Text)
      'edit by 2004/07/22 nick  檢查該申請人或代理人狀態，若為不再使用則停在原地
      Dim oState As Boolean
      oState = True
      'textSP08_2 = GetCustomerName(textSP08, 0)
      textSP08_2 = GetCustomerNameAndState(textSP08, 0, oState)
      If oState = False Then
             Cancel = True
             Exit Sub
      End If
      If textSP08_2 = Empty Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "申請人代碼<" & textSP08 & ">不正確"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textSP08_GotFocus
      End If
   End If
   'Add By Cheng 2002/08/22
   If Cancel = False Then
      If Me.textSP08.Text <> m_strCust1 Then
         If Not PUB_EditCustOk(m_CP09, m_TM01, m_TM02, m_TM03, m_TM04) Then Cancel = True
      End If
   End If
   If Cancel = True Then textSP08_GotFocus

End Sub

Private Sub textSP48_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

' 是否發行
Private Sub textSP48_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textSP48) = False Then
      Select Case textSP48
         Case " ", "Y":
         Case Else
            Cancel = True
            strTit = "資料檢核"
            strMsg = "只可輸入空白或Y"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textSP48_GotFocus
      End Select
   End If
End Sub

Private Sub textSP58_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

' 申請人2
Private Sub textSP58_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   Cancel = False
   textSP58_2 = Empty
   If IsEmptyText(textSP58) = False Then
        Me.textSP58.Text = ChangeCustomerL(Me.textSP58.Text)
      'edit by 2004/07/22 nick  檢查該申請人或代理人狀態，若為不再使用則停在原地
      Dim oState As Boolean
      oState = True
      'textSP58_2 = GetCustomerName(textSP58, 0)
      textSP58_2 = GetCustomerNameAndState(textSP58, 0, oState)
      If oState = False Then
            Cancel = True
            Exit Sub
      End If
      If textSP58_2 = Empty Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "申請人代碼<" & textSP58 & ">不正確"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textSP58_GotFocus
      End If
   End If
   'Add By Cheng 2002/08/22
   If Cancel = False Then
      If Me.textSP58.Text <> m_strCust2 Then
         If Not PUB_EditCustOk(m_CP09, m_TM01, m_TM02, m_TM03, m_TM04) Then Cancel = True
      End If
   End If
   If Cancel = True Then textSP58_GotFocus
   
End Sub

Private Sub textSP59_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

' 申請人3
Private Sub textSP59_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   Cancel = False
   textSP59_2 = Empty
   If IsEmptyText(textSP59) = False Then
        Me.textSP59.Text = ChangeCustomerL(Me.textSP59.Text)
      'edit by 2004/07/22 nick  檢查該申請人或代理人狀態，若為不再使用則停在原地
      Dim oState As Boolean
      'textSP59_2 = GetCustomerName(textSP59)
      textSP59_2 = GetCustomerNameAndState(textSP59, "0", oState)
      If oState = False Then
           Cancel = True
           Exit Sub
      End If
      If textSP59_2 = Empty Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "申請人代碼<" & textSP59 & ">不正確"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textSP59_GotFocus
      End If
   End If
   'Add By Cheng 2002/08/22
   If Cancel = False Then
      If Me.textSP59.Text <> m_strCust3 Then
         If Not PUB_EditCustOk(m_CP09, m_TM01, m_TM02, m_TM03, m_TM04) Then Cancel = True
      End If
   End If
   If Cancel = True Then textSP59_GotFocus

End Sub

' 催審期限
Private Sub textUargeDate_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textUargeDate) = False Then
      If CheckIsDate(textUargeDate, False) = False Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "催審期限日期不正確"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textUargeDate_GotFocus
      End If
   End If
End Sub

Private Function CheckDataValid() As Boolean
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   CheckDataValid = False
   'add by nickc 2008/05/01
   If IsDebt(m_TM10, textCP09) Then
        strTit = "警告！禁止發文！"
        strMsg = "未收款且無 預定收款日 請轉告智權同仁！！"
        nResponse = MsgBox(strMsg, vbOKOnly, strTit)
        GoTo EXITSUB
   End If
   
   'Add By Sindy 2012/4/17
   If m_blnClkChgButton = False Then
      MsgBox "請輸入變更事項!!!", vbExclamation + vbOKOnly
      Me.cmdMod.SetFocus
      GoTo EXITSUB
   End If
   
   ' 案件名稱(中, 英, 日)
   If IsEmptyText(textSP05) = True And IsEmptyText(textSP06) = True And IsEmptyText(textSP07) = True Then
      strTit = "檢核資料"
      strMsg = "請輸入案件名稱"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textSP05.SetFocus
      GoTo EXITSUB
   End If
   ' 發文日
   If IsEmptyText(textCP27) = True Then
      strTit = "檢核資料"
      strMsg = "請輸入發文日"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textCP27.SetFocus
      GoTo EXITSUB
   End If
   ' 代理人
   If IsEmptyText(textCP44) = True Then
      strTit = "檢核資料"
      strMsg = "請輸入代理人"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textCP44.SetFocus
      GoTo EXITSUB
   End If
   
    'add by nickc 2006/03/17 加入驗證
    Dim Cancel As Boolean
    Cancel = False
    textCP27_Validate Cancel
    If Cancel = True Then GoTo EXITSUB
    
   CheckDataValid = True
EXITSUB:
End Function

Private Sub textUargeDate_GotFocus()
   InverseTextBox textUargeDate
End Sub

Private Sub textPrint_GotFocus()
   InverseTextBox textPrint
End Sub

Private Sub textAttach_GotFocus()
   InverseTextBox textAttach
End Sub

Private Sub textPrintLetter_GotFocus()
   InverseTextBox textPrintLetter
End Sub

Private Sub textSP05_GotFocus()
   InverseTextBox textSP05
End Sub

Private Sub textSP06_GotFocus()
   InverseTextBox textSP06
End Sub

Private Sub textSP07_GotFocus()
   InverseTextBox textSP07
End Sub

Private Sub textSP08_GotFocus()
   InverseTextBox textSP08
End Sub

Private Sub textSP18_GotFocus()
   InverseTextBox textSP18
End Sub

Private Sub textSP42_GotFocus()
   InverseTextBox textSP42
End Sub

Private Sub textSP46_GotFocus()
   InverseTextBox textSP46
End Sub

Private Sub textSP48_GotFocus()
   InverseTextBox textSP48
End Sub

Private Sub textSP58_GotFocus()
   InverseTextBox textSP58
End Sub

Private Sub textSP59_GotFocus()
   InverseTextBox textSP59
End Sub

Private Sub textCP18_GotFocus()
   InverseTextBox textCP18
End Sub

Private Sub textCP26_GotFocus()
   InverseTextBox textCP26
End Sub

Private Sub textCP27_GotFocus()
   InverseTextBox textCP27
End Sub

Private Sub textCP44_GotFocus()
   InverseTextBox textCP44
End Sub

Private Sub textCP64_GotFocus()
   InverseTextBox textCP64
End Sub

Private Sub textCF09_GotFocus()
   InverseTextBox textCF09
End Sub

' 列印定稿前將例外欄位加入到列印定稿例外欄位檔案中
Private Sub InsExpField()
   Dim strSql As String
   Dim strTemp As String
   
   If checkAttach(0).Value = 1 Then
      If strTemp <> Empty Then: strTemp = strTemp & "及"
      strTemp = strTemp & "申請書"
   End If
   If checkAttach(1).Value = 1 Then
      If strTemp <> Empty Then: strTemp = strTemp & "及"
      strTemp = strTemp & "照片"
   End If
   If checkAttach(2).Value = 1 Then
      If strTemp <> Empty Then: strTemp = strTemp & "及"
      strTemp = strTemp & "圖示"
   End If
   If checkAttach(3).Value = 1 Then
      If strTemp <> Empty Then: strTemp = strTemp & "及"
      strTemp = strTemp & "書"
   End If
   If IsEmptyText(strTemp) = False Then
      strTemp = "，隨函附上" & strTemp
      strTemp = strTemp & "各一份，以供查存"
   End If
   
   ' 系統類別為CFC
   If m_TM01 = "CFC" Then
      ' 案件性質為著作權申請
      If m_CP10 = "806" Then
         ' 清除定稿例外欄位檔原有資料
         EndLetter "01", m_CP09, "04", strUserNum
         ' 函知附件
         strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                  "VALUES ('" & "01" & "','" & m_CP09 & "','" & "04" & "','" & strUserNum & _
                  "','函知附件','" & strTemp & "')"
         cnnConnection.Execute strSql
      End If
   End If
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 列印定稿
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub PrintLetter()
   ' 先呼叫定稿程式的清除原定稿資料的函式去清除之前殘留在例外欄位檔中的資料
   InsExpField
   
   ' 系統類別為CFC
   If m_TM01 = "CFC" Then
      ' 案件性質為著作權申請
      If m_CP10 = "806" Then
         ' 列印定稿
         NowPrint m_CP09, "01", "04", False, strUserNum, 0
      End If
   End If
End Sub

' 列印指示信前將例外欄位加入到列印定稿例外欄位檔案中
Private Sub InsExpField_2()
   Dim strSql As String
   ' 發行否
   Dim strPublish As String
      
   If textSP48 = "Y" Then
      strPublish = "published"
   Else
      strPublish = "not yet published"
   End If
   
   ' 系統類別為CFC
   If m_TM01 = "CFC" Then
      ' 案件性質為著作權申請
      If m_CP10 = "806" Then
         ' 清除定稿例外欄位檔原有資料
         EndLetter "01", m_CP09, "31", strUserNum
         ' 附件種類
         strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                  "VALUES ('" & "01" & "','" & m_CP09 & "','" & "31" & "','" & strUserNum & _
                  "','附件種類','" & textAttach & "')"
         cnnConnection.Execute strSql
         ' 發行否
         strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                  "VALUES ('" & "01" & "','" & m_CP09 & "','" & "31" & "','" & strUserNum & _
                  "','發行否','" & strPublish & "')"
         cnnConnection.Execute strSql
      End If
   End If
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 列印指示信
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub PrintLetter_2()
   ' 先呼叫定稿程式的清除原定稿資料的函式去清除之前殘留在例外欄位檔中的資料
   InsExpField
   
   ' 系統類別為CFC
   If m_TM01 = "CFC" Then
      ' 案件性質為著作權申請
      If m_CP10 = "806" Then
         ' 列印只指示信
         NowPrint m_CP09, "01", "31", False, strUserNum, 0
      End If
   End If
End Sub

'Add By Cheng 2002/05/23
Private Function TxtValidate() As Boolean
Dim objTxt As Object
Dim ii As Integer
Dim Cancel As Boolean
   
   TxtValidate = False
   If Me.textCP18.Enabled = True Then
      Cancel = False
      textCP18_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textCP26.Enabled = True Then
      Cancel = False
      textCP26_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textCP27.Enabled = True Then
      Cancel = False
      textCP27_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textCP44.Enabled = True Then
      Cancel = False
      textCP44_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textCP64.Enabled = True Then
      Cancel = False
      textCP64_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textPrint.Enabled = True Then
      Cancel = False
      textPrint_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textPrintLetter.Enabled = True Then
      Cancel = False
      textPrintLetter_Validate Cancel
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
   
   If Me.textSP06.Enabled = True Then
      Cancel = False
      textSP06_Validate Cancel
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
   
   If Me.textSP18.Enabled = True Then
      Cancel = False
      textSP18_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textSP48.Enabled = True Then
      Cancel = False
      textSP48_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textSP58.Enabled = True Then
      Cancel = False
      textSP58_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textSP59.Enabled = True Then
      Cancel = False
      textSP59_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textUargeDate.Enabled = True Then
      Cancel = False
      textUargeDate_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   'Add By Sindy 2016/12/20
   '檢查有設定副本收受人需提醒並新增信函副本B類收文
   m_990CP09 = ""
   If textPrint = "N" Then '不印定稿
      If PUB_ChkCC(m_TM01, m_TM02, m_TM03, m_TM04, m_CP09, m_990CP09) = False Then
         SSTab1.Tab = 0
         Exit Function
      End If
   End If
   '2016/12/20 END
   
   'Added by Lydia 2021/08/20 檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
   If PUB_ChkUniText(Me, , True, "TextBox") = False Then
        Exit Function
   End If
      
   TxtValidate = True
End Function

