VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm030001_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "CF 指示信"
   ClientHeight    =   5700
   ClientLeft      =   50
   ClientTop       =   340
   ClientWidth     =   9180
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5700
   ScaleWidth      =   9180
   Begin VB.CommandButton cmdPath 
      Height          =   315
      Left            =   8592
      Picture         =   "frm030001_1.frx":0000
      Style           =   1  '圖片外觀
      TabIndex        =   64
      Top             =   5328
      Width           =   330
   End
   Begin VB.TextBox txtPath 
      Height          =   315
      Left            =   1944
      TabIndex        =   62
      Top             =   5328
      Width           =   6636
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "商品及服務資料輸入(&I)"
      Height          =   345
      Index           =   3
      Left            =   3570
      Style           =   1  '圖片外觀
      TabIndex        =   11
      Top             =   30
      Width           =   2235
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "結束(&X)"
      Height          =   345
      Index           =   2
      Left            =   8070
      TabIndex        =   14
      Top             =   30
      Width           =   1005
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "回前畫面(&U)"
      Height          =   345
      Index           =   1
      Left            =   6900
      TabIndex        =   13
      Top             =   30
      Width           =   1125
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   345
      Index           =   0
      Left            =   5850
      TabIndex        =   12
      Top             =   30
      Width           =   1005
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3735
      Left            =   60
      TabIndex        =   27
      Top             =   1530
      Width           =   9045
      _ExtentX        =   15963
      _ExtentY        =   6579
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "基本資料"
      TabPicture(0)   =   "frm030001_1.frx":0102
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label4(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label6"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label9"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label1(7)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label20"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label19"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label8"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label7"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label4(1)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label4(2)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "textCP44_2"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "textTM23"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "textTM23_2C"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "textTM05"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "textTM09"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "textTM23_2E"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "textTM23_2J"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "textTM06"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "textTM07"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "textTM05_1"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "textPrintTNT"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "textCP44"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "txtLetterHead"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "cmdSuggest"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "textPrintDHL"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).ControlCount=   25
      TabCaption(1)   =   "代表人"
      TabPicture(1)   =   "frm030001_1.frx":011E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdPriority"
      Tab(1).Control(1)=   "Combo2(1)"
      Tab(1).Control(2)=   "Combo2(0)"
      Tab(1).Control(3)=   "textTM67"
      Tab(1).Control(4)=   "textTM52"
      Tab(1).Control(5)=   "textTM51"
      Tab(1).Control(6)=   "textTM49"
      Tab(1).Control(7)=   "textTM48"
      Tab(1).Control(8)=   "textTM47"
      Tab(1).Control(9)=   "textTM50"
      Tab(1).Control(10)=   "Label26"
      Tab(1).Control(11)=   "Label21"
      Tab(1).Control(12)=   "Label5(8)"
      Tab(1).Control(13)=   "Label5(7)"
      Tab(1).Control(14)=   "Label5(6)"
      Tab(1).Control(15)=   "Label5(5)"
      Tab(1).Control(16)=   "Label5(4)"
      Tab(1).Control(17)=   "Label5(3)"
      Tab(1).Control(18)=   "Label14(1)"
      Tab(1).Control(19)=   "Label18(2)"
      Tab(1).ControlCount=   20
      Begin VB.TextBox textPrintDHL 
         Height          =   300
         Left            =   4080
         MaxLength       =   1
         TabIndex        =   5
         Top             =   3240
         Width           =   375
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
         Height          =   315
         Left            =   660
         TabIndex        =   60
         Top             =   450
         Width           =   300
      End
      Begin VB.TextBox txtLetterHead 
         Enabled         =   0   'False
         Height          =   300
         Left            =   7035
         MaxLength       =   1
         TabIndex        =   6
         Top             =   3240
         Width           =   375
      End
      Begin VB.CommandButton cmdPriority 
         Caption         =   "輸入(&V)"
         Height          =   300
         Left            =   -73830
         TabIndex        =   8
         Top             =   750
         Width           =   1032
      End
      Begin VB.TextBox textCP44 
         Height          =   300
         Left            =   990
         MaxLength       =   12
         TabIndex        =   0
         Top             =   450
         Width           =   1245
      End
      Begin VB.TextBox textPrintTNT 
         Height          =   300
         Left            =   1440
         MaxLength       =   1
         TabIndex        =   4
         Top             =   3240
         Width           =   372
      End
      Begin MSForms.ComboBox Combo2 
         Height          =   330
         Index           =   1
         Left            =   -73920
         TabIndex        =   10
         Top             =   2370
         Width           =   7395
         VariousPropertyBits=   679495707
         DisplayStyle    =   3
         Size            =   "13039;582"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.ComboBox Combo2 
         Height          =   330
         Index           =   0
         Left            =   -73920
         TabIndex        =   9
         Top             =   1110
         Width           =   7395
         VariousPropertyBits=   679495707
         DisplayStyle    =   3
         Size            =   "13039;582"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM05_1 
         Height          =   960
         Left            =   1290
         TabIndex        =   54
         Top             =   1920
         Width           =   7635
         VariousPropertyBits=   679493659
         ScrollBars      =   2
         Size            =   "13467;1693"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM07 
         Height          =   300
         Left            =   1290
         TabIndex        =   58
         Top             =   2600
         Width           =   7605
         VariousPropertyBits=   671105051
         MaxLength       =   40
         ScrollBars      =   2
         Size            =   "13414;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM06 
         Height          =   300
         Left            =   1290
         TabIndex        =   57
         Top             =   2260
         Width           =   7605
         VariousPropertyBits=   671105051
         MaxLength       =   40
         ScrollBars      =   2
         Size            =   "13414;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM23_2J 
         Height          =   300
         Left            =   2085
         TabIndex        =   53
         TabStop         =   0   'False
         Top             =   1500
         Width           =   6810
         VariousPropertyBits=   671105055
         BackColor       =   -2147483633
         ForeColor       =   -2147483641
         Size            =   "12012;529"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM23_2E 
         Height          =   300
         Left            =   2085
         TabIndex        =   52
         TabStop         =   0   'False
         Top             =   1170
         Width           =   6810
         VariousPropertyBits=   671105055
         BackColor       =   -2147483633
         ForeColor       =   -2147483641
         Size            =   "12012;529"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM67 
         Height          =   300
         Left            =   -73860
         TabIndex        =   7
         Top             =   420
         Width           =   7695
         VariousPropertyBits=   671105051
         MaxLength       =   200
         Size            =   "13573;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM52 
         Height          =   300
         Left            =   -73920
         TabIndex        =   39
         Top             =   3300
         Width           =   7392
         VariousPropertyBits=   671105055
         BackColor       =   -2147483644
         MaxLength       =   40
         Size            =   "13039;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM51 
         Height          =   300
         Left            =   -73920
         TabIndex        =   38
         Top             =   3000
         Width           =   7392
         VariousPropertyBits=   671105055
         BackColor       =   -2147483644
         MaxLength       =   60
         Size            =   "13039;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM49 
         Height          =   300
         Left            =   -73920
         TabIndex        =   37
         Top             =   2040
         Width           =   7392
         VariousPropertyBits=   671105055
         BackColor       =   -2147483644
         MaxLength       =   40
         Size            =   "13039;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM48 
         Height          =   300
         Left            =   -73920
         TabIndex        =   36
         Top             =   1740
         Width           =   7392
         VariousPropertyBits=   671105055
         BackColor       =   -2147483644
         MaxLength       =   60
         Size            =   "13039;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM47 
         Height          =   300
         Left            =   -73920
         TabIndex        =   35
         Top             =   1440
         Width           =   7392
         VariousPropertyBits=   671105055
         BackColor       =   -2147483644
         MaxLength       =   40
         Size            =   "13039;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM50 
         Height          =   300
         Left            =   -73920
         TabIndex        =   34
         Top             =   2700
         Width           =   7392
         VariousPropertyBits=   671105055
         BackColor       =   -2147483644
         MaxLength       =   40
         Size            =   "13039;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM09 
         Height          =   300
         Left            =   990
         TabIndex        =   3
         Top             =   2940
         Width           =   7905
         VariousPropertyBits=   671105051
         MaxLength       =   395
         Size            =   "13944;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM05 
         Height          =   300
         Left            =   1290
         TabIndex        =   2
         Top             =   1920
         Width           =   7605
         VariousPropertyBits=   671105051
         MaxLength       =   40
         ScrollBars      =   2
         Size            =   "13414;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM23_2C 
         Height          =   300
         Left            =   2280
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   840
         Width           =   6675
         VariousPropertyBits=   671105055
         BackColor       =   -2147483633
         ForeColor       =   -2147483641
         Size            =   "11774;529"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM23 
         Height          =   300
         Left            =   990
         TabIndex        =   1
         Top             =   810
         Width           =   1245
         VariousPropertyBits=   671105051
         MaxLength       =   9
         Size            =   "2037;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCP44_2 
         Height          =   300
         Left            =   2280
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   480
         Width           =   6675
         VariousPropertyBits=   671105055
         BackColor       =   -2147483633
         MaxLength       =   20
         Size            =   "11774;529"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label4 
         Caption         =   "是否列印DHL：        （Y:印）"
         Height          =   255
         Index           =   2
         Left            =   2880
         TabIndex        =   61
         Top             =   3270
         Width           =   2610
      End
      Begin VB.Label Label4 
         Caption         =   "是否印信頭：        （N:不印）"
         Height          =   255
         Index           =   1
         Left            =   5910
         TabIndex        =   59
         Top             =   3285
         Width           =   2370
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "案件日文名稱 :"
         Height          =   180
         Left            =   90
         TabIndex        =   56
         Top             =   2660
         Width           =   1170
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "案件英文名稱 :"
         Height          =   180
         Left            =   90
         TabIndex        =   55
         Top             =   2260
         Width           =   1170
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         Caption         =   "優先權資料 :"
         Height          =   180
         Left            =   -74910
         TabIndex        =   51
         Top             =   750
         Width           =   990
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "放棄專用權 :"
         Height          =   180
         Left            =   -74940
         TabIndex        =   50
         Top             =   450
         Width           =   990
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "(Y:印)"
         Height          =   180
         Left            =   1920
         TabIndex        =   49
         Top             =   3300
         Width           =   465
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "是否列印TNT :"
         Height          =   180
         Left            =   90
         TabIndex        =   48
         Top             =   3300
         Width           =   1140
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(日):"
         Height          =   180
         Index           =   8
         Left            =   -74400
         TabIndex        =   47
         Top             =   3300
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(英):"
         Height          =   180
         Index           =   7
         Left            =   -74400
         TabIndex        =   46
         Top             =   3000
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(中):"
         Height          =   180
         Index           =   6
         Left            =   -74400
         TabIndex        =   45
         Top             =   2700
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(日):"
         Height          =   180
         Index           =   5
         Left            =   -74400
         TabIndex        =   44
         Top             =   2040
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(英):"
         Height          =   180
         Index           =   4
         Left            =   -74400
         TabIndex        =   43
         Top             =   1740
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(中):"
         Height          =   180
         Index           =   3
         Left            =   -74400
         TabIndex        =   42
         Top             =   1440
         Width           =   345
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "代表人1"
         Height          =   180
         Index           =   1
         Left            =   -74760
         TabIndex        =   41
         Top             =   1140
         Width           =   630
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "代表人2"
         Height          =   180
         Index           =   2
         Left            =   -74760
         TabIndex        =   40
         Top             =   2370
         Width           =   630
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "商品類別 :"
         Height          =   180
         Index           =   7
         Left            =   90
         TabIndex        =   33
         Top             =   2940
         Width           =   810
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "案件名稱 :"
         Height          =   180
         Left            =   90
         TabIndex        =   32
         Top             =   1920
         Width           =   810
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "申請人 :"
         Height          =   180
         Left            =   90
         TabIndex        =   31
         Top             =   810
         Width           =   630
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "代理人 :"
         Height          =   180
         Index           =   0
         Left            =   90
         TabIndex        =   29
         Top             =   450
         Width           =   630
      End
   End
   Begin VB.TextBox textCP10 
      BackColor       =   &H80000004&
      BorderStyle     =   0  '沒有框線
      Height          =   300
      Left            =   1170
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   825
      Width           =   2532
   End
   Begin VB.TextBox textCP09 
      BackColor       =   &H80000004&
      BorderStyle     =   0  '沒有框線
      Height          =   300
      Left            =   1170
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   480
      Width           =   2532
   End
   Begin VB.TextBox textTMKey 
      BackColor       =   &H80000004&
      BorderStyle     =   0  '沒有框線
      Height          =   300
      Left            =   5220
      Locked          =   -1  'True
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   480
      Width           =   2532
   End
   Begin VB.TextBox textCP12 
      BackColor       =   &H80000004&
      BorderStyle     =   0  '沒有框線
      Height          =   300
      Left            =   5220
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   825
      Width           =   2532
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "代表圖下載存放路徑："
      Height          =   180
      Index           =   3
      Left            =   72
      TabIndex        =   63
      Top             =   5376
      Width           =   1800
   End
   Begin MSForms.TextBox textCP13 
      Height          =   300
      Left            =   5220
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   1170
      Width           =   2532
      VariousPropertyBits=   16415
      BackColor       =   -2147483644
      MaxLength       =   20
      Size            =   "4466;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textCP14 
      Height          =   300
      Left            =   1170
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   1170
      Width           =   2532
      VariousPropertyBits=   16415
      BackColor       =   -2147483644
      MaxLength       =   20
      Size            =   "4466;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "智權人員 :"
      Height          =   180
      Index           =   11
      Left            =   4350
      TabIndex        =   26
      Top             =   1170
      Width           =   810
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "案件性質 :"
      Height          =   180
      Index           =   6
      Left            =   270
      TabIndex        =   25
      Top             =   825
      Width           =   810
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "收文號 :"
      Height          =   180
      Index           =   1
      Left            =   450
      TabIndex        =   24
      Top             =   510
      Width           =   630
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "本所案號 :"
      Height          =   180
      Index           =   0
      Left            =   4350
      TabIndex        =   23
      Top             =   480
      Width           =   810
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "業務區別 :"
      Height          =   180
      Index           =   2
      Left            =   4350
      TabIndex        =   22
      Top             =   825
      Width           =   810
   End
   Begin VB.Label Label24 
      AutoSize        =   -1  'True
      Caption         =   "承辦人 :"
      Height          =   180
      Left            =   450
      TabIndex        =   21
      Top             =   1170
      Width           =   630
   End
End
Attribute VB_Name = "frm030001_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/08/03 改成Form2.0 ; textCP14、textCP13、textCP44_2、textTM23_2C、textTM23_2E、textTM23_2J、
                                                               'textTM05、textTM06、textTM07、textTM05_1、Combo2(index)、textTM67、
                                                               'textTM47、textTM48、textTM49、textTM50、textTM51、textTM52)
'Memo By Sindy 2012/12/4 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/29 員工編號欄已修改
'Memo By Sindy 2010/8/10 日期欄已修改
Option Explicit

Dim m_strCust1 As String '申請人1
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
' 申請人
Dim m_TM23 As String
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
' 優先權畫面所使用的變數
Dim m_Pa(1 To 4) As String '本所案號
Dim m_Priority(1 To 6) As String
Public ChkTG As Boolean
Dim LetterStyle As Integer

Private Sub cmdOK_Click(Index As Integer)
Select Case Index
Case 0
        LetterStyle = 0
        If CheckDataValid = True Then
            If TxtValidate = False Then Exit Sub
            'Add by Sindy 2013/5/23
            If CheckCP44 = False Then
               textCP44.SetFocus
               Exit Sub
            End If
            '2013/5/23 End
            ' 設定滑鼠游標為等待狀態
            If m_TM10 = "101" Then '101.美國
                LetterStyle = MsgBox("此為美國案，Y 為已使用；N 為未使用；請回答...", vbYesNo, "選擇指示信樣式...")
            'Add By Sindy 2018/1/25
            ElseIf m_TM10 = "112" And m_CP10 = "101" Then '112.波多黎各
                LetterStyle = MsgBox("波多黎各案，申請時是否已提使用宣誓？Y 為已使用；N 為未使用；請回答...", vbYesNo, "選擇...")
            '2018/1/25 END
            End If
            Screen.MousePointer = vbHourglass
            ' 更新欄位輸入的內容
            OnUpdateField
            ' 存檔
            If OnSaveData = False Then MsgBox "存檔失敗，請洽系統管理員 !", vbCritical: Screen.MousePointer = vbDefault: Exit Sub
            '印 TNT
            Select Case UCase(textPrintTNT)
            Case "Y"
                    Screen.MousePointer = vbDefault
                    frm060321.Show
                    bolToEndByNick = False
                    frm060321.Hide
                    '傳收文號
                    frm060321.GetCP09 = m_CP09
                    frm060321.txt1(0).Text = m_TM01
                    frm060321.txt1(1).Text = m_TM02
                    frm060321.txt1(2).Text = m_TM03
                    frm060321.txt1(3).Text = m_TM04
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
            Case Else
            End Select
            
            'Add by Lydia 2014/12/30 + DHL
            Select Case UCase(textPrintDHL)
            Case "Y"
                    Screen.MousePointer = vbDefault
                    frm060330.Show
                    bolToEndByNick = False
                    frm060330.Hide
                    '傳收文號
                    'frm060330.GetCP09 = m_CP09 'mark by lydia 2022/03/28
                    frm060330.txt1(0).Text = m_TM01
                    frm060330.txt1(1).Text = m_TM02
                    frm060330.txt1(2).Text = m_TM03
                    frm060330.txt1(3).Text = m_TM04
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
            Case Else
            End Select
            
            '地址條
            pub_AddressListSN = pub_AddressListSN + 1
            PUB_AddNewAddressList strUserNum, m_TM01, m_TM02, m_TM03, m_TM04, "" & pub_AddressListSN, "0"
            ' 設定滑鼠游標為預設
            Screen.MousePointer = vbDefault
            frm030001.Show
            frm030001.txt1(0).Text = ""
            frm030001.txt1(1).Text = ""
            frm030001.txt1(2).Text = ""
            frm030001.txt1(3).Text = ""
            Unload Me
            Exit Sub
        End If
Case 1
        Unload Me
        frm030001.Show
Case 2
        Unload frm030001
        Unload Me
Case 3
        frm03010303_04.Hide
        Set frm03010303_04.UpForm = Me
        frm03010303_04.TGKey = frm030001.txt1(0).Text & "-" & frm030001.txt1(1).Text & "-" & IIf(Trim(frm030001.txt1(2).Text) = "", "0", Trim(frm030001.txt1(2).Text)) & "-" & IIf(Trim(frm030001.txt1(3).Text) = "", "00", Trim(frm030001.txt1(3).Text))
        frm03010303_04.AllClass = textTM09.Text
        'edit by nick 2005/01/26 有輸一種就可以了
        'frm03010303_04.ChkEng = True
        'frm03010303_04.PubMsg = "此為英文 Order Letter 所以英文必須要輸入！"
        Me.Hide
        frm03010303_04.QueryData
        frm03010303_04.Show vbModal 'Modify By Sindy 2009/09/17 改為強制回應表單
Case Else
End Select
End Sub

Private Sub OnUpdateField()
Dim strCP64 As String
Dim intI As Integer
   
   ' 代理人
   If IsEmptyText(textCP44) = False Then
      'Add By Sindy 2013/5/23 加判斷是否為聯絡人
      intI = InStr(textCP44, "-")
      If intI > 0 Then
         SetCPFieldNewData "CP44", Left(textCP44, intI - 1) & String(9 - Len(Left(textCP44, intI - 1)), "0")
         SetCPFieldNewData "CP116", Mid(textCP44, intI + 1)
      Else
      '2013/5/23 End
         SetCPFieldNewData "CP44", textCP44 & String(9 - Len(textCP44), "0")
         SetCPFieldNewData "CP116", ""
      End If
   Else
      SetCPFieldNewData "CP44", "" 'textCP44
      SetCPFieldNewData "CP116", ""
   End If
   
   Select Case m_TM01
      ' 系統類別為CFT的為更新商標基本檔
      Case "CFT":
         ' 案件名稱
         SetTMSPFieldNewData "TM05", textTM05_1
         ' 商品類別
         SetTMSPFieldNewData "TM09", textTM09
         ' 申請人
         If IsEmptyText(textTM23) = False Then
            SetTMSPFieldNewData "TM23", textTM23 & String(9 - Len(textTM23), "0")
         Else
            SetTMSPFieldNewData "TM23", textTM23
         End If
         ' 代表人
         SetTMSPFieldNewData "TM47", textTM47
         ' 代表人
         SetTMSPFieldNewData "TM48", textTM48
         ' 代表人
         SetTMSPFieldNewData "TM49", textTM49
         ' 代表人
         SetTMSPFieldNewData "TM50", textTM50
         ' 代表人
         SetTMSPFieldNewData "TM51", textTM51
         ' 代表人
         SetTMSPFieldNewData "TM52", textTM52
         ' 放棄專用權
         SetTMSPFieldNewData "TM67", textTM67
      Case Else:
         ' 案件中文名稱
         SetTMSPFieldNewData "SP05", textTM05
         ' 案件英文名稱
         SetTMSPFieldNewData "SP06", textTM06
         ' 案件日文名稱
         SetTMSPFieldNewData "SP07", textTM07
         ' 申請人
         If IsEmptyText(textTM23) = False Then
            SetTMSPFieldNewData "SP08", textTM23 & String(9 - Len(textTM23), "0")
         Else
            SetTMSPFieldNewData "SP08", textTM23
         End If
   End Select
      
End Sub

Public Function OnSaveData() As Boolean
OnSaveData = False
   Dim strTmp As String
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   Dim nIndex As Integer
   Dim bFirst As Boolean
   Dim bDifference As Boolean
   Dim strNP07 As String
   Dim strNP08 As String
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
   
   'Add By Sindy 2018/1/25
   If m_TM10 = "112" And m_CP10 = "101" And LetterStyle > 0 Then '112.波多黎各
      If bDifference = True Then
         strSql = strSql & ","
      End If
      bDifference = True
      strSql = strSql & "CP64=CP64||'" & IIf(LetterStyle = 6, "已提使用宣誓;", "未提使用宣誓;") & "'"
   End If
   '2018/1/25 END
   
   ' 設定SQL語法更新的條件
   strSql = strSql & " " & _
                  "WHERE CP09 = '" & m_CP09 & "' "
   ' 執行SQL指令
   If bDifference = True Then: cnnConnection.Execute strSql
   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   ' 更新基本檔
   Select Case m_TM01
      ' 系統類別為CFT的為更新商標基本檔
      Case "CFT":
        'Modify By Sindy 2018/2/13 Mark:在前畫面會開啟申請人案件地址供其修改資料
        'add by nick 2004/12/28
'        CheckOC3
'        With AdoRecordSet3
'            .CursorLocation = adUseClient
'            .Open "select * from customer where cu01='" & Mid(ChangeCustomerL(Me.textTM23.Text), 1, 8) & "' and cu02='" & Mid(ChangeCustomerL(Me.textTM23.Text), 9, 1) & "' ", cnnConnection, adOpenStatic, adLockReadOnly
'            If .RecordCount <> 0 Then
'                SetTMSPFieldNewData "TM24", "" & .Fields("cu23").Value
'                SetTMSPFieldNewData "TM25", "" & .Fields("cu24").Value & IIf(IsNull(.Fields("cu25").Value), "", " " & "" & .Fields("cu25").Value) & IIf(IsNull(.Fields("cu26").Value), "", " " & "" & .Fields("cu26").Value) & IIf(IsNull(.Fields("cu27").Value), "", " " & "" & .Fields("cu27").Value) & IIf(IsNull(.Fields("cu28").Value), "", " " & "" & .Fields("cu28").Value)
'                SetTMSPFieldNewData "TM26", "" & .Fields("cu29").Value
'            End If
'        End With
'        CheckOC3
         OnUpdateTradeMark
      Case Else:
         OnUpdateServicePractice
   End Select
   
   ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   ' 儲存優先權資料
   'edit by nickc 2007/02/06 不用 dll 了 objPublicData.SavePriority m_Pa, m_Priority(1), m_Priority(2), m_Priority(3)
   'Modify by Amy 2014/04/17 +, m_Priority(4), m_Priority(5)
   'Modify by Sindy 2017/10/12 +, m_Priority(6)
   ClsPDSavePriority m_Pa, m_Priority(1), m_Priority(2), m_Priority(3), m_Priority(4), m_Priority(5), m_Priority(6)

   ' 列印定稿
   PrintLetter
   
   Set rsTmp = Nothing
    cnnConnection.CommitTrans
    OnSaveData = True
    Exit Function
CheckingErr:
    MsgBox (Err.Description)
     cnnConnection.RollbackTrans
End Function

' 列印定稿前將例外欄位加入到列印定稿例外欄位檔案中
Private Sub InsExpField()
   Dim strSql As String
   Dim tmpDBStr As String   'tmp
   Dim arrTM09 As Variant   'tmp
   Dim i_931101 As Integer   'tmp
   Dim o1_34 As Boolean     'goods
   Dim o35_45 As Boolean   'services
   Dim letter_1 As String    '處理狀況
   Dim LetterSpace As Integer '空幾格
   
   ' 系統類別為CFT
   If m_TM01 = "CFT" Then
        Select Case m_TM10
           Case "101":
              Select Case LetterStyle
              Case 6
                    letter_1 = "02"
              Case 7
                    letter_1 = "01"
              Case Else
              End Select
           Case Else:
              letter_1 = "00"
        End Select
        EndLetter "15", textCP09.Text, letter_1, strUserNum
        ' Goods or Service
        arrTM09 = Split(textTM09, ",")
        tmpDBStr = ""
        o1_34 = False
        o35_45 = False
        For i_931101 = 0 To UBound(arrTM09)
            If o1_34 = False Then
                If Val(arrTM09(i_931101)) >= 1 And Val(arrTM09(i_931101)) <= 34 Then
                    o1_34 = True
                End If
            End If
            If o35_45 = False Then
                If Val(arrTM09(i_931101)) >= 35 And Val(arrTM09(i_931101)) <= 45 Then
                    o35_45 = True
                End If
            End If
            If o1_34 = True And o35_45 = True Then
                Exit For
            End If
        Next i_931101
        If o1_34 = True And o35_45 = True Then
            tmpDBStr = "Goods & Services"
        ElseIf o1_34 = True And o35_45 = False Then
            tmpDBStr = "Goods"
        ElseIf o1_34 = False And o35_45 = True Then
            tmpDBStr = "Services"
        End If
        LetterSpace = Len(tmpDBStr)
        strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                 "VALUES ('" & "15" & "','" & textCP09.Text & "','" & letter_1 & "','" & strUserNum & _
                 "','列印備註','" & tmpDBStr & "')"
        cnnConnection.Execute strSql
        
        If letter_1 = "00" Then Exit Sub 'Modify by Amy 2021/07/16 定稿別15處理/狀況00 無「附件」tag不需run 下列程式,因CFT-022383 29類商品資料過長會error,定稿會以|?TMGoods抓其資料-與Morgan討論過以此方式處理
        
        For i_931101 = 0 To UBound(arrTM09)
            arrTM09(i_931101) = "'" & arrTM09(i_931101) & "'"
        Next i_931101
        tmpDBStr = ""
        strSql = "select tg05,tg07 from tmgoods where tg05 in (" & Join(arrTM09, ",") & ") and tg01='" & m_TM01 & "' and tg02='" & m_TM02 & "' and tg03='" & m_TM03 & "' and tg04='" & m_TM04 & "' order by tg05 "
        CheckOC3
        With AdoRecordSet3
            .CursorLocation = adUseClient
            .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
            If .RecordCount <> 0 Then
                .MoveFirst
                Do While .EOF = False
                    tmpDBStr = tmpDBStr & "Class " & CheckStr(.Fields("tg05").Value) & vbCrLf & CheckStr(.Fields("tg07").Value)
                    .MoveNext
                    If .EOF = False Then
                        tmpDBStr = tmpDBStr & vbCrLf & vbCrLf
                    End If
                Loop
            End If
        End With
        CheckOC3
        '商品及服務資料
        strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                 "VALUES ('" & "15" & "','" & textCP09.Text & "','" & letter_1 & "','" & strUserNum & _
                 "','附件','" & ChgSQL(tmpDBStr) & "')"
        cnnConnection.Execute strSql
   End If
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 列印定稿
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub PrintLetter()
   ' 先呼叫定稿程式的清除原定稿資料的函式去清除之前殘留在例外欄位檔中的資料
   InsExpField
   Dim letter_1 As String    '處理狀況
   ' 系統類別為CFT
   If m_TM01 = "CFT" Then
        Select Case m_TM10
           Case "101":
              Select Case LetterStyle
              Case 6
                    letter_1 = "02"
              Case 7
                    letter_1 = "01"
              Case Else
              End Select
           Case Else:
              letter_1 = "00"
        End Select
            'Modified by Lydia 2016/03/01 只有除第一頁出本所信頭
            'NowPrint textCP09.Text, "15", letter_1, True, strUserNum, 0, , , , , , IIf(txtLetterHead = "N", "", True)
            'Modify by Amy 2018/07/27 +Me.Name
            NowPrint textCP09.Text, "15", letter_1, True, strUserNum, 0, , , , , , IIf(txtLetterHead = "N", "", True), , , , 1, , , , , , Me.Name
   End If
End Sub
' 更新商標基本檔的相關欄位
Private Sub OnUpdateTradeMark()
   Dim strTmp As String
   Dim strSql As String
   Dim nIndex As Integer
   Dim bFirst As Boolean
   Dim bDifference As Boolean
      
   ' 更新案件進度檔
   strSql = "UPDATE TradeMark SET "
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
   ' 執行SQL指令
   If bDifference = True Then
      ' 設定SQL語法更新的條件
      strSql = strSql & " " & _
                     "WHERE TM01 = '" & m_TM01 & "' AND " & _
                           "TM02 = '" & m_TM02 & "' AND " & _
                           "TM03 = '" & m_TM03 & "' AND " & _
                           "TM04 = '" & m_TM04 & "' "
      cnnConnection.Execute strSql
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
   ' 執行SQL指令
   If bDifference = True Then
      ' 設定SQL語法更新的條件
      strSql = strSql & " " & _
                     "WHERE SP01 = '" & m_TM01 & "' AND " & _
                           "SP02 = '" & m_TM02 & "' AND " & _
                           "SP03 = '" & m_TM03 & "' AND " & _
                           "SP04 = '" & m_TM04 & "' "
      cnnConnection.Execute strSql
   End If
End Sub

Private Function CheckDataValid() As Boolean
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   CheckDataValid = False
   ' 商品及服務
   If ChkTG = False Then
      strTit = "檢核資料"
      strMsg = "請輸入商品及服務"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      cmdok(3).SetFocus
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
   ' 申請人
   If IsEmptyText(textTM23) = True Then
      strTit = "檢核資料"
      strMsg = "請輸入申請人"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textTM23.SetFocus
      GoTo EXITSUB
   End If
   ' 案件名稱
   If m_TM01 = "CFT" Then
        If IsEmptyText(textTM05_1) = True Then
           strTit = "檢核資料"
           strMsg = "請輸入案件名稱"
           nResponse = MsgBox(strMsg, vbOKOnly, strTit)
           textTM05_1.SetFocus
           GoTo EXITSUB
        End If
   Else
        If IsEmptyText(textTM05) = True Then
           strTit = "檢核資料"
           strMsg = "請輸入案件名稱"
           nResponse = MsgBox(strMsg, vbOKOnly, strTit)
           textTM05.SetFocus
           GoTo EXITSUB
        End If
    End If
   ' 商品類別
   If IsEmptyText(textTM09) = True Then
      strTit = "檢核資料"
      strMsg = "請輸入商品類別"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textTM09.SetFocus
      GoTo EXITSUB
   End If
   
   'Added by Morgan 2023/7/13
   If ChkPath() = False Then
      txtPath.SetFocus
      txtPath_GotFocus
      Exit Function
   Else
      strExc(1) = ""
      GetImgByteFile_Case m_TM01, m_TM02, m_TM03, m_TM04, strExc(1)
      If strExc(1) <> "" Then
         strExc(2) = txtPath & "\" & m_TM01 & m_TM02 & IIf(m_TM04 <> "00", "-" & m_TM03 & "-" & m_TM04, IIf(m_TM03 <> "0", "-" & m_TM03, "")) & Right(strExc(1), 4)
         If Dir(strExc(2)) <> "" Then
            If MsgBox("[ " & strExc(2) & " ]圖檔已存在，是否要覆蓋？", vbYesNo + vbQuestion + vbDefaultButton2) = vbYes Then
               Kill strExc(2)
            Else
               GoTo EXITSUB
            End If
         End If
         Name strExc(1) As strExc(2)
         'MsgBox "圖檔已下載[ " & strExc(2) & " ]。", vbInformation
      End If
   End If
   'end 2023/7/13
   
   CheckDataValid = True
EXITSUB:
End Function
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
' 取得商標基本檔的欄位內容
Private Sub QueryTradeMark()
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   
   strSql = "SELECT * FROM TradeMark " & _
            "WHERE TM01 = '" & m_TM01 & "' AND " & _
                  "TM02 = '" & m_TM02 & "' AND " & _
                  "TM03 = '" & m_TM03 & "' AND " & _
                  "TM04 = '" & m_TM04 & "' "
                  
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      ' 案件名稱
      textTM05_1 = Empty
      If IsNull(rsTmp.Fields("TM05")) = False Then
         textTM05_1 = rsTmp.Fields("TM05")
      End If
      SetTMSPFieldOldData "TM05", textTM05_1, 0
      textTM05_1.Visible = True
      textTM05.Visible = False
      textTM06.Visible = False
      textTM07.Visible = False
      Label7.Visible = False
      Label8.Visible = False
      textTM05_1.Enabled = True
      textTM05.Enabled = False
      textTM06.Enabled = False
      textTM07.Enabled = False
      Label7.Enabled = False
      Label8.Enabled = False
      ' 商品類別
      textTM09 = Empty
      If IsNull(rsTmp.Fields("TM09")) = False Then
         textTM09 = rsTmp.Fields("TM09")
      End If
      SetTMSPFieldOldData "TM09", textTM09, 0
      ' 申請人
      Dim oState As Boolean
      oState = True
      m_TM23 = Empty
      If IsNull(rsTmp.Fields("TM23")) = False Then
         m_TM23 = rsTmp.Fields("TM23")
         textTM23 = rsTmp.Fields("TM23")
         strSql = GetCustomerNameAndState_030001(textTM23, 0, oState)
      End If
      m_strCust1 = "" & Me.textTM23.Text
      SetTMSPFieldOldData "TM23", textTM23, 0
      'add by nick 2004/12/28
      SetTMSPFieldOldData "TM24", "" & rsTmp.Fields("TM24"), 0
      SetTMSPFieldOldData "TM25", "" & rsTmp.Fields("TM25"), 0
      SetTMSPFieldOldData "TM26", "" & rsTmp.Fields("TM26"), 0
      '代表人
      Dim i As Integer, j As Integer
      For i = 0 To 1
         Combo2(i).AddItem ""
      Next
      
      If rsTmp.Fields("TM23").Value <> "" Then
         strExc(0) = "SELECT CU40,CU43,CU46,CU49,CU52,CU55 FROM CUSTOMER WHERE " & ChgCustomer(rsTmp.Fields("TM23").Value)
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))   'edit by nickc 2007/02/06 不用 dll 了   = objLawDll.ReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            For j = 1 To 6
               If IsNull(RsTemp.Fields(j - 1)) Then
                  strExc(0) = ""
               Else
                  strExc(0) = "-" & RsTemp.Fields(j - 1)
               End If
               Combo2(0).AddItem rsTmp.Fields("TM23").Value & "-" & j & strExc(0)
               Combo2(1).AddItem rsTmp.Fields("TM23").Value & "-" & j & strExc(0)
            Next
         End If
      End If
      
      'Morgan 2003/11/20 -- end
      
      ' 代表人1(中)
      textTM47 = Empty
      If IsNull(rsTmp.Fields("TM47")) = False Then
         textTM47 = rsTmp.Fields("TM47")
      End If
      SetTMSPFieldOldData "TM47", textTM47, 0
      ' 代表人1(英)
      textTM48 = Empty
      If IsNull(rsTmp.Fields("TM48")) = False Then
         textTM48 = rsTmp.Fields("TM48")
      End If
      SetTMSPFieldOldData "TM48", textTM48, 0
      ' 代表人1(日)
      textTM49 = Empty
      If IsNull(rsTmp.Fields("TM49")) = False Then
         textTM49 = rsTmp.Fields("TM49")
      End If
      SetTMSPFieldOldData "TM49", textTM49, 0
      ' 代表人2(中)
      textTM50 = Empty
      If IsNull(rsTmp.Fields("TM50")) = False Then
         textTM50 = rsTmp.Fields("TM50")
      End If
      SetTMSPFieldOldData "TM50", textTM50, 0
      ' 代表人2(英)
      textTM51 = Empty
      If IsNull(rsTmp.Fields("TM51")) = False Then
         textTM51 = rsTmp.Fields("TM51")
      End If
      SetTMSPFieldOldData "TM51", textTM51, 0
      ' 代表人2(日)
      textTM52 = Empty
      If IsNull(rsTmp.Fields("TM52")) = False Then
         textTM52 = rsTmp.Fields("TM52")
      End If
      SetTMSPFieldOldData "TM52", textTM52, 0
      ' 放棄專用權
      textTM67 = Empty
      If IsNull(rsTmp.Fields("TM67")) = False Then
         textTM67 = rsTmp.Fields("TM67")
      End If
      SetTMSPFieldOldData "TM67", textTM67, 0
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Sub

Private Sub cmdOpen_Click()
   Dim stFileName As String, stFolderPath As String, stFullName As String
   
   '讀取前次設定路徑
   stFolderPath = GetSetting("TAIE", "P", UCase(Me.Name) & "Dir", "")
   If stFolderPath <> "" Then
      If PUB_ChkDir(stFolderPath) = False Then
         stFolderPath = PUB_Getdesktop
      End If
   Else
      stFolderPath = PUB_Getdesktop
   End If
   stFolderPath = PUB_GetFolder(Me.hWnd, stFolderPath, "請選取資料夾:")
   If Trim(stFolderPath) <> "" Then 'they did not hit cancel
      SaveSetting "TAIE", "P", UCase(Me.Name) & "Dir", stFolderPath
   Else
      Exit Sub
   End If
   If Right(Trim(stFolderPath), 1) <> "\" Then
      stFolderPath = Trim(stFolderPath) & "\"
   End If
   
End Sub

Private Sub cmdPath_Click()
   Dim fName As String, strStartFolder As String
   
   If Dir(txtPath & "\", vbDirectory) <> "" Then strStartFolder = txtPath
   
   fName = PUB_GetFolder(Me.hWnd, strStartFolder, "請選取資料夾:")
   If fName <> "" Then 'they did not hit cancel
      txtPath = fName
      SaveSetting "TAIE", "CFT", UCase(Me.Name) & "Dir", txtPath
   End If
End Sub

Private Sub cmdPriority_Click()
   ' 修改優先權資料
   'Modify by Amy 2014/04/17 +, m_Priority(4), m_Priority(5)
   'Modify by Sindy 2017/10/12 +, m_Priority(6)
   'Modify by Sindy 2019/1/23 + m_TM01 & m_TM02 & m_TM03 & m_TM04
   ModifyPriority m_Priority(1), m_Priority(2), m_Priority(3), , , m_TM01 & m_TM02 & m_TM03 & m_TM04, , , m_Priority(4), m_Priority(5), m_Priority(6)
End Sub

'Add By Sindy 2013/5/23
Private Sub cmdSuggest_Click()
   If m_TM10 = "" Then
      MsgBox "無申請國家，無法使用建議代理人！"
   Else
      ShowSuggest
   End If
End Sub

'Add By Sindy 2013/5/23
Private Sub ShowSuggest(Optional bNoMsg As Boolean)
   Dim bCancel As Boolean
   Dim strDetail As String 'Added by Lydia 2018/01/16
   If m_TM10 <> "" And m_CP10 = "101" Then
      'Modified by Lydia 2018/01/16 + strdetail
      'If PUB_ReadFTList_CFT(m_TM01, m_TM10, RsTemp) = True Then
      If PUB_ReadFTList_CFT(m_TM01, m_TM10, RsTemp, , strDetail) = True Then
         Set frm880012.grdDataList.Recordset = RsTemp
         Set frm880012.fmParent = Me
         frm880012.lblDate.Caption = "FC給案：" & ChangeTStringToTDateString(TransDate(strDetail, 1)) & "~至今之FCT申請案"
         frm880012.Show vbModal
         If Me.Tag <> "" Then
            textCP44 = Me.Tag
            textCP44_Validate bCancel
         End If
      ElseIf bNoMsg = False Then
         MsgBox "該申請國無建議代理人！"
      End If
   End If
End Sub

'Add By Sindy 2013/5/23
'檢查給案量是否超過
Private Function CheckCP44() As Boolean
   Dim stDate As String, stFC05 As String, stDate1 As String, stDate2 As String, stYear As String
   Dim iPos As Integer, stCon As String, stVTB As String, stConCP As String
   Dim bolRtn As Boolean
   
   bolRtn = True
   If textCP44 <> "" And m_TM10 <> "" And m_CP10 = "101" Then
      stDate = strSrvDate(1)
      stYear = Left(stDate, 4)
      '下半年
      If Val(Mid(stDate, 5, 2)) > 6 Then
         stFC05 = "2"
         stDate1 = stYear & "0701"
         stDate2 = stYear & "1231"
      '上半年
      Else
         stFC05 = "1"
         stDate1 = stYear & "0101"
         stDate2 = stYear & "0630"
      End If
      stCon = " AND FC04=" & (stYear - 1911) & " and FC05='" & stFC05 & "'"
      stConCP = " and cp10='101' and cp27>=" & stDate1 & " and cp27<=" & stDate2
      
      If InStr(textCP44, "-") > 0 Then
         iPos = InStr(textCP44, "-")
         stCon = stCon & " and FC01||FC02='" & ChangeCustomerL(Left(textCP44, iPos - 1)) & "' and FC03='" & Mid(textCP44, iPos + 1) & "'"
         stConCP = stConCP & " and CP44='" & ChangeCustomerL(Left(textCP44, iPos - 1)) & "' and CP116='" & Mid(textCP44, iPos + 1) & "'"
      Else
         stCon = stCon & " and FC01||FC02='" & ChangeCustomerL(textCP44) & "' AND FC03 IS NULL"
         stConCP = stConCP & " and CP44='" & ChangeCustomerL(textCP44) & "' and CP116 IS NULL"
      End If
      
      stVTB = "select nvl(count(*),0) Q1 from caseprogress where CP01='CFT'" & stConCP
      strExc(0) = "select FC07,Q1" & _
      " From fagentConfig,(" & stVTB & ") X" & _
      " where  FC06='CFT'" & stCon
      
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

'Morgan 2003/11/20
Private Sub Combo2_Click(Index As Integer)

   Dim i As Integer, strTmp As String
   
   If (Combo2(Index).Text = "") Then
      For i = 0 To 2
         Me.Controls("textTM" & Format(47 + i + 3 * Index, "#")).Text = ""
      Next i
      Exit Sub
   End If
   
   strTmp = Mid(Combo2(Index).Text, InStr(Combo2(Index).Text, "-") + 1, 1)
   strExc(1) = "CU" & 39 + (Val(strTmp) - 1) * 3 & ",CU" & 40 + (Val(strTmp) - 1) * 3 & ",CU" & 41 + (Val(strTmp) - 1) * 3
   strExc(0) = "SELECT " & strExc(1) & " FROM CUSTOMER WHERE " & ChgCustomer(Left(Combo2(Index).Text, InStr(Combo2(Index).Text, "-") - 1))
   
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))   'edit by nickc 2007/02/06 不用 dll 了   = objLawDll.ReadRstMsg(intI, strExc(0))
   If intI = 1 Then
   
      For i = 0 To 2
         
         If Not IsNull(RsTemp.Fields(i)) Then
            Me.Controls("textTM" & Format(47 + i + 3 * Index, "#")).Text = RsTemp.Fields(i)
         Else
            Me.Controls("textTM" & Format(47 + i + 3 * Index, "#")).Text = ""
         End If
         
      Next
   End If
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
      ' 案件中文名稱
      textTM05 = Empty
        If IsNull(rsTmp.Fields("SP05")) = False Then
           textTM05 = rsTmp.Fields("SP05")
        End If
        SetTMSPFieldOldData "SP05", textTM05, 0
      textTM05_1.Visible = False
      textTM05.Visible = True
      textTM06.Visible = True
      textTM07.Visible = True
      Label7.Visible = True
      Label8.Visible = True
      textTM05_1.Enabled = False
      textTM05.Enabled = True
      textTM06.Enabled = True
      textTM07.Enabled = True
      Label7.Enabled = True
      Label8.Enabled = True
      ' 案件英文名稱
      textTM06 = Empty
      If IsNull(rsTmp.Fields("SP06")) = False Then
         textTM06 = rsTmp.Fields("SP06")
      End If
      SetTMSPFieldOldData "SP06", textTM06, 0
      ' 案件日文名稱
      textTM07 = Empty
      If IsNull(rsTmp.Fields("SP07")) = False Then
         textTM07 = rsTmp.Fields("SP07")
      End If
      SetTMSPFieldOldData "SP07", textTM07, 0
      ' 申請人
      Dim oState As Boolean
      oState = True
      m_TM23 = Empty
      If IsNull(rsTmp.Fields("SP08")) = False Then
         m_TM23 = rsTmp.Fields("SP08")
         textTM23 = rsTmp.Fields("TM23")
         strSql = GetCustomerNameAndState_030001(textTM23, 0, oState)
      End If
      'Add By Cheng 2002/08/23
      m_strCust1 = "" & Me.textTM23.Text
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
   Dim strCP27 As String
   Dim strCP44 As String
   Dim strCP45 As String
   Dim nIndex As Integer
   Dim bFind As Boolean
   'Dim oState As Boolean
   Dim strTempName As String   '2010/11/24 add by sonia
   
   strSql = "SELECT * FROM CaseProgress " & _
            "WHERE cp09 in (select nvl(A.cp09,B.cp09) as cp09 from (select cp09 from caseprogress where cp01 = '" & m_TM01 & "' AND cp02 = '" & m_TM02 & "' AND cp03 = '" & m_TM03 & "' AND cp04 = '" & m_TM04 & "'  and cp10='101' and cp27 is null and cp57 is null) A,(select min(cp09) as cp09 from caseprogress where cp01 = '" & m_TM01 & "' AND cp02 = '" & m_TM02 & "' AND cp03 = '" & m_TM03 & "' AND cp04 = '" & m_TM04 & "'  and cp10='107' and cp27 is null and cp57 is null) B) "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
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
      'Add By Sindy 2013/5/23
      If m_CP10 = "101" Then
         textCP44.MaxLength = 12
         cmdSuggest.Visible = True
      Else
         textCP44.MaxLength = 9
         cmdSuggest.Visible = False
      End If
      '2013/5/23 End
      ' 業務區別
      If IsNull(rsTmp.Fields("CP12")) = False Then
         textCP12 = GetDepartmentName(rsTmp.Fields("CP12"))
      End If
      ' 智權人員
      If IsNull(rsTmp.Fields("CP13")) = False Then
         textCP13 = GetStaffName(rsTmp.Fields("CP13"), True)
      End If
      ' 承辦人員
      If IsNull(rsTmp.Fields("CP14")) = False Then
         '92.10.6 ADD BY SONIA
         m_CP14 = rsTmp.Fields("CP14")
         '92'10'6 END
         textCP14 = GetStaffName(rsTmp.Fields("CP14"), True)
      End If
      ' 代理人
      textCP44 = Empty
      If IsNull(rsTmp.Fields("CP44")) = False Then
         textCP44 = rsTmp.Fields("CP44")
         SetCPFieldOldData "CP44", textCP44, 0 'Modify By Sindy 2013/5/23
      Else
         SetCPFieldOldData "CP44", "", 0 'Modify By Sindy 2013/5/23
      End If
      'Add By Sindy 2013/5/23
      If IsNull(rsTmp.Fields("CP116")) = False Then
         textCP44 = textCP44 & "-" & rsTmp.Fields("CP116")
         SetCPFieldOldData "CP116", rsTmp.Fields("CP116"), 0
      Else
         SetCPFieldOldData "CP116", "", 0
      End If
      '2013/5/23 End
      '2010/11/24 modify by sonia 取消basQuery的GetFAgentNameAndState
      'oState = True
      'textCP44_2 = GetFAgentNameAndState(textCP44, oState)
'      If PUB_GetAgentNameAndState(m_TM01, textCP44.Text, strTempName) Then
'         textCP44_2 = strTempName
'      Else
'         textCP44_2.Text = ""
'      End If
      '2010/11/24 end
      Call textCP44_Validate(False) 'Modify By Sindy 2013/5/23
    End If
End Sub

Public Sub SetData(ByVal nType As Integer, ByVal strData As String, Optional ByVal bClear As Boolean = False)
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
      Case 0: m_TM01 = strData
      Case 1: m_TM02 = strData
      Case 2: m_TM03 = strData
      Case 3: m_TM04 = strData
   End Select
End Sub
Private Function TxtValidate() As Boolean
Dim objTxt As Object
Dim ii As Integer
Dim Cancel As Boolean

TxtValidate = False

'If Me.textCP26.Enabled = True Then
'   Cancel = False
'   textCP26_Validate Cancel
'   If Cancel = True Then
'      Exit Function
'   End If
'End If

If Me.textCP44.Enabled = True Then
   Cancel = False
   textCP44_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textPrintTNT.Enabled = True Then
   Cancel = False
   textPrintTNT_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If
'Add by Lydia 2014/12/30 + DHL
If Me.textPrintDHL.Enabled = True Then
   Cancel = False
   textPrintDHL_Validate Cancel
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
If Me.textTM05_1.Enabled = True Then
   Cancel = False
   textTM05_1_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If
If Me.textTM06.Enabled = True Then
   Cancel = False
   textTM06_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textTM07.Enabled = True Then
   Cancel = False
   textTM07_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textTM09.Enabled = True Then
   Cancel = False
   textTM09_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textTM23.Enabled = True Then
   Cancel = False
   textTM23_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

TxtValidate = True
End Function


' 清除商標基本檔檔案欄位串列
Private Sub ClearTMSPFieldList()
   If m_TMSPCount > 0 Then
      Erase m_TMSPList
   End If
   m_TMSPCount = 0
End Sub

' 讀取資料庫
Public Function QueryData() As Boolean
   Dim strDay As String
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
QueryData = False
   ' 先清除商標基本檔或服務業務基本檔欄位串列
   ClearTMSPFieldList
   ' 先清除案件進度檔欄位串列
   ClearCPFieldList
   
   ' 先取得本所案號
   strSql = "SELECT * FROM CaseProgress " & _
            "WHERE cp09 in (select nvl(A.cp09,B.cp09) as cp09 from (select cp09 from caseprogress where cp01 = '" & m_TM01 & "' AND cp02 = '" & m_TM02 & "' AND cp03 = '" & m_TM03 & "' AND cp04 = '" & m_TM04 & "'  and cp10='101' and cp27 is null and cp57 is null) A,(select min(cp09) as cp09 from caseprogress where cp01 = '" & m_TM01 & "' AND cp02 = '" & m_TM02 & "' AND cp03 = '" & m_TM03 & "' AND cp04 = '" & m_TM04 & "'  and cp10='107' and cp27 is null and cp57 is null) B) "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      ' 本所案號
      If IsNull(rsTmp.Fields("CP01")) = False Then: m_TM01 = rsTmp.Fields("CP01")
      If IsNull(rsTmp.Fields("CP02")) = False Then: m_TM02 = rsTmp.Fields("CP02")
      If IsNull(rsTmp.Fields("CP03")) = False Then: m_TM03 = rsTmp.Fields("CP03")
      If IsNull(rsTmp.Fields("CP04")) = False Then: m_TM04 = rsTmp.Fields("CP04")
      If IsNull(rsTmp.Fields("CP09")) = False Then: m_CP09 = rsTmp.Fields("CP09")
   Else
        MsgBox "查無資料！", , "錯誤！"
        Exit Function
   End If
   rsTmp.Close
   ' 本所案號
'   textTMKey = m_TM01 & m_TM02 & m_TM03 & m_TM04
   textTMKey.Text = m_TM01 & "-" & m_TM02 & "-" & IIf(Len("" & m_TM03) <= 0, "0", m_TM03) & "-" & IIf(Len("" & m_TM04) <= 0, "00", m_TM04)

   ' 收文號
   textCP09 = m_CP09
   
   'Modify By Sindy 2024/3/20 陳金蓮提:CFT申請案之指示信函必須先會稿才能送件,現因智權人員若有指定送件方式
   'CFT承辦人員即無法由系統叫出指示信函,故建議取消指示信函階段之送件方式管控,以利CFT承辦人員之作業
   '故不檢查
'   'Modify By Sindy 2024/1/23 檢查送件方式
'   If PUB_ChkCP141IsSend(m_CP09, , "處理指示信") = False Then
'      Unload Me
'      Exit Function
'   End If
'   '2024/1/23 END
   
   ' 取得國家代碼
   m_TM10 = GetNationNo(m_TM01, m_TM02, m_TM03, m_TM04)
   
   ' 取得案件進度檔的欄位
   QueryCaseProgress
   
   ' 取得基本檔
   Select Case m_TM01
      ' 系統類別為CFT的為讀取商標基本檔
      Case "CFT":
         QueryTradeMark
        'add by nick 2004/10/29 檢查是否有TG
         frm03010303_04.Hide
        Set frm03010303_04.UpForm = Me
        frm03010303_04.TGKey = m_TM01 & "-" & m_TM02 & "-" & m_TM03 & "-" & m_TM04
        frm03010303_04.AllClass = textTM09 'edit by nickc 2006/06/30 " "
        frm03010303_04.ChkEng = True
        frm03010303_04.Hide
        frm03010303_04.QueryData
        Unload frm03010303_04
        If ChkTG = True Then
            cmdok(3).BackColor = &H8000000F
        Else
            cmdok(3).BackColor = &HFF&
        End If
      Case Else:
         QueryServicePractice
   End Select
   
   ' 讀取優先權資料
   m_Pa(1) = m_TM01
   m_Pa(2) = m_TM02
   m_Pa(3) = m_TM03
   m_Pa(4) = m_TM04
   'edit by nickc 2007/02/06 不用 dll 了 objPublicData.ReadPriority m_Pa, m_Priority(1), m_Priority(2), m_Priority(3)
   'Modify by Amy 2014/04/17 +, m_Priority(4), m_Priority(5)
   'Modify by Sindy 2017/10/12 +, m_Priority(6)
   ClsPDReadPriority m_Pa, m_Priority(1), m_Priority(2), m_Priority(3), m_Priority(4), m_Priority(5), m_Priority(6)
   Set rsTmp = Nothing
QueryData = True
End Function
Private Sub Form_Load()
   'Added by Lydia 2016/09/10 設定代表人中文名稱和英文名稱長度
    textTM47.MaxLength = Pub_MaxCEL10
    textTM48.MaxLength = Pub_MaxCEL11
    textTM50.MaxLength = Pub_MaxCEL10
    textTM51.MaxLength = Pub_MaxCEL11
   'end 2016/09/10
   
   MoveFormToCenter Me
   Me.Tag = "frm030001_1"
   
   Me.SSTab1.Tab = 0 'Added by Lydia 2021/09/03
   
   SetPath 'Added by Morgan 2023/7/13
End Sub

Private Sub Form_Unload(Cancel As Integer)

Set frm030001_1 = Nothing
End Sub

Private Sub textCP44_GotFocus()
InverseTextBox textCP44
End Sub

Private Sub textCP44_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textCP44_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   'Dim oState As Boolean
   Dim strTempName As String   '2010/11/24 add by sonia
   
   If IsEmptyText(textCP44) = False Then
      '2010/11/24 modify by sonia 取消basQuery的GetFAgentNameAndState
      'oState = True
      'textCP44_2 = GetFAgentNameAndState(textCP44, oState)
      'If oState = False Then
      '      Cancel = True
      '      Exit Sub
      'End If
      
      'Add By Sindy 2013/5/23 加判斷是否為聯絡人
      If InStr(textCP44, "-") > 0 Then
         textCP44_2.Text = ""
         If ClsPDGetContact(textCP44, strTempName) Then
            textCP44_2 = strTempName
         End If
      Else
      '2013/5/23 End
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
      End If
      If IsEmptyText(textCP44_2) = True Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "代理人不存在"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP44_GotFocus
         Exit Sub
      End If
    End If
End Sub

Private Sub textPrintTNT_GotFocus()
InverseTextBox textPrintTNT
End Sub

Private Sub textPrintTNT_KeyPress(KeyAscii As Integer)
KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textPrintTNT_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textPrintTNT) = False Then
      Select Case textPrintTNT
       'Add by Lydia 2014/12/30 +DHL
       '  Case " ", "Y":
         Case " "
         Case "Y"
            textPrintDHL = ""
         Case Else
            Cancel = True
            strTit = "資料檢核"
            strMsg = "只可輸入空白或Y"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textPrintTNT_GotFocus
      End Select
   End If
End Sub

Private Sub textTM05_1_GotFocus()
InverseTextBox textTM05_1
'edit by nickc 2007/06/06 切換輸入法改用API
OpenIme
End Sub

Private Sub textTM05_GotFocus()
InverseTextBox textTM05
End Sub

Private Sub textTM06_GotFocus()
InverseTextBox textTM06
End Sub

Private Sub textTM07_GotFocus()
InverseTextBox textTM07
End Sub

Private Sub textTM09_GotFocus()
InverseTextBox textTM09
End Sub

Private Sub textTM09_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim strTemp As String
   Dim nResponse
   Dim nCount As Integer
   Dim nIndex As Integer
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   Cancel = False
   ' 無資料時不做任何檢查
   If IsEmptyText(textTM09) = True Then
      GoTo EXITSUB
   End If
   
   nCount = GetSubStringCount(textTM09)
   For nIndex = 1 To nCount
      strTemp = GetSubString(textTM09, nIndex)
      If Len(strTemp) > 6 Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "商品類別<" & strTemp & ">不正確"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textTM09_GotFocus
         GoTo EXITSUB
      End If
   Next nIndex
   
   For nIndex = 1 To nCount
      strTemp = GetSubString(textTM09, nIndex)
      For nCount = 1 To nCount
         If nIndex <> nCount Then
            If strTemp = GetSubString(textTM09, nCount) Then
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
textTM09 = Replace(textTM09, " ", "")
EXITSUB:
   Set rsTmp = Nothing
End Sub

Private Sub textTM23_GotFocus()
InverseTextBox textTM23
End Sub
'Modified by Lydia 2021/08/03  改成Form 2.0
'Private Sub textTM23_KeyPress(KeyAscii As Integer)
Private Sub textTM23_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textTM23_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   Cancel = False
   textTM23_2C = Empty
   textTM23_2E = Empty
   textTM23_2J = Empty
   If IsEmptyText(textTM23) = False Then
        Me.textTM23.Text = ChangeCustomerL(Me.textTM23.Text)
      'edit by 2004/07/22 nick  檢查該申請人或代理人狀態，若為不再使用則停在原地
      Dim oState As Boolean
      oState = True
      'textTM23_2 = GetCustomerName(textTM23, 0)
      strMsg = GetCustomerNameAndState_030001(textTM23, 0, oState)
      If oState = False Then
            Cancel = True
            Exit Sub
      End If
      If textTM23_2C = Empty And textTM23_2E = Empty And textTM23_2J = Empty Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "申請人代碼<" & textTM23 & ">不正確"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textTM23_GotFocus
      End If
   End If
   'Add By Cheng 2002/08/22
   If Cancel = False Then
      If Me.textTM23.Text <> m_strCust1 Then
         If Not PUB_EditCustOk(m_CP09, m_TM01, m_TM02, m_TM03, m_TM04) Then Cancel = True
      End If
   End If
   If Cancel = True Then textTM23_GotFocus
End Sub

Private Sub textTM67_GotFocus()
InverseTextBox textTM67
End Sub

Public Function GetCustomerNameAndState_030001(ByVal strCustomer As String, Optional ByVal nLanguage As String = "0", Optional oState As Boolean) As String
   Dim rsTmp As New ADODB.Recordset
   Dim strKey As String
   Dim strSql As String
   
  
   GetCustomerNameAndState_030001 = Empty
   
   If Len(strCustomer) < 9 Then: strCustomer = strCustomer & String(9 - Len(strCustomer), "0")
   
   If Len(strCustomer) > 8 Then
      strSql = "SELECT * FROM Customer " & _
               "WHERE CU01 = '" & Mid(strCustomer, 1, 8) & "' AND " & _
                     "CU02 = '" & Mid(strCustomer, 9, 1) & "'"
   Else
      strSql = "SELECT * FROM Customer " & _
               "WHERE CU01 = '" & Mid(strCustomer, 1, 8) & "' AND " & _
                     "CU02 = '0' "
   End If
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
            If IsNull(rsTmp.Fields("CU04")) = False Then
               textTM23_2C = rsTmp.Fields("CU04")
            End If
            If IsNull(rsTmp.Fields("CU05")) = False Then
               textTM23_2E = rsTmp.Fields("CU05")
            End If
            If IsNull(rsTmp.Fields("CU88")) = False Then
               textTM23_2E = textTM23_2E & " " & rsTmp.Fields("CU88")
            End If
            If IsNull(rsTmp.Fields("CU89")) = False Then
               textTM23_2E = textTM23_2E & " " & rsTmp.Fields("CU89")
            End If
            If IsNull(rsTmp.Fields("CU90")) = False Then
               textTM23_2E = textTM23_2E & " " & rsTmp.Fields("CU90")
            End If
            If IsNull(rsTmp.Fields("CU06")) = False Then
               textTM23_2J = rsTmp.Fields("CU06")
            End If
      If CheckStr(rsTmp.Fields("cu80").Value) = "不再使用" Then
             MsgBox "此申請人資料已不再使用，請確認！！", , MsgText(5)
             oState = False
      Else
             oState = True
      End If
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Function

Private Sub textTM05_1_Validate(Cancel As Boolean)
Dim strTit As String
Dim strMsg As String
Dim nResponse
   
    Cancel = False
    If CheckLengthIsOK(textTM05_1, 140) = False Then
        Cancel = True
        strTit = "檢核資料"
        strMsg = "案件名稱內容太長"
        textTM05_1_GotFocus
    End If
    'edit by nickc 2007/06/06 切換輸入法改用API
    If Cancel = False Then CloseIme
End Sub

' 案件中文名稱
Private Sub textTM05_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If CheckLengthIsOK(textTM05, 40) = False Then
      Cancel = True
      strTit = "檢核資料"
      strMsg = "案件中文名稱內容太長"
      'nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textTM05_GotFocus
   End If
End Sub

' 案件英文名稱
Private Sub textTM06_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If CheckLengthIsOK(textTM06, 60) = False Then
      Cancel = True
      strTit = "檢核資料"
      strMsg = "案件英文名稱內容太長"
      'nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textTM06_GotFocus
   End If
End Sub

' 案件日文名稱
Private Sub textTM07_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If CheckLengthIsOK(textTM07, 40) = False Then
      Cancel = True
      strTit = "檢核資料"
      strMsg = "案件日文名稱內容太長"
      'nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textTM07_GotFocus
   End If
End Sub
'Add by Morgan 2011/7/14 +可控制是否印信頭
Private Sub txtLetterHead_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 8 And KeyAscii <> Asc("N") Then
      KeyAscii = 0
   End If
End Sub
'Add by Lydia 2014/12/30 +DHL
Private Sub textPrintDHL_GotFocus()
InverseTextBox textPrintDHL
End Sub

Private Sub textPrintDHL_KeyPress(KeyAscii As Integer)
KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textPrintDHL_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textPrintDHL) = False Then
      Select Case textPrintDHL
         Case " "
         Case "Y"
            textPrintTNT = ""
         Case Else
            Cancel = True
            strTit = "資料檢核"
            strMsg = "只可輸入空白或Y"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textPrintDHL_GotFocus
      End Select
   End If
End Sub

'Add by Sindy 2013/5/23 CFT使用的是fagentConfig檔
'Move by Lydia 2018/01/15 從basQuery搬來
'Modified by Lydia 2018/01/16 +stDate3
Public Function PUB_ReadFTList_CFT(ByVal p_Sys As String, ByVal p_Cty As String, ByRef adoRst As ADODB.Recordset, Optional ByVal p_Date As String, Optional ByRef stDate3 As String) As Boolean
   Dim stYear As String, stDate1 As String, stDate2 As String, stFC05 As String
   Dim intR As Integer, stSQL As String
   Dim stVTB As String, stCon As String
   Dim stVTB2 As String, stCon2 As String 'Added by Lydia 2018/01/15
   If p_Date = "" Then
      p_Date = strSrvDate(1)
   Else
      p_Date = DBDATE(p_Date)
   End If
   stYear = Left(p_Date, 4)
   '下半年
   If Val(Mid(p_Date, 5, 2)) > 6 Then
      stFC05 = "2"
      stDate1 = stYear & "0701"
      stDate2 = stYear & "1231"
      stDate3 = stYear & "0101" 'Added by Lydia 2018/01/15 抓前半年
   '上半年
   Else
      stFC05 = "1"
      stDate1 = stYear & "0101"
      stDate2 = stYear & "0630"
      stDate3 = Val(stYear) - 1 & "0701" 'Added by Lydia 2018/01/15 抓前半年
   End If
   
   stCon = " AND FC04=" & (stYear - 1911) & " and FC05='" & stFC05 & "' and FC06='" & p_Sys & "'"
   stCon2 = Mid(stCon, InStr(stCon, "FC04")) 'Added by Lydia 2018/01/15
   
   '申請國家為歐盟時帶出所有歐洲代理人
   If p_Cty = "239" Then
      stCon = stCon & " and substr(fa10,1,1)='2'"
   Else
      stCon = stCon & " and substr(fa10,1,3)='" & p_Cty & "'"
   End If
   
   'Added by Lydia 2022/07/19 給案量開放可以輸入0，與之配合
   'CFT指示信frm030001_1的PUB_ReadFTList_CFT，讀取互惠代理人資料時請剔除建議給案量為0之代理人。
   'CFP案暫時不用改，他們是用FAgentTarget而不是fagentconfig，而且CFP只用了97及98二年，現在都不用了。
   stCon = stCon & " and nvl(fc07,0) > 0 "
   
   '已給案量統計
   stVTB = "select FC01||FC02||FC03 Cx1,nvl(count(*),0) Q1 From fagentConfig, fagent, caseprogress" & _
      " where FA01(+)=FC01 AND FA02(+)=FC02" & stCon & _
      " and cp44(+)=FC01||FC02 and cp44||cp116=FC01||FC02||FC03 and cp01(+)=FC06" & _
      " and cp10='101'" & _
      " and cp04='00' and cp27>=" & stDate1 & " and cp27<=" & stDate2 & _
      " group by FC01,FC02,FC03"
   
   'Added by Lydia 2018/01/15  +FC給案量
   'Modified by Lydia 2018/01/17  為了加速查詢O8速度,改抓CP139在半年內應該不會改代理人
   '                                             實測上不抓互惠記錄,依自動索引/*+ INDEX(CASEPROGRESS IDXCP010510132757) */ 為最快
   'stVTB2 = "select  substr(tm44,1,8) FC_A1,substr(tm44,9,1) FC_A2,count(*) FC_A3 " & _
                   "from (select DISTINCT FC01,'' FC03 from FAGENTCONFIG where " & stCon2 & ") VTB2,trademark,caseprogress " & _
                    "where tm44=FC01||'0' AND tm01||''='FCT' and cp01(+)=tm01 and cp02(+)=tm02 and cp03(+)=tm03 and cp04(+)=tm04 " & _
                    "and cp09<'B' and cp10='101' and cp159=0 and cp05>=" & stDate3 & " group by substr(tm44,1,8),substr(tm44,9,1) "
   stVTB2 = "select substr(tm44,1,8) FC_A1,substr(tm44,9,1) FC_A2,count(*) FC_A3 " & _
            "From caseprogress, Trademark " & _
            "where cp01='FCT' and cp09<'B' and cp10='101' and nvl(cp57,0)=0 " & _
            "and cp01=tm01(+) and cp02=tm02(+) and cp03=tm03(+) and cp04=tm04(+) " & _
            "and cp05>=" & stDate3 & " group by substr(tm44,1,8),substr(tm44,9,1) "
   'end 2018/01/17
   
   'Modified by Lydia 2018/01/15 +FC給案量
   'stSQL = "select '' C1,FC01||FC02||decode(FC03,null,'','-'||FC03) C2" & _
      ",decode(FC03,null,RTRIM(FA05||' '||FA63||' '||FA64||' '||FA65)||' '||FA06||' '||FA04" & _
      ",PCC03||' '||PCC04||' '||PCC05) C3,nvl(Q1,0) C4,FC07 C5,round(100*nvl(Q1,0)/FC07)||'%' C6,0 as C8,round(100*nvl(Q1,0)/FC07) C7" & _
      " From fagentConfig, fagent, PotCustCont, (" & stVTB & ") X" & _
      " where FA01(+)=FC01 AND FA02(+)=FC02 AND PCC01(+)=FC01 AND PCC02(+)=FC03" & stCon & _
      " and Cx1(+)=FC01||FC02||FC03 and nvl(Q1,0)<FC07" & _
      " order by 7,4,5,2,3"
   stSQL = "select '' C1,FC01||FC02||decode(FC03,null,'','-'||FC03) C2" & _
      ",decode(FC03,null,RTRIM(FA05||' '||FA63||' '||FA64||' '||FA65)||' '||FA06||' '||FA04" & _
      ",PCC03||' '||PCC04||' '||PCC05) C3,nvl(Q1,0) C4,FC07 C5,round(100*nvl(Q1,0)/FC07)||'%' C6,NVL(FC_A3,0) C8,round(100*nvl(Q1,0)/FC07) C7" & _
      " From fagentConfig, fagent, PotCustCont, (" & stVTB & ") X , (" & stVTB2 & ") N " & _
      " where FA01(+)=FC01 AND FA02(+)=FC02 AND PCC01(+)=FC01 AND PCC02(+)=FC03" & _
      " AND FA01=FC_A1(+) AND FA02=FC_A2(+) " & stCon & _
      " and Cx1(+)=FC01||FC02||FC03 and nvl(Q1,0)<FC07" & _
      " order by C7,C4,C5,C2,C3"
   'end 2018/01/15
   intR = 1
   Set adoRst = ClsLawReadRstMsg(intR, stSQL)
   If intR = 1 Then
      PUB_ReadFTList_CFT = True
   End If
End Function

Private Sub SetPath()
   '讀取前次設定路徑
   txtPath.Text = GetSetting("TAIE", "CFT", UCase(Me.Name) & "Dir", "")
   If txtPath <> "" Then ChkPath
End Sub

Private Function ChkPath() As Boolean
   If txtPath = "" Then
      MsgBox "[ 代表圖下載存放路徑 ] 尚未設定！", vbExclamation
   Else
      If PUB_ChkDir(txtPath) = True Then
         ChkPath = True
      Else
         MsgBox "代表圖下載存放路徑 [ " & txtPath & " ] 不存在，請重新設定！", vbCritical
         txtPath = ""
      End If
   End If
End Function

Private Sub txtPath_GotFocus()
   TextInverse txtPath
End Sub
