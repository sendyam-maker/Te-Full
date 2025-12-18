VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm090801_Q 
   BorderStyle     =   1  '單線固定
   Caption         =   "檢視接洽單"
   ClientHeight    =   8900
   ClientLeft      =   3400
   ClientTop       =   2720
   ClientWidth     =   8950
   ControlBox      =   0   'False
   HasDC           =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8900
   ScaleWidth      =   8950
   Begin VB.CheckBox Check10 
      Caption         =   "規費調整"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   195
      Left            =   2160
      TabIndex        =   707
      Top             =   30
      Width           =   1100
   End
   Begin VB.Frame Frame31 
      Height          =   2410
      Left            =   0
      TabIndex        =   668
      Top             =   6480
      Width           =   8920
      Begin VB.CommandButton cmdOK 
         Caption         =   "清除畫面(&C)"
         Height          =   350
         Index           =   2
         Left            =   7530
         TabIndex        =   681
         Top             =   1960
         Width           =   1110
      End
      Begin VB.Frame Frame44 
         Height          =   285
         Left            =   4290
         TabIndex        =   679
         Top             =   1140
         Visible         =   0   'False
         Width           =   8000
         Begin VB.Label Label25 
            Caption         =   "案件屬性："
            Height          =   180
            Left            =   0
            TabIndex        =   680
            Top             =   60
            Width           =   945
         End
      End
      Begin VB.ComboBox Combo1 
         Height          =   260
         Index           =   3
         Left            =   480
         TabIndex        =   678
         Top             =   450
         Width           =   2235
      End
      Begin VB.ComboBox Combo1 
         Height          =   260
         Index           =   2
         Left            =   480
         TabIndex        =   677
         Top             =   120
         Width           =   2235
      End
      Begin VB.ComboBox Combo1 
         Height          =   260
         Index           =   4
         Left            =   480
         TabIndex        =   676
         Top             =   780
         Width           =   2235
      End
      Begin VB.CheckBox Check6 
         Caption         =   "現金"
         Height          =   255
         Index           =   0
         Left            =   180
         TabIndex        =   675
         Top             =   1090
         Width           =   735
      End
      Begin VB.CheckBox Check6 
         Caption         =   "支票"
         Height          =   240
         Index           =   1
         Left            =   2160
         TabIndex        =   674
         Top             =   1090
         Width           =   975
      End
      Begin VB.CommandButton cmdSend 
         Caption         =   "待收文區(10)"
         Height          =   350
         Left            =   5640
         TabIndex        =   673
         Top             =   1570
         Width           =   1170
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "新增"
         Height          =   350
         Index           =   0
         Left            =   7020
         TabIndex        =   672
         Top             =   1570
         Width           =   825
      End
      Begin VB.CommandButton cmdUpd 
         BackColor       =   &H0080FF80&
         Caption         =   "加入"
         Height          =   320
         Left            =   5430
         Style           =   1  '圖片外觀
         TabIndex        =   671
         Top             =   1960
         Width           =   525
      End
      Begin VB.CommandButton cmdDel 
         BackColor       =   &H0080FF80&
         Caption         =   "刪除"
         Height          =   320
         Left            =   6030
         Style           =   1  '圖片外觀
         TabIndex        =   670
         Top             =   1960
         Width           =   525
      End
      Begin VB.CommandButton cmdClear2 
         BackColor       =   &H0080FF80&
         Caption         =   "清除"
         Height          =   320
         Left            =   6630
         Style           =   1  '圖片外觀
         TabIndex        =   669
         Top             =   1960
         Width           =   525
      End
      Begin VB.Label Label1 
         Caption         =   "這組(2)物件目前程式有用到"
         Height          =   410
         Index           =   49
         Left            =   7440
         TabIndex        =   706
         Top             =   90
         Width           =   1130
      End
      Begin VB.Label Label1 
         Caption         =   "4."
         Height          =   260
         Index           =   98
         Left            =   90
         TabIndex        =   705
         Top             =   810
         Width           =   1010
      End
      Begin VB.Label Label1 
         Caption         =   "3."
         Height          =   260
         Index           =   97
         Left            =   90
         TabIndex        =   704
         Top             =   480
         Width           =   1010
      End
      Begin VB.Label Label1 
         Caption         =   "2."
         Height          =   260
         Index           =   96
         Left            =   90
         TabIndex        =   703
         Top             =   150
         Width           =   1010
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   111
         Left            =   3900
         TabIndex        =   702
         Top             =   780
         Width           =   1100
         VariousPropertyBits=   671107099
         Size            =   "1931;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   112
         Left            =   4980
         TabIndex        =   701
         Top             =   780
         Width           =   920
         VariousPropertyBits=   671107099
         Size            =   "1614;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   110
         Left            =   2760
         TabIndex        =   700
         Top             =   780
         Width           =   1130
         VariousPropertyBits=   671107099
         Size            =   "1984;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   105
         Left            =   3900
         TabIndex        =   699
         Top             =   120
         Width           =   1100
         VariousPropertyBits=   671107099
         Size            =   "1931;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   106
         Left            =   4980
         TabIndex        =   698
         Top             =   120
         Width           =   920
         VariousPropertyBits=   671107099
         Size            =   "1614;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   108
         Left            =   3900
         TabIndex        =   697
         Top             =   450
         Width           =   1100
         VariousPropertyBits=   671107099
         Size            =   "1931;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   109
         Left            =   4980
         TabIndex        =   696
         Top             =   450
         Width           =   920
         VariousPropertyBits=   671107099
         Size            =   "1614;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   104
         Left            =   2760
         TabIndex        =   695
         Top             =   120
         Width           =   1130
         VariousPropertyBits=   671107099
         Size            =   "1984;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   107
         Left            =   2760
         TabIndex        =   694
         Top             =   450
         Width           =   1130
         VariousPropertyBits=   671107099
         Size            =   "1984;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.ComboBox Combo2 
         Height          =   300
         Index           =   3
         Left            =   5880
         TabIndex        =   693
         Top             =   780
         Width           =   1520
         VariousPropertyBits=   679495707
         DisplayStyle    =   3
         Size            =   "2672;529"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.ComboBox Combo2 
         Height          =   300
         Index           =   1
         Left            =   5880
         TabIndex        =   692
         Top             =   120
         Width           =   1520
         VariousPropertyBits=   679495707
         DisplayStyle    =   3
         Size            =   "2672;529"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.ComboBox Combo2 
         Height          =   300
         Index           =   2
         Left            =   5880
         TabIndex        =   691
         Top             =   450
         Width           =   1520
         VariousPropertyBits=   679495707
         DisplayStyle    =   3
         Size            =   "2672;529"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   116
         Left            =   960
         TabIndex        =   690
         Top             =   1060
         Width           =   1130
         VariousPropertyBits=   671107099
         Size            =   "1984;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   118
         Left            =   1350
         TabIndex        =   689
         Top             =   1360
         Width           =   1130
         VariousPropertyBits=   671107099
         Size            =   "1984;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         Caption         =   "預定收款日期："
         Height          =   290
         Index           =   105
         Left            =   90
         TabIndex        =   688
         Top             =   1390
         Width           =   1280
      End
      Begin VB.Label Label28 
         Caption         =   "民國年月日"
         ForeColor       =   &H00FF0000&
         Height          =   290
         Index           =   2
         Left            =   2520
         TabIndex        =   687
         Top             =   1360
         Width           =   950
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   144
         Left            =   4710
         TabIndex        =   686
         Top             =   1510
         Width           =   1130
         VariousPropertyBits=   671107099
         Size            =   "1984;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         Caption         =   "關連表單編號："
         Height          =   290
         Index           =   136
         Left            =   3450
         TabIndex        =   685
         Top             =   1540
         Width           =   1280
      End
      Begin VB.Label Label1 
         Caption         =   "到期日："
         Height          =   290
         Index           =   139
         Left            =   270
         TabIndex        =   684
         Top             =   1850
         Width           =   770
      End
      Begin VB.Label Label28 
         Caption         =   "Ex: 930820 "
         Height          =   310
         Index           =   0
         Left            =   2250
         TabIndex        =   683
         Top             =   1820
         Width           =   950
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   117
         Left            =   1050
         TabIndex        =   682
         Top             =   1790
         Width           =   1130
         VariousPropertyBits=   671107099
         Size            =   "1984;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
   End
   Begin VB.CheckBox ChkCRL152 
      Caption         =   "自行送簽核"
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
      Left            =   1350
      TabIndex        =   649
      Top             =   300
      Width           =   1305
   End
   Begin VB.Frame FrameCRL66 
      Caption         =   "Frame31"
      Height          =   255
      Left            =   0
      TabIndex        =   637
      Top             =   0
      Width           =   1305
      Begin VB.CheckBox ChkCRL66 
         BackColor       =   &H0000C000&
         Caption         =   "對造已簽准"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   638
         Top             =   0
         Width           =   1305
      End
   End
   Begin VB.Frame FrameCRL90 
      Caption         =   "Frame32"
      Height          =   255
      Left            =   1350
      TabIndex        =   635
      Top             =   0
      Width           =   765
      Begin VB.CheckBox Check11 
         BackColor       =   &H008080FF&
         Caption         =   "急件"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   636
         Top             =   0
         Width           =   765
      End
   End
   Begin VB.Frame FrameCRL147 
      Caption         =   "Frame31"
      Height          =   255
      Left            =   0
      TabIndex        =   633
      Top             =   270
      Width           =   1305
      Begin VB.CheckBox Check12 
         BackColor       =   &H0080FFFF&
         Caption         =   "費用已核准"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   634
         Top             =   0
         Width           =   1305
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000009&
      Height          =   315
      Left            =   -7680
      ScaleHeight     =   280
      ScaleWidth      =   290
      TabIndex        =   50
      Top             =   30
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "查詢"
      CausesValidation=   0   'False
      Height          =   350
      Index           =   4
      Left            =   7140
      TabIndex        =   47
      Top             =   30
      Visible         =   0   'False
      Width           =   770
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   350
      Index           =   1
      Left            =   7980
      TabIndex        =   46
      Top             =   30
      Width           =   770
   End
   Begin VB.TextBox Text5 
      Height          =   315
      Left            =   4620
      Locked          =   -1  'True
      TabIndex        =   45
      Top             =   0
      Visible         =   0   'False
      Width           =   1070
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H0000C0C0&
      Caption         =   "案源輸入(&I)"
      Height          =   350
      Index           =   5
      Left            =   6030
      Style           =   1  '圖片外觀
      TabIndex        =   119
      Top             =   30
      Visible         =   0   'False
      Width           =   1050
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5940
      Left            =   12
      TabIndex        =   48
      Top             =   540
      Width           =   8892
      _ExtentX        =   15699
      _ExtentY        =   10478
      _Version        =   393216
      Tabs            =   6
      Tab             =   4
      TabsPerRow      =   6
      TabHeight       =   520
      ForeColor       =   -2147483633
      TabCaption(0)   =   "案件/收費項目"
      TabPicture(0)   =   "frm090801_Q.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame33(0)"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "來源及申請人"
      TabPicture(1)   =   "frm090801_Q.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "SSTab2"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "說明處理事項"
      TabPicture(2)   =   "frm090801_Q.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "SSTab3"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "發明人資料"
      TabPicture(3)   =   "frm090801_Q.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "ChkAddress(9)"
      Tab(3).Control(1)=   "ChkAddress(8)"
      Tab(3).Control(2)=   "ChkAddress(7)"
      Tab(3).Control(3)=   "ChkAddress(6)"
      Tab(3).Control(4)=   "ChkAddress(5)"
      Tab(3).Control(5)=   "ChkAddress(4)"
      Tab(3).Control(6)=   "ChkAddress(3)"
      Tab(3).Control(7)=   "ChkAddress(2)"
      Tab(3).Control(8)=   "ChkAddress(1)"
      Tab(3).Control(9)=   "ChkAddress(0)"
      Tab(3).Control(10)=   "Combo3(9)"
      Tab(3).Control(11)=   "Text3(9)"
      Tab(3).Control(12)=   "Combo3(8)"
      Tab(3).Control(13)=   "Text3(8)"
      Tab(3).Control(14)=   "Combo3(7)"
      Tab(3).Control(15)=   "Text3(7)"
      Tab(3).Control(16)=   "Combo3(6)"
      Tab(3).Control(17)=   "Text3(6)"
      Tab(3).Control(18)=   "Combo3(5)"
      Tab(3).Control(19)=   "Text3(5)"
      Tab(3).Control(20)=   "Combo3(4)"
      Tab(3).Control(21)=   "Text3(4)"
      Tab(3).Control(22)=   "Combo3(3)"
      Tab(3).Control(23)=   "Text3(3)"
      Tab(3).Control(24)=   "Combo3(2)"
      Tab(3).Control(25)=   "Text3(2)"
      Tab(3).Control(26)=   "Combo3(1)"
      Tab(3).Control(27)=   "Text3(1)"
      Tab(3).Control(28)=   "Combo3(0)"
      Tab(3).Control(29)=   "Text3(0)"
      Tab(3).Control(30)=   "Text4(0)"
      Tab(3).Control(31)=   "Text2(9)"
      Tab(3).Control(32)=   "Text2(8)"
      Tab(3).Control(33)=   "Text2(7)"
      Tab(3).Control(34)=   "Text2(6)"
      Tab(3).Control(35)=   "Text2(5)"
      Tab(3).Control(36)=   "Text2(4)"
      Tab(3).Control(37)=   "Text2(3)"
      Tab(3).Control(38)=   "Text2(2)"
      Tab(3).Control(39)=   "Text2(1)"
      Tab(3).Control(40)=   "Text2(0)"
      Tab(3).Control(41)=   "Label6(1)"
      Tab(3).Control(42)=   "Label10"
      Tab(3).Control(43)=   "Label5(9)"
      Tab(3).Control(44)=   "Label5(8)"
      Tab(3).Control(45)=   "Label5(7)"
      Tab(3).Control(46)=   "Label5(6)"
      Tab(3).Control(47)=   "Label5(5)"
      Tab(3).Control(48)=   "Label5(4)"
      Tab(3).Control(49)=   "Label5(3)"
      Tab(3).Control(50)=   "Label5(2)"
      Tab(3).Control(51)=   "Label5(1)"
      Tab(3).Control(52)=   "Label5(0)"
      Tab(3).Control(53)=   "Label9(0)"
      Tab(3).Control(54)=   "Label8(0)"
      Tab(3).Control(55)=   "Label7(0)"
      Tab(3).Control(56)=   "Label6(0)"
      Tab(3).Control(57)=   "Text4(9)"
      Tab(3).Control(58)=   "Text4(8)"
      Tab(3).Control(59)=   "Text4(7)"
      Tab(3).Control(60)=   "Text4(6)"
      Tab(3).Control(61)=   "Text4(5)"
      Tab(3).Control(62)=   "Text4(4)"
      Tab(3).Control(63)=   "Text4(3)"
      Tab(3).Control(64)=   "Text4(2)"
      Tab(3).Control(65)=   "Text4(1)"
      Tab(3).ControlCount=   66
      TabCaption(4)   =   "商標圖及說明"
      TabPicture(4)   =   "frm090801_Q.frx":0070
      Tab(4).ControlEnabled=   -1  'True
      Tab(4).Control(0)=   "PicText"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).Control(1)=   "G_SeekPicColor"
      Tab(4).Control(1).Enabled=   0   'False
      Tab(4).Control(2)=   "Frame47"
      Tab(4).Control(2).Enabled=   0   'False
      Tab(4).Control(3)=   "Frame5"
      Tab(4).Control(3).Enabled=   0   'False
      Tab(4).Control(4)=   "tmpPic"
      Tab(4).Control(4).Enabled=   0   'False
      Tab(4).Control(5)=   "cmdPic"
      Tab(4).Control(5).Enabled=   0   'False
      Tab(4).Control(6)=   "Frame15"
      Tab(4).Control(6).Enabled=   0   'False
      Tab(4).Control(7)=   "Frame17"
      Tab(4).Control(7).Enabled=   0   'False
      Tab(4).Control(8)=   "pic1"
      Tab(4).Control(8).Enabled=   0   'False
      Tab(4).Control(9)=   "FRTMQ"
      Tab(4).Control(9).Enabled=   0   'False
      Tab(4).Control(10)=   "Frame29"
      Tab(4).Control(10).Enabled=   0   'False
      Tab(4).Control(11)=   "cmdSavePic"
      Tab(4).Control(11).Enabled=   0   'False
      Tab(4).ControlCount=   12
      TabCaption(5)   =   "簽核狀況"
      TabPicture(5)   =   "frm090801_Q.frx":008C
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "txtF0309"
      Tab(5).Control(1)=   "txtF0310"
      Tab(5).Control(2)=   "Frame57"
      Tab(5).Control(3)=   "GRD1"
      Tab(5).Control(4)=   "Label29"
      Tab(5).Control(5)=   "LblReason"
      Tab(5).Control(6)=   "Label36"
      Tab(5).Control(7)=   "Label37"
      Tab(5).Control(8)=   "txtF0310_2"
      Tab(5).Control(9)=   "txtF0407"
      Tab(5).Control(10)=   "Label30"
      Tab(5).Control(11)=   "Label32"
      Tab(5).Control(12)=   "Label1(128)"
      Tab(5).Control(13)=   "txtF0306"
      Tab(5).Control(14)=   "Label33"
      Tab(5).Control(15)=   "Label1(130)"
      Tab(5).Control(16)=   "Label35"
      Tab(5).Control(17)=   "txtCRL69"
      Tab(5).Control(18)=   "LblRecved"
      Tab(5).ControlCount=   19
      Begin VB.CommandButton cmdSavePic 
         Caption         =   "下載"
         Height          =   300
         Left            =   7464
         TabIndex        =   657
         Top             =   5112
         Visible         =   0   'False
         Width           =   1080
      End
      Begin VB.Frame Frame29 
         Height          =   375
         Left            =   390
         TabIndex        =   642
         Top             =   480
         Width           =   4335
         Begin MSForms.TextBox Text1 
            Height          =   300
            Index           =   129
            Left            =   720
            TabIndex        =   643
            Top             =   0
            Width           =   1305
            VariousPropertyBits=   671107099
            Size            =   "3016;529"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox Text1 
            Height          =   300
            Index           =   132
            Left            =   2910
            TabIndex        =   644
            Top             =   0
            Width           =   1305
            VariousPropertyBits=   671107099
            Size            =   "3016;529"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin VB.Label Label1 
            Caption         =   "申請案號："
            Height          =   255
            Index           =   23
            Left            =   2050
            TabIndex        =   646
            Top             =   60
            Width           =   1005
         End
         Begin VB.Label Label1 
            Caption         =   "審定號："
            Height          =   255
            Index           =   126
            Left            =   0
            TabIndex        =   645
            Top             =   60
            Width           =   765
         End
      End
      Begin VB.Frame FRTMQ 
         Caption         =   "委查結果"
         Height          =   2655
         Left            =   330
         TabIndex        =   312
         Top             =   2430
         Width           =   4335
         Begin VB.CommandButton cmdTMQApp 
            Caption         =   "查名單輸入"
            Height          =   300
            Left            =   1850
            TabIndex        =   314
            Top             =   240
            Width           =   1080
         End
         Begin VB.CommandButton cmdTMQ 
            Caption         =   "查覆區"
            Height          =   300
            Left            =   3000
            TabIndex        =   313
            Top             =   240
            Width           =   1080
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid GrdTMQ 
            Height          =   1935
            Left            =   120
            TabIndex        =   315
            Top             =   600
            Width           =   4095
            _ExtentX        =   7214
            _ExtentY        =   3404
            _Version        =   393216
            Cols            =   5
            FixedCols       =   0
            AllowUserResizing=   3
            FormatString    =   "V|已讀(Y)|申請編號|委查編號|結果"
            _NumberOfBands  =   1
            _Band(0).Cols   =   5
         End
         Begin VB.Label Label20 
            Caption         =   "勾選""V"" 表示收文"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   360
            TabIndex        =   316
            Top             =   300
            Width           =   1935
         End
      End
      Begin VB.PictureBox pic1 
         Height          =   255
         Left            =   4560
         ScaleHeight     =   220
         ScaleWidth      =   220
         TabIndex        =   338
         Top             =   4710
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Frame Frame17 
         Height          =   255
         Left            =   2340
         TabIndex        =   335
         Top             =   1530
         Visible         =   0   'False
         Width           =   2160
         Begin VB.OptionButton optColor 
            Caption         =   "黑白/灰階"
            Height          =   180
            Index           =   0
            Left            =   0
            TabIndex        =   337
            Top             =   30
            Value           =   -1  'True
            Width           =   1125
         End
         Begin VB.OptionButton optColor 
            Caption         =   "彩色"
            Height          =   180
            Index           =   1
            Left            =   1380
            TabIndex        =   336
            Top             =   45
            Width           =   675
         End
      End
      Begin VB.Frame Frame15 
         Height          =   465
         Left            =   390
         TabIndex        =   329
         Top             =   870
         Width           =   4515
         Begin VB.OptionButton opt1 
            Caption         =   "圖形附檔"
            Height          =   255
            Index           =   1
            Left            =   780
            TabIndex        =   334
            Top             =   180
            Width           =   1065
         End
         Begin VB.OptionButton opt1 
            Caption         =   "文字"
            Height          =   255
            Index           =   0
            Left            =   30
            TabIndex        =   333
            Top             =   180
            Width           =   675
         End
         Begin VB.OptionButton opt1 
            Caption         =   "圖後補"
            Height          =   255
            Index           =   2
            Left            =   1890
            TabIndex        =   332
            Top             =   180
            Width           =   915
         End
         Begin VB.OptionButton opt1 
            Caption         =   "其他"
            Height          =   255
            Index           =   4
            Left            =   3750
            TabIndex        =   331
            Top             =   180
            Width           =   675
         End
         Begin VB.OptionButton opt1 
            Caption         =   "同卷號"
            Height          =   255
            Index           =   3
            Left            =   2820
            TabIndex        =   330
            Top             =   180
            Width           =   915
         End
      End
      Begin VB.CommandButton cmdPic 
         Caption         =   "選擇圖檔"
         Enabled         =   0   'False
         Height          =   465
         Left            =   1140
         TabIndex        =   328
         Top             =   1410
         Width           =   1125
      End
      Begin VB.PictureBox tmpPic 
         Height          =   4455
         Left            =   4950
         ScaleHeight     =   442
         ScaleMode       =   3  '像素
         ScaleWidth      =   355
         TabIndex        =   327
         Top             =   570
         Width           =   3588
         Begin VB.Image tmpImg 
            Height          =   1770
            Left            =   1425
            Stretch         =   -1  'True
            Top             =   1095
            Visible         =   0   'False
            Width           =   1890
         End
      End
      Begin VB.Frame Frame5 
         Height          =   550
         Left            =   285
         TabIndex        =   321
         Top             =   1812
         Width           =   4430
         Begin VB.TextBox TxtC1 
            Height          =   300
            Index           =   0
            Left            =   1320
            TabIndex        =   325
            Top             =   120
            Width           =   465
         End
         Begin VB.TextBox TxtC1 
            Height          =   300
            Index           =   1
            Left            =   1830
            TabIndex        =   324
            Top             =   120
            Width           =   765
         End
         Begin VB.TextBox TxtC1 
            Height          =   300
            Index           =   2
            Left            =   2640
            TabIndex        =   323
            Top             =   120
            Width           =   225
         End
         Begin VB.TextBox TxtC1 
            Height          =   300
            Index           =   3
            Left            =   2930
            TabIndex        =   322
            Top             =   120
            Width           =   345
         End
         Begin VB.Line Line22 
            X1              =   1740
            X2              =   3200
            Y1              =   270
            Y2              =   270
         End
         Begin VB.Label Label1 
            Caption         =   "圖同本所案號："
            Height          =   255
            Index           =   24
            Left            =   50
            TabIndex        =   326
            Top             =   165
            Width           =   1305
         End
      End
      Begin VB.Frame Frame47 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   0  '沒有框線
         Height          =   1215
         Left            =   390
         TabIndex        =   318
         Top             =   3780
         Width           =   3945
         Begin VB.Label fra47Title 
            Caption         =   "商標說明："
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.5
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   30
            TabIndex        =   320
            Top             =   0
            Width           =   2535
         End
         Begin MSForms.TextBox Text7 
            Height          =   915
            Left            =   0
            TabIndex        =   319
            Top             =   270
            Width           =   3855
            VariousPropertyBits=   671107099
            Size            =   "6800;1614"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
      End
      Begin VB.PictureBox G_SeekPicColor 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
         Enabled         =   0   'False
         Height          =   300
         Left            =   210
         ScaleHeight     =   26
         ScaleMode       =   3  '像素
         ScaleWidth      =   20
         TabIndex        =   317
         Top             =   420
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.CheckBox ChkAddress 
         Height          =   480
         Index           =   9
         Left            =   -66510
         TabIndex        =   235
         Top             =   4815
         Width           =   285
      End
      Begin VB.CheckBox ChkAddress 
         Height          =   480
         Index           =   8
         Left            =   -66510
         TabIndex        =   234
         Top             =   4395
         Width           =   285
      End
      Begin VB.CheckBox ChkAddress 
         Height          =   480
         Index           =   7
         Left            =   -66510
         TabIndex        =   233
         Top             =   4035
         Width           =   285
      End
      Begin VB.CheckBox ChkAddress 
         Height          =   480
         Index           =   6
         Left            =   -66510
         TabIndex        =   232
         Top             =   3615
         Width           =   285
      End
      Begin VB.CheckBox ChkAddress 
         Height          =   480
         Index           =   5
         Left            =   -66510
         TabIndex        =   231
         Top             =   3225
         Width           =   285
      End
      Begin VB.CheckBox ChkAddress 
         Height          =   480
         Index           =   4
         Left            =   -66510
         TabIndex        =   230
         Top             =   2865
         Width           =   285
      End
      Begin VB.CheckBox ChkAddress 
         Height          =   480
         Index           =   3
         Left            =   -66510
         TabIndex        =   229
         Top             =   2505
         Width           =   285
      End
      Begin VB.CheckBox ChkAddress 
         Height          =   480
         Index           =   2
         Left            =   -66510
         TabIndex        =   228
         Top             =   2085
         Width           =   285
      End
      Begin VB.CheckBox ChkAddress 
         Height          =   480
         Index           =   1
         Left            =   -66510
         TabIndex        =   227
         Top             =   1635
         Width           =   285
      End
      Begin VB.CheckBox ChkAddress 
         Height          =   480
         Index           =   0
         Left            =   -66510
         TabIndex        =   226
         Top             =   1275
         Width           =   285
      End
      Begin VB.ComboBox Combo3 
         Height          =   260
         Index           =   9
         Left            =   -73320
         TabIndex        =   225
         Top             =   4875
         Width           =   1575
      End
      Begin VB.TextBox Text3 
         Height          =   315
         Index           =   9
         Left            =   -71700
         TabIndex        =   224
         Top             =   4875
         Width           =   1335
      End
      Begin VB.ComboBox Combo3 
         Height          =   260
         Index           =   8
         Left            =   -73320
         TabIndex        =   223
         Top             =   4485
         Width           =   1575
      End
      Begin VB.TextBox Text3 
         Height          =   315
         Index           =   8
         Left            =   -71700
         TabIndex        =   222
         Top             =   4485
         Width           =   1335
      End
      Begin VB.ComboBox Combo3 
         Height          =   260
         Index           =   7
         Left            =   -73320
         TabIndex        =   221
         Top             =   4095
         Width           =   1575
      End
      Begin VB.TextBox Text3 
         Height          =   315
         Index           =   7
         Left            =   -71700
         TabIndex        =   220
         Top             =   4095
         Width           =   1335
      End
      Begin VB.ComboBox Combo3 
         Height          =   260
         Index           =   6
         Left            =   -73320
         TabIndex        =   219
         Top             =   3705
         Width           =   1575
      End
      Begin VB.TextBox Text3 
         Height          =   315
         Index           =   6
         Left            =   -71700
         TabIndex        =   218
         Top             =   3705
         Width           =   1335
      End
      Begin VB.ComboBox Combo3 
         Height          =   260
         Index           =   5
         Left            =   -73320
         TabIndex        =   217
         Top             =   3315
         Width           =   1575
      End
      Begin VB.TextBox Text3 
         Height          =   315
         Index           =   5
         Left            =   -71700
         TabIndex        =   216
         Top             =   3315
         Width           =   1335
      End
      Begin VB.ComboBox Combo3 
         Height          =   260
         Index           =   4
         Left            =   -73320
         TabIndex        =   215
         Top             =   2925
         Width           =   1575
      End
      Begin VB.TextBox Text3 
         Height          =   315
         Index           =   4
         Left            =   -71700
         TabIndex        =   214
         Top             =   2925
         Width           =   1335
      End
      Begin VB.ComboBox Combo3 
         Height          =   260
         Index           =   3
         Left            =   -73320
         TabIndex        =   213
         Top             =   2535
         Width           =   1575
      End
      Begin VB.TextBox Text3 
         Height          =   315
         Index           =   3
         Left            =   -71700
         TabIndex        =   212
         Top             =   2535
         Width           =   1335
      End
      Begin VB.ComboBox Combo3 
         Height          =   260
         Index           =   2
         Left            =   -73320
         TabIndex        =   211
         Top             =   2145
         Width           =   1575
      End
      Begin VB.TextBox Text3 
         Height          =   315
         Index           =   2
         Left            =   -71700
         TabIndex        =   210
         Top             =   2145
         Width           =   1335
      End
      Begin VB.ComboBox Combo3 
         Height          =   260
         Index           =   1
         Left            =   -73320
         TabIndex        =   209
         Top             =   1755
         Width           =   1575
      End
      Begin VB.TextBox Text3 
         Height          =   315
         Index           =   1
         Left            =   -71700
         TabIndex        =   208
         Top             =   1755
         Width           =   1335
      End
      Begin VB.ComboBox Combo3 
         Height          =   260
         Index           =   0
         Left            =   -73320
         TabIndex        =   207
         Top             =   1365
         Width           =   1575
      End
      Begin VB.TextBox Text3 
         Height          =   315
         Index           =   0
         Left            =   -71700
         TabIndex        =   206
         Top             =   1365
         Width           =   1335
      End
      Begin VB.TextBox txtF0309 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  '沒有框線
         ForeColor       =   &H000000C0&
         Height          =   240
         Left            =   -73605
         Locked          =   -1  'True
         TabIndex        =   193
         Text            =   "txtF0309"
         Top             =   840
         Width           =   1845
      End
      Begin VB.TextBox txtF0310 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  '沒有框線
         BeginProperty Font 
            Name            =   "新細明體-ExtB"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000012&
         Height          =   255
         Left            =   -73995
         Locked          =   -1  'True
         TabIndex        =   192
         Text            =   "txtF0310"
         Top             =   540
         Width           =   645
      End
      Begin VB.Frame Frame57 
         Appearance      =   0  '平面
         Caption         =   "Frame57"
         ForeColor       =   &H80000008&
         Height          =   1425
         Left            =   -69765
         TabIndex        =   187
         Top             =   840
         Visible         =   0   'False
         Width           =   3555
         Begin MSForms.TextBox Text1 
            Height          =   630
            Index           =   136
            Left            =   60
            TabIndex        =   626
            Top             =   750
            Width           =   3435
            VariousPropertyBits=   -1476376557
            ForeColor       =   16512
            BorderStyle     =   1
            ScrollBars      =   2
            Size            =   "6059;1111"
            Value           =   "計件值(2)"
            SpecialEffect   =   0
            FontName        =   "新細明體-ExtB"
            FontEffects     =   1073741825
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
            FontWeight      =   700
         End
         Begin MSForms.TextBox Text1 
            Height          =   300
            Index           =   134
            Left            =   1500
            TabIndex        =   191
            Top             =   150
            Width           =   1965
            VariousPropertyBits=   671107099
            Size            =   "3466;529"
            Value           =   "CRL67"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin VB.Label Label1 
            Caption         =   "一案兩請："
            Height          =   255
            Index           =   134
            Left            =   60
            TabIndex        =   190
            Top             =   210
            Width           =   1005
         End
         Begin MSForms.TextBox Text1 
            Height          =   300
            Index           =   135
            Left            =   1500
            TabIndex        =   189
            Top             =   450
            Width           =   1965
            VariousPropertyBits=   671107099
            Size            =   "3466;529"
            Value           =   "CRL68"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin VB.Label Label1 
            Caption         =   "擬制喪失新穎性："
            Height          =   255
            Index           =   135
            Left            =   60
            TabIndex        =   188
            Top             =   510
            Width           =   1485
         End
      End
      Begin VB.Frame Frame33 
         Height          =   5565
         Index           =   0
         Left            =   -74940
         TabIndex        =   89
         Top             =   330
         Width           =   8805
         Begin VB.Frame Frame33 
            BackColor       =   &H00FFFFC0&
            BorderStyle     =   0  '沒有框線
            Height          =   804
            Index           =   1
            Left            =   96
            TabIndex        =   658
            Top             =   4776
            Width           =   3492
            Begin VB.CheckBox Check7 
               Caption         =   "是"
               ForeColor       =   &H00FF0000&
               Height          =   225
               Index           =   2
               Left            =   840
               TabIndex        =   660
               Top             =   576
               Width           =   495
            End
            Begin VB.CheckBox Check7 
               Caption         =   "否"
               ForeColor       =   &H00FF0000&
               Height          =   225
               Index           =   3
               Left            =   1800
               TabIndex        =   659
               Top             =   576
               Width           =   495
            End
            Begin MSForms.OptionButton optDB 
               Height          =   228
               Index           =   0
               Left            =   96
               TabIndex        =   662
               Top             =   48
               Width           =   2532
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               DisplayStyle    =   5
               Size            =   "4466;402"
               Value           =   "0"
               Caption         =   "立即開立DEBIT NOTE"
               FontName        =   "新細明體"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.OptionButton optDB 
               Height          =   300
               Index           =   1
               Left            =   96
               TabIndex        =   661
               Top             =   288
               Width           =   3360
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               DisplayStyle    =   5
               Size            =   "5927;529"
               Value           =   "0"
               Caption         =   "待通知後開立，是否要加印國內收據：　　　"
               FontName        =   "新細明體"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
         End
         Begin VB.CommandButton cmdAddAtt 
            BackColor       =   &H00C0C0C0&
            Caption         =   "文件匯入"
            CausesValidation=   0   'False
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   10
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   7560
            Style           =   1  '圖片外觀
            TabIndex        =   17
            Top             =   630
            Width           =   1275
         End
         Begin VB.CommandButton cmdCRL119 
            BackColor       =   &H00C0C0C0&
            Caption         =   "特殊收據"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   6450
            Style           =   1  '圖片外觀
            TabIndex        =   44
            Top             =   5280
            Visible         =   0   'False
            Width           =   1005
         End
         Begin VB.Frame FrameCRC 
            BackColor       =   &H00C0FFFF&
            Caption         =   "案件性質區"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000040C0&
            Height          =   2175
            Left            =   30
            TabIndex        =   120
            Top             =   2280
            Width           =   7035
            Begin VB.CheckBox chkEnglish 
               BackColor       =   &H00C0FFFF&
               Caption         =   "同時申請三國(含)以上之美日德"
               ForeColor       =   &H00FF0000&
               Height          =   255
               Left            =   4020
               TabIndex        =   641
               Top             =   120
               Visible         =   0   'False
               Width           =   2800
            End
            Begin VB.ComboBox Combo1 
               Height          =   260
               Index           =   1
               Left            =   240
               TabIndex        =   30
               Top             =   540
               Width           =   2145
            End
            Begin MSHierarchicalFlexGridLib.MSHFlexGrid GridCase 
               Height          =   1275
               Left            =   60
               TabIndex        =   35
               Top             =   840
               Width           =   6915
               _ExtentX        =   12188
               _ExtentY        =   2258
               _Version        =   393216
               Cols            =   6
               FixedCols       =   0
               AllowUserResizing=   2
               FormatString    =   "順序|案件性質|費用|規費|點數|備註"
               _NumberOfBands  =   1
               _Band(0).Cols   =   6
            End
            Begin MSForms.ComboBox Combo2 
               Height          =   300
               Index           =   0
               Left            =   5460
               TabIndex        =   34
               Top             =   540
               Width           =   1515
               VariousPropertyBits=   679495707
               DisplayStyle    =   3
               Size            =   "2672;529"
               MatchEntry      =   1
               ShowDropButtonWhen=   2
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.TextBox Text1 
               Height          =   300
               Index           =   101
               Left            =   2370
               TabIndex        =   31
               Top             =   540
               Width           =   1125
               VariousPropertyBits=   671107099
               Size            =   "1984;529"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.TextBox Text1 
               Height          =   300
               Index           =   103
               Left            =   4560
               TabIndex        =   33
               Top             =   540
               Width           =   915
               VariousPropertyBits=   671107099
               Size            =   "1614;529"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.TextBox Text1 
               Height          =   300
               Index           =   102
               Left            =   3510
               TabIndex        =   32
               Top             =   540
               Width           =   1065
               VariousPropertyBits=   671107099
               Size            =   "1879;529"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin VB.Label LblCnt 
               BackColor       =   &H00C0FFFF&
               Caption         =   "1"
               BeginProperty Font 
                  Name            =   "新細明體"
                  Size            =   9
                  Charset         =   136
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   255
               Left            =   30
               TabIndex        =   126
               Top             =   570
               Width           =   225
            End
            Begin VB.Label Label1 
               BackColor       =   &H00C0FFFF&
               Caption         =   "備註："
               Height          =   225
               Index           =   4
               Left            =   5520
               TabIndex        =   125
               Top             =   360
               Width           =   555
            End
            Begin VB.Label Label1 
               BackColor       =   &H00C0FFFF&
               Caption         =   "點數："
               Height          =   225
               Index           =   101
               Left            =   4650
               TabIndex        =   124
               Top             =   360
               Width           =   765
            End
            Begin VB.Label Label1 
               BackColor       =   &H00C0FFFF&
               Caption         =   "規費："
               Height          =   225
               Index           =   100
               Left            =   3540
               TabIndex        =   123
               Top             =   360
               Width           =   765
            End
            Begin VB.Label Label1 
               BackColor       =   &H00C0FFFF&
               Caption         =   "費用："
               Height          =   225
               Index           =   99
               Left            =   2400
               TabIndex        =   122
               Top             =   360
               Width           =   765
            End
            Begin VB.Label Label1 
               BackColor       =   &H00C0FFFF&
               Caption         =   "案件性質："
               Height          =   225
               Index           =   10
               Left            =   270
               TabIndex        =   121
               Top             =   360
               Width           =   1005
            End
         End
         Begin VB.CheckBox Check1 
            Caption         =   "請貼印花"
            Height          =   285
            Left            =   6240
            TabIndex        =   41
            Top             =   4770
            Width           =   1065
         End
         Begin VB.Frame Frame14 
            Height          =   285
            Left            =   2820
            TabIndex        =   97
            Top             =   360
            Width           =   2835
            Begin VB.OptionButton OptChoose 
               Caption         =   "外至台"
               Height          =   315
               Index           =   1
               Left            =   1935
               TabIndex        =   4
               Top             =   0
               Width           =   855
            End
            Begin VB.OptionButton OptChoose 
               Caption         =   "國內"
               Height          =   315
               Index           =   0
               Left            =   1200
               TabIndex        =   3
               Top             =   0
               Value           =   -1  'True
               Width           =   705
            End
            Begin VB.Label Label1 
               Caption         =   "接洽單種類："
               Height          =   255
               Index           =   120
               Left            =   0
               TabIndex        =   98
               Top             =   60
               Width           =   1185
            End
         End
         Begin VB.Frame Frame1 
            Height          =   315
            Left            =   3960
            TabIndex        =   95
            Top             =   630
            Width           =   1545
            Begin VB.OptionButton Option1 
               Caption         =   "舊案"
               Height          =   315
               Index           =   1
               Left            =   810
               TabIndex        =   7
               Top             =   30
               Width           =   705
            End
            Begin VB.OptionButton Option1 
               Caption         =   "新案"
               Height          =   315
               Index           =   0
               Left            =   60
               TabIndex        =   6
               Top             =   30
               Value           =   -1  'True
               Width           =   705
            End
         End
         Begin VB.ComboBox Combo1 
            Height          =   260
            Index           =   0
            Left            =   1050
            TabIndex        =   13
            Top             =   1380
            Width           =   2625
         End
         Begin VB.CheckBox Check7 
            Caption         =   "是"
            ForeColor       =   &H00FF0000&
            Height          =   225
            Index           =   0
            Left            =   7680
            TabIndex        =   24
            Top             =   2550
            Width           =   495
         End
         Begin VB.CheckBox Check7 
            Caption         =   "否"
            ForeColor       =   &H00FF0000&
            Height          =   225
            Index           =   1
            Left            =   8220
            TabIndex        =   25
            Top             =   2550
            Width           =   495
         End
         Begin VB.CheckBox Check2 
            Caption         =   "收據暫不列印"
            Height          =   285
            Left            =   6240
            TabIndex        =   42
            Top             =   5010
            Width           =   1395
         End
         Begin VB.ComboBox Combo4 
            Height          =   260
            ItemData        =   "frm090801_Q.frx":00A8
            Left            =   5070
            List            =   "frm090801_Q.frx":00B5
            Style           =   2  '單純下拉式
            TabIndex        =   40
            Top             =   5100
            Width           =   1125
         End
         Begin VB.TextBox txtPrintType 
            Height          =   270
            Left            =   3240
            TabIndex        =   93
            Text            =   "2"
            Top             =   1110
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.ComboBox Combo5 
            Height          =   260
            ItemData        =   "frm090801_Q.frx":00CF
            Left            =   7310
            List            =   "frm090801_Q.frx":00DC
            TabIndex        =   15
            Top             =   1380
            Visible         =   0   'False
            Width           =   1365
         End
         Begin VB.CheckBox ChkPCT 
            Caption         =   "是否PCT案"
            Height          =   180
            Left            =   4050
            TabIndex        =   12
            Top             =   960
            Width           =   1155
         End
         Begin VB.ComboBox Combo6 
            Height          =   260
            ItemData        =   "frm090801_Q.frx":0100
            Left            =   4920
            List            =   "frm090801_Q.frx":0102
            TabIndex        =   14
            Top             =   1380
            Visible         =   0   'False
            Width           =   1365
         End
         Begin VB.Frame Frame26 
            Height          =   1065
            Left            =   7035
            TabIndex        =   91
            Top             =   3180
            Width           =   1635
            Begin VB.CheckBox Check8 
               Caption         =   "送件日"
               Height          =   255
               Index           =   0
               Left            =   60
               TabIndex        =   27
               Top             =   210
               Width           =   1395
            End
            Begin VB.CheckBox Check8 
               Caption         =   "代理人請款日"
               Height          =   225
               Index           =   1
               Left            =   60
               TabIndex        =   28
               Top             =   450
               Width           =   1395
            End
            Begin VB.CheckBox Check8 
               Caption         =   "代理人請款之匯款日"
               Height          =   375
               Index           =   2
               Left            =   60
               TabIndex        =   29
               Top             =   660
               Width           =   1395
            End
            Begin VB.Label Label1 
               Caption         =   "收據自動列印時間點"
               Height          =   255
               Index           =   103
               Left            =   60
               TabIndex        =   92
               Top             =   30
               Width           =   1635
            End
         End
         Begin VB.Frame Frame28 
            Height          =   285
            Left            =   30
            TabIndex        =   90
            Top             =   1680
            Width           =   4425
            Begin VB.OptionButton OptEntity 
               Caption         =   "法人"
               Enabled         =   0   'False
               Height          =   315
               Index           =   0
               Left            =   30
               TabIndex        =   18
               TabStop         =   0   'False
               Top             =   0
               Width           =   1020
            End
            Begin VB.OptionButton OptEntity 
               Caption         =   "法人小個體"
               Enabled         =   0   'False
               Height          =   315
               Index           =   1
               Left            =   1080
               TabIndex        =   19
               TabStop         =   0   'False
               Top             =   0
               Width           =   1350
            End
            Begin VB.OptionButton OptEntity 
               Caption         =   "個人/新創公司"
               Enabled         =   0   'False
               Height          =   315
               Index           =   2
               Left            =   2445
               TabIndex        =   20
               TabStop         =   0   'False
               Top             =   0
               Width           =   1455
            End
         End
         Begin VB.CheckBox Check4 
            Caption         =   "電子送件"
            Height          =   180
            Left            =   7725
            TabIndex        =   22
            Top             =   1740
            Visible         =   0   'False
            Width           =   1065
         End
         Begin VB.CheckBox Check5 
            Caption         =   "附英文摘要"
            Height          =   180
            Left            =   6390
            TabIndex        =   21
            Top             =   1740
            Visible         =   0   'False
            Width           =   1245
         End
         Begin VB.CheckBox Check9 
            Caption         =   "特殊收據"
            ForeColor       =   &H000000C0&
            Height          =   255
            Left            =   6240
            TabIndex        =   43
            Top             =   5280
            Width           =   1155
         End
         Begin VB.Frame Frame2 
            Height          =   450
            Left            =   990
            TabIndex        =   96
            Top             =   4335
            Width           =   4755
            Begin VB.OptionButton Option2 
               Caption         =   "以 DEBIT NOTE 請款"
               Height          =   255
               Index           =   2
               Left            =   1140
               TabIndex        =   37
               Top             =   150
               Width           =   1935
            End
            Begin VB.OptionButton Option2 
               Caption         =   "其他(請寫全銜)"
               Height          =   255
               Index           =   1
               Left            =   3180
               TabIndex        =   38
               Top             =   150
               Width           =   1545
            End
            Begin VB.OptionButton Option2 
               Caption         =   "同申請人"
               Height          =   255
               Index           =   0
               Left            =   60
               TabIndex        =   36
               Top             =   150
               Width           =   1215
            End
         End
         Begin VB.Frame Frame605 
            Caption         =   "Frame605"
            Height          =   315
            Left            =   4080
            TabIndex        =   587
            Top             =   1650
            Width           =   2265
            Begin MSForms.TextBox Text1 
               Height          =   300
               Index           =   143
               Left            =   1830
               TabIndex        =   590
               Top             =   0
               Width           =   375
               VariousPropertyBits=   671107097
               Size            =   "661;529"
               FontName        =   "新細明體-ExtB"
               FontEffects     =   1073750016
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.TextBox Text1 
               Height          =   300
               Index           =   142
               Left            =   1290
               TabIndex        =   589
               Top             =   0
               Width           =   375
               VariousPropertyBits=   671107097
               Size            =   "661;529"
               FontName        =   "新細明體-ExtB"
               FontEffects     =   1073750016
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin VB.Label labelYF 
               AutoSize        =   -1  'True
               Caption         =   "繳費年度：         ~"
               ForeColor       =   &H000000FF&
               Height          =   180
               Left            =   390
               TabIndex        =   588
               Top             =   60
               Width           =   1395
            End
         End
         Begin VB.Label Label1 
            Caption         =   "分割成  案"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   12
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   260
            Index           =   123
            Left            =   3060
            TabIndex        =   663
            Top             =   30
            Width           =   2420
         End
         Begin VB.Label Label38 
            Caption         =   "(有文件)"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   12
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   240
            Left            =   7650
            TabIndex        =   640
            Top             =   690
            Width           =   1035
         End
         Begin MSForms.TextBox Text1 
            Height          =   300
            Index           =   145
            Left            =   4572
            TabIndex        =   631
            Top             =   4776
            Width           =   348
            VariousPropertyBits=   671107099
            Size            =   "609;529"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin VB.Label Label28 
            Caption         =   "(1.電子 2.紙本)"
            Height          =   312
            Index           =   1
            Left            =   4968
            TabIndex        =   630
            Top             =   4836
            Width           =   1212
         End
         Begin VB.Label Label1 
            Caption         =   "證書形式："
            Height          =   288
            Index           =   141
            Left            =   3648
            TabIndex        =   629
            Top             =   4836
            Width           =   948
         End
         Begin MSForms.TextBox Text1 
            Height          =   300
            Index           =   11
            Left            =   1050
            TabIndex        =   23
            Top             =   1950
            Width           =   7545
            VariousPropertyBits=   671107099
            Size            =   "13309;529"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "對造為本所客戶："
            ForeColor       =   &H00FF0000&
            Height          =   180
            Index           =   86
            Left            =   7110
            TabIndex        =   103
            Top             =   2295
            Width           =   1440
         End
         Begin MSForms.ComboBox cboTitle 
            Height          =   300
            Left            =   5760
            TabIndex        =   39
            Top             =   4440
            Width           =   2745
            VariousPropertyBits=   679495707
            DisplayStyle    =   3
            Size            =   "4842;529"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox Text1 
            Height          =   300
            Index           =   6
            Left            =   5520
            TabIndex        =   8
            Top             =   690
            Width           =   465
            VariousPropertyBits=   675301403
            Size            =   "820;529"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox Text1 
            Height          =   300
            Index           =   0
            Left            =   1050
            TabIndex        =   0
            Top             =   0
            Width           =   615
            VariousPropertyBits=   671107097
            Size            =   "1085;529"
            FontName        =   "新細明體-ExtB"
            FontEffects     =   1073750016
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox Text1 
            Height          =   300
            Index           =   1
            Left            =   1050
            TabIndex        =   1
            Top             =   330
            Width           =   1125
            VariousPropertyBits=   671107099
            Size            =   "1984;529"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox Text1 
            Height          =   300
            Index           =   3
            Left            =   1056
            TabIndex        =   2
            Top             =   696
            Width           =   1128
            VariousPropertyBits=   671107099
            Size            =   "1984;529"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox Text1 
            Height          =   300
            Index           =   7
            Left            =   6060
            TabIndex        =   9
            Top             =   690
            Width           =   765
            VariousPropertyBits=   675301401
            Size            =   "1349;529"
            FontName        =   "新細明體-ExtB"
            FontEffects     =   1073750016
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox Text1 
            Height          =   300
            Index           =   8
            Left            =   6900
            TabIndex        =   10
            Top             =   690
            Width           =   225
            VariousPropertyBits=   675301401
            Size            =   "397;529"
            FontName        =   "新細明體-ExtB"
            FontEffects     =   1073750016
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox Text1 
            Height          =   300
            Index           =   9
            Left            =   7230
            TabIndex        =   11
            Top             =   690
            Width           =   345
            VariousPropertyBits=   675301401
            Size            =   "609;529"
            FontName        =   "新細明體-ExtB"
            FontEffects     =   1073750016
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox Text1 
            Height          =   300
            Index           =   10
            Left            =   6804
            TabIndex        =   5
            Top             =   336
            Width           =   708
            VariousPropertyBits=   671107099
            Size            =   "1244;529"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox Text1 
            Height          =   300
            Index           =   113
            Left            =   7590
            TabIndex        =   26
            Top             =   2820
            Width           =   1140
            VariousPropertyBits=   671107099
            Size            =   "2011;529"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox Text1 
            Height          =   300
            Index           =   18
            Left            =   7310
            TabIndex        =   16
            Top             =   1080
            Width           =   1430
            VariousPropertyBits=   671107099
            Size            =   "2514;529"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox Text1 
            Height          =   300
            Index           =   115
            Left            =   5805
            TabIndex        =   94
            Top             =   4455
            Visible         =   0   'False
            Width           =   1590
            VariousPropertyBits=   671107099
            Size            =   "2805;529"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin VB.Label Label1 
            Caption         =   "區　　別："
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   118
            Top             =   45
            Width           =   1005
         End
         Begin VB.Label Label1 
            Caption         =   "法定期限："
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   117
            Top             =   390
            Width           =   1005
         End
         Begin VB.Label Label1 
            Caption         =   "本所期限："
            Height          =   252
            Index           =   3
            Left            =   120
            TabIndex        =   116
            Top             =   732
            Width           =   1008
         End
         Begin VB.Label Label1 
            Caption         =   "申  請  國："
            Height          =   255
            Index           =   9
            Left            =   120
            TabIndex        =   115
            Top             =   1410
            Width           =   1005
         End
         Begin VB.Label Label1 
            Caption         =   "主　　題："
            Height          =   255
            Index           =   11
            Left            =   120
            TabIndex        =   114
            Top             =   1980
            Width           =   1005
         End
         Begin VB.Label lblZone 
            BackColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   1710
            TabIndex        =   113
            Top             =   30
            Width           =   1005
         End
         Begin VB.Label Label1 
            Caption         =   "本所案號："
            Height          =   255
            Index           =   6
            Left            =   3000
            TabIndex        =   112
            Top             =   720
            Width           =   1005
         End
         Begin VB.Line Line1 
            X1              =   5820
            X2              =   7530
            Y1              =   840
            Y2              =   840
         End
         Begin VB.Label Label1 
            Caption         =   "填表日期："
            Height          =   255
            Index           =   7
            Left            =   5760
            TabIndex        =   111
            Top             =   30
            Width           =   1005
         End
         Begin VB.Label lblDate 
            BackColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   6810
            TabIndex        =   110
            Top             =   0
            Width           =   1725
         End
         Begin VB.Label Label1 
            Caption         =   "員工代號："
            Height          =   255
            Index           =   8
            Left            =   5760
            TabIndex        =   109
            Top             =   390
            Width           =   1005
         End
         Begin MSForms.Label lblStaffName 
            Height          =   255
            Left            =   7560
            TabIndex        =   108
            Top             =   360
            Width           =   1155
            BackColor       =   16777215
            VariousPropertyBits=   27
            Size            =   "2037;450"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "後金："
            Height          =   180
            Index           =   102
            Left            =   7065
            TabIndex        =   107
            Top             =   2880
            Width           =   540
         End
         Begin VB.Label Label1 
            Caption         =   "收據抬頭："
            Height          =   285
            Index           =   104
            Left            =   90
            TabIndex        =   106
            Top             =   4500
            Width           =   1005
         End
         Begin VB.Label Label1 
            Caption         =   "客戶案號(型號)："
            Height          =   360
            Index           =   16
            Left            =   6510
            TabIndex        =   105
            Top             =   1020
            Width           =   810
         End
         Begin VB.Label Label2 
            Caption         =   "按此鈕可列出此客戶↑所有收據抬頭"
            ForeColor       =   &H000000FF&
            Height          =   645
            Index           =   2
            Left            =   7650
            TabIndex        =   104
            Top             =   4800
            Width           =   1080
         End
         Begin VB.Label Label1 
            Caption         =   "收據公司："
            Height          =   285
            Index           =   90
            Left            =   4170
            TabIndex        =   102
            Top             =   5160
            Width           =   915
         End
         Begin VB.Label Label1 
            Caption         =   "輸出方式： 　   ( 1.螢幕 2.印表機 )"
            Height          =   180
            Index           =   122
            Left            =   2250
            TabIndex        =   101
            Top             =   1170
            Visible         =   0   'False
            Width           =   2745
         End
         Begin VB.Label Label1 
            Caption         =   "案件屬性:"
            Height          =   180
            Index           =   168
            Left            =   6510
            TabIndex        =   100
            Top             =   1410
            Visible         =   0   'False
            Width           =   765
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "商標種類:"
            Height          =   180
            Index           =   124
            Left            =   4110
            TabIndex        =   99
            Top             =   1410
            Visible         =   0   'False
            Width           =   770
         End
      End
      Begin TabDlg.SSTab SSTab3 
         Height          =   5505
         Left            =   -74910
         TabIndex        =   88
         Top             =   360
         Width           =   8700
         _ExtentX        =   15346
         _ExtentY        =   9701
         _Version        =   393216
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "一般"
         TabPicture(0)   =   "frm090801_Q.frx":0104
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "txtCRL70"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Text1(119)"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "Label16"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "Frame13"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "FrameT"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "FrameP"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "Frame18"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).Control(7)=   "FrameL"
         Tab(0).Control(7).Enabled=   0   'False
         Tab(0).ControlCount=   8
         TabCaption(1)   =   "舉發聲明"
         TabPicture(1)   =   "frm090801_Q.frx":0120
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "txtMonth(0)"
         Tab(1).Control(1)=   "txtYear(0)"
         Tab(1).Control(2)=   "txtDay(0)"
         Tab(1).Control(3)=   "txtYear(1)"
         Tab(1).Control(4)=   "txtMonth(1)"
         Tab(1).Control(5)=   "txtDay(1)"
         Tab(1).Control(6)=   "txtItemCount"
         Tab(1).Control(7)=   "chkItem(0)"
         Tab(1).Control(8)=   "chkItem(1)"
         Tab(1).Control(9)=   "chkItem(4)"
         Tab(1).Control(10)=   "chkItem(3)"
         Tab(1).Control(11)=   "chkItem(5)"
         Tab(1).Control(12)=   "chkItem(2)"
         Tab(1).Control(13)=   "txtItemList"
         Tab(1).Control(14)=   "chkItem(6)"
         Tab(1).Control(15)=   "Label18"
         Tab(1).Control(16)=   "Label15"
         Tab(1).Control(17)=   "Label17"
         Tab(1).ControlCount=   18
         Begin VB.Frame FrameL 
            BackColor       =   &H00FFFFC0&
            Caption         =   "法務"
            Height          =   2295
            Left            =   480
            TabIndex        =   295
            Top             =   3990
            Visible         =   0   'False
            Width           =   8595
            Begin VB.Frame Frame19 
               Height          =   525
               Left            =   60
               TabIndex        =   299
               Top             =   150
               Visible         =   0   'False
               Width           =   8060
               Begin VB.CheckBox Check3 
                  Caption         =   "公平交易法"
                  Height          =   255
                  Index           =   6
                  Left            =   5220
                  TabIndex        =   306
                  Top             =   240
                  Width           =   1300
               End
               Begin VB.CheckBox Check3 
                  Caption         =   "營業秘密法"
                  Height          =   255
                  Index           =   5
                  Left            =   3780
                  TabIndex        =   305
                  Top             =   240
                  Width           =   1300
               End
               Begin VB.CheckBox Check3 
                  Caption         =   "專利"
                  Height          =   255
                  Index           =   0
                  Left            =   3780
                  TabIndex        =   304
                  Top             =   -30
                  Width           =   710
               End
               Begin VB.CheckBox Check3 
                  Caption         =   "商標"
                  Height          =   255
                  Index           =   1
                  Left            =   4500
                  TabIndex        =   303
                  Top             =   -30
                  Width           =   710
               End
               Begin VB.CheckBox Check3 
                  Caption         =   "著作權"
                  Height          =   255
                  Index           =   2
                  Left            =   5220
                  TabIndex        =   302
                  Top             =   -30
                  Width           =   855
               End
               Begin VB.CheckBox Check3 
                  Caption         =   "其他智財權"
                  Height          =   255
                  Index           =   3
                  Left            =   6090
                  TabIndex        =   301
                  Top             =   -30
                  Width           =   1215
               End
               Begin VB.CheckBox Check3 
                  Caption         =   "一般"
                  Height          =   255
                  Index           =   4
                  Left            =   7320
                  TabIndex        =   300
                  Top             =   -30
                  Width           =   675
               End
               Begin MSForms.TextBox Text1 
                  Height          =   450
                  Index           =   127
                  Left            =   930
                  TabIndex        =   308
                  Top             =   0
                  Width           =   2775
                  VariousPropertyBits=   671107103
                  BackColor       =   -2147483633
                  Size            =   "4895;794"
                  FontName        =   "新細明體-ExtB"
                  FontHeight      =   180
                  FontCharSet     =   136
                  FontPitchAndFamily=   34
               End
               Begin VB.Label Label31 
                  Caption         =   "案件屬性："
                  Height          =   180
                  Left            =   0
                  TabIndex        =   307
                  Top             =   45
                  Width           =   945
               End
            End
            Begin VB.Frame Frame41 
               Height          =   705
               Left            =   60
               TabIndex        =   296
               Top             =   630
               Visible         =   0   'False
               Width           =   3525
               Begin VB.OptionButton Option9 
                  Caption         =   "提供書面分析"
                  Height          =   195
                  Index           =   0
                  Left            =   150
                  TabIndex        =   298
                  Top             =   180
                  Width           =   1755
               End
               Begin VB.OptionButton Option9 
                  Caption         =   "請律師向當事人說明"
                  Height          =   195
                  Index           =   1
                  Left            =   150
                  TabIndex        =   297
                  Top             =   420
                  Width           =   3255
               End
            End
            Begin MSForms.TextBox Text1 
               Height          =   300
               Index           =   139
               Left            =   1140
               TabIndex        =   309
               Top             =   1380
               Width           =   1125
               VariousPropertyBits=   671107099
               Size            =   "1984;529"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.TextBox Text1 
               Height          =   300
               Index           =   140
               Left            =   2760
               TabIndex        =   310
               Top             =   1380
               Width           =   1125
               VariousPropertyBits=   671107099
               Size            =   "1984;529"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin VB.Label Label1 
               Caption         =   "聘任期間：   　                        ～"
               Height          =   255
               Index           =   140
               Left            =   120
               TabIndex        =   311
               Top             =   1440
               Width           =   3795
            End
         End
         Begin VB.Frame Frame18 
            Height          =   315
            Left            =   120
            TabIndex        =   651
            Top             =   1980
            Visible         =   0   'False
            Width           =   8120
            Begin VB.Frame Frame3 
               Height          =   340
               Left            =   5160
               TabIndex        =   664
               Top             =   -60
               Width           =   2290
               Begin VB.OptionButton OptCP164 
                  Caption         =   "之後"
                  Height          =   195
                  Index           =   2
                  Left            =   1410
                  TabIndex        =   667
                  Top             =   95
                  Width           =   705
               End
               Begin VB.OptionButton OptCP164 
                  Caption         =   "當天"
                  Height          =   195
                  Index           =   0
                  Left            =   30
                  TabIndex        =   666
                  Top             =   95
                  Width           =   705
               End
               Begin VB.OptionButton OptCP164 
                  Caption         =   "之前"
                  Height          =   195
                  Index           =   1
                  Left            =   720
                  TabIndex        =   665
                  Top             =   95
                  Width           =   705
               End
            End
            Begin VB.OptionButton OptSendType 
               Caption         =   "指定日期"
               Height          =   315
               Index           =   2
               Left            =   3060
               TabIndex        =   654
               Top             =   -30
               Width           =   1040
            End
            Begin VB.OptionButton OptSendType 
               Caption         =   "收款後"
               Height          =   315
               Index           =   1
               Left            =   2130
               TabIndex        =   653
               Top             =   -30
               Width           =   1220
            End
            Begin VB.OptionButton OptSendType 
               Caption         =   "立即"
               Height          =   315
               Index           =   0
               Left            =   900
               TabIndex        =   652
               Top             =   -30
               Width           =   1220
            End
            Begin VB.Label Label1 
               Caption         =   "送件方式："
               Height          =   255
               Index           =   121
               Left            =   30
               TabIndex        =   656
               Top             =   30
               Width           =   915
            End
            Begin MSForms.TextBox Text1 
               Height          =   300
               Index           =   126
               Left            =   4110
               TabIndex        =   655
               Top             =   0
               Width           =   1040
               VariousPropertyBits=   671107099
               Size            =   "1826;529"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
         End
         Begin VB.Frame FrameP 
            BackColor       =   &H00C0FFC0&
            Caption         =   "專利"
            Height          =   2295
            Left            =   930
            TabIndex        =   276
            Top             =   2820
            Visible         =   0   'False
            Width           =   8595
            Begin VB.Frame Frame16 
               Height          =   705
               Left            =   90
               TabIndex        =   288
               Top             =   660
               Visible         =   0   'False
               Width           =   3525
               Begin VB.OptionButton Option3 
                  Caption         =   "切勿超頁超項，本案無法向客戶收款"
                  Height          =   195
                  Index           =   1
                  Left            =   150
                  TabIndex        =   290
                  Top             =   420
                  Width           =   3255
               End
               Begin VB.OptionButton Option3 
                  Caption         =   "可以收超頁超項費"
                  Height          =   195
                  Index           =   0
                  Left            =   150
                  TabIndex        =   289
                  Top             =   180
                  Width           =   1755
               End
            End
            Begin VB.Frame Frame25 
               Height          =   465
               Left            =   3990
               TabIndex        =   285
               Top             =   750
               Width           =   2190
               Begin MSForms.TextBox Text1 
                  Height          =   300
                  Index           =   128
                  Left            =   1080
                  TabIndex        =   287
                  Top             =   60
                  Width           =   1005
                  VariousPropertyBits=   671107099
                  Size            =   "3016;529"
                  FontName        =   "新細明體-ExtB"
                  FontHeight      =   180
                  FontCharSet     =   136
                  FontPitchAndFamily=   34
               End
               Begin VB.Label Label14 
                  Caption         =   "優惠期事實發生日期："
                  Height          =   375
                  Left            =   45
                  TabIndex        =   286
                  Top             =   60
                  Width           =   915
               End
            End
            Begin VB.Frame Frame48 
               Height          =   315
               Left            =   90
               TabIndex        =   281
               Top             =   390
               Width           =   2445
               Begin VB.OptionButton OptNewDrug 
                  Caption         =   "是"
                  Height          =   180
                  Index           =   1
                  Left            =   1470
                  TabIndex        =   283
                  TabStop         =   0   'False
                  Top             =   90
                  Width           =   510
               End
               Begin VB.OptionButton OptNewDrug 
                  Caption         =   "否"
                  Height          =   180
                  Index           =   0
                  Left            =   2010
                  TabIndex        =   282
                  TabStop         =   0   'False
                  Top             =   90
                  Width           =   555
               End
               Begin VB.Label Label27 
                  AutoSize        =   -1  'True
                  Caption         =   "是否為新藥專利："
                  Height          =   180
                  Left            =   60
                  TabIndex        =   284
                  Top             =   90
                  Width           =   1440
               End
            End
            Begin VB.Frame Frame45 
               Height          =   345
               Left            =   90
               TabIndex        =   277
               Top             =   60
               Visible         =   0   'False
               Width           =   2295
               Begin VB.OptionButton Opt45 
                  Caption         =   "否"
                  Height          =   195
                  Index           =   1
                  Left            =   1800
                  TabIndex        =   279
                  Top             =   120
                  Width           =   585
               End
               Begin VB.OptionButton Opt45 
                  Caption         =   "是"
                  Height          =   195
                  Index           =   0
                  Left            =   1260
                  TabIndex        =   278
                  Top             =   120
                  Width           =   555
               End
               Begin VB.Label Label26 
                  AutoSize        =   -1  'True
                  Caption         =   "是否會簡體版："
                  Height          =   180
                  Left            =   30
                  TabIndex        =   280
                  Top             =   120
                  Width           =   1260
               End
            End
            Begin MSForms.TextBox Text1 
               Height          =   300
               Index           =   4
               Left            =   5505
               TabIndex        =   291
               Top             =   120
               Width           =   375
               VariousPropertyBits=   671107099
               Size            =   "661;529"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.TextBox Text1 
               Height          =   300
               Index           =   114
               Left            =   5685
               TabIndex        =   294
               Top             =   420
               Width           =   375
               VariousPropertyBits=   671107099
               Size            =   "661;529"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "申請技術報告項數："
               Height          =   180
               Index           =   125
               Left            =   4005
               TabIndex        =   293
               Top             =   510
               Width           =   1620
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "申請優先權證明書          份"
               Height          =   180
               Index           =   88
               Left            =   4005
               TabIndex        =   292
               Top             =   210
               Width           =   2190
            End
         End
         Begin VB.Frame FrameT 
            BackColor       =   &H00C0E0FF&
            Caption         =   "商標"
            Height          =   2295
            Left            =   60
            TabIndex        =   156
            Top             =   2280
            Visible         =   0   'False
            Width           =   8595
            Begin VB.Frame Frame27 
               Height          =   765
               Left            =   60
               TabIndex        =   272
               Top             =   1440
               Visible         =   0   'False
               Width           =   3435
               Begin VB.OptionButton Option8 
                  Caption         =   "有近似"
                  Height          =   195
                  Index           =   0
                  Left            =   2310
                  TabIndex        =   274
                  Top             =   120
                  Width           =   915
               End
               Begin VB.OptionButton Option8 
                  Caption         =   "無近似"
                  Height          =   195
                  Index           =   1
                  Left            =   2310
                  TabIndex        =   273
                  Top             =   360
                  Width           =   915
               End
               Begin VB.Label Label19 
                  Caption         =   "是否在該國有無近似案件："
                  Height          =   225
                  Left            =   60
                  TabIndex        =   275
                  Top             =   120
                  Width           =   2175
               End
            End
            Begin VB.Frame Frame21 
               Height          =   1005
               Left            =   60
               TabIndex        =   170
               Top             =   420
               Visible         =   0   'False
               Width           =   3135
               Begin VB.Frame Frame22 
                  BorderStyle     =   0  '沒有框線
                  Height          =   285
                  Left            =   60
                  TabIndex        =   183
                  Top             =   150
                  Width           =   2235
                  Begin VB.OptionButton OptEP06 
                     Caption         =   "否"
                     Height          =   255
                     Index           =   1
                     Left            =   1740
                     TabIndex        =   185
                     Top             =   0
                     Width           =   435
                  End
                  Begin VB.OptionButton OptEP06 
                     Caption         =   "是"
                     Height          =   255
                     Index           =   0
                     Left            =   1290
                     TabIndex        =   184
                     Top             =   0
                     Width           =   435
                  End
                  Begin VB.Label Label23 
                     AutoSize        =   -1  'True
                     Caption         =   "資料是否齊備："
                     Height          =   210
                     Left            =   0
                     TabIndex        =   186
                     Top             =   0
                     Width           =   1260
                  End
               End
               Begin VB.Frame Frame23 
                  BorderStyle     =   0  '沒有框線
                  Height          =   285
                  Left            =   60
                  TabIndex        =   179
                  Top             =   390
                  Width           =   2235
                  Begin VB.OptionButton OptEP34 
                     Caption         =   "否"
                     Height          =   255
                     Index           =   1
                     Left            =   1740
                     TabIndex        =   181
                     Top             =   30
                     Width           =   435
                  End
                  Begin VB.OptionButton OptEP34 
                     Caption         =   "是"
                     Height          =   255
                     Index           =   0
                     Left            =   1290
                     TabIndex        =   180
                     Top             =   30
                     Width           =   435
                  End
                  Begin VB.Label Label12 
                     Alignment       =   1  '靠右對齊
                     AutoSize        =   -1  'True
                     Caption         =   "是否會稿："
                     Height          =   180
                     Left            =   0
                     TabIndex        =   182
                     Top             =   30
                     Width           =   1260
                  End
               End
               Begin VB.Frame Frame42 
                  BackColor       =   &H00FFFFC0&
                  BorderStyle     =   0  '沒有框線
                  Height          =   285
                  Left            =   870
                  TabIndex        =   175
                  Top             =   390
                  Width           =   2235
                  Begin VB.OptionButton OptCP143 
                     Caption         =   "是"
                     Height          =   255
                     Index           =   0
                     Left            =   1290
                     TabIndex        =   177
                     Top             =   0
                     Width           =   435
                  End
                  Begin VB.OptionButton OptCP143 
                     Caption         =   "否"
                     Height          =   255
                     Index           =   1
                     Left            =   1740
                     TabIndex        =   176
                     Top             =   0
                     Width           =   435
                  End
                  Begin VB.Label Label24 
                     AutoSize        =   -1  'True
                     Caption         =   "查名是否齊備："
                     Height          =   180
                     Left            =   0
                     TabIndex        =   178
                     Top             =   30
                     Width           =   1260
                  End
               End
               Begin VB.Frame Frame43 
                  BackColor       =   &H80000004&
                  BorderStyle     =   0  '沒有框線
                  Height          =   285
                  Left            =   60
                  TabIndex        =   171
                  Top             =   690
                  Width           =   2235
                  Begin VB.OptionButton OptCRL133 
                     Caption         =   "否"
                     Height          =   255
                     Index           =   1
                     Left            =   1755
                     TabIndex        =   173
                     Top             =   0
                     Value           =   -1  'True
                     Width           =   435
                  End
                  Begin VB.OptionButton OptCRL133 
                     Caption         =   "是"
                     Height          =   255
                     Index           =   0
                     Left            =   1290
                     TabIndex        =   172
                     Top             =   0
                     Width           =   435
                  End
                  Begin VB.Label Label21 
                     AutoSize        =   -1  'True
                     Caption         =   "可否延期："
                     Height          =   180
                     Left            =   360
                     TabIndex        =   174
                     Top             =   30
                     Width           =   900
                  End
               End
            End
            Begin VB.Frame Frame58 
               Caption         =   "T,TF,CFT"
               Height          =   375
               Left            =   30
               TabIndex        =   167
               Top             =   60
               Visible         =   0   'False
               Width           =   5205
               Begin MSForms.TextBox Text1 
                  Height          =   300
                  Index           =   138
                  Left            =   990
                  TabIndex        =   169
                  Top             =   30
                  Width           =   2900
                  VariousPropertyBits=   671107099
                  Size            =   "5115;529"
                  FontName        =   "新細明體-ExtB"
                  FontHeight      =   180
                  FontCharSet     =   136
                  FontPitchAndFamily=   34
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  Caption         =   "商品類別：                                                                  (以逗號分隔 )"
                  ForeColor       =   &H000000C0&
                  Height          =   180
                  Index           =   138
                  Left            =   60
                  TabIndex        =   168
                  Top             =   60
                  Width           =   4935
               End
            End
            Begin VB.Frame Frame20 
               Caption         =   "加註"
               Height          =   1785
               Left            =   3570
               TabIndex        =   157
               Top             =   450
               Visible         =   0   'False
               Width           =   4845
               Begin VB.OptionButton Option6 
                  Caption         =   "相同文字／圖形，申請人已註冊"
                  Height          =   195
                  Index           =   5
                  Left            =   1020
                  TabIndex        =   163
                  Top             =   1080
                  Width           =   2925
               End
               Begin VB.OptionButton Option6 
                  Caption         =   "不查名，請逕提出申請"
                  Height          =   195
                  Index           =   4
                  Left            =   1020
                  TabIndex        =   162
                  Top             =   1560
                  Width           =   2925
               End
               Begin VB.OptionButton Option6 
                  Caption         =   "尚待查名"
                  Height          =   195
                  Index           =   3
                  Left            =   1020
                  TabIndex        =   161
                  Top             =   852
                  Width           =   2925
               End
               Begin VB.OptionButton Option6 
                  Caption         =   "已查名，無近似"
                  Height          =   195
                  Index           =   0
                  Left            =   1020
                  TabIndex        =   160
                  Top             =   150
                  Width           =   3255
               End
               Begin VB.OptionButton Option6 
                  Caption         =   "已查名，有近似，建議客戶申請"
                  Height          =   195
                  Index           =   2
                  Left            =   1020
                  TabIndex        =   159
                  Top             =   618
                  Width           =   2925
               End
               Begin VB.OptionButton Option6 
                  Caption         =   "已查名，有近似，客戶自願嘗試"
                  Height          =   195
                  Index           =   1
                  Left            =   1020
                  TabIndex        =   158
                  Top             =   384
                  Width           =   3255
               End
               Begin MSForms.TextBox Text1 
                  Height          =   300
                  Index           =   133
                  Left            =   1740
                  TabIndex        =   164
                  Top             =   1260
                  Width           =   2325
                  VariousPropertyBits=   671107099
                  Size            =   "4101;529"
                  FontName        =   "新細明體-ExtB"
                  FontHeight      =   180
                  FontCharSet     =   136
                  FontPitchAndFamily=   34
               End
               Begin VB.Label Label22 
                  Caption         =   "有第                                                        號"
                  Height          =   180
                  Left            =   1275
                  TabIndex        =   166
                  Top             =   1320
                  Width           =   3255
               End
               Begin VB.Label Label11 
                  Caption         =   "查名狀態："
                  Height          =   180
                  Left            =   90
                  TabIndex        =   165
                  Top             =   210
                  Width           =   945
               End
            End
         End
         Begin VB.Frame Frame13 
            Caption         =   "代理人"
            Height          =   795
            Left            =   60
            TabIndex        =   148
            Top             =   4650
            Width           =   8565
            Begin VB.CommandButton cmdSerach 
               Caption         =   "搜尋(&1)"
               Height          =   255
               Index           =   5
               Left            =   2640
               TabIndex        =   149
               Top             =   210
               Width           =   945
            End
            Begin VB.Label Label1 
               Caption         =   "彼所案號："
               ForeColor       =   &H000000C0&
               Height          =   225
               Index           =   137
               Left            =   3900
               TabIndex        =   155
               Top             =   240
               Width           =   915
            End
            Begin MSForms.TextBox Text1 
               Height          =   300
               Index           =   137
               Left            =   4860
               TabIndex        =   154
               Top             =   150
               Width           =   3495
               VariousPropertyBits=   671107099
               Size            =   "6165;529"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.TextBox Text1 
               Height          =   300
               Index           =   130
               Left            =   1410
               TabIndex        =   153
               Top             =   480
               Width           =   6945
               VariousPropertyBits=   671107099
               Size            =   "12250;529"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.TextBox Text1 
               Height          =   300
               Index           =   5
               Left            =   1410
               TabIndex        =   152
               Top             =   150
               Width           =   1095
               VariousPropertyBits=   671107099
               Size            =   "1931;529"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin VB.Label Label1 
               Caption         =   "代理人名稱："
               Height          =   255
               Index           =   119
               Left            =   180
               TabIndex        =   151
               Top             =   510
               Width           =   1305
            End
            Begin VB.Label Label1 
               Caption         =   "代理人編號："
               Height          =   255
               Index           =   118
               Left            =   180
               TabIndex        =   150
               Top             =   240
               Width           =   1215
            End
         End
         Begin VB.TextBox txtMonth 
            Alignment       =   2  '置中對齊
            Enabled         =   0   'False
            Height          =   285
            Index           =   0
            Left            =   -72660
            TabIndex        =   134
            Top             =   1620
            Width           =   285
         End
         Begin VB.TextBox txtYear 
            Alignment       =   2  '置中對齊
            Enabled         =   0   'False
            Height          =   285
            Index           =   0
            Left            =   -73380
            TabIndex        =   133
            Top             =   1620
            Width           =   420
         End
         Begin VB.TextBox txtDay 
            Alignment       =   2  '置中對齊
            Enabled         =   0   'False
            Height          =   285
            Index           =   0
            Left            =   -72120
            TabIndex        =   129
            Top             =   1620
            Width           =   285
         End
         Begin VB.TextBox txtYear 
            Alignment       =   2  '置中對齊
            Enabled         =   0   'False
            Height          =   285
            Index           =   1
            Left            =   -70995
            TabIndex        =   132
            Top             =   1620
            Width           =   420
         End
         Begin VB.TextBox txtMonth 
            Alignment       =   2  '置中對齊
            Enabled         =   0   'False
            Height          =   285
            Index           =   1
            Left            =   -70320
            TabIndex        =   131
            Top             =   1620
            Width           =   285
         End
         Begin VB.TextBox txtDay 
            Alignment       =   2  '置中對齊
            Enabled         =   0   'False
            Height          =   285
            Index           =   1
            Left            =   -69735
            TabIndex        =   130
            Top             =   1620
            Width           =   285
         End
         Begin VB.TextBox txtItemCount 
            Enabled         =   0   'False
            Height          =   270
            Left            =   -72345
            TabIndex        =   128
            Top             =   660
            Width           =   375
         End
         Begin VB.CheckBox chkItem 
            Caption         =   "請求撤銷全部請求項：共計　　　項"
            Height          =   210
            Index           =   0
            Left            =   -74820
            TabIndex        =   142
            Top             =   690
            Width           =   3225
         End
         Begin VB.CheckBox chkItem 
            Caption         =   "請求撤銷部分之請求項："
            Height          =   210
            Index           =   1
            Left            =   -74820
            TabIndex        =   141
            Top             =   900
            Width           =   2400
         End
         Begin VB.CheckBox chkItem 
            Caption         =   "共有專利申請權非由全體共有人提出申請者"
            Height          =   210
            Index           =   4
            Left            =   -71580
            TabIndex        =   140
            Top             =   1110
            Width           =   4335
         End
         Begin VB.CheckBox chkItem 
            Caption         =   "專利權人為非專利申請權人者"
            Height          =   210
            Index           =   3
            Left            =   -71580
            TabIndex        =   139
            Top             =   900
            Width           =   4335
         End
         Begin VB.CheckBox chkItem 
            Caption         =   "專利權人所屬國家對中華民國申請專利不予受理者"
            Height          =   210
            Index           =   5
            Left            =   -71580
            TabIndex        =   138
            Top             =   1320
            Width           =   4335
         End
         Begin VB.CheckBox chkItem 
            Caption         =   "請求撤銷設計專利權"
            Enabled         =   0   'False
            Height          =   210
            Index           =   2
            Left            =   -71580
            TabIndex        =   137
            Top             =   690
            Width           =   4335
         End
         Begin VB.TextBox txtItemList 
            Enabled         =   0   'False
            Height          =   300
            Left            =   -74550
            TabIndex        =   136
            Text            =   "第項"
            Top             =   1110
            Width           =   2715
         End
         Begin VB.CheckBox chkItem 
            Caption         =   "請求撤銷自「        年      月      日」至「        年      月      日」之專利權期間延長"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   6
            Left            =   -74820
            TabIndex        =   135
            Top             =   1650
            Width           =   7440
         End
         Begin VB.Label Label16 
            Caption         =   "訊息區："
            Height          =   180
            Left            =   4980
            TabIndex        =   591
            Top             =   390
            Width           =   945
         End
         Begin MSForms.TextBox Text1 
            Height          =   1575
            Index           =   119
            Left            =   60
            TabIndex        =   147
            Top             =   360
            Width           =   4875
            VariousPropertyBits=   -1466941413
            ScrollBars      =   2
            Size            =   "8599;2778"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox txtCRL70 
            Height          =   1335
            Left            =   4950
            TabIndex        =   146
            Top             =   600
            Width           =   3675
            VariousPropertyBits=   -1466941409
            BackColor       =   -2147483633
            BorderStyle     =   1
            ScrollBars      =   2
            Size            =   "6482;2355"
            SpecialEffect   =   0
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "請求撤銷發明(新型)專利權"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Left            =   -74820
            TabIndex        =   145
            Top             =   450
            Width           =   2295
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "請求撤銷全部專利權"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Left            =   -71580
            TabIndex        =   144
            Top             =   450
            Width           =   1755
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "( 例如：第 1,3,5-12 項 )"
            Height          =   180
            Left            =   -74550
            TabIndex        =   143
            Top             =   1410
            Width           =   1800
         End
      End
      Begin TabDlg.SSTab SSTab2 
         Height          =   5505
         Left            =   -74850
         TabIndex        =   49
         Top             =   375
         Width           =   8625
         _ExtentX        =   15222
         _ExtentY        =   9701
         _Version        =   393216
         Tabs            =   5
         TabsPerRow      =   5
         TabHeight       =   520
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "申請人1"
         TabPicture(0)   =   "frm090801_Q.frx":013C
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Frame36"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "FrameCase"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).ControlCount=   2
         TabCaption(1)   =   "申請人2"
         TabPicture(1)   =   "frm090801_Q.frx":0158
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Frame37"
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "申請人3"
         TabPicture(2)   =   "frm090801_Q.frx":0174
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "Frame38"
         Tab(2).ControlCount=   1
         TabCaption(3)   =   "申請人4"
         TabPicture(3)   =   "frm090801_Q.frx":0190
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "Frame39"
         Tab(3).ControlCount=   1
         TabCaption(4)   =   "申請人5"
         TabPicture(4)   =   "frm090801_Q.frx":01AC
         Tab(4).ControlEnabled=   0   'False
         Tab(4).Control(0)=   "Frame40"
         Tab(4).ControlCount=   1
         Begin VB.Frame FrameCase 
            BackColor       =   &H00FFC0C0&
            Caption         =   "編輯 本案與總號"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   1485
            Left            =   1110
            TabIndex        =   612
            Top             =   3720
            Visible         =   0   'False
            Width           =   4275
            Begin VB.CommandButton cmdRemove 
               Caption         =   "移除"
               Height          =   315
               Left            =   2490
               TabIndex        =   619
               Top             =   960
               Width           =   705
            End
            Begin VB.CommandButton cmdAdd 
               Caption         =   "加入"
               Height          =   315
               Left            =   2490
               TabIndex        =   618
               Top             =   600
               Width           =   705
            End
            Begin VB.ListBox lstCase 
               Height          =   220
               ItemData        =   "frm090801_Q.frx":01C8
               Left            =   480
               List            =   "frm090801_Q.frx":01CF
               TabIndex        =   621
               Top             =   600
               Width           =   1905
            End
            Begin VB.TextBox txtSystem 
               Height          =   300
               Left            =   1080
               TabIndex        =   613
               Top             =   210
               Width           =   525
            End
            Begin VB.TextBox txtCode 
               Height          =   300
               Index           =   2
               Left            =   3090
               TabIndex        =   616
               Top             =   210
               Width           =   492
            End
            Begin VB.TextBox txtCode 
               Height          =   300
               Index           =   1
               Left            =   2670
               TabIndex        =   615
               Top             =   210
               Width           =   372
            End
            Begin VB.TextBox txtCode 
               Height          =   300
               Index           =   0
               Left            =   1650
               TabIndex        =   614
               Top             =   210
               Width           =   975
            End
            Begin VB.CommandButton cmdCRL55_OK 
               Caption         =   "確定"
               Height          =   315
               Left            =   3360
               TabIndex        =   620
               Top             =   600
               Width           =   825
            End
            Begin VB.Label Label2 
               BackColor       =   &H00FFC0C0&
               Caption         =   "本所案號："
               Height          =   255
               Index           =   1
               Left            =   150
               TabIndex        =   617
               Top             =   240
               Width           =   975
            End
         End
         Begin VB.Frame Frame40 
            Caption         =   "Frame40"
            Height          =   4575
            Left            =   -74940
            TabIndex        =   532
            Top             =   330
            Width           =   8505
            Begin VB.Frame Frame54 
               Height          =   735
               Left            =   -30
               TabIndex        =   540
               Top             =   1650
               Width           =   8535
               Begin MSForms.ComboBox cboContact 
                  Height          =   300
                  Index           =   5
                  Left            =   2940
                  TabIndex        =   555
                  Top             =   90
                  Width           =   1230
                  VariousPropertyBits=   679495707
                  DisplayStyle    =   3
                  Size            =   "2170;529"
                  MatchEntry      =   1
                  ShowDropButtonWhen=   2
                  FontName        =   "新細明體-ExtB"
                  FontHeight      =   180
                  FontCharSet     =   136
                  FontPitchAndFamily=   34
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  Caption         =   "手機："
                  Height          =   180
                  Index           =   64
                  Left            =   6630
                  TabIndex        =   554
                  Top             =   390
                  Width           =   540
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  Caption         =   "傳真："
                  Height          =   180
                  Index           =   67
                  Left            =   6630
                  TabIndex        =   553
                  Top             =   120
                  Width           =   540
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  Caption         =   "電話："
                  Height          =   180
                  Index           =   68
                  Left            =   4200
                  TabIndex        =   552
                  Top             =   120
                  Width           =   540
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  Caption         =   "接洽人："
                  Height          =   180
                  Index           =   69
                  Left            =   2190
                  TabIndex        =   551
                  Top             =   120
                  Width           =   720
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  Caption         =   "客戶編號："
                  Height          =   180
                  Index           =   70
                  Left            =   90
                  TabIndex        =   550
                  Top             =   120
                  Width           =   900
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  Caption         =   "ID No.："
                  Height          =   180
                  Index           =   81
                  Left            =   90
                  TabIndex        =   549
                  Top             =   420
                  Width           =   660
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  Caption         =   "LINE ID："
                  BeginProperty Font 
                     Name            =   "新細明體"
                     Size            =   8.5
                     Charset         =   136
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   165
                  Index           =   93
                  Left            =   4200
                  TabIndex        =   548
                  Top             =   420
                  Width           =   690
               End
               Begin MSForms.TextBox Text1 
                  Height          =   300
                  Index           =   84
                  Left            =   7170
                  TabIndex        =   547
                  Top             =   390
                  Width           =   1335
                  VariousPropertyBits=   671107099
                  Size            =   "2355;529"
                  FontName        =   "新細明體-ExtB"
                  FontHeight      =   180
                  FontCharSet     =   136
                  FontPitchAndFamily=   34
               End
               Begin MSForms.TextBox Text1 
                  Height          =   300
                  Index           =   76
                  Left            =   1020
                  TabIndex        =   546
                  Top             =   90
                  Width           =   1095
                  VariousPropertyBits=   671107099
                  Size            =   "1931;529"
                  FontName        =   "新細明體-ExtB"
                  FontHeight      =   180
                  FontCharSet     =   136
                  FontPitchAndFamily=   34
               End
               Begin MSForms.TextBox Text1 
                  Height          =   300
                  Index           =   78
                  Left            =   4935
                  TabIndex        =   545
                  Top             =   90
                  Width           =   1650
                  VariousPropertyBits=   671107099
                  Size            =   "2910;529"
                  FontName        =   "新細明體-ExtB"
                  FontHeight      =   180
                  FontCharSet     =   136
                  FontPitchAndFamily=   34
               End
               Begin MSForms.TextBox Text1 
                  Height          =   300
                  Index           =   79
                  Left            =   4935
                  TabIndex        =   544
                  Top             =   390
                  Width           =   1650
                  VariousPropertyBits=   671107099
                  Size            =   "2910;529"
                  FontName        =   "新細明體-ExtB"
                  FontHeight      =   180
                  FontCharSet     =   136
                  FontPitchAndFamily=   34
               End
               Begin MSForms.TextBox Text1 
                  Height          =   300
                  Index           =   80
                  Left            =   7170
                  TabIndex        =   543
                  Top             =   90
                  Width           =   1335
                  VariousPropertyBits=   671107099
                  Size            =   "2355;529"
                  FontName        =   "新細明體-ExtB"
                  FontHeight      =   180
                  FontCharSet     =   136
                  FontPitchAndFamily=   34
               End
               Begin MSForms.TextBox Text1 
                  Height          =   300
                  Index           =   77
                  Left            =   2940
                  TabIndex        =   542
                  Top             =   390
                  Visible         =   0   'False
                  Width           =   1230
                  VariousPropertyBits=   671107099
                  Size            =   "2170;529"
                  FontName        =   "新細明體-ExtB"
                  FontHeight      =   180
                  FontCharSet     =   136
                  FontPitchAndFamily=   34
               End
               Begin MSForms.TextBox Text1 
                  Height          =   300
                  Index           =   96
                  Left            =   1020
                  TabIndex        =   541
                  Top             =   390
                  Width           =   1755
                  VariousPropertyBits=   671107099
                  Size            =   "3096;529"
                  FontName        =   "新細明體-ExtB"
                  FontHeight      =   180
                  FontCharSet     =   136
                  FontPitchAndFamily=   34
               End
            End
            Begin VB.Frame Frame4 
               Height          =   405
               Index           =   4
               Left            =   1365
               TabIndex        =   562
               Top             =   2295
               Width           =   1695
               Begin VB.OptionButton optCP815 
                  Caption         =   "是"
                  Enabled         =   0   'False
                  Height          =   255
                  Index           =   0
                  Left            =   120
                  TabIndex        =   564
                  Top             =   120
                  Width           =   615
               End
               Begin VB.OptionButton optCP815 
                  Caption         =   "否"
                  Enabled         =   0   'False
                  Height          =   255
                  Index           =   1
                  Left            =   930
                  TabIndex        =   563
                  Top             =   120
                  Width           =   660
               End
            End
            Begin VB.Frame Frame10 
               Height          =   1740
               Left            =   30
               TabIndex        =   556
               Top             =   -60
               Width           =   8430
               Begin VB.Frame Frame30 
                  Caption         =   "特例簽核"
                  ForeColor       =   &H00000000&
                  Height          =   495
                  Index           =   4
                  Left            =   2130
                  TabIndex        =   605
                  Top             =   60
                  Width           =   2085
                  Begin VB.CheckBox ChkCRA26 
                     Caption         =   "有對造"
                     Enabled         =   0   'False
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
                     Height          =   270
                     Index           =   4
                     Left            =   90
                     TabIndex        =   607
                     Top             =   180
                     Width           =   885
                  End
                  Begin VB.CheckBox ChkCRA27 
                     Caption         =   "有跨所"
                     Enabled         =   0   'False
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
                     Height          =   270
                     Index           =   4
                     Left            =   1080
                     TabIndex        =   606
                     Top             =   180
                     Width           =   975
                  End
               End
               Begin VB.OptionButton Option35 
                  Caption         =   "新客戶"
                  Height          =   285
                  Index           =   0
                  Left            =   1020
                  TabIndex        =   558
                  Top             =   240
                  Width           =   1755
               End
               Begin VB.OptionButton Option35 
                  Caption         =   "舊客戶"
                  Height          =   285
                  Index           =   1
                  Left            =   5580
                  TabIndex        =   557
                  Top             =   225
                  Value           =   -1  'True
                  Width           =   975
               End
               Begin MSForms.TextBox Text1 
                  Height          =   300
                  Index           =   598
                  Left            =   4665
                  TabIndex        =   559
                  Top             =   645
                  Width           =   2505
                  VariousPropertyBits=   671107099
                  Size            =   "4419;529"
                  FontName        =   "新細明體-ExtB"
                  FontHeight      =   180
                  FontCharSet     =   136
                  FontPitchAndFamily=   34
               End
               Begin VB.Shape Shape5 
                  Height          =   1515
                  Left            =   60
                  Top             =   180
                  Width           =   8325
               End
               Begin VB.Label Label1 
                  Caption         =   "與                                                            為關係企業"
                  Height          =   260
                  Index           =   109
                  Left            =   4280
                  TabIndex        =   561
                  Top             =   680
                  Width           =   4130
               End
               Begin VB.Line Line18 
                  X1              =   4245
                  X2              =   4245
                  Y1              =   570
                  Y2              =   1680
               End
               Begin VB.Line Line19 
                  X1              =   495
                  X2              =   8355
                  Y1              =   555
                  Y2              =   555
               End
               Begin VB.Line Line20 
                  X1              =   495
                  X2              =   495
                  Y1              =   195
                  Y2              =   1680
               End
               Begin VB.Label Label2 
                  Alignment       =   2  '置中對齊
                  Caption         =   "案件來源說明"
                  Height          =   1110
                  Index           =   6
                  Left            =   135
                  TabIndex        =   560
                  Top             =   465
                  Width           =   315
               End
               Begin VB.Line Line21 
                  X1              =   510
                  X2              =   8370
                  Y1              =   975
                  Y2              =   975
               End
            End
            Begin VB.CommandButton cmdTW 
               Caption         =   "申請地址："
               BeginProperty Font 
                  Name            =   "新細明體"
                  Size            =   8.5
                  Charset         =   136
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Index           =   9
               Left            =   30
               TabIndex        =   539
               TabStop         =   0   'False
               Top             =   3915
               Width           =   980
            End
            Begin VB.CommandButton cmdTW 
               Caption         =   "聯絡地址："
               BeginProperty Font 
                  Name            =   "新細明體"
                  Size            =   8.5
                  Charset         =   136
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Index           =   8
               Left            =   30
               TabIndex        =   538
               TabStop         =   0   'False
               Top             =   3615
               Width           =   980
            End
            Begin VB.CommandButton cmdSearchZip 
               Caption         =   "Zip"
               BeginProperty Font 
                  Name            =   "新細明體"
                  Size            =   8.5
                  Charset         =   136
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   270
               Index           =   9
               Left            =   7995
               TabIndex        =   537
               Top             =   3915
               Width           =   480
            End
            Begin VB.CommandButton cmdSearchZip 
               Caption         =   "Zip"
               BeginProperty Font 
                  Name            =   "新細明體"
                  Size            =   8.5
                  Charset         =   136
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   270
               Index           =   8
               Left            =   7995
               TabIndex        =   536
               Top             =   3615
               Width           =   480
            End
            Begin VB.CommandButton cmdSerach 
               Caption         =   "搜尋(&5)"
               Height          =   585
               Index           =   4
               Left            =   4515
               TabIndex        =   535
               Top             =   3015
               Width           =   345
            End
            Begin VB.CommandButton cmdQual 
               Caption         =   "中小企業減免資格"
               Height          =   285
               Index           =   5
               Left            =   1350
               TabIndex        =   534
               Top             =   2715
               Visible         =   0   'False
               Width           =   1725
            End
            Begin VB.CommandButton CmdSame 
               Caption         =   "同上"
               BeginProperty Font 
                  Name            =   "新細明體"
                  Size            =   8.5
                  Charset         =   136
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Index           =   5
               Left            =   1005
               TabIndex        =   533
               Top             =   3915
               Visible         =   0   'False
               Width           =   550
            End
            Begin MSForms.TextBox Text1 
               Height          =   300
               Index           =   189
               Left            =   1005
               TabIndex        =   569
               Top             =   4215
               Width           =   7305
               VariousPropertyBits=   671107099
               Size            =   "12885;529"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.TextBox Text1 
               Height          =   300
               Index           =   91
               Left            =   1005
               TabIndex        =   566
               Top             =   3915
               Width           =   6285
               VariousPropertyBits=   679495707
               Size            =   "11086;529"
               FontName        =   "新細明體"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.TextBox Text1 
               Height          =   300
               Index           =   90
               Left            =   1005
               TabIndex        =   565
               Top             =   3615
               Width           =   6285
               VariousPropertyBits=   671107099
               Size            =   "11086;529"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.TextBox Text1 
               Height          =   300
               Index           =   86
               Left            =   1365
               TabIndex        =   568
               Top             =   3315
               Width           =   3075
               VariousPropertyBits=   671107099
               Size            =   "5424;529"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.TextBox Text1 
               Height          =   300
               Index           =   85
               Left            =   1365
               TabIndex        =   567
               Top             =   3015
               Width           =   3075
               VariousPropertyBits=   671107099
               Size            =   "5424;529"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.TextBox Text1 
               Height          =   300
               Index           =   83
               Left            =   3945
               TabIndex        =   571
               Top             =   2385
               Width           =   4425
               VariousPropertyBits=   671107099
               Size            =   "7805;529"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.TextBox Text1 
               Height          =   300
               Index           =   98
               Left            =   3945
               TabIndex        =   570
               Top             =   2685
               Width           =   4425
               VariousPropertyBits=   671107099
               Size            =   "7805;529"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.TextBox Text1 
               Height          =   285
               Index           =   81
               Left            =   6990
               TabIndex        =   586
               Top             =   2460
               Visible         =   0   'False
               Width           =   1335
               VariousPropertyBits=   671107097
               BackColor       =   -2147483648
               Size            =   "2355;503"
               FontName        =   "新細明體-ExtB"
               FontEffects     =   1073750016
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin VB.Label Label1 
               Caption         =   "聯絡地址："
               Height          =   255
               Index           =   95
               Left            =   105
               TabIndex        =   585
               Top             =   3630
               Width           =   945
            End
            Begin VB.Label Label1 
               Caption         =   "申請地址："
               Height          =   255
               Index           =   76
               Left            =   105
               TabIndex        =   584
               Top             =   3930
               Width           =   945
            End
            Begin MSForms.TextBox Text1 
               Height          =   300
               Index           =   89
               Left            =   7275
               TabIndex        =   583
               Top             =   3615
               Width           =   705
               VariousPropertyBits=   671107099
               Size            =   "1244;529"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.TextBox Text1 
               Height          =   300
               Index           =   88
               Left            =   6465
               TabIndex        =   582
               Top             =   3315
               Width           =   1905
               VariousPropertyBits=   671107099
               Size            =   "3360;529"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.TextBox Text1 
               Height          =   300
               Index           =   87
               Left            =   6465
               TabIndex        =   581
               Top             =   3015
               Width           =   1905
               VariousPropertyBits=   671107099
               Size            =   "3360;529"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.TextBox Text1 
               Height          =   300
               Index           =   124
               Left            =   7275
               TabIndex        =   580
               Top             =   3915
               Width           =   705
               VariousPropertyBits=   671107099
               Size            =   "1244;529"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin VB.Label Label1 
               Caption         =   "代表人(中文)"
               Height          =   255
               Index           =   73
               Left            =   4965
               TabIndex        =   579
               Top             =   3045
               Width           =   1455
            End
            Begin VB.Label Label1 
               Caption         =   "　　　(英文)："
               Height          =   255
               Index           =   72
               Left            =   105
               TabIndex        =   578
               Top             =   3375
               Width           =   1305
            End
            Begin VB.Label Label1 
               Caption         =   "申請人(中文)："
               Height          =   255
               Index           =   71
               Left            =   105
               TabIndex        =   577
               Top             =   3045
               Width           =   1305
            End
            Begin VB.Label Label1 
               Caption         =   "E-Mail："
               Height          =   255
               Index           =   65
               Left            =   3255
               TabIndex        =   576
               Top             =   2415
               Width           =   765
            End
            Begin VB.Label Label1 
               Caption         =   "　　　(英文)："
               Height          =   255
               Index           =   74
               Left            =   4965
               TabIndex        =   575
               Top             =   3330
               Width           =   1455
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "符合年費減免："
               Height          =   180
               Index           =   66
               Left            =   105
               TabIndex        =   574
               Top             =   2445
               Width           =   1260
            End
            Begin VB.Label Label1 
               Caption         =   "國籍："
               Height          =   255
               Index           =   113
               Left            =   3255
               TabIndex        =   573
               Top             =   2715
               Width           =   705
            End
            Begin VB.Label Label1 
               Caption         =   "英文地址："
               Height          =   255
               Index           =   117
               Left            =   105
               TabIndex        =   572
               Top             =   4245
               Width           =   1305
            End
         End
         Begin VB.Frame Frame39 
            Caption         =   "Frame39"
            Height          =   4575
            Left            =   -74940
            TabIndex        =   477
            Top             =   330
            Width           =   8505
            Begin VB.Frame Frame53 
               Height          =   705
               Left            =   0
               TabIndex        =   485
               Top             =   1650
               Width           =   8535
               Begin MSForms.ComboBox cboContact 
                  Height          =   300
                  Index           =   4
                  Left            =   2970
                  TabIndex        =   500
                  Top             =   60
                  Width           =   1230
                  VariousPropertyBits=   679495707
                  DisplayStyle    =   3
                  Size            =   "2170;529"
                  MatchEntry      =   1
                  ShowDropButtonWhen=   2
                  FontName        =   "新細明體-ExtB"
                  FontHeight      =   180
                  FontCharSet     =   136
                  FontPitchAndFamily=   34
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  Caption         =   "手機："
                  Height          =   180
                  Index           =   51
                  Left            =   6630
                  TabIndex        =   499
                  Top             =   360
                  Width           =   540
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  Caption         =   "傳真："
                  Height          =   180
                  Index           =   54
                  Left            =   6630
                  TabIndex        =   498
                  Top             =   90
                  Width           =   540
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  Caption         =   "電話："
                  Height          =   180
                  Index           =   55
                  Left            =   4230
                  TabIndex        =   497
                  Top             =   90
                  Width           =   540
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  Caption         =   "接洽人："
                  Height          =   180
                  Index           =   56
                  Left            =   2220
                  TabIndex        =   496
                  Top             =   90
                  Width           =   720
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  Caption         =   "客戶編號："
                  Height          =   180
                  Index           =   57
                  Left            =   120
                  TabIndex        =   495
                  Top             =   90
                  Width           =   900
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  Caption         =   "ID No.："
                  Height          =   180
                  Index           =   80
                  Left            =   120
                  TabIndex        =   494
                  Top             =   390
                  Width           =   660
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  Caption         =   "LINE ID："
                  BeginProperty Font 
                     Name            =   "新細明體"
                     Size            =   8.5
                     Charset         =   136
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   165
                  Index           =   91
                  Left            =   4230
                  TabIndex        =   493
                  Top             =   390
                  Width           =   690
               End
               Begin MSForms.TextBox Text1 
                  Height          =   300
                  Index           =   68
                  Left            =   7170
                  TabIndex        =   492
                  Top             =   360
                  Width           =   1335
                  VariousPropertyBits=   671107099
                  Size            =   "2355;529"
                  FontName        =   "新細明體-ExtB"
                  FontHeight      =   180
                  FontCharSet     =   136
                  FontPitchAndFamily=   34
               End
               Begin MSForms.TextBox Text1 
                  Height          =   300
                  Index           =   60
                  Left            =   1050
                  TabIndex        =   491
                  Top             =   60
                  Width           =   1095
                  VariousPropertyBits=   671107099
                  Size            =   "1931;529"
                  FontName        =   "新細明體-ExtB"
                  FontHeight      =   180
                  FontCharSet     =   136
                  FontPitchAndFamily=   34
               End
               Begin MSForms.TextBox Text1 
                  Height          =   300
                  Index           =   62
                  Left            =   4935
                  TabIndex        =   490
                  Top             =   60
                  Width           =   1650
                  VariousPropertyBits=   671107099
                  Size            =   "2910;529"
                  FontName        =   "新細明體-ExtB"
                  FontHeight      =   180
                  FontCharSet     =   136
                  FontPitchAndFamily=   34
               End
               Begin MSForms.TextBox Text1 
                  Height          =   300
                  Index           =   63
                  Left            =   4935
                  TabIndex        =   489
                  Top             =   360
                  Width           =   1650
                  VariousPropertyBits=   671107099
                  Size            =   "2910;529"
                  FontName        =   "新細明體-ExtB"
                  FontHeight      =   180
                  FontCharSet     =   136
                  FontPitchAndFamily=   34
               End
               Begin MSForms.TextBox Text1 
                  Height          =   300
                  Index           =   64
                  Left            =   7170
                  TabIndex        =   488
                  Top             =   60
                  Width           =   1335
                  VariousPropertyBits=   671107099
                  Size            =   "2355;529"
                  FontName        =   "新細明體-ExtB"
                  FontHeight      =   180
                  FontCharSet     =   136
                  FontPitchAndFamily=   34
               End
               Begin MSForms.TextBox Text1 
                  Height          =   300
                  Index           =   95
                  Left            =   1050
                  TabIndex        =   487
                  Top             =   360
                  Width           =   1755
                  VariousPropertyBits=   671107099
                  Size            =   "3096;529"
                  FontName        =   "新細明體-ExtB"
                  FontHeight      =   180
                  FontCharSet     =   136
                  FontPitchAndFamily=   34
               End
               Begin MSForms.TextBox Text1 
                  Height          =   300
                  Index           =   61
                  Left            =   2970
                  TabIndex        =   486
                  Top             =   360
                  Visible         =   0   'False
                  Width           =   1230
                  VariousPropertyBits=   671107099
                  Size            =   "2170;529"
                  FontName        =   "新細明體-ExtB"
                  FontHeight      =   180
                  FontCharSet     =   136
                  FontPitchAndFamily=   34
               End
            End
            Begin VB.Frame Frame4 
               Height          =   405
               Index           =   3
               Left            =   1335
               TabIndex        =   507
               Top             =   2295
               Width           =   1695
               Begin VB.OptionButton optCP814 
                  Caption         =   "是"
                  Enabled         =   0   'False
                  Height          =   255
                  Index           =   0
                  Left            =   120
                  TabIndex        =   509
                  Top             =   120
                  Width           =   615
               End
               Begin VB.OptionButton optCP814 
                  Caption         =   "否"
                  Enabled         =   0   'False
                  Height          =   255
                  Index           =   1
                  Left            =   930
                  TabIndex        =   508
                  Top             =   120
                  Width           =   660
               End
            End
            Begin VB.Frame Frame9 
               Height          =   1740
               Left            =   30
               TabIndex        =   501
               Top             =   -60
               Width           =   8430
               Begin VB.Frame Frame30 
                  Caption         =   "特例簽核"
                  ForeColor       =   &H00000000&
                  Height          =   495
                  Index           =   3
                  Left            =   2130
                  TabIndex        =   602
                  Top             =   60
                  Width           =   2085
                  Begin VB.CheckBox ChkCRA27 
                     Caption         =   "有跨所"
                     Enabled         =   0   'False
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
                     Height          =   270
                     Index           =   3
                     Left            =   1080
                     TabIndex        =   604
                     Top             =   180
                     Width           =   975
                  End
                  Begin VB.CheckBox ChkCRA26 
                     Caption         =   "有對造"
                     Enabled         =   0   'False
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
                     Height          =   270
                     Index           =   3
                     Left            =   90
                     TabIndex        =   603
                     Top             =   180
                     Width           =   885
                  End
               End
               Begin VB.OptionButton Option34 
                  Caption         =   "新客戶"
                  Height          =   285
                  Index           =   0
                  Left            =   990
                  TabIndex        =   503
                  Top             =   240
                  Width           =   1785
               End
               Begin VB.OptionButton Option34 
                  Caption         =   "舊客戶"
                  Height          =   285
                  Index           =   1
                  Left            =   5580
                  TabIndex        =   502
                  Top             =   225
                  Value           =   -1  'True
                  Width           =   975
               End
               Begin MSForms.TextBox Text1 
                  Height          =   300
                  Index           =   498
                  Left            =   4665
                  TabIndex        =   504
                  Top             =   645
                  Width           =   2505
                  VariousPropertyBits=   671107099
                  Size            =   "4419;529"
                  FontName        =   "新細明體-ExtB"
                  FontHeight      =   180
                  FontCharSet     =   136
                  FontPitchAndFamily=   34
               End
               Begin VB.Shape Shape4 
                  Height          =   1515
                  Left            =   60
                  Top             =   180
                  Width           =   8325
               End
               Begin VB.Label Label1 
                  Caption         =   "與                                                            為關係企業"
                  Height          =   260
                  Index           =   108
                  Left            =   4280
                  TabIndex        =   506
                  Top             =   680
                  Width           =   4130
               End
               Begin VB.Line Line14 
                  X1              =   4245
                  X2              =   4245
                  Y1              =   570
                  Y2              =   1680
               End
               Begin VB.Line Line15 
                  X1              =   495
                  X2              =   8355
                  Y1              =   555
                  Y2              =   555
               End
               Begin VB.Line Line16 
                  X1              =   495
                  X2              =   495
                  Y1              =   195
                  Y2              =   1680
               End
               Begin VB.Label Label2 
                  Alignment       =   2  '置中對齊
                  Caption         =   "案件來源說明"
                  Height          =   1110
                  Index           =   5
                  Left            =   135
                  TabIndex        =   505
                  Top             =   465
                  Width           =   315
               End
               Begin VB.Line Line17 
                  X1              =   510
                  X2              =   8370
                  Y1              =   975
                  Y2              =   975
               End
            End
            Begin VB.CommandButton cmdTW 
               Caption         =   "申請地址："
               BeginProperty Font 
                  Name            =   "新細明體"
                  Size            =   8.5
                  Charset         =   136
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Index           =   7
               Left            =   0
               TabIndex        =   484
               TabStop         =   0   'False
               Top             =   3915
               Width           =   980
            End
            Begin VB.CommandButton cmdTW 
               Caption         =   "聯絡地址："
               BeginProperty Font 
                  Name            =   "新細明體"
                  Size            =   8.5
                  Charset         =   136
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Index           =   6
               Left            =   0
               TabIndex        =   483
               TabStop         =   0   'False
               Top             =   3615
               Width           =   980
            End
            Begin VB.CommandButton cmdSearchZip 
               Caption         =   "Zip"
               BeginProperty Font 
                  Name            =   "新細明體"
                  Size            =   8.5
                  Charset         =   136
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   270
               Index           =   7
               Left            =   7965
               TabIndex        =   482
               Top             =   3915
               Width           =   480
            End
            Begin VB.CommandButton cmdSearchZip 
               Caption         =   "Zip"
               BeginProperty Font 
                  Name            =   "新細明體"
                  Size            =   8.5
                  Charset         =   136
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   270
               Index           =   6
               Left            =   7965
               TabIndex        =   481
               Top             =   3615
               Width           =   480
            End
            Begin VB.CommandButton cmdSerach 
               Caption         =   "搜尋(&4)"
               Height          =   585
               Index           =   3
               Left            =   4485
               TabIndex        =   480
               Top             =   3015
               Width           =   345
            End
            Begin VB.CommandButton cmdQual 
               Caption         =   "中小企業減免資格"
               Height          =   285
               Index           =   4
               Left            =   1320
               TabIndex        =   479
               Top             =   2715
               Visible         =   0   'False
               Width           =   1725
            End
            Begin VB.CommandButton CmdSame 
               Caption         =   "同上"
               BeginProperty Font 
                  Name            =   "新細明體"
                  Size            =   8.5
                  Charset         =   136
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Index           =   4
               Left            =   975
               TabIndex        =   478
               Top             =   3915
               Visible         =   0   'False
               Width           =   550
            End
            Begin MSForms.TextBox Text1 
               Height          =   300
               Index           =   70
               Left            =   1335
               TabIndex        =   514
               Top             =   3315
               Width           =   3075
               VariousPropertyBits=   671107099
               Size            =   "5424;529"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.TextBox Text1 
               Height          =   300
               Index           =   69
               Left            =   1335
               TabIndex        =   513
               Top             =   3015
               Width           =   3075
               VariousPropertyBits=   671107099
               Size            =   "5424;529"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.TextBox Text1 
               Height          =   300
               Index           =   74
               Left            =   975
               TabIndex        =   512
               Top             =   3615
               Width           =   6285
               VariousPropertyBits=   671107099
               Size            =   "11086;529"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.TextBox Text1 
               Height          =   300
               Index           =   173
               Left            =   975
               TabIndex        =   511
               Top             =   4215
               Width           =   7305
               VariousPropertyBits=   671107099
               Size            =   "12885;529"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.TextBox Text1 
               Height          =   300
               Index           =   75
               Left            =   975
               TabIndex        =   510
               Top             =   3915
               Width           =   6285
               VariousPropertyBits=   679495707
               Size            =   "11086;529"
               FontName        =   "新細明體"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.TextBox Text1 
               Height          =   300
               Index           =   67
               Left            =   3915
               TabIndex        =   516
               Top             =   2385
               Width           =   4425
               VariousPropertyBits=   671107099
               Size            =   "7805;529"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.TextBox Text1 
               Height          =   300
               Index           =   82
               Left            =   3915
               TabIndex        =   515
               Top             =   2685
               Width           =   4425
               VariousPropertyBits=   671107099
               Size            =   "7805;529"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.TextBox Text1 
               Height          =   300
               Index           =   65
               Left            =   6960
               TabIndex        =   531
               Top             =   2400
               Visible         =   0   'False
               Width           =   1335
               VariousPropertyBits=   671107097
               BackColor       =   -2147483648
               Size            =   "2355;529"
               FontName        =   "新細明體-ExtB"
               FontEffects     =   1073750016
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin VB.Label Label1 
               Caption         =   "聯絡地址："
               Height          =   255
               Index           =   131
               Left            =   75
               TabIndex        =   530
               Top             =   3660
               Width           =   945
            End
            Begin VB.Label Label1 
               Caption         =   "申請地址："
               Height          =   255
               Index           =   129
               Left            =   75
               TabIndex        =   529
               Top             =   3960
               Width           =   945
            End
            Begin MSForms.TextBox Text1 
               Height          =   300
               Index           =   73
               Left            =   7245
               TabIndex        =   528
               Top             =   3615
               Width           =   705
               VariousPropertyBits=   671107099
               Size            =   "1244;529"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.TextBox Text1 
               Height          =   300
               Index           =   72
               Left            =   6435
               TabIndex        =   527
               Top             =   3315
               Width           =   1905
               VariousPropertyBits=   671107099
               Size            =   "3360;529"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.TextBox Text1 
               Height          =   300
               Index           =   71
               Left            =   6435
               TabIndex        =   526
               Top             =   3015
               Width           =   1905
               VariousPropertyBits=   671107099
               Size            =   "3360;529"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.TextBox Text1 
               Height          =   300
               Index           =   123
               Left            =   7245
               TabIndex        =   525
               Top             =   3915
               Width           =   705
               VariousPropertyBits=   671107099
               Size            =   "1244;529"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin VB.Label Label1 
               Caption         =   "代表人(中文)"
               Height          =   255
               Index           =   60
               Left            =   4935
               TabIndex        =   524
               Top             =   3045
               Width           =   1455
            End
            Begin VB.Label Label1 
               Caption         =   "　　　(英文)："
               Height          =   255
               Index           =   59
               Left            =   75
               TabIndex        =   523
               Top             =   3375
               Width           =   1305
            End
            Begin VB.Label Label1 
               Caption         =   "申請人(中文)："
               Height          =   255
               Index           =   58
               Left            =   75
               TabIndex        =   522
               Top             =   3045
               Width           =   1305
            End
            Begin VB.Label Label1 
               Caption         =   "E-Mail："
               Height          =   255
               Index           =   52
               Left            =   3225
               TabIndex        =   521
               Top             =   2415
               Width           =   705
            End
            Begin VB.Label Label1 
               Caption         =   "　　　(英文)："
               Height          =   255
               Index           =   61
               Left            =   4935
               TabIndex        =   520
               Top             =   3330
               Width           =   1455
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "符合年費減免："
               Height          =   180
               Index           =   53
               Left            =   75
               TabIndex        =   519
               Top             =   2445
               Width           =   1260
            End
            Begin VB.Label Label1 
               Caption         =   "國籍："
               Height          =   255
               Index           =   112
               Left            =   3225
               TabIndex        =   518
               Top             =   2715
               Width           =   705
            End
            Begin VB.Label Label1 
               Caption         =   "英文地址："
               Height          =   255
               Index           =   116
               Left            =   75
               TabIndex        =   517
               Top             =   4245
               Width           =   1305
            End
         End
         Begin VB.Frame Frame38 
            Caption         =   "Frame38"
            Height          =   4575
            Left            =   -74940
            TabIndex        =   422
            Top             =   330
            Width           =   8505
            Begin VB.Frame Frame52 
               Height          =   735
               Left            =   0
               TabIndex        =   423
               Top             =   1650
               Width           =   8535
               Begin MSForms.ComboBox cboContact 
                  Height          =   300
                  Index           =   3
                  Left            =   2940
                  TabIndex        =   438
                  Top             =   90
                  Width           =   1230
                  VariousPropertyBits=   679495707
                  DisplayStyle    =   3
                  Size            =   "2170;529"
                  MatchEntry      =   1
                  ShowDropButtonWhen=   2
                  FontName        =   "新細明體-ExtB"
                  FontHeight      =   180
                  FontCharSet     =   136
                  FontPitchAndFamily=   34
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  Caption         =   "手機："
                  Height          =   180
                  Index           =   38
                  Left            =   6630
                  TabIndex        =   437
                  Top             =   390
                  Width           =   540
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  Caption         =   "傳真："
                  Height          =   180
                  Index           =   41
                  Left            =   6630
                  TabIndex        =   436
                  Top             =   120
                  Width           =   540
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  Caption         =   "電話："
                  Height          =   180
                  Index           =   42
                  Left            =   4200
                  TabIndex        =   435
                  Top             =   120
                  Width           =   540
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  Caption         =   "接洽人："
                  Height          =   180
                  Index           =   43
                  Left            =   2220
                  TabIndex        =   434
                  Top             =   120
                  Width           =   720
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  Caption         =   "客戶編號："
                  Height          =   180
                  Index           =   44
                  Left            =   90
                  TabIndex        =   433
                  Top             =   120
                  Width           =   900
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  Caption         =   "ID No.："
                  Height          =   180
                  Index           =   79
                  Left            =   90
                  TabIndex        =   432
                  Top             =   420
                  Width           =   660
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  Caption         =   "LINE ID："
                  BeginProperty Font 
                     Name            =   "新細明體"
                     Size            =   8.5
                     Charset         =   136
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   165
                  Index           =   89
                  Left            =   4200
                  TabIndex        =   431
                  Top             =   420
                  Width           =   690
               End
               Begin MSForms.TextBox Text1 
                  Height          =   300
                  Index           =   52
                  Left            =   7170
                  TabIndex        =   430
                  Top             =   390
                  Width           =   1335
                  VariousPropertyBits=   671107099
                  Size            =   "2355;529"
                  FontName        =   "新細明體-ExtB"
                  FontHeight      =   180
                  FontCharSet     =   136
                  FontPitchAndFamily=   34
               End
               Begin MSForms.TextBox Text1 
                  Height          =   300
                  Index           =   44
                  Left            =   1020
                  TabIndex        =   429
                  Top             =   90
                  Width           =   1095
                  VariousPropertyBits=   671107099
                  Size            =   "1931;529"
                  FontName        =   "新細明體-ExtB"
                  FontHeight      =   180
                  FontCharSet     =   136
                  FontPitchAndFamily=   34
               End
               Begin MSForms.TextBox Text1 
                  Height          =   300
                  Index           =   46
                  Left            =   4905
                  TabIndex        =   428
                  Top             =   90
                  Width           =   1650
                  VariousPropertyBits=   671107099
                  Size            =   "2910;529"
                  FontName        =   "新細明體-ExtB"
                  FontHeight      =   180
                  FontCharSet     =   136
                  FontPitchAndFamily=   34
               End
               Begin MSForms.TextBox Text1 
                  Height          =   300
                  Index           =   47
                  Left            =   4905
                  TabIndex        =   427
                  Top             =   390
                  Width           =   1650
                  VariousPropertyBits=   671107099
                  Size            =   "2910;529"
                  FontName        =   "新細明體-ExtB"
                  FontHeight      =   180
                  FontCharSet     =   136
                  FontPitchAndFamily=   34
               End
               Begin MSForms.TextBox Text1 
                  Height          =   300
                  Index           =   48
                  Left            =   7170
                  TabIndex        =   426
                  Top             =   90
                  Width           =   1335
                  VariousPropertyBits=   671107099
                  Size            =   "2355;529"
                  FontName        =   "新細明體-ExtB"
                  FontHeight      =   180
                  FontCharSet     =   136
                  FontPitchAndFamily=   34
               End
               Begin MSForms.TextBox Text1 
                  Height          =   300
                  Index           =   94
                  Left            =   1020
                  TabIndex        =   425
                  Top             =   390
                  Width           =   1755
                  VariousPropertyBits=   671107099
                  Size            =   "3096;529"
                  FontName        =   "新細明體-ExtB"
                  FontHeight      =   180
                  FontCharSet     =   136
                  FontPitchAndFamily=   34
               End
               Begin MSForms.TextBox Text1 
                  Height          =   300
                  Index           =   45
                  Left            =   2940
                  TabIndex        =   424
                  Top             =   390
                  Visible         =   0   'False
                  Width           =   1230
                  VariousPropertyBits=   671107099
                  Size            =   "2170;529"
                  FontName        =   "新細明體-ExtB"
                  FontHeight      =   180
                  FontCharSet     =   136
                  FontPitchAndFamily=   34
               End
            End
            Begin VB.Frame Frame8 
               Height          =   1710
               Left            =   30
               TabIndex        =   449
               Top             =   -30
               Width           =   8430
               Begin VB.Frame Frame30 
                  Caption         =   "特例簽核"
                  ForeColor       =   &H00000000&
                  Height          =   495
                  Index           =   2
                  Left            =   2130
                  TabIndex        =   599
                  Top             =   30
                  Width           =   2085
                  Begin VB.CheckBox ChkCRA27 
                     Caption         =   "有跨所"
                     Enabled         =   0   'False
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
                     Height          =   270
                     Index           =   2
                     Left            =   1080
                     TabIndex        =   601
                     Top             =   180
                     Width           =   975
                  End
                  Begin VB.CheckBox ChkCRA26 
                     Caption         =   "有對造"
                     Enabled         =   0   'False
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
                     Height          =   270
                     Index           =   2
                     Left            =   90
                     TabIndex        =   600
                     Top             =   180
                     Width           =   885
                  End
               End
               Begin VB.OptionButton Option33 
                  Caption         =   "新客戶"
                  Height          =   285
                  Index           =   0
                  Left            =   990
                  TabIndex        =   451
                  Top             =   240
                  Width           =   1785
               End
               Begin VB.OptionButton Option33 
                  Caption         =   "舊客戶"
                  Height          =   285
                  Index           =   1
                  Left            =   5580
                  TabIndex        =   450
                  Top             =   225
                  Value           =   -1  'True
                  Width           =   975
               End
               Begin MSForms.TextBox Text1 
                  Height          =   300
                  Index           =   398
                  Left            =   4665
                  TabIndex        =   452
                  Top             =   615
                  Width           =   2505
                  VariousPropertyBits=   671107099
                  Size            =   "4419;529"
                  FontName        =   "新細明體-ExtB"
                  FontHeight      =   180
                  FontCharSet     =   136
                  FontPitchAndFamily=   34
               End
               Begin VB.Shape Shape3 
                  Height          =   1515
                  Left            =   60
                  Top             =   150
                  Width           =   8325
               End
               Begin VB.Label Label1 
                  Caption         =   "與                                                            為關係企業"
                  Height          =   260
                  Index           =   106
                  Left            =   4280
                  TabIndex        =   454
                  Top             =   650
                  Width           =   4130
               End
               Begin VB.Line Line10 
                  X1              =   4245
                  X2              =   4245
                  Y1              =   540
                  Y2              =   1650
               End
               Begin VB.Line Line11 
                  X1              =   495
                  X2              =   8355
                  Y1              =   525
                  Y2              =   525
               End
               Begin VB.Line Line12 
                  X1              =   495
                  X2              =   495
                  Y1              =   165
                  Y2              =   1650
               End
               Begin VB.Label Label2 
                  Alignment       =   2  '置中對齊
                  Caption         =   "案件來源說明"
                  Height          =   1110
                  Index           =   4
                  Left            =   135
                  TabIndex        =   453
                  Top             =   465
                  Width           =   315
               End
               Begin VB.Line Line13 
                  X1              =   510
                  X2              =   8370
                  Y1              =   945
                  Y2              =   945
               End
            End
            Begin VB.CommandButton cmdTW 
               Caption         =   "申請地址："
               BeginProperty Font 
                  Name            =   "新細明體"
                  Size            =   8.5
                  Charset         =   136
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Index           =   5
               Left            =   30
               TabIndex        =   448
               TabStop         =   0   'False
               Top             =   3915
               Width           =   980
            End
            Begin VB.CommandButton cmdTW 
               Caption         =   "聯絡地址："
               BeginProperty Font 
                  Name            =   "新細明體"
                  Size            =   8.5
                  Charset         =   136
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Index           =   4
               Left            =   30
               TabIndex        =   447
               TabStop         =   0   'False
               Top             =   3615
               Width           =   980
            End
            Begin VB.CommandButton cmdSearchZip 
               Caption         =   "Zip"
               BeginProperty Font 
                  Name            =   "新細明體"
                  Size            =   8.5
                  Charset         =   136
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   270
               Index           =   5
               Left            =   7995
               TabIndex        =   446
               Top             =   3915
               Width           =   480
            End
            Begin VB.CommandButton cmdSearchZip 
               Caption         =   "Zip"
               BeginProperty Font 
                  Name            =   "新細明體"
                  Size            =   8.5
                  Charset         =   136
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   270
               Index           =   4
               Left            =   7995
               TabIndex        =   445
               Top             =   3615
               Width           =   480
            End
            Begin VB.CommandButton cmdSerach 
               Caption         =   "搜尋(&3)"
               Height          =   585
               Index           =   2
               Left            =   4515
               TabIndex        =   444
               Top             =   3015
               Width           =   345
            End
            Begin VB.Frame Frame4 
               Height          =   405
               Index           =   2
               Left            =   1365
               TabIndex        =   441
               Top             =   2295
               Width           =   1695
               Begin VB.OptionButton optCP813 
                  Caption         =   "是"
                  Enabled         =   0   'False
                  Height          =   255
                  Index           =   0
                  Left            =   120
                  TabIndex        =   443
                  Top             =   120
                  Width           =   615
               End
               Begin VB.OptionButton optCP813 
                  Caption         =   "否"
                  Enabled         =   0   'False
                  Height          =   255
                  Index           =   1
                  Left            =   930
                  TabIndex        =   442
                  Top             =   120
                  Width           =   660
               End
            End
            Begin VB.CommandButton cmdQual 
               Caption         =   "中小企業減免資格"
               Height          =   285
               Index           =   3
               Left            =   1350
               TabIndex        =   440
               Top             =   2715
               Visible         =   0   'False
               Width           =   1725
            End
            Begin VB.CommandButton CmdSame 
               Caption         =   "同上"
               BeginProperty Font 
                  Name            =   "新細明體"
                  Size            =   8.5
                  Charset         =   136
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Index           =   3
               Left            =   1005
               TabIndex        =   439
               Top             =   3915
               Visible         =   0   'False
               Width           =   550
            End
            Begin MSForms.TextBox Text1 
               Height          =   300
               Index           =   157
               Left            =   1005
               TabIndex        =   456
               Top             =   4215
               Width           =   7305
               VariousPropertyBits=   671107099
               Size            =   "12885;529"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin VB.Label Label1 
               Caption         =   "英文地址："
               Height          =   255
               Index           =   115
               Left            =   105
               TabIndex        =   457
               Top             =   4245
               Width           =   1305
            End
            Begin MSForms.TextBox Text1 
               Height          =   300
               Index           =   54
               Left            =   1365
               TabIndex        =   462
               Top             =   3315
               Width           =   3075
               VariousPropertyBits=   671107099
               Size            =   "5424;529"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.TextBox Text1 
               Height          =   300
               Index           =   53
               Left            =   1365
               TabIndex        =   461
               Top             =   3015
               Width           =   3075
               VariousPropertyBits=   671107099
               Size            =   "5424;529"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.TextBox Text1 
               Height          =   300
               Index           =   58
               Left            =   1005
               TabIndex        =   460
               Top             =   3615
               Width           =   6285
               VariousPropertyBits=   671107099
               Size            =   "11086;529"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.TextBox Text1 
               Height          =   300
               Index           =   59
               Left            =   1005
               TabIndex        =   455
               Top             =   3915
               Width           =   6285
               VariousPropertyBits=   679495707
               Size            =   "11086;529"
               FontName        =   "新細明體"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.TextBox Text1 
               Height          =   300
               Index           =   51
               Left            =   3915
               TabIndex        =   459
               Top             =   2385
               Width           =   4455
               VariousPropertyBits=   671107099
               Size            =   "7858;529"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.TextBox Text1 
               Height          =   300
               Index           =   66
               Left            =   3915
               TabIndex        =   458
               Top             =   2685
               Width           =   4455
               VariousPropertyBits=   671107099
               Size            =   "7858;529"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.TextBox Text1 
               Height          =   300
               Index           =   49
               Left            =   6960
               TabIndex        =   476
               Top             =   2430
               Visible         =   0   'False
               Width           =   1335
               VariousPropertyBits=   671107097
               BackColor       =   -2147483648
               Size            =   "2355;529"
               FontName        =   "新細明體-ExtB"
               FontEffects     =   1073750016
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin VB.Label Label1 
               Caption         =   "聯絡地址："
               Height          =   255
               Index           =   133
               Left            =   105
               TabIndex        =   475
               Top             =   3720
               Width           =   945
            End
            Begin VB.Label Label1 
               Caption         =   "申請地址："
               Height          =   255
               Index           =   132
               Left            =   105
               TabIndex        =   474
               Top             =   4020
               Width           =   945
            End
            Begin MSForms.TextBox Text1 
               Height          =   300
               Index           =   57
               Left            =   7275
               TabIndex        =   473
               Top             =   3615
               Width           =   705
               VariousPropertyBits=   671107099
               Size            =   "1244;529"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.TextBox Text1 
               Height          =   300
               Index           =   56
               Left            =   6465
               TabIndex        =   472
               Top             =   3315
               Width           =   1905
               VariousPropertyBits=   671107099
               Size            =   "3360;529"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.TextBox Text1 
               Height          =   300
               Index           =   55
               Left            =   6465
               TabIndex        =   471
               Top             =   3015
               Width           =   1905
               VariousPropertyBits=   671107099
               Size            =   "3360;529"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.TextBox Text1 
               Height          =   300
               Index           =   122
               Left            =   7275
               TabIndex        =   470
               Top             =   3915
               Width           =   705
               VariousPropertyBits=   671107099
               Size            =   "1244;529"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin VB.Label Label1 
               Caption         =   "代表人(中文)"
               Height          =   255
               Index           =   47
               Left            =   4965
               TabIndex        =   469
               Top             =   3045
               Width           =   1455
            End
            Begin VB.Label Label1 
               Caption         =   "　　　(英文)："
               Height          =   255
               Index           =   46
               Left            =   105
               TabIndex        =   468
               Top             =   3375
               Width           =   1305
            End
            Begin VB.Label Label1 
               Caption         =   "申請人(中文)："
               Height          =   255
               Index           =   45
               Left            =   105
               TabIndex        =   467
               Top             =   3045
               Width           =   1305
            End
            Begin VB.Label Label1 
               Caption         =   "E-Mail："
               Height          =   255
               Index           =   39
               Left            =   3255
               TabIndex        =   466
               Top             =   2415
               Width           =   765
            End
            Begin VB.Label Label1 
               Caption         =   "　　　(英文)："
               Height          =   255
               Index           =   48
               Left            =   4965
               TabIndex        =   465
               Top             =   3330
               Width           =   1455
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "符合年費減免："
               Height          =   180
               Index           =   40
               Left            =   105
               TabIndex        =   464
               Top             =   2445
               Width           =   1260
            End
            Begin VB.Label Label1 
               Caption         =   "國籍："
               Height          =   255
               Index           =   111
               Left            =   3255
               TabIndex        =   463
               Top             =   2715
               Width           =   705
            End
         End
         Begin VB.Frame Frame37 
            Caption         =   "Frame37"
            Height          =   4575
            Left            =   -74940
            TabIndex        =   367
            Top             =   330
            Width           =   8505
            Begin VB.Frame Frame51 
               Height          =   705
               Left            =   30
               TabIndex        =   369
               Top             =   1680
               Width           =   8445
               Begin MSForms.ComboBox cboContact 
                  Height          =   300
                  Index           =   2
                  Left            =   2910
                  TabIndex        =   384
                  Top             =   60
                  Width           =   1230
                  VariousPropertyBits=   679495707
                  DisplayStyle    =   3
                  Size            =   "2170;529"
                  MatchEntry      =   1
                  ShowDropButtonWhen=   2
                  FontName        =   "新細明體-ExtB"
                  FontHeight      =   180
                  FontCharSet     =   136
                  FontPitchAndFamily=   34
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  Caption         =   "手機："
                  Height          =   180
                  Index           =   25
                  Left            =   6540
                  TabIndex        =   383
                  Top             =   360
                  Width           =   540
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  Caption         =   "傳真："
                  Height          =   180
                  Index           =   28
                  Left            =   6540
                  TabIndex        =   382
                  Top             =   90
                  Width           =   540
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  Caption         =   "電話："
                  Height          =   180
                  Index           =   29
                  Left            =   4170
                  TabIndex        =   381
                  Top             =   90
                  Width           =   540
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  Caption         =   "接洽人："
                  Height          =   180
                  Index           =   30
                  Left            =   2190
                  TabIndex        =   380
                  Top             =   90
                  Width           =   720
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  Caption         =   "客戶編號："
                  Height          =   180
                  Index           =   31
                  Left            =   60
                  TabIndex        =   379
                  Top             =   90
                  Width           =   900
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  Caption         =   "ID No.："
                  Height          =   180
                  Index           =   78
                  Left            =   60
                  TabIndex        =   378
                  Top             =   390
                  Width           =   660
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  Caption         =   "LINE ID："
                  BeginProperty Font 
                     Name            =   "新細明體"
                     Size            =   8.5
                     Charset         =   136
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   165
                  Index           =   87
                  Left            =   4170
                  TabIndex        =   377
                  Top             =   390
                  Width           =   690
               End
               Begin MSForms.TextBox Text1 
                  Height          =   300
                  Index           =   36
                  Left            =   7080
                  TabIndex        =   376
                  Top             =   360
                  Width           =   1335
                  VariousPropertyBits=   671107099
                  Size            =   "2355;529"
                  FontName        =   "新細明體-ExtB"
                  FontHeight      =   180
                  FontCharSet     =   136
                  FontPitchAndFamily=   34
               End
               Begin MSForms.TextBox Text1 
                  Height          =   300
                  Index           =   28
                  Left            =   990
                  TabIndex        =   375
                  Top             =   60
                  Width           =   1095
                  VariousPropertyBits=   671107099
                  Size            =   "1931;529"
                  FontName        =   "新細明體-ExtB"
                  FontHeight      =   180
                  FontCharSet     =   136
                  FontPitchAndFamily=   34
               End
               Begin MSForms.TextBox Text1 
                  Height          =   300
                  Index           =   30
                  Left            =   4845
                  TabIndex        =   374
                  Top             =   60
                  Width           =   1650
                  VariousPropertyBits=   671107099
                  Size            =   "2910;529"
                  FontName        =   "新細明體-ExtB"
                  FontHeight      =   180
                  FontCharSet     =   136
                  FontPitchAndFamily=   34
               End
               Begin MSForms.TextBox Text1 
                  Height          =   300
                  Index           =   31
                  Left            =   4845
                  TabIndex        =   373
                  Top             =   360
                  Width           =   1650
                  VariousPropertyBits=   671107099
                  Size            =   "2910;529"
                  FontName        =   "新細明體-ExtB"
                  FontHeight      =   180
                  FontCharSet     =   136
                  FontPitchAndFamily=   34
               End
               Begin MSForms.TextBox Text1 
                  Height          =   300
                  Index           =   32
                  Left            =   7080
                  TabIndex        =   372
                  Top             =   60
                  Width           =   1335
                  VariousPropertyBits=   671107099
                  Size            =   "2355;529"
                  FontName        =   "新細明體-ExtB"
                  FontHeight      =   180
                  FontCharSet     =   136
                  FontPitchAndFamily=   34
               End
               Begin MSForms.TextBox Text1 
                  Height          =   300
                  Index           =   93
                  Left            =   990
                  TabIndex        =   371
                  Top             =   360
                  Width           =   1755
                  VariousPropertyBits=   671107099
                  Size            =   "3096;529"
                  FontName        =   "新細明體-ExtB"
                  FontHeight      =   180
                  FontCharSet     =   136
                  FontPitchAndFamily=   34
               End
               Begin MSForms.TextBox Text1 
                  Height          =   300
                  Index           =   29
                  Left            =   2910
                  TabIndex        =   370
                  Top             =   360
                  Visible         =   0   'False
                  Width           =   1230
                  VariousPropertyBits=   671107099
                  Size            =   "2170;529"
                  FontName        =   "新細明體-ExtB"
                  FontHeight      =   180
                  FontCharSet     =   136
                  FontPitchAndFamily=   34
               End
            End
            Begin VB.Frame Frame6 
               Height          =   1740
               Left            =   30
               TabIndex        =   394
               Top             =   -30
               Width           =   8430
               Begin VB.Frame Frame30 
                  Caption         =   "特例簽核"
                  ForeColor       =   &H00000000&
                  Height          =   495
                  Index           =   1
                  Left            =   2130
                  TabIndex        =   596
                  Top             =   60
                  Width           =   2085
                  Begin VB.CheckBox ChkCRA26 
                     Caption         =   "有對造"
                     Enabled         =   0   'False
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
                     Height          =   270
                     Index           =   1
                     Left            =   90
                     TabIndex        =   598
                     Top             =   180
                     Width           =   885
                  End
                  Begin VB.CheckBox ChkCRA27 
                     Caption         =   "有跨所"
                     Enabled         =   0   'False
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
                     Height          =   270
                     Index           =   1
                     Left            =   1080
                     TabIndex        =   597
                     Top             =   180
                     Width           =   975
                  End
               End
               Begin VB.OptionButton Option32 
                  Caption         =   "舊客戶"
                  Height          =   285
                  Index           =   1
                  Left            =   5580
                  TabIndex        =   396
                  Top             =   225
                  Value           =   -1  'True
                  Width           =   975
               End
               Begin VB.OptionButton Option32 
                  Caption         =   "新客戶"
                  Height          =   285
                  Index           =   0
                  Left            =   990
                  TabIndex        =   395
                  Top             =   240
                  Width           =   1785
               End
               Begin MSForms.TextBox Text1 
                  Height          =   300
                  Index           =   298
                  Left            =   4665
                  TabIndex        =   397
                  Top             =   645
                  Width           =   2505
                  VariousPropertyBits=   671107099
                  Size            =   "4419;529"
                  FontName        =   "新細明體-ExtB"
                  FontHeight      =   180
                  FontCharSet     =   136
                  FontPitchAndFamily=   34
               End
               Begin VB.Line Line5 
                  X1              =   510
                  X2              =   8370
                  Y1              =   975
                  Y2              =   975
               End
               Begin VB.Label Label2 
                  Alignment       =   2  '置中對齊
                  Caption         =   "案件來源說明"
                  Height          =   1110
                  Index           =   3
                  Left            =   135
                  TabIndex        =   399
                  Top             =   465
                  Width           =   315
               End
               Begin VB.Line Line6 
                  X1              =   495
                  X2              =   495
                  Y1              =   195
                  Y2              =   1680
               End
               Begin VB.Line Line8 
                  X1              =   495
                  X2              =   8355
                  Y1              =   555
                  Y2              =   555
               End
               Begin VB.Line Line9 
                  X1              =   4245
                  X2              =   4245
                  Y1              =   570
                  Y2              =   1680
               End
               Begin VB.Label Label1 
                  Caption         =   "與                                                            為關係企業"
                  Height          =   260
                  Index           =   107
                  Left            =   4250
                  TabIndex        =   398
                  Top             =   680
                  Width           =   4130
               End
               Begin VB.Shape Shape2 
                  Height          =   1515
                  Left            =   60
                  Top             =   180
                  Width           =   8325
               End
            End
            Begin VB.CommandButton cmdTW 
               Caption         =   "聯絡地址："
               BeginProperty Font 
                  Name            =   "新細明體"
                  Size            =   8.5
                  Charset         =   136
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Index           =   2
               Left            =   60
               TabIndex        =   393
               TabStop         =   0   'False
               Top             =   3615
               Width           =   980
            End
            Begin VB.CommandButton cmdTW 
               Caption         =   "申請地址："
               BeginProperty Font 
                  Name            =   "新細明體"
                  Size            =   8.5
                  Charset         =   136
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Index           =   3
               Left            =   60
               TabIndex        =   392
               TabStop         =   0   'False
               Top             =   3915
               Width           =   980
            End
            Begin VB.CommandButton cmdSearchZip 
               Caption         =   "Zip"
               BeginProperty Font 
                  Name            =   "新細明體"
                  Size            =   8.5
                  Charset         =   136
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   270
               Index           =   3
               Left            =   7995
               TabIndex        =   391
               Top             =   3915
               Width           =   480
            End
            Begin VB.CommandButton cmdSearchZip 
               Caption         =   "Zip"
               BeginProperty Font 
                  Name            =   "新細明體"
                  Size            =   8.5
                  Charset         =   136
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   270
               Index           =   2
               Left            =   7995
               TabIndex        =   390
               Top             =   3615
               Width           =   480
            End
            Begin VB.CommandButton cmdSerach 
               Caption         =   "搜尋(&2)"
               Height          =   585
               Index           =   1
               Left            =   4545
               TabIndex        =   389
               Top             =   3015
               Width           =   345
            End
            Begin VB.Frame Frame4 
               Height          =   405
               Index           =   1
               Left            =   1395
               TabIndex        =   386
               Top             =   2295
               Width           =   1695
               Begin VB.OptionButton optCP812 
                  Caption         =   "是"
                  Enabled         =   0   'False
                  Height          =   255
                  Index           =   0
                  Left            =   120
                  TabIndex        =   388
                  Top             =   120
                  Width           =   615
               End
               Begin VB.OptionButton optCP812 
                  Caption         =   "否"
                  Enabled         =   0   'False
                  Height          =   255
                  Index           =   1
                  Left            =   930
                  TabIndex        =   387
                  Top             =   120
                  Width           =   660
               End
            End
            Begin VB.CommandButton cmdQual 
               Caption         =   "中小企業減免資格"
               Height          =   285
               Index           =   2
               Left            =   1380
               TabIndex        =   385
               Top             =   2715
               Visible         =   0   'False
               Width           =   1725
            End
            Begin VB.CommandButton CmdSame 
               Caption         =   "同上"
               BeginProperty Font 
                  Name            =   "新細明體"
                  Size            =   8.5
                  Charset         =   136
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Index           =   2
               Left            =   1035
               TabIndex        =   368
               Top             =   3915
               Visible         =   0   'False
               Width           =   550
            End
            Begin MSForms.TextBox Text1 
               Height          =   300
               Index           =   38
               Left            =   1395
               TabIndex        =   406
               Top             =   3315
               Width           =   3075
               VariousPropertyBits=   671107099
               Size            =   "5424;529"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.TextBox Text1 
               Height          =   300
               Index           =   37
               Left            =   1395
               TabIndex        =   405
               Top             =   3015
               Width           =   3075
               VariousPropertyBits=   671107099
               Size            =   "5424;529"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.TextBox Text1 
               Height          =   300
               Index           =   35
               Left            =   3975
               TabIndex        =   402
               Top             =   2385
               Width           =   4425
               VariousPropertyBits=   671107099
               Size            =   "7805;529"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.TextBox Text1 
               Height          =   300
               Index           =   50
               Left            =   3975
               TabIndex        =   401
               Top             =   2685
               Width           =   4425
               VariousPropertyBits=   671107099
               Size            =   "7805;529"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.TextBox Text1 
               Height          =   300
               Index           =   42
               Left            =   1035
               TabIndex        =   404
               Top             =   3615
               Width           =   6255
               VariousPropertyBits=   671107099
               Size            =   "11033;529"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.TextBox Text1 
               Height          =   300
               Index           =   141
               Left            =   1035
               TabIndex        =   403
               Top             =   4215
               Width           =   7305
               VariousPropertyBits=   671107099
               Size            =   "12885;529"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.TextBox Text1 
               Height          =   300
               Index           =   43
               Left            =   1035
               TabIndex        =   400
               Top             =   3915
               Width           =   6255
               VariousPropertyBits=   679495707
               Size            =   "11033;529"
               FontName        =   "新細明體"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.TextBox Text1 
               Height          =   300
               Index           =   33
               Left            =   7050
               TabIndex        =   421
               Top             =   2430
               Visible         =   0   'False
               Width           =   1335
               VariousPropertyBits=   671107097
               BackColor       =   -2147483648
               Size            =   "2355;529"
               FontName        =   "新細明體-ExtB"
               FontEffects     =   1073750016
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin VB.Label Label1 
               Caption         =   "聯絡地址："
               Height          =   255
               Index           =   62
               Left            =   135
               TabIndex        =   420
               Top             =   3615
               Width           =   945
            End
            Begin VB.Label Label1 
               Caption         =   "申請地址："
               Height          =   255
               Index           =   50
               Left            =   135
               TabIndex        =   419
               Top             =   3915
               Width           =   945
            End
            Begin MSForms.TextBox Text1 
               Height          =   300
               Index           =   41
               Left            =   7275
               TabIndex        =   418
               Top             =   3615
               Width           =   705
               VariousPropertyBits=   671107099
               Size            =   "1244;529"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.TextBox Text1 
               Height          =   300
               Index           =   40
               Left            =   6495
               TabIndex        =   417
               Top             =   3315
               Width           =   1905
               VariousPropertyBits=   671107099
               Size            =   "3360;529"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.TextBox Text1 
               Height          =   300
               Index           =   39
               Left            =   6495
               TabIndex        =   416
               Top             =   3015
               Width           =   1905
               VariousPropertyBits=   671107099
               Size            =   "3360;529"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.TextBox Text1 
               Height          =   300
               Index           =   121
               Left            =   7275
               TabIndex        =   415
               Top             =   3915
               Width           =   705
               VariousPropertyBits=   671107099
               Size            =   "1244;529"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin VB.Label Label1 
               Caption         =   "代表人(中文)"
               Height          =   255
               Index           =   34
               Left            =   4995
               TabIndex        =   414
               Top             =   3045
               Width           =   1455
            End
            Begin VB.Label Label1 
               Caption         =   "　　　(英文)："
               Height          =   255
               Index           =   33
               Left            =   135
               TabIndex        =   413
               Top             =   3375
               Width           =   1305
            End
            Begin VB.Label Label1 
               Caption         =   "申請人(中文)："
               Height          =   255
               Index           =   32
               Left            =   135
               TabIndex        =   412
               Top             =   3045
               Width           =   1305
            End
            Begin VB.Label Label1 
               Caption         =   "E-Mail："
               Height          =   255
               Index           =   26
               Left            =   3285
               TabIndex        =   411
               Top             =   2415
               Width           =   675
            End
            Begin VB.Label Label1 
               Caption         =   "　　　(英文)："
               Height          =   255
               Index           =   35
               Left            =   4995
               TabIndex        =   410
               Top             =   3330
               Width           =   1455
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "符合年費減免："
               Height          =   180
               Index           =   27
               Left            =   135
               TabIndex        =   409
               Top             =   2445
               Width           =   1260
            End
            Begin VB.Label Label1 
               Caption         =   "國籍："
               Height          =   255
               Index           =   110
               Left            =   3285
               TabIndex        =   408
               Top             =   2715
               Width           =   705
            End
            Begin VB.Label Label1 
               Caption         =   "英文地址："
               Height          =   255
               Index           =   114
               Left            =   135
               TabIndex        =   407
               Top             =   4245
               Width           =   945
            End
         End
         Begin VB.Frame Frame36 
            Caption         =   "Frame36"
            Height          =   5145
            Left            =   30
            TabIndex        =   340
            Top             =   330
            Width           =   8535
            Begin VB.Frame Frame50 
               Height          =   705
               Left            =   30
               TabIndex        =   341
               Top             =   2280
               Width           =   8445
               Begin MSForms.ComboBox cboContact 
                  Height          =   300
                  Index           =   1
                  Left            =   2850
                  TabIndex        =   61
                  Top             =   60
                  Width           =   1230
                  VariousPropertyBits=   679495707
                  DisplayStyle    =   3
                  Size            =   "2170;529"
                  MatchEntry      =   1
                  ShowDropButtonWhen=   2
                  FontName        =   "新細明體-ExtB"
                  FontHeight      =   180
                  FontCharSet     =   136
                  FontPitchAndFamily=   34
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  Caption         =   "ID No.："
                  Height          =   180
                  Index           =   77
                  Left            =   60
                  TabIndex        =   348
                  Top             =   390
                  Width           =   660
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  Caption         =   "客戶編號："
                  Height          =   180
                  Index           =   12
                  Left            =   60
                  TabIndex        =   347
                  Top             =   90
                  Width           =   900
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  Caption         =   "接洽人："
                  Height          =   180
                  Index           =   13
                  Left            =   2130
                  TabIndex        =   346
                  Top             =   90
                  Width           =   720
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  Caption         =   "電話："
                  Height          =   180
                  Index           =   14
                  Left            =   4110
                  TabIndex        =   345
                  Top             =   90
                  Width           =   540
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  Caption         =   "傳真："
                  Height          =   180
                  Index           =   15
                  Left            =   6480
                  TabIndex        =   344
                  Top             =   90
                  Width           =   540
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  Caption         =   "手機："
                  Height          =   180
                  Index           =   18
                  Left            =   6480
                  TabIndex        =   343
                  Top             =   360
                  Width           =   540
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  Caption         =   "LINE ID："
                  BeginProperty Font 
                     Name            =   "新細明體"
                     Size            =   8.5
                     Charset         =   136
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   165
                  Index           =   85
                  Left            =   4080
                  TabIndex        =   342
                  Top             =   390
                  Width           =   690
               End
               Begin MSForms.TextBox Text1 
                  Height          =   300
                  Index           =   20
                  Left            =   7050
                  TabIndex        =   67
                  Top             =   360
                  Width           =   1335
                  VariousPropertyBits=   671107099
                  Size            =   "2355;529"
                  FontName        =   "新細明體-ExtB"
                  FontHeight      =   180
                  FontCharSet     =   136
                  FontPitchAndFamily=   34
               End
               Begin MSForms.TextBox Text1 
                  Height          =   300
                  Index           =   12
                  Left            =   990
                  TabIndex        =   60
                  Top             =   60
                  Width           =   1095
                  VariousPropertyBits=   671107099
                  Size            =   "1931;529"
                  FontName        =   "新細明體-ExtB"
                  FontHeight      =   180
                  FontCharSet     =   136
                  FontPitchAndFamily=   34
               End
               Begin MSForms.TextBox Text1 
                  Height          =   300
                  Index           =   14
                  Left            =   4785
                  TabIndex        =   62
                  Top             =   60
                  Width           =   1650
                  VariousPropertyBits=   671107099
                  Size            =   "2910;529"
                  FontName        =   "新細明體-ExtB"
                  FontHeight      =   180
                  FontCharSet     =   136
                  FontPitchAndFamily=   34
               End
               Begin MSForms.TextBox Text1 
                  Height          =   300
                  Index           =   15
                  Left            =   4785
                  TabIndex        =   66
                  Top             =   360
                  Width           =   1650
                  VariousPropertyBits=   671107099
                  Size            =   "2910;529"
                  FontName        =   "新細明體-ExtB"
                  FontHeight      =   180
                  FontCharSet     =   136
                  FontPitchAndFamily=   34
               End
               Begin MSForms.TextBox Text1 
                  Height          =   300
                  Index           =   16
                  Left            =   7050
                  TabIndex        =   63
                  Top             =   60
                  Width           =   1335
                  VariousPropertyBits=   671107099
                  Size            =   "2355;529"
                  FontName        =   "新細明體-ExtB"
                  FontHeight      =   180
                  FontCharSet     =   136
                  FontPitchAndFamily=   34
               End
               Begin MSForms.TextBox Text1 
                  Height          =   300
                  Index           =   92
                  Left            =   990
                  TabIndex        =   64
                  Top             =   360
                  Width           =   1785
                  VariousPropertyBits=   671107099
                  Size            =   "3149;529"
                  FontName        =   "新細明體-ExtB"
                  FontHeight      =   180
                  FontCharSet     =   136
                  FontPitchAndFamily=   34
               End
               Begin MSForms.TextBox Text1 
                  Height          =   300
                  Index           =   13
                  Left            =   2850
                  TabIndex        =   65
                  Top             =   360
                  Visible         =   0   'False
                  Width           =   1230
                  VariousPropertyBits=   671107099
                  Size            =   "2170;529"
                  FontName        =   "新細明體-ExtB"
                  FontHeight      =   180
                  FontCharSet     =   136
                  FontPitchAndFamily=   34
               End
            End
            Begin VB.Frame Frame4 
               Height          =   405
               Index           =   0
               Left            =   1395
               TabIndex        =   349
               Top             =   2895
               Width           =   1695
               Begin VB.OptionButton optCP811 
                  Caption         =   "否"
                  Enabled         =   0   'False
                  Height          =   255
                  Index           =   1
                  Left            =   930
                  TabIndex        =   69
                  Top             =   120
                  Width           =   660
               End
               Begin VB.OptionButton optCP811 
                  Caption         =   "是"
                  Enabled         =   0   'False
                  Height          =   255
                  Index           =   0
                  Left            =   120
                  TabIndex        =   68
                  Top             =   120
                  Width           =   615
               End
            End
            Begin VB.Frame Frame7 
               Height          =   2280
               Left            =   30
               TabIndex        =   350
               Top             =   -30
               Width           =   8430
               Begin VB.Frame Frame12 
                  Caption         =   "Frame12"
                  Height          =   825
                  Left            =   4350
                  TabIndex        =   608
                  Top             =   1350
                  Width           =   3975
                  Begin VB.OptionButton Option5 
                     Caption         =   "相同"
                     ForeColor       =   &H000000FF&
                     Height          =   315
                     Index           =   0
                     Left            =   90
                     TabIndex        =   623
                     Top             =   240
                     Width           =   675
                  End
                  Begin VB.OptionButton Option5 
                     Caption         =   "有關"
                     ForeColor       =   &H000000FF&
                     Height          =   315
                     Index           =   1
                     Left            =   90
                     TabIndex        =   622
                     Top             =   480
                     Width           =   675
                  End
                  Begin VB.CommandButton cmdCRL55 
                     BackColor       =   &H00C0C0C0&
                     Caption         =   "本案與總號"
                     Height          =   255
                     Left            =   30
                     Style           =   1  '圖片外觀
                     TabIndex        =   609
                     Top             =   -30
                     Width           =   1125
                  End
                  Begin MSForms.TextBox Text1 
                     Height          =   660
                     Index           =   100
                     Left            =   1200
                     TabIndex        =   610
                     Top             =   60
                     Width           =   2385
                     VariousPropertyBits=   -1467987937
                     ScrollBars      =   2
                     Size            =   "4207;1164"
                     FontName        =   "新細明體-ExtB"
                     FontHeight      =   180
                     FontCharSet     =   136
                     FontPitchAndFamily=   34
                  End
                  Begin VB.Label Label1 
                     Caption         =   "本案與總號："
                     ForeColor       =   &H00000000&
                     Height          =   255
                     Index           =   84
                     Left            =   60
                     TabIndex        =   611
                     Top             =   0
                     Width           =   1635
                  End
               End
               Begin VB.Frame Frame30 
                  Caption         =   "特例簽核"
                  ForeColor       =   &H00000000&
                  Height          =   495
                  Index           =   0
                  Left            =   2130
                  TabIndex        =   593
                  Top             =   60
                  Width           =   2085
                  Begin VB.CheckBox ChkCRA27 
                     Caption         =   "有跨所"
                     Enabled         =   0   'False
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
                     Height          =   270
                     Index           =   0
                     Left            =   1080
                     TabIndex        =   595
                     Top             =   180
                     Width           =   975
                  End
                  Begin VB.CheckBox ChkCRA26 
                     Caption         =   "有對造"
                     Enabled         =   0   'False
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
                     Height          =   270
                     Index           =   0
                     Left            =   90
                     TabIndex        =   594
                     Top             =   180
                     Width           =   885
                  End
               End
               Begin VB.OptionButton Option4 
                  Caption         =   "自動來所"
                  Height          =   345
                  Index           =   3
                  Left            =   2460
                  TabIndex        =   54
                  Top             =   615
                  Width           =   1275
               End
               Begin VB.OptionButton Option4 
                  Caption         =   "同仁"
                  Height          =   345
                  Index           =   2
                  Left            =   540
                  TabIndex        =   57
                  Top             =   1290
                  Width           =   705
               End
               Begin VB.OptionButton Option4 
                  Caption         =   "客戶"
                  Height          =   255
                  Index           =   1
                  Left            =   540
                  TabIndex        =   56
                  Top             =   1050
                  Width           =   1275
               End
               Begin VB.OptionButton Option4 
                  Caption         =   "主動開拓"
                  Height          =   345
                  Index           =   0
                  Left            =   540
                  TabIndex        =   53
                  Top             =   600
                  Width           =   1275
               End
               Begin VB.Frame Frame11 
                  Height          =   255
                  Left            =   1140
                  TabIndex        =   351
                  Top             =   270
                  Width           =   5760
                  Begin VB.OptionButton Option31 
                     Caption         =   "舊客戶"
                     Height          =   285
                     Index           =   1
                     Left            =   4470
                     TabIndex        =   52
                     Top             =   -30
                     Value           =   -1  'True
                     Width           =   975
                  End
                  Begin VB.OptionButton Option31 
                     Caption         =   "新客戶"
                     Height          =   285
                     Index           =   0
                     Left            =   0
                     TabIndex        =   51
                     Top             =   -15
                     Width           =   975
                  End
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  Caption         =   "(74001)"
                  Height          =   180
                  Index           =   142
                  Left            =   1410
                  TabIndex        =   632
                  Top             =   1680
                  Visible         =   0   'False
                  Width           =   570
               End
               Begin MSForms.TextBox Text1 
                  Height          =   300
                  Index           =   198
                  Left            =   4640
                  TabIndex        =   55
                  Top             =   650
                  Width           =   2600
                  VariousPropertyBits=   671107099
                  Size            =   "4586;529"
                  FontName        =   "新細明體-ExtB"
                  FontHeight      =   180
                  FontCharSet     =   136
                  FontPitchAndFamily=   34
               End
               Begin MSForms.TextBox Text1 
                  Height          =   300
                  Index           =   99
                  Left            =   4875
                  TabIndex        =   59
                  Top             =   1020
                  Width           =   3435
                  VariousPropertyBits=   671107099
                  Size            =   "6059;529"
                  FontName        =   "新細明體-ExtB"
                  FontHeight      =   180
                  FontCharSet     =   136
                  FontPitchAndFamily=   34
               End
               Begin MSForms.TextBox Text1 
                  Height          =   300
                  Index           =   97
                  Left            =   1380
                  TabIndex        =   58
                  Top             =   1350
                  Width           =   1245
                  VariousPropertyBits=   671107099
                  Size            =   "2196;529"
                  FontName        =   "新細明體-ExtB"
                  FontHeight      =   180
                  FontCharSet     =   136
                  FontPitchAndFamily=   34
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  Caption         =   "介紹"
                  Height          =   180
                  Index           =   37
                  Left            =   2730
                  TabIndex        =   355
                  Top             =   1380
                  Width           =   360
               End
               Begin VB.Line Line7 
                  X1              =   510
                  X2              =   8370
                  Y1              =   975
                  Y2              =   975
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  Caption         =   "其他："
                  Height          =   180
                  Index           =   83
                  Left            =   4275
                  TabIndex        =   354
                  Top             =   1050
                  Width           =   540
               End
               Begin VB.Label Label2 
                  Alignment       =   2  '置中對齊
                  Caption         =   "案件來源說明"
                  Height          =   1110
                  Index           =   0
                  Left            =   135
                  TabIndex        =   353
                  Top             =   465
                  Width           =   315
               End
               Begin VB.Line Line2 
                  X1              =   495
                  X2              =   495
                  Y1              =   195
                  Y2              =   2220
               End
               Begin VB.Line Line3 
                  X1              =   495
                  X2              =   8355
                  Y1              =   555
                  Y2              =   555
               End
               Begin VB.Line Line4 
                  X1              =   4245
                  X2              =   4245
                  Y1              =   570
                  Y2              =   2220
               End
               Begin VB.Label Label1 
                  Caption         =   "與                                                        為關係企業"
                  Height          =   260
                  Index           =   82
                  Left            =   4340
                  TabIndex        =   352
                  Top             =   680
                  Width           =   4010
               End
               Begin VB.Shape Shape1 
                  Height          =   2055
                  Left            =   60
                  Top             =   180
                  Width           =   8325
               End
            End
            Begin VB.CommandButton cmdTW 
               Caption         =   "聯絡地址："
               BeginProperty Font 
                  Name            =   "新細明體"
                  Size            =   8.5
                  Charset         =   136
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Index           =   0
               Left            =   30
               TabIndex        =   78
               TabStop         =   0   'False
               Top             =   4215
               Width           =   980
            End
            Begin VB.CommandButton cmdTW 
               Caption         =   "申請地址："
               BeginProperty Font 
                  Name            =   "新細明體"
                  Size            =   8.5
                  Charset         =   136
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Index           =   1
               Left            =   30
               TabIndex        =   82
               TabStop         =   0   'False
               Top             =   4515
               Width           =   980
            End
            Begin VB.CommandButton cmdSearchZip 
               Caption         =   "Zip"
               BeginProperty Font 
                  Name            =   "新細明體"
                  Size            =   8.5
                  Charset         =   136
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   270
               Index           =   0
               Left            =   8010
               TabIndex        =   81
               Top             =   4215
               Width           =   480
            End
            Begin VB.CommandButton cmdSearchZip 
               Caption         =   "Zip"
               BeginProperty Font 
                  Name            =   "新細明體"
                  Size            =   8.5
                  Charset         =   136
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   270
               Index           =   1
               Left            =   8010
               TabIndex        =   86
               Top             =   4515
               Width           =   480
            End
            Begin VB.CommandButton cmdSerach 
               Caption         =   "搜尋(&1)"
               Height          =   585
               Index           =   0
               Left            =   4545
               TabIndex        =   75
               Top             =   3615
               Width           =   345
            End
            Begin VB.CommandButton cmdQual 
               Caption         =   "中小企業減免資格"
               Height          =   285
               Index           =   1
               Left            =   1380
               TabIndex        =   70
               Top             =   3315
               Visible         =   0   'False
               Width           =   1725
            End
            Begin VB.CommandButton CmdSame 
               Caption         =   "同上"
               BeginProperty Font 
                  Name            =   "新細明體"
                  Size            =   8.5
                  Charset         =   136
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Index           =   1
               Left            =   1005
               TabIndex        =   83
               Top             =   4515
               Visible         =   0   'False
               Width           =   550
            End
            Begin MSForms.TextBox Text1 
               Height          =   300
               Index           =   22
               Left            =   1395
               TabIndex        =   74
               Top             =   3915
               Width           =   3075
               VariousPropertyBits=   671107099
               Size            =   "5424;529"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.TextBox Text1 
               Height          =   300
               Index           =   21
               Left            =   1395
               TabIndex        =   73
               Top             =   3615
               Width           =   3075
               VariousPropertyBits=   671107099
               Size            =   "5424;529"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.TextBox Text1 
               Height          =   300
               Index           =   125
               Left            =   1005
               TabIndex        =   87
               Top             =   4815
               Width           =   7305
               VariousPropertyBits=   671107099
               Size            =   "12885;529"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.TextBox Text1 
               Height          =   300
               Index           =   26
               Left            =   1005
               TabIndex        =   79
               Top             =   4215
               Width           =   6300
               VariousPropertyBits=   671107099
               Size            =   "11112;529"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.TextBox Text1 
               Height          =   300
               Index           =   27
               Left            =   1005
               TabIndex        =   84
               Top             =   4515
               Width           =   6300
               VariousPropertyBits=   679495707
               Size            =   "11112;529"
               FontName        =   "新細明體"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.TextBox Text1 
               Height          =   300
               Index           =   19
               Left            =   3975
               TabIndex        =   71
               Top             =   2985
               Width           =   4425
               VariousPropertyBits=   671107099
               Size            =   "7805;529"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.TextBox Text1 
               Height          =   300
               Index           =   34
               Left            =   3975
               TabIndex        =   72
               Top             =   3285
               Width           =   4425
               VariousPropertyBits=   671107099
               Size            =   "7805;529"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.TextBox Text1 
               Height          =   300
               Index           =   17
               Left            =   6990
               TabIndex        =   366
               Top             =   3030
               Visible         =   0   'False
               Width           =   1335
               VariousPropertyBits=   671107097
               BackColor       =   -2147483648
               Size            =   "2355;529"
               FontName        =   "新細明體-ExtB"
               FontEffects     =   1073750016
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin VB.Label Label1 
               Caption         =   "英文地址："
               Height          =   255
               Index           =   92
               Left            =   75
               TabIndex        =   365
               Top             =   4845
               Width           =   1305
            End
            Begin VB.Label Label1 
               Caption         =   "聯絡地址："
               Height          =   255
               Index           =   75
               Left            =   75
               TabIndex        =   364
               Top             =   4350
               Width           =   945
            End
            Begin VB.Label Label1 
               Caption         =   "申請地址："
               Height          =   255
               Index           =   63
               Left            =   75
               TabIndex        =   363
               Top             =   4650
               Width           =   945
            End
            Begin MSForms.TextBox Text1 
               Height          =   300
               Index           =   25
               Left            =   7305
               TabIndex        =   80
               Top             =   4215
               Width           =   705
               VariousPropertyBits=   671107099
               Size            =   "1244;529"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.TextBox Text1 
               Height          =   300
               Index           =   24
               Left            =   6495
               TabIndex        =   77
               Top             =   3915
               Width           =   1905
               VariousPropertyBits=   671107099
               Size            =   "3360;529"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.TextBox Text1 
               Height          =   300
               Index           =   23
               Left            =   6495
               TabIndex        =   76
               Top             =   3615
               Width           =   1905
               VariousPropertyBits=   671107099
               Size            =   "3360;529"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.TextBox Text1 
               Height          =   300
               Index           =   120
               Left            =   7305
               TabIndex        =   85
               Top             =   4515
               Width           =   705
               VariousPropertyBits=   671107099
               Size            =   "1244;529"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin VB.Label Label1 
               Caption         =   "代表人(中文)"
               Height          =   255
               Index           =   21
               Left            =   4995
               TabIndex        =   362
               Top             =   3645
               Width           =   1455
            End
            Begin VB.Label Label1 
               Caption         =   "　　　(英文)："
               Height          =   255
               Index           =   20
               Left            =   135
               TabIndex        =   361
               Top             =   3945
               Width           =   1305
            End
            Begin VB.Label Label1 
               Caption         =   "申請人(中文)："
               Height          =   255
               Index           =   19
               Left            =   135
               TabIndex        =   360
               Top             =   3615
               Width           =   1305
            End
            Begin VB.Label Label1 
               Caption         =   "E-Mail："
               Height          =   255
               Index           =   17
               Left            =   3285
               TabIndex        =   359
               Top             =   3015
               Width           =   675
            End
            Begin VB.Label Label1 
               Caption         =   "　　　(英文)："
               Height          =   255
               Index           =   22
               Left            =   4995
               TabIndex        =   358
               Top             =   3930
               Width           =   1455
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "符合年費減免："
               Height          =   180
               Index           =   5
               Left            =   135
               TabIndex        =   357
               Top             =   3045
               Width           =   1260
            End
            Begin VB.Label Label1 
               Caption         =   "國籍："
               Height          =   255
               Index           =   94
               Left            =   3285
               TabIndex        =   356
               Top             =   3315
               Width           =   675
            End
         End
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid GRD1 
         Bindings        =   "frm090801_Q.frx":01DC
         Height          =   2430
         Left            =   -70425
         TabIndex        =   194
         Top             =   3180
         Width           =   4185
         _ExtentX        =   7391
         _ExtentY        =   4269
         _Version        =   393216
         FixedCols       =   0
         ScrollTrack     =   -1  'True
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
         _Band(0).Cols   =   2
      End
      Begin MSForms.Label Label29 
         Height          =   260
         Left            =   -73005
         TabIndex        =   650
         Top             =   5640
         Width           =   6700
         Caption         =   "Create ID:           Date         Time             Update ID:                Date                  Time"
         Size            =   "11818;459"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label LblReason 
         Appearance      =   0  '平面
         BorderStyle     =   1  '單線固定
         Caption         =   "LblReason"
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   -73605
         TabIndex        =   625
         Top             =   2820
         Width           =   2505
      End
      Begin VB.Label Label36 
         AutoSize        =   -1  'True
         Caption         =   "(系統產生)"
         Height          =   240
         Left            =   -74895
         TabIndex        =   624
         Top             =   1350
         Width           =   840
      End
      Begin VB.Label Label37 
         AutoSize        =   -1  'True
         Caption         =   "(一般簽核)"
         Height          =   180
         Left            =   -74895
         TabIndex        =   592
         Top             =   1620
         Width           =   840
      End
      Begin MSForms.TextBox PicText 
         Height          =   300
         Left            =   420
         TabIndex        =   339
         Top             =   2070
         Width           =   3855
         VariousPropertyBits=   671107099
         Size            =   "6800;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text4 
         Height          =   315
         Index           =   0
         Left            =   -70320
         TabIndex        =   237
         Top             =   1365
         Width           =   3705
         VariousPropertyBits=   671107099
         Size            =   "6535;556"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text2 
         Height          =   315
         Index           =   9
         Left            =   -74670
         TabIndex        =   254
         Top             =   4875
         Width           =   1335
         VariousPropertyBits=   671107099
         Size            =   "2355;556"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text2 
         Height          =   315
         Index           =   8
         Left            =   -74670
         TabIndex        =   252
         Top             =   4485
         Width           =   1335
         VariousPropertyBits=   671107099
         Size            =   "2355;556"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text2 
         Height          =   315
         Index           =   7
         Left            =   -74670
         TabIndex        =   250
         Top             =   4095
         Width           =   1335
         VariousPropertyBits=   671107099
         Size            =   "2355;556"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text2 
         Height          =   315
         Index           =   6
         Left            =   -74670
         TabIndex        =   248
         Top             =   3705
         Width           =   1335
         VariousPropertyBits=   671107099
         Size            =   "2355;556"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text2 
         Height          =   315
         Index           =   5
         Left            =   -74670
         TabIndex        =   246
         Top             =   3315
         Width           =   1335
         VariousPropertyBits=   671107099
         Size            =   "2355;556"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text2 
         Height          =   315
         Index           =   4
         Left            =   -74670
         TabIndex        =   244
         Top             =   2925
         Width           =   1335
         VariousPropertyBits=   671107099
         Size            =   "2355;556"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text2 
         Height          =   315
         Index           =   3
         Left            =   -74670
         TabIndex        =   242
         Top             =   2535
         Width           =   1335
         VariousPropertyBits=   671107099
         Size            =   "2355;556"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text2 
         Height          =   315
         Index           =   2
         Left            =   -74670
         TabIndex        =   240
         Top             =   2145
         Width           =   1335
         VariousPropertyBits=   671107099
         Size            =   "2355;556"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text2 
         Height          =   315
         Index           =   1
         Left            =   -74670
         TabIndex        =   238
         Top             =   1755
         Width           =   1335
         VariousPropertyBits=   671107099
         Size            =   "2355;556"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text2 
         Height          =   315
         Index           =   0
         Left            =   -74670
         TabIndex        =   236
         Top             =   1365
         Width           =   1335
         VariousPropertyBits=   671107099
         Size            =   "2355;556"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label6 
         Caption         =   "PS：發明人未建檔者，標題顯示黃色"
         ForeColor       =   &H000000FF&
         Height          =   615
         Index           =   1
         Left            =   -72780
         TabIndex        =   271
         Top             =   420
         Width           =   5895
      End
      Begin VB.Label Label10 
         Caption         =   "同申請地址"
         Height          =   405
         Left            =   -66780
         TabIndex        =   270
         Top             =   915
         Width           =   585
      End
      Begin VB.Label Label5 
         Caption         =   "10."
         Height          =   285
         Index           =   9
         Left            =   -74910
         TabIndex        =   269
         Top             =   4905
         Width           =   285
      End
      Begin VB.Label Label5 
         Caption         =   "9."
         Height          =   285
         Index           =   8
         Left            =   -74910
         TabIndex        =   268
         Top             =   4545
         Width           =   285
      End
      Begin VB.Label Label5 
         Caption         =   "8."
         Height          =   285
         Index           =   7
         Left            =   -74910
         TabIndex        =   267
         Top             =   4155
         Width           =   285
      End
      Begin VB.Label Label5 
         Caption         =   "7."
         Height          =   285
         Index           =   6
         Left            =   -74910
         TabIndex        =   266
         Top             =   3765
         Width           =   285
      End
      Begin VB.Label Label5 
         Caption         =   "6."
         Height          =   285
         Index           =   5
         Left            =   -74910
         TabIndex        =   265
         Top             =   3375
         Width           =   285
      End
      Begin VB.Label Label5 
         Caption         =   "5."
         Height          =   285
         Index           =   4
         Left            =   -74910
         TabIndex        =   264
         Top             =   2985
         Width           =   285
      End
      Begin VB.Label Label5 
         Caption         =   "4."
         Height          =   285
         Index           =   3
         Left            =   -74910
         TabIndex        =   263
         Top             =   2595
         Width           =   285
      End
      Begin VB.Label Label5 
         Caption         =   "3."
         Height          =   285
         Index           =   2
         Left            =   -74910
         TabIndex        =   262
         Top             =   2205
         Width           =   285
      End
      Begin VB.Label Label5 
         Caption         =   "2."
         Height          =   285
         Index           =   1
         Left            =   -74910
         TabIndex        =   261
         Top             =   1815
         Width           =   285
      End
      Begin VB.Label Label5 
         Caption         =   "1."
         Height          =   285
         Index           =   0
         Left            =   -74910
         TabIndex        =   260
         Top             =   1425
         Width           =   285
      End
      Begin VB.Label Label9 
         Caption         =   "地址"
         Height          =   285
         Index           =   0
         Left            =   -70320
         TabIndex        =   259
         Top             =   1185
         Width           =   795
      End
      Begin VB.Label Label8 
         Caption         =   "國籍"
         Height          =   285
         Index           =   0
         Left            =   -73290
         TabIndex        =   258
         Top             =   1185
         Width           =   795
      End
      Begin VB.Label Label7 
         Caption         =   "ID"
         Height          =   285
         Index           =   0
         Left            =   -71610
         TabIndex        =   257
         Top             =   1155
         Width           =   795
      End
      Begin VB.Label Label6 
         Caption         =   "發明人姓名"
         Height          =   255
         Index           =   0
         Left            =   -74670
         TabIndex        =   256
         Top             =   1155
         Width           =   1215
      End
      Begin MSForms.TextBox Text4 
         Height          =   315
         Index           =   9
         Left            =   -70320
         TabIndex        =   255
         Top             =   4875
         Width           =   3705
         VariousPropertyBits=   671107099
         Size            =   "6535;556"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text4 
         Height          =   315
         Index           =   8
         Left            =   -70320
         TabIndex        =   253
         Top             =   4485
         Width           =   3705
         VariousPropertyBits=   671107099
         Size            =   "6535;556"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text4 
         Height          =   315
         Index           =   7
         Left            =   -70320
         TabIndex        =   251
         Top             =   4095
         Width           =   3705
         VariousPropertyBits=   671107099
         Size            =   "6535;556"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text4 
         Height          =   315
         Index           =   6
         Left            =   -70320
         TabIndex        =   249
         Top             =   3705
         Width           =   3705
         VariousPropertyBits=   671107099
         Size            =   "6535;556"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text4 
         Height          =   315
         Index           =   5
         Left            =   -70320
         TabIndex        =   247
         Top             =   3315
         Width           =   3705
         VariousPropertyBits=   671107099
         Size            =   "6535;556"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text4 
         Height          =   315
         Index           =   4
         Left            =   -70320
         TabIndex        =   245
         Top             =   2925
         Width           =   3705
         VariousPropertyBits=   671107099
         Size            =   "6535;556"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text4 
         Height          =   315
         Index           =   3
         Left            =   -70320
         TabIndex        =   243
         Top             =   2535
         Width           =   3705
         VariousPropertyBits=   671107099
         Size            =   "6535;556"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text4 
         Height          =   315
         Index           =   2
         Left            =   -70320
         TabIndex        =   241
         Top             =   2145
         Width           =   3705
         VariousPropertyBits=   671107099
         Size            =   "6535;556"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text4 
         Height          =   315
         Index           =   1
         Left            =   -70320
         TabIndex        =   239
         Top             =   1755
         Width           =   3705
         VariousPropertyBits=   671107099
         Size            =   "6535;556"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtF0310_2 
         Height          =   255
         Left            =   -73335
         TabIndex        =   205
         Top             =   510
         Width           =   1035
         VariousPropertyBits=   679495711
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ScrollBars      =   3
         Size            =   "1826;450"
         Value           =   "txtF0310_2"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtF0407 
         Height          =   2430
         Left            =   -73995
         TabIndex        =   204
         Top             =   3180
         Width           =   3525
         VariousPropertyBits=   -1466939361
         ScrollBars      =   3
         Size            =   "6218;4286"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         Caption         =   "簽核歷程："
         Height          =   180
         Left            =   -74895
         TabIndex        =   203
         Top             =   3180
         Width           =   900
      End
      Begin VB.Label Label32 
         AutoSize        =   -1  'True
         Caption         =   "目前表單狀態："
         ForeColor       =   &H000000C0&
         Height          =   180
         Left            =   -74910
         TabIndex        =   202
         Top             =   840
         Width           =   1260
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "填單人員："
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   128
         Left            =   -74895
         TabIndex        =   201
         Top             =   540
         Width           =   900
      End
      Begin MSForms.TextBox txtF0306 
         Height          =   510
         Left            =   -73995
         TabIndex        =   200
         Top             =   2280
         Width           =   7755
         VariousPropertyBits=   -1466939365
         BackColor       =   -2147483633
         ScrollBars      =   3
         Size            =   "13679;900"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label33 
         AutoSize        =   -1  'True
         Caption         =   "智權備註："
         Height          =   180
         Left            =   -74895
         TabIndex        =   199
         Top             =   2280
         Width           =   900
      End
      Begin VB.Label Label1 
         Caption         =   "櫃檯退回原因："
         Height          =   180
         Index           =   130
         Left            =   -74895
         TabIndex        =   198
         Top             =   2850
         Width           =   1455
      End
      Begin VB.Label Label35 
         AutoSize        =   -1  'True
         Caption         =   "呈報主管："
         Height          =   180
         Left            =   -74895
         TabIndex        =   197
         Top             =   1140
         Width           =   900
      End
      Begin MSForms.TextBox txtCRL69 
         Height          =   1140
         Left            =   -73995
         TabIndex        =   196
         Top             =   1110
         Width           =   4215
         VariousPropertyBits=   -1466939361
         BackColor       =   12632319
         BorderStyle     =   1
         ScrollBars      =   2
         Size            =   "7435;2011"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label LblRecved 
         Caption         =   "已收文"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   -71625
         TabIndex        =   195
         Top             =   450
         Visible         =   0   'False
         Width           =   855
      End
   End
   Begin MSForms.TextBox txtNote 
      Height          =   660
      Left            =   1200
      TabIndex        =   648
      Top             =   9360
      Width           =   7755
      VariousPropertyBits=   -1466939365
      ScrollBars      =   3
      Size            =   "13679;1164"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label34 
      AutoSize        =   -1  'True
      Caption         =   "您的意見："
      Height          =   180
      Left            =   300
      TabIndex        =   647
      Top             =   9360
      Width           =   900
   End
   Begin VB.Label Label3 
      Caption         =   "(可由方向鍵切換頁籤)"
      ForeColor       =   &H00FF0000&
      Height          =   200
      Left            =   3900
      TabIndex        =   639
      Top             =   360
      Width           =   2570
   End
   Begin VB.Label lblAPPLQ 
      AutoSize        =   -1  'True
      Caption         =   "有對造"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   180
      Left            =   3300
      TabIndex        =   628
      Top             =   210
      Visible         =   0   'False
      Width           =   680
   End
   Begin VB.Label lblZip 
      AutoSize        =   -1  'True
      Caption         =   "有跨所"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   180
      Left            =   3300
      TabIndex        =   627
      Top             =   0
      Visible         =   0   'False
      Width           =   680
   End
   Begin VB.Label LblText5 
      Caption         =   "編號："
      Height          =   260
      Left            =   4080
      TabIndex        =   127
      Top             =   30
      Visible         =   0   'False
      Width           =   1010
   End
End
Attribute VB_Name = "frm090801_Q"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Modify By Sindy 2022/9/8 案件性質原寫固定4筆,改寫用Grid方式,呈現/輸入多筆
'Memo by Morgan 2022/1/20 改成Form2.0 (Text1,Text6,Text2,Text4,Text7,Combo2,cboTitle,cboContact,lblStaffName)
'Memo by Lydia 2019/07/01 表單名稱:國內案件接洽記錄單=>案件接洽單
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/8/3 日期欄已修改
'Modified by Morgan 2022/3/10 跳行控制統一用 vbCrLf 取代(原來程式還有 chr(13),chr(13)+chr(10),chr(10)+chr(13) 3種寫法，但在 2.0物件會有不同的顯示，正確應該是 chr(13)+chr(10)才對)
Option Explicit

Dim m_intTotPage As Integer '總頁數
Public m_strCustCode As String '申請人代號
Public m_blnOneRec As Boolean  '一筆資料
Dim m_strPA08 As String
Dim m_strPA16 As String
Dim m_strPA10 As String     '2010/3/31 ADD BY SONIA
Dim m_intCustCnt As Integer '計算申請人個數
Dim m_strST06 As String '所別
'Add by Morgan 2004/1/8
Dim intCopies As Integer
Dim m_strPA14 As String     '2010/8/17 add by sonia

'Add by Morgan 2004/5/20
Dim stCountry As String '申請國家
Dim stCustNo1 As String  '申請人1
'add by nick 2004/08/05
Dim stCustNo2 As String  '申請人2
Dim stCustNo3 As String  '申請人3
Dim stCustNo4 As String  '申請人4
Dim stCustNo5 As String  '申請人5
'每行往下移多少
Const CustMove = 80
Dim iiiii As Integer
Dim PrinterPages As Integer
'add by nick 2004/10/05
Dim IsoptCP81 As Boolean
'add by nickc 2005/05/27
Dim Is307or308 As Boolean
Dim CountBy307308 As String
Dim NowCount As Integer
'add by nickc 2006/07/19
Dim Is308New As Boolean
Dim Is308Old As Boolean
Dim Is308Child As Boolean
'add by nickc 2007/11/12 加入特殊客戶檢查
Dim IsSpecCu As Boolean
Dim SpecCUName As String
Dim SpecMemo As String
'add by nickc 2008/01/18 加入智權人員的客戶備註
Dim IsCuMemo As Boolean
Dim CuMemoName As String
Dim CuMemo As String
Dim m_strTM12 As String  '2008/3/21 ADD BY SONIA
Dim m_strTM15 As String  '2008/3/21 ADD BY SONIA
Dim strInventorNo As String         '2008/8/25  ADD BY TONI
Dim strInventorName As String     '2011/1/31  ADD BY Sindy
Dim strPetition  As String              '2008/8/29  ADD BY TONI

Public IsWmf As Boolean 'Add By Sindy 2009/08/31
Dim m_416Fee             '2010/1/6 ADD BY SONIA 檢查台灣發明實審
Dim m_Note1 As String, m_Note2 As String 'Add By Sindy 2010/4/6
Dim m_strGetNP01 As String 'Add By Sindy 2015/9/17
Dim pa() As String                'Add By Sindy 2010/7/9
Dim m_strTitleName As String
'Add By Sindy 2010/6/21
Dim PicRs As ADODB.Recordset
Dim file_num As Integer
Dim bytes() As Byte
Dim m_Image As New cImage
Dim m_Jpeg  As cJpeg
Dim strCRL() As String
Dim strCRA(1 To 27) As String
'2010/6/21 End
'Add By Sindy 2022/8/29
Dim strCRC(1 To 8) As String
Dim m_strSys As String
'2022/8/29 End
Public m_blnCallPrint As Boolean  'Add By Sindy 2010/6/21 外部呼叫查詢/列印
Public m_blnCallPrint_CRL119 As Boolean 'Add By Sindy 2014/2/7 是否列印特殊收據頁
Dim m_Device 'Add By Sindy 2010/7/7
Public bolPrint As Boolean, intPCnt As Integer 'Add By Sindy 2010/7/7 從frm12040152傳來的變數值
Dim bolRuleFeeErr As Boolean 'Add By Sindy 2010/11/22
Dim m_bCaseClosed As Boolean 'Add by Morgan 2010/12/10 是否已閉卷
Dim m_CP05 As String, m_CP09 As String, m_CP14 As String 'Add by Morgan 2010/12/20
Dim m_strCRL57 As String 'Add by Morgan 2011/1/13 暫存案件說明事項處理情形
Public m_bolPrintMark As Boolean 'Add by Morgan 2010/1/14 是否列印非整批發文字樣
'Add By Sindy 2011/1/24
Dim m_ChkOCaseAndCAddrNotAlike As Boolean
Public rsAddrNotAlike As New ADODB.Recordset
Dim PLeft(1 To 9) As Integer
Dim iLine As Integer
Dim strTemp(1 To 9) As String
'2011/1/24 End
'Add by Morgan 2011/3/2
Dim m_bolDueDayAlert As Boolean '當日法限提醒
Dim m_lngRefund As Long 'Add by Morgan 2010/10/12
Dim m_lstCaseNo As String 'Add by Morgan 2011/4/7
Dim bolPrintNewCase As Boolean 'Add By Sindy 2011/6/7 為新案
Public m_AppAddr As String, m_Zipcode As String 'Add By Sindy 2011/7/8
Public m_AppAddrChange As Boolean, m_CaseAndCAddrNotAlikePrint As Boolean 'Add By Sindy 2011/7/11
'2011/9/5 ADD BY SONIA
Dim stCustCAddr(1 To 5) As String  '申請人客戶檔中文地址
Dim stCustEAddr(1 To 5) As String  '申請人客戶檔英文地址
'2011/9/5 END
Dim m_CRL02 As String '填表日期 Add By Sindy 2011/10/18
Dim strCAddr As String, strEAddr As String, strZip As String   '2011/10/21add by sonia 自SetCustTxt移過來
Dim m_NA59 As String 'Add By Sindy 2011/11/11
Dim dblMoney As Double 'Add By Sindy 2012/3/16
Dim arrCaseProperty, arrNation 'Add By Sindy 2012/6/28
Dim m_strLiveProofMemo As String 'Added by Morgan 2012/11/23 存活證明備註
Dim dblAmt As Double, dblPFee As Double, dblTFee As Double 'Add By Sindy 2012/11/08
Dim dblChkAmt As Double 'Add By Sindy 2012/12/10
'Added by Lydia 2020/02/03
Dim dblCu183 As Double '個人之應收帳款上限
Dim dblAmtR As Double, dblPFeeR As Double, dblTFeeR As Double '關係企業之應收帳款金額
'end 2020/02/03

Dim oChk As CheckBox 'Added by Morgan 2013/1/15
Dim m_bolAdd803Memo As Boolean 'Added by Morgan 2013/1/15
'Added by Morgan 2013/4/2
Public m_stAD15 As String, m_stAD16 As String
Public m_stAD10 As String 'Added by Morgan 2019/4/12
Dim arrAD1516(5, 7) As String 'Modified by Morgan 2019/4/12 第2維 3,4,5,6,7要放AD10(新),CU15,AD10(原),AD15(原),AD16(原)
Dim m_CU143 As String 'Add By Sindy 2013/11/20 預定收款日放寬月數
Dim m_CU144(1 To 5) As String 'Add By Sindy 2013/12/16 不可開立發票
'Add By Sindy 2014/2/6
Public m_stCRL01 As String
Public m_stCRL97 As String, m_stCRL118 As String
Public m_stCRL98 As String, m_stCRL99 As String, m_stCRL100 As String, m_stCRL101 As String
Public m_stCRL102 As String, m_stCRL103 As String, m_stCRL104 As String, m_stCRL105 As String
Public m_stCRL106 As String, m_stCRL107 As String, m_stCRL108 As String, m_stCRL109 As String
Public m_stCRL110 As String, m_stCRL111 As String, m_stCRL112 As String, m_stCRL113 As String
Public m_stCRL120 As String, m_stCRL121 As String, m_stCRL122 As String, m_stCRL123 As String
'2014/2/6 END
'Add By Sindy 2015/8/28
'電話,傳真,E-Mail
Public m_stCRL114 As String, m_stCRL115 As String, m_stCRL116 As String
Public m_stCRL117 As String, m_stCRL124 As String, m_stCRL126 As String
Public m_stCRL127 As String, m_stCRL128 As String, m_stCRL129 As String
Public m_stCRL130 As String, m_stCRL131 As String, m_stCRL132 As String
'2015/8/28 END
'Public m_stCRL134 As String, m_stCRL135 As String '對造 陳報主管員編/原因 'Add by Amy 2016/09/02
Dim m_strCRL146 As Double '點數低於底價
Dim strTawT102MClassFee As String 'Add By Sindy 2014/2/19 舊案時檢查台灣商標延展跨類的規費
Public bolExternalCall As Boolean 'Add By Sindy 2014/5/23 外部程式(寄件查詢)呼叫此作業
Dim m_PrevForm As Form '前一畫面 Add By Sindy 2014/6/3
Public m_strSaveFiles As String 'Add By Sindy 2014/7/21 新增附件1區
Public m_strSaveFiles2 As String 'Add By Sindy 2022/10/11 新增附件2區
Dim strCase(1 To 4) As String, bolCaseIsExists As Boolean 'Add By Sindy 2014/7/28
Dim intCN308Field As Integer 'Add by Amy 2014/10/08 案件性質308欄位-for 大陸T分割案
Dim mPYFee As Boolean 'Add by Lydia 2014/12/22
'Add by Lydia 2015/01/14 分所保留判斷的性質=>屬於印2份,不屬於印3份
Private Const 分所保留 = "401,405,406,412,414,416,417,421,429,436,601,604,605,608,609,610,611,701,702,703,704,705,706,707,708,709,807,904,905,908,915,919,920,921,928"
'Added by Lydia 2015/10/14
Dim TMQList As String
'Add by Amy 2015/10/22
'Added by Lydia 2016/03/21 增加查名單輸入
Public Tmpfrm090126 As Form
Dim mTQC01 As String 'Added by Lydia 2016/04/06 記錄查名代號(民國年+流水號6碼)
Dim pTMQList As String 'Added by Lydia 2016/05/05記錄已列印的委查結果
Dim m_UseTmqTma As String 'Added by Lydia 2024/11/11 查名單(網中)：使用原查名單TMQ=1／查名單(網中)TMA=2做為查名的來源
Private Const m_TAutoRecv = "303,201,206,211" 'Add By Sindy 2016/5/11 T案可以自動收文的案件性質
'Dim strMemoPrt As String 'Added by Lydia 2016/05/27
'Add by Amy 2016/06/06 for 臺灣地址判斷
Dim bolNotChk As Boolean '列印時不檢查
'Add by Amy 2016/09/01
Public strCaseNo1 As String, strCaseNo2 As String, StrCaseNo3 As String, strCaseNo4 As String '商標相同號數之本所案號 'Memo by Lydia 2020/12/15 增加：CFT緬甸重新申請案
Public strTM28 As String '商標相同號數之本所案號的卷宗性質
Dim bolNotShow As Boolean 'for 當601異議/603評定/605廢止原新商標案改收舊案時不重抓資料
Dim bolCusCAddr(1 To 5) As Boolean 'Add by Amy 2016/12/23 是否有客戶中文地址-舊案用
'*****************************************************************************************
'Added by Lydia 2019/01/30 記錄代理人D/N備註
Dim m_FA45 As String    '專利D/N備註
Dim m_FA110 As String   '商標D/N備註
Dim m_Tuser As String 'Added by Lydia 2019/02/14 創新業務部預設收文人員
Dim m_LAmsg As String 'Added by Lydia 2019/04/10 顧問服務件數
'Add by Amy 2020/02/14
Public bolNotClsVal As Boolean '不清特殊收據資訊
Dim stChReceipt As String '收據抬頭顯示 1.收據抬頭/2.特殊收據
'Added by Lydia 2020/03/30 收據公司別(簡稱)
Dim m_CompName1 As String '1公司
Dim m_Comp1forIdx As Integer  '1公司的combo4.index
Dim m_CompName2 As String '2公司
Dim m_Comp2forIdx As Integer  '1公司的combo4.index
Dim m_CompNameJ As String 'J公司
Dim m_CompJforIdx As Integer  'J公司的combo4.index
Dim m_CompNameL As String 'L公司
Dim m_CompLforIdx As Integer  'L公司的combo4.index

'Added by Morgan 2020/4/21
'介紹案源回傳欄位
Public iReturn As Integer, strLawMan As String, strIntroducer As String, strCtrlDate As String, strTTCP09 As String, strTTCP10 As String, strTTSaveFiles As String
Public strLSourceType As String, strLC47 As String '案源類型: A,B1,B2,C; 法務案類型

Dim strCheckStatus As String '案源待補輸狀態
Dim bolIsPTCCase As Boolean, bolIsIPCase As Boolean, bolIsSuitCase As Boolean, bolSuYuan As Boolean '是否PTC案,是否智財權案,是否訴訟案,是否訴願案
Dim bolLawOfficeCase As Boolean, bolLawOfficeCase2 As Boolean '是否為介紹案源的法律案接洽單,是否為介紹案源的法律案接洽單2
Dim strLOS15 As String, strLOS17 As String, strLOS18 As String '案源單號,法務案接洽單號,P/T案接洽單號
Dim strLCaseCP10 As String '法律案預設收文案件性質
Dim strPTCP10 As String 'PT案預設收文案件性質
'P/T案:系統別,原說明事項,是否新案,規費
Dim strPTSysCode As String, strPTMemo As String, bolPTIsNew As Boolean, lngPTOFee As Long
'出庭費提醒,是否為補收B2案源的出庭費,法務案號
Dim bolCourtFeeAlert As Boolean, bolIsB2CourtFee As Boolean, strLCaseNo(4) As String
Dim strCUNo(5) As String
Dim strSalesDep As String '智權人員部門
Dim strLang As String '代理人查詢語文
'end 2020/4/21
Public bolIsTmp As Boolean 'Added by Morgan 2020/8/3
Dim strCaseNA239 As String, bolCase201 As Boolean 'Added by Lydia 2020/11/19 CFP和CFT英國脫歐案管制：歐盟案案號、新案第一次先不清除畫面
Dim strCaseNA239new As String 'Added by Lydia 2021/03/05 CFT歐盟尚未註冊案轉換英國申請案：記錄歐盟案案號
Private Const strACSdate1 = "20210428"  'Added by Lydia 2021/03/29 ACS案件收文與點數及營業稅：啟用日控制 'Memo by Lydia 2021/04/27  先上線
Dim strGetNp15 As String 'Added by Lydia 2021/04/15 CFP和CFT英國脫歐委任代理之後續處理：將下一程序備註欄之「脫歐英國案代理人：Y…」印在案件說明處理事項欄
Dim m_bol919Received As Boolean, m_lng605Fee As Long, m_lngRefunded As Long 'Added by Morgan 2021/4/29
Dim m_ACS112msg As String 'Added by Lydia 2021/05/07 ACS智財顧問專業分配比例管制：各專業部門、服務次數、工作時數小計、未發文次數
Dim m_ACS112chk As String 'Added by Lydia 2021/05/07 ACS智財顧問專業分配比例管制：已過該案之最大智財顧問112期間、總工作時數>=有效智財顧問112之CP15簽約時數，則檢查是否收文ACS智財顧問112

Dim strSameTrade As String 'Add by Amy 2021/11/29 國內同業msg
Dim m_strContactList(5) As String 'Added by Morgan 2022/1/20
'Added by Morgan 2022/1/25
Dim m_bWordFormat As Boolean '是否以Word輸出
Dim m_WordVar() As String '4 x N 1=變數名稱, 2=輸出內容, 3=非變數(Y), 4=旗標名稱(從此處開始尋找)
Dim m_iSizeD2 As Integer 'Word變數數量
'end 2022/1/25
Dim m_bolText1SetFocus As Boolean, m_objControl As Control 'Added by Morgan 2022/3/17
'Modified by Lydia 2022/09/06 改抓特殊設定
'Private Const cnt應收帳款檢查排除 As String = "74018,70005" 'Added by Lydia 2022/06/15 應收帳款上限檢查排除特定人員: 如果人員有異動, 請一併修改接洽單frm090801和收文frm010004~frm010007
Dim m_ChkAmtExcept As String   'cnt應收帳款檢查排除=>m_ChkAmtExcept（程式碼用取代的）
Dim dblPrevRow As Double
'Add By Sindy 2022/9/6
Dim m_strCaseCPM As String, m_bolNewCase As Boolean, m_strNewCP10 As String, m_bolHad10Point As Boolean, m_dblTotOFee As Double, m_dblTotRvFee As Double, m_dbltotPoint As Double
Dim ff1 As Integer
Dim m_Msgbox As String
Public m_SignFlowEmp As String '簽核人員(因有可能人員休假職代代為操作)
Dim m_F0316 As String, m_F0307 As String, m_F0308 As String, m_F0309 As String
'2022/9/6
Dim m_strTransCase As String 'Add By Sindy 2023/3/24

'Add By Sindy 2014/6/3
Public Sub SetParent(ByRef fm As Form)
   Set m_PrevForm = fm
End Sub

'檢查減免身分
'edit by nick 2004/08/05
Private Sub setCP811()
   Dim stCustID As String
   Dim stAD10 As String, stAD15 As String, stAD16 As String 'Added by Morgan 2013/4/2
   Dim stCU15 As String 'Added by Morgan 2019/4/12

   'Added by Morgan 2013/4/2
   arrAD1516(1, 0) = ""
   arrAD1516(1, 1) = ""
   arrAD1516(1, 2) = ""
   arrAD1516(1, 3) = "" 'Added by Morgan 2019/4/12
   arrAD1516(1, 4) = "" 'Added by Morgan 2019/4/12
   arrAD1516(1, 5) = "" 'Added by Morgan 2019/4/15
   arrAD1516(1, 6) = "" 'Added by Morgan 2019/4/15
   arrAD1516(1, 7) = "" 'Added by Morgan 2019/9/25
   'end 2013/4/2

On Error GoTo ErrHnd
   
      optCP811(0).Value = 0: optCP811(1).Value = 0
      If stCountry <> "" And stCustNo1 <> "" Then
         'edit by nickc 2005/04/07 皆以客戶個人為主
         'Modified by Morgan 2013/4/2+stAD10, stAD15, stAD16
         stCustID = PUB_GetAD03(stCustNo1, stCountry, stAD10, stCU15, stAD15, stAD16)
         arrAD1516(1, 4) = stCU15 'Added by Morgan 2019/4/23 日本案減免資格選項要用
         'end 2019/4/23
         If stCustID = "Y" Then
            optCP811(0).Value = 1
            'Added by Morgan 2013/4/2
            '台灣中小企業可減免紀錄原設定資格並設為可更改
            arrAD1516(1, 0) = "Y" 'Modified by Morgan 2019/4/12 +日本案也可減免(不限定中小企業，個人/學校也要設定資格)
            'Modified by Morgan 2019/4/12 +日本案也可減免(不限定中小企業，個人/學校也要設定資格)
            'If stCountry = "000" And stAD10 = "3" Then
            If (stCountry = "000" And stAD10 = "3") Or stCountry = "011" Then
               arrAD1516(1, 1) = stAD15
               arrAD1516(1, 2) = stAD16
               'Added by Morgan 2019/4/15
               arrAD1516(1, 3) = stAD10
               '記錄原設定已便判斷是否有更改
               arrAD1516(1, 5) = stAD10
               arrAD1516(1, 6) = stAD15
               arrAD1516(1, 7) = stAD16
               'end 2019/4/15
            End If
            'end 2013/4/2
         'edit by nick 2004/08/18
         '有設定過在給值
         ElseIf stCustID = "N" Then
         'EDIT BY NICK 2004/07/30 要給預設值
         'ElseIf stCustID = "N" Then
         'Else
            optCP811(1).Value = 1
            arrAD1516(1, 0) = "N" 'Added by Morgan 2013/4/9
         End If
      End If
   
ErrHnd:
   'Modified by Lydia 2019/08/12 +Titile
   'If Err.Number <> 0 Then MsgBox Err.Description, vbCritical
   If Err.Number <> 0 Then MsgBox Err.Description, vbCritical, "SetCP811"
End Sub

'檢查減免身分
'add by nick 2004/08/05
Private Sub setCP812()

   Dim stCustID As String
   Dim stAD10 As String, stAD15 As String, stAD16 As String 'Added by Morgan 2013/4/2
   Dim stCU15 As String 'Added by Morgan 2019/4/12
   
   'Added by Morgan 2013/4/2
   arrAD1516(2, 0) = ""
   arrAD1516(2, 1) = ""
   arrAD1516(2, 2) = ""
   arrAD1516(2, 3) = "" 'Added by Morgan 2019/4/12
   arrAD1516(2, 4) = "" 'Added by Morgan 2019/4/12
   arrAD1516(2, 5) = "" 'Added by Morgan 2019/4/15
   arrAD1516(2, 6) = "" 'Added by Morgan 2019/4/15
   arrAD1516(2, 7) = "" 'Added by Morgan 2019/9/25
   'end 2013/4/2


On Error GoTo ErrHnd
      
      optCP812(0).Value = 0: optCP812(1).Value = 0
      If stCountry <> "" And stCustNo2 <> "" Then
         'edit by nickc 2005/04/07 皆以客戶個人為主
         'Modified by Morgan 2013/4/2+stAD10, stAD15, stAD16
         stCustID = PUB_GetAD03(stCustNo2, stCountry, stAD10, stCU15, stAD15, stAD16)
         arrAD1516(2, 4) = stCU15 'Added by Morgan 2019/4/23 日本案減免資格選項要用
         If stCustID = "Y" Then
            optCP812(0).Value = 1
            'Added by Morgan 2013/4/2
            arrAD1516(2, 0) = "Y"
            'Modified by Morgan 2019/9/25 修正ad10會被清除問題
            'If (stCountry = "000" And stAD10 = "3") Then
            '   arrAD1516(2, 1) = stAD15
            '   arrAD1516(2, 2) = stAD16
            ''Added by Morgan 2019/4/12
            'ElseIf stCountry = "011" Then
            If (stCountry = "000" And stAD10 = "3") Or stCountry = "011" Then
             'end 2019/9/25
               '記錄原設定已便判斷是否有更改
               arrAD1516(2, 1) = stAD15
               arrAD1516(2, 2) = stAD16
               arrAD1516(2, 3) = stAD10
               arrAD1516(2, 5) = stAD10
               arrAD1516(2, 6) = stAD15
               arrAD1516(2, 7) = stAD16 'Added by Morgan 2019/9/25
            End If
            'end 2013/4/2
            
         'edit by nick 2004/08/18
         '有設定過在給值
         ElseIf stCustID = "N" Then
         'Else
            optCP812(1).Value = 1
            arrAD1516(2, 0) = "N" 'Added by Morgan 2013/4/9
         End If
      End If

ErrHnd:
   'Modified by Lydia 2019/08/12 +Titile
   'If Err.Number <> 0 Then MsgBox Err.Description, vbCritical
   If Err.Number <> 0 Then MsgBox Err.Description, vbCritical, "SetCP812"
End Sub

'檢查減免身分
'add by nick 2004/08/05
Private Sub setCP813()

   Dim stCustID As String
   Dim stAD10 As String, stAD15 As String, stAD16 As String 'Added by Morgan 2013/4/2
   Dim stCU15 As String 'Added by Morgan 2019/4/12
   
   'Added by Morgan 2013/4/2
   arrAD1516(3, 0) = ""
   arrAD1516(3, 1) = ""
   arrAD1516(3, 2) = ""
   arrAD1516(3, 3) = "" 'Added by Morgan 2019/4/12
   arrAD1516(3, 4) = "" 'Added by Morgan 2019/4/12
   arrAD1516(3, 5) = "" 'Added by Morgan 2019/4/15
   arrAD1516(3, 6) = "" 'Added by Morgan 2019/4/15
   arrAD1516(3, 7) = "" 'Added by Morgan 2019/9/25
   'end 2013/4/2


On Error GoTo ErrHnd
   
      optCP813(0).Value = 0: optCP813(1).Value = 0
      If stCountry <> "" And stCustNo3 <> "" Then
         'edit by nickc 2005/04/07 皆以客戶個人為主
         'Modified by Morgan 2013/4/2+stAD10, stAD15, stAD16
         stCustID = PUB_GetAD03(stCustNo3, stCountry, stAD10, stCU15, stAD15, stAD16)
         arrAD1516(3, 4) = stCU15 'Added by Morgan 2019/4/23 日本案減免資格選項要用
         If stCustID = "Y" Then
            optCP813(0).Value = 1
            
            'Added by Morgan 2013/4/2
            arrAD1516(3, 0) = "Y"
            'Modified by Morgan 2019/9/25 修正ad10會被清除問題
            'If stCountry = "000" And stAD10 = "3" Then
            '   arrAD1516(3, 1) = stAD15
            '   arrAD1516(3, 2) = stAD16
            ''Added by Morgan 2019/4/12
            'ElseIf stCountry = "011" Then
            If (stCountry = "000" And stAD10 = "3") Or stCountry = "011" Then
            'end 2019/9/25
               arrAD1516(3, 1) = stAD15
               arrAD1516(3, 2) = stAD16
               arrAD1516(3, 3) = stAD10
               '記錄原減免身分&資格
               arrAD1516(3, 5) = stAD10
               arrAD1516(3, 6) = stAD15
               arrAD1516(3, 7) = stAD16 'Added by Morgan 2019/9/25
            End If
            'end 2013/4/2
            
         'edit by nick 2004/08/18
         '有設定過在給值
         ElseIf stCustID = "N" Then
         'Else
            optCP813(1).Value = 1
            arrAD1516(3, 0) = "N" 'Added by Morgan 2013/4/9
         End If
      End If
  
ErrHnd:
   'Modified by Lydia 2019/08/12 +Titile
   'If Err.Number <> 0 Then MsgBox Err.Description, vbCritical
   If Err.Number <> 0 Then MsgBox Err.Description, vbCritical, "SetCP813"
End Sub

'檢查減免身分
'add by nick 2004/08/05
Private Sub setCP814()

   Dim stCustID As String
   Dim stAD10 As String, stAD15 As String, stAD16 As String 'Added by Morgan 2013/4/2
   Dim stCU15 As String 'Added by Morgan 2019/4/12
   
   'Added by Morgan 2013/4/2
   arrAD1516(4, 0) = ""
   arrAD1516(4, 1) = ""
   arrAD1516(4, 2) = ""
   arrAD1516(4, 3) = "" 'Added by Morgan 2019/4/12
   arrAD1516(4, 4) = "" 'Added by Morgan 2019/4/12
   arrAD1516(4, 5) = "" 'Added by Morgan 2019/4/15
   arrAD1516(4, 6) = "" 'Added by Morgan 2019/4/15
   arrAD1516(4, 7) = "" 'Added by Morgan 2019/9/25
   'end 2013/4/2

On Error GoTo ErrHnd
   
      optCP814(0).Value = 0: optCP814(1).Value = 0
      If stCountry <> "" And stCustNo4 <> "" Then
         'edit by nickc 2005/04/07 皆以客戶個人為主
         'Modified by Morgan 2013/4/2+stAD10, stAD15, stAD16
         stCustID = PUB_GetAD03(stCustNo4, stCountry, stAD10, stCU15, stAD15, stAD16)
         arrAD1516(4, 4) = stCU15 'Added by Morgan 2019/4/23 日本案減免資格選項要用
         If stCustID = "Y" Then
            optCP814(0).Value = 1
            
            'Added by Morgan 2013/4/2
            arrAD1516(4, 0) = "Y"
            'Modified by Morgan 2019/9/25 修正ad10會被清除問題
            'If stCountry = "000" And stAD10 = "3" Then
            '   arrAD1516(4, 1) = stAD15
            '   arrAD1516(4, 2) = stAD16
            ''Added by Morgan 2019/4/12
            'ElseIf stCountry = "011" Then
            If (stCountry = "000" And stAD10 = "3") Or stCountry = "011" Then
            'end 2019/9/25
               arrAD1516(4, 1) = stAD15
               arrAD1516(4, 2) = stAD16
               arrAD1516(4, 3) = stAD10
               '記錄原減免身分&資格
               arrAD1516(4, 5) = stAD10
               arrAD1516(4, 6) = stAD15
               arrAD1516(4, 7) = stAD16 'Added by Morgan 2019/9/25
            End If
            'end 2013/4/2
            
         'edit by nick 2004/08/18
         '有設定過在給值
         ElseIf stCustID = "N" Then
         'Else
            optCP814(1).Value = 1
            arrAD1516(4, 0) = "N" 'Added by Morgan 2013/4/9
         End If
      End If

ErrHnd:
   'Modified by Lydia 2019/08/12 +Titile
   'If Err.Number <> 0 Then MsgBox Err.Description, vbCritical
   If Err.Number <> 0 Then MsgBox Err.Description, vbCritical, "SetCP814"
End Sub

'檢查減免身分
'add by nick 2004/08/05
Private Sub setCP815()

   Dim stCustID As String
   Dim stAD10 As String, stAD15 As String, stAD16 As String 'Added by Morgan 2013/4/2
   Dim stCU15 As String 'Added by Morgan 2019/4/12
   
   'Added by Morgan 2013/4/2
   arrAD1516(5, 0) = ""
   arrAD1516(5, 1) = ""
   arrAD1516(5, 2) = ""
   arrAD1516(5, 3) = "" 'Added by Morgan 2019/4/12
   arrAD1516(5, 4) = "" 'Added by Morgan 2019/4/12
   arrAD1516(5, 5) = "" 'Added by Morgan 2019/4/15
   arrAD1516(5, 6) = "" 'Added by Morgan 2019/4/15
   arrAD1516(5, 7) = "" 'Added by Morgan 2019/9/25
   'end 2013/4/2
   
On Error GoTo ErrHnd
   
      optCP815(0).Value = 0: optCP815(1).Value = 0
      If stCountry <> "" And stCustNo5 <> "" Then
         'edit by nickc 2005/04/07 皆以客戶個人為主
         'Modified by Morgan 2013/4/2+stAD10, stAD15, stAD16
         stCustID = PUB_GetAD03(stCustNo5, stCountry, stAD10, stCU15, stAD15, stAD16)
         arrAD1516(5, 4) = stCU15 'Added by Morgan 2019/4/23 日本案減免資格選項要用
         If stCustID = "Y" Then
            optCP815(0).Value = 1
            'Added by Morgan 2013/4/2
            arrAD1516(5, 0) = "Y"
            'Modified by Morgan 2019/9/25 修正ad10會被清除問題
            'If stCountry = "000" And stAD10 = "3" Then
            '   arrAD1516(5, 1) = stAD15
            '   arrAD1516(5, 2) = stAD16
            ''Added by Morgan 2019/4/12
            'ElseIf stCountry = "011" Then
            If (stCountry = "000" And stAD10 = "3") Or stCountry = "011" Then
             'end 2019/9/25
               arrAD1516(5, 1) = stAD15
               arrAD1516(5, 2) = stAD16
               arrAD1516(5, 3) = stAD10
               '記錄原減免身分&資格
               arrAD1516(5, 5) = stAD10
               arrAD1516(5, 6) = stAD15
               arrAD1516(5, 7) = stAD16 'Added by Morgan 2019/9/25
            End If
            'end 2013/4/2
            
         'edit by nick 2004/08/18
         '有設定過在給值
         ElseIf stCustID = "N" Then
         'Else
            optCP815(1).Value = 1
            arrAD1516(5, 0) = "N" 'Added by Morgan 2013/4/9
         End If
      End If
   
ErrHnd:
   'Modified by Lydia 2019/08/12 +Titile
   'If Err.Number <> 0 Then MsgBox Err.Description, vbCritical
   If Err.Number <> 0 Then MsgBox Err.Description, vbCritical, "SetCP815"
End Sub

'Add by Morgan 2008/9/16
Private Sub cboContact_Change(Index As Integer)
   'Modify By Sindy 2010/6/21
   If m_blnCallPrint = True Then
      '外部呼叫列印
      '不執行下段else程式
   '2010/6/21 End
   Else
      If cboContact(Index).Tag <> "1" Then
         If cboContact(Index).ListIndex = -1 Then 'Added by Morgan 2022/1/21 2.0 點選也會觸發,增加判斷 .ListIndex = -1
            cboContact(Index) = ""
         End If
      End If
   End If
End Sub

Private Sub cboContact_Click(Index As Integer)
   SetAddress Index
End Sub

Private Sub SetAddress(p_index As Integer)
   Dim stCU01 As String, stCU02 As String, stContNo As String
   Dim oText1 As Object, oText2 As Object, oText3 As Object
   'Modified by Morgan 2022/1/20 改2.0
   'stContNo = Format(cboContact(p_index).ItemData(cboContact(p_index).ListIndex), "00")
   stContNo = Format(PUB_GetItemData(m_strContactList(p_index), cboContact(p_index).ListIndex), "00")
   'end 2022/1/20
   Select Case p_index
      Case 1
         Set oText1 = Text1(12)
         Set oText2 = Text1(25)
         Set oText3 = Text1(26)
      Case 2
         Set oText1 = Text1(28)
         Set oText2 = Text1(41)
         Set oText3 = Text1(42)
      Case 3
         Set oText1 = Text1(44)
         Set oText2 = Text1(57)
         Set oText3 = Text1(58)
      Case 4
         Set oText1 = Text1(60)
         Set oText2 = Text1(73)
         Set oText3 = Text1(74)
      Case 5
         Set oText1 = Text1(76)
         Set oText2 = Text1(89)
         Set oText3 = Text1(90)
   End Select
   If oText1.Text <> "" Then
      stCU01 = Left(oText1.Text, 8)
      stCU02 = Mid(oText1.Text, 9, 1)
      'Modify by Amy 2016/01/08 +IsPcc 是否抓potcustcont for 臺灣地址判斷
      strExc(0) = "select nvl(pcc21,cu30),nvl(pcc22,cu31),Nvl(pcc21,'Y') as IsPcc from customer,potcustcont where cu01='" & stCU01 & "' and cu02='" & stCU02 & "' and pcc01(+)=cu01 and pcc02(+)='" & stContNo & "'"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         oText2 = "" & RsTemp.Fields(0)
         oText3 = "" & RsTemp.Fields(1)
      End If
   End If
End Sub

'Add by Amy 2016/06/24
Private Sub cboTitle_Change()
    '修改cboTitle欄清除Check9特殊收據的CheckBox
    '特殊收據frm090801_7的欄位也一併清除
    'Modify by  Amy 2020/02/14 +if
    If bolNotClsVal = False Then
        Check9.Value = 0
        m_stCRL01 = ""
        m_stCRL97 = ""
        m_stCRL98 = ""
        m_stCRL99 = ""
        m_stCRL100 = ""
        m_stCRL101 = ""
        m_stCRL102 = ""
        m_stCRL103 = ""
        m_stCRL104 = ""
        m_stCRL105 = ""
        m_stCRL106 = ""
        m_stCRL107 = ""
        m_stCRL108 = ""
        m_stCRL109 = ""
        m_stCRL110 = ""
        m_stCRL111 = ""
        m_stCRL112 = ""
        m_stCRL113 = ""
        m_stCRL114 = ""
        m_stCRL115 = ""
        m_stCRL116 = ""
        m_stCRL117 = ""
        m_stCRL118 = ""
        m_stCRL120 = ""
        m_stCRL121 = ""
        m_stCRL122 = ""
        m_stCRL123 = ""
        m_stCRL124 = ""
        m_stCRL126 = ""
        m_stCRL127 = ""
        m_stCRL128 = ""
        m_stCRL129 = ""
        m_stCRL130 = ""
        m_stCRL131 = ""
        m_stCRL132 = ""
    End If
End Sub

'Add By Sindy 2024/11/13
Private Sub cboTitle_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   cboTitle.ToolTipText = cboTitle.Text
End Sub

'Add By Sindy 2013/1/14
Private Sub cboTitle_Validate(Cancel As Boolean)
   If cboTitle.Text = "" Then Exit Sub
   cboTitle.Text = PUB_StringFilter(cboTitle.Text) 'Add By Sindy 2017/12/5 去掉跳行符號
   If Not CheckLengthIsOK(cboTitle, 100) Then
      Cancel = True
   End If
End Sub

'Add By Sindy 2022/11/4
Private Sub Check11_Click()
   If Check11.Value = 1 Then
      Check11.BackColor = &H8080FF '紅色
   Else
      Check11.BackColor = &H8000000F '灰色
   End If
End Sub
Private Sub Check12_Click()
   If Check12.Value = 1 Then
      Check12.BackColor = &H80FFFF '黃色
   Else
      Check12.BackColor = &H8000000F '灰色
   End If
End Sub

Private Sub ChkCRL66_Click()
   If ChkCRL66.Value = 1 Then
      ChkCRL66.BackColor = &HC000& '綠色
   Else
      ChkCRL66.BackColor = &H8000000F '灰色
   End If
End Sub

'Add By Sindy 2012/11/12
Private Sub Check2_LostFocus()
   m_strCaseCPM = GetAllCaseCPM(, , , , , m_dblTotRvFee) 'Add By Sindy 2022/9/9 取得案件性質代碼
   
   'If m_blnCallPrint = True Then Exit Sub 'Add By Sindy 2014/7/25
   'If Val(Text1(101)) = 0 And Val(Text1(104)) = 0 And Val(Text1(107)) = 0 And Val(Text1(110)) = 0 Then Exit Sub
   If Val(m_dblTotRvFee) = 0 Then Exit Sub
   
   '收據暫不列印時
   If Check2.Value = 1 Then
      '系統預設為1.送件日
      If Check8(0).Value = 0 And Check8(1).Value = 0 And Check8(2).Value = 0 Then
         Check8(0).Value = 1
         'Modify By Sindy 2023/4/18
         Check8(1).Value = 0
         Check8(2).Value = 0
         '2023/4/18 END
      End If
      arrNation = Split(Me.Combo1(0).Text, " ") '申請國家
      '台灣案,則為送件日
      If arrNation(0) = "000" Then
         Check8(0).Value = 1
         'Modify By Sindy 2023/4/18
         Check8(1).Value = 0
         Check8(2).Value = 0
         '2023/4/18 END
      End If
   Else
      Check8(0).Value = 0
      Check8(1).Value = 0
      Check8(2).Value = 0
   End If
End Sub

'Add By Sindy 2011/6/7
Private Sub Check3_Click(Index As Integer)
   If Check3(Index).Value = 1 Then
      If InStr(Text1(127).Text, Trim(Check3(Index).Caption)) = 0 Then
         If Text1(127).Text = "" Then
            Text1(127).Text = Trim(Check3(Index).Caption)
         Else
            Text1(127).Text = Text1(127).Text & "," & Trim(Check3(Index).Caption)
         End If
      End If
   Else
      '案件屬性=xx,xx,xx
      If Left(Text1(127), Len(Trim(Check3(Index).Caption))) = Trim(Check3(Index).Caption) Then
         Text1(127).Text = Replace(Text1(127).Text, Trim(Check3(Index).Caption) & ",", "")
         Text1(127).Text = Replace(Text1(127).Text, Trim(Check3(Index).Caption), "")
      Else
         Text1(127).Text = Replace(Text1(127).Text, "," & Trim(Check3(Index).Caption), "")
      End If
   End If
End Sub

Private Sub Check4_Click()
   If Check4.Value = 1 Then
      'Modify By Sindy 2015/1/26 經理說生醫案大家均可做電子送件
      'Modify By Sindy 2014/1/10 若智權人員為專利處人員時, 開放生醫案可勾選電子送件
'      If Not (PUB_GetStaffST15(Text1(10), "1") >= "P10" And PUB_GetStaffST15(Text1(10), "1") <= "P14") _
'         And Left(Combo5, 1) = "3" Then
'         MsgBox "生醫案不可電子送件！", vbExclamation
'         Check4.Value = 0
'      End If
   End If
End Sub

'2011/4/22 ADD BY SONIA 現金或支票計算預定收款日
Private Sub Check6_Click(Index As Integer)
   If Check6(0) = 1 Or Check6(1) = 1 Then
      If PUB_GetST06(Text1(10)) <> "1" Then
         Text1(118) = PUB_GetWorkDayAfterSysDate(CDbl(strSrvDate(1)), 6)
      Else
         Text1(118) = PUB_GetWorkDayAfterSysDate(CDbl(strSrvDate(1)), 5)
      End If
      Text1(118).Locked = True
      Text1(118).Enabled = False
   Else
      Text1(118).Locked = False
      Text1(118).Enabled = True
   End If
End Sub
'2011/4/22 END

'Add By Sindy 2013/11/25
Private Sub Check8_GotFocus(Index As Integer)
   Check8(0).Value = 0
   Check8(1).Value = 0
   Check8(2).Value = 0
End Sub

'Add By Sindy 2014/2/6
Private Sub Check9_Click()
   'Modify by Amy 2016/07/07 +勾DEBIT NOTE 請款,不可勾選特殊收據控制
   If Check9.Value = 1 And Option2(2).Value = False Then
      cmdCRL119.Visible = True
      If m_blnCallPrint = False And Check9.Tag <> "NotShow" Then
         Call cmdCRL119_Click
      End If
      Check9.Tag = ""
   Else
      cmdCRL119.Visible = False
   End If
End Sub

Private Sub ChkAddress_Click(Index As Integer)
Dim i As Integer
  
If ChkAddress(Index).Value = 1 Then
   If InStr(Text1(12), Text3(Index).Tag) > 0 Then
      Text4(Index).Text = "發明人地址同申請人地址"
   ElseIf InStr(Text1(28), Text3(Index).Tag) > 0 Then
      Text4(Index).Text = "發明人地址同申請人地址"
   ElseIf InStr(Text1(44), Text3(Index).Tag) > 0 Then
      Text4(Index).Text = "發明人地址同申請人地址"
   ElseIf InStr(Text1(60), Text3(Index).Tag) > 0 Then
      Text4(Index).Text = "發明人地址同申請人地址"
   ElseIf InStr(Text1(76), Text3(Index).Tag) > 0 Then
      Text4(Index).Text = "發明人地址同申請人地址"
   End If
Else
   Text4(Index).Text = Text4(Index).Tag
End If
End Sub

'Added by Morgan 2013/1/15
Private Sub chkItem_Click(Index As Integer)
   Dim ii As Integer
   
   'Add By Sindy 2023/3/23 Form Load Read Error : Mark
   If Not Me.ActiveControl Is Nothing Then
   '2023/3/23 END
      If Me.ActiveControl <> chkItem(Index) Then Exit Sub
   End If
   
   txtItemCount.Enabled = False
   txtItemList.Enabled = False
   txtYear(0).Enabled = False
   txtYear(1).Enabled = False
   txtMonth(0).Enabled = False
   txtMonth(1).Enabled = False
   txtDay(0).Enabled = False
   txtDay(1).Enabled = False
   
   If Index = 0 Or Index = 1 Then
      If chkItem(Index).Value = vbChecked Then
         For Each oChk In chkItem
            If oChk.Index <> Index Then
               oChk.Value = vbUnchecked
            End If
         Next
         
         Select Case Index
         Case 0
            txtItemCount.Enabled = True
'            txtItemCount.SetFocus
         Case 1
            txtItemList.Enabled = True
'            txtItemList.SetFocus
            If Left(txtItemList, 1) = "第" Then
               txtItemList.SelStart = 1
               txtItemList.SelLength = 0
            End If
         End Select
      End If
   ElseIf Index = 6 Then
      If chkItem(Index).Value = vbChecked Then
         For Each oChk In chkItem
            If oChk.Index <> Index Then
               oChk.Value = vbUnchecked
            End If
         Next
         txtYear(0).Enabled = True
'         txtYear(0).SetFocus
         txtYear(1).Enabled = True
         txtMonth(0).Enabled = True
         txtMonth(1).Enabled = True
         txtDay(0).Enabled = True
         txtDay(1).Enabled = True
      End If
   ElseIf chkItem(Index).Value = vbChecked Then
      chkItem(0).Value = vbUnchecked
      chkItem(1).Value = vbUnchecked
      chkItem(6).Value = vbUnchecked
   End If
End Sub

'Add By Sindy 2014/7/18 新增附件
'Modify By Sindy 2022/10/11
Private Sub cmdAddAtt_Click()
   '若為舊案
   If Me.Option1(1).Value = True And Me.Text1(6).Text <> "" And Me.Text1(7).Text <> "" Then
       If Me.Text1(8).Text = "" Then Me.Text1(8).Text = "0"
       If Me.Text1(9).Text = "" Then Me.Text1(9).Text = "00"
   End If
   
   Call frm090801_13.SetParent(Me)
   If Option1(0).Value = True Then
      frm090801_13.bolNewCase = True '是新案
   Else
      frm090801_13.bolNewCase = False
   End If
   If Text5.Visible = True Then
      frm090801_13.m_strCRL01 = Text5
   Else
      frm090801_13.m_strCRL01 = ""
   End If
   frm090801_13.m_blnCallQuery = m_blnCallPrint '查詢
   frm090801_13.strCaseNA239 = strCaseNA239
   frm090801_13.lblCaseNo = Text1(6) & "-" & Text1(7) & "-" & Text1(8) & "-" & Text1(9)
   frm090801_13.m_strSaveFiles = Me.m_strSaveFiles
   frm090801_13.m_strSaveFiles2 = Me.m_strSaveFiles2
   If UCase(TypeName(m_PrevForm)) = UCase("frm210148") Then
      frm090801_13.cmdAddAtt(0).Enabled = False
      frm090801_13.cmdAddAtt(1).Enabled = False
      frm090801_13.cmdRemAtt(0).Enabled = False
      frm090801_13.cmdRemAtt(1).Enabled = False
   End If
   Call frm090801_13.QueryData(True)
   frm090801_13.Show vbModal
   If Me.m_strSaveFiles = "" And Me.m_strSaveFiles2 = "" Then
      Me.cmdAddAtt.BackColor = &H808080 '灰色
   Else
      Me.cmdAddAtt.BackColor = &HC0C0FF '粉紅色
   End If
End Sub

'Add By Sindy 2014/2/6
Private Sub cmdCRL119_Click()
   frm090801_7.SetParent Me
   'Add by Amy 2016/09/01
   frm090801_7.stCaseNo1 = Text1(6)
   frm090801_7.stCaseNo2 = Text1(7)
   frm090801_7.stCaseNo3 = Text1(8)
   frm090801_7.stCaseNo4 = Text1(9)
   'end 2016/09/01
   frm090801_7.m_stCRL01 = m_stCRL01
   frm090801_7.m_stCRL97 = m_stCRL97
   frm090801_7.m_stCRL98 = m_stCRL98
   frm090801_7.m_stCRL99 = m_stCRL99
   frm090801_7.m_stCRL100 = m_stCRL100
   frm090801_7.m_stCRL101 = m_stCRL101
   frm090801_7.m_stCRL102 = m_stCRL102
   frm090801_7.m_stCRL103 = m_stCRL103
   frm090801_7.m_stCRL104 = m_stCRL104
   frm090801_7.m_stCRL105 = m_stCRL105
   frm090801_7.m_stCRL106 = m_stCRL106
   frm090801_7.m_stCRL107 = m_stCRL107
   frm090801_7.m_stCRL108 = m_stCRL108
   frm090801_7.m_stCRL109 = m_stCRL109
   frm090801_7.m_stCRL110 = m_stCRL110
   frm090801_7.m_stCRL111 = m_stCRL111
   frm090801_7.m_stCRL112 = m_stCRL112
   frm090801_7.m_stCRL113 = m_stCRL113
   frm090801_7.m_stCRL114 = m_stCRL114
   frm090801_7.m_stCRL115 = m_stCRL115
   frm090801_7.m_stCRL116 = m_stCRL116
   frm090801_7.m_stCRL117 = m_stCRL117
   frm090801_7.m_stCRL118 = m_stCRL118
   frm090801_7.m_stCRL120 = m_stCRL120
   frm090801_7.m_stCRL121 = m_stCRL121
   frm090801_7.m_stCRL122 = m_stCRL122
   frm090801_7.m_stCRL123 = m_stCRL123
   frm090801_7.m_stCRL124 = m_stCRL124
   frm090801_7.m_stCRL126 = m_stCRL126
   frm090801_7.m_stCRL127 = m_stCRL127
   frm090801_7.m_stCRL128 = m_stCRL128
   frm090801_7.m_stCRL129 = m_stCRL129
   frm090801_7.m_stCRL130 = m_stCRL130
   frm090801_7.m_stCRL131 = m_stCRL131
   frm090801_7.m_stCRL132 = m_stCRL132
   'Modify By Sindy 2019/6/24 隱藏的順序有差
   'Modify By Sindy 2023/11/27 因卷宗區查詢接洽單的特殊收據會出現錯誤
   '                           改MDIChild屬性=False
   'frm090801_7.Show 'vbModal
   'Modify By Sindy 2024/3/26 因收據開立作業是非強制表單的形式開啟特殊收據,若已開故,再按下鍵時增加此判斷
   If PUB_CheckFormExist("frm090801_7") = True Then
      frm090801_7.Show
   Else
   '2024/3/26 END
      frm090801_7.Show vbModal
   End If
   '2023/11/27 END
   frm090801_7.ZOrder
   'If InStr(TypeName(m_PrevForm), "Frmacc112") = 0 Then Me.Hide
   '2019/6/24 END
End Sub

'Added by Morgan 2023/7/6
'下載商標圖檔
Private Sub cmdSavePic_Click()
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
   
   strExc(0) = "Select crif05,crl07||crl08||decode(crl10,'00',decode(crl09,'0','','-'||crl09),'-'||crl09||'-'||crl10) CNO " & _
                 "From consultrecimagef,CONSULTRECORDLIST " & _
               "Where crif01 ='" & Trim(Text5) & "' and crl01(+)=crif01"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      stFileName = stFolderPath & RsTemp("CNO") & ".jpg"
      If Dir(stFileName) <> "" Then
         If MsgBox("[ " & stFileName & " ]圖檔已存在，是否要覆蓋？", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then
            Exit Sub
         End If
      End If
      If PUB_GetFtpFile(RsTemp.Fields("crif05"), stFileName, UCase("consultrecimagef")) = True Then
         MsgBox "圖檔已下載[ " & stFileName & " ]。", vbInformation
      Else
         MsgBox "圖檔下載失敗！", vbCritical
      End If
   Else
      MsgBox "圖檔讀取失敗！", vbCritical
   End If
End Sub

'Add By Sindy 2022/12/16
Private Sub SSTab1_KeyDown(KeyCode As Integer, Shift As Integer)
   'Debug.Print KeyCode
   '39:右 38:上 40:下 37:左
   If SSTab1.Tab = 1 Then
      If KeyCode = 40 Then
         SSTab2.SetFocus
      End If
   Else
      If KeyCode = 40 Then
         SSTab3.SetFocus
      End If
   End If
End Sub
Private Sub SSTab2_KeyDown(KeyCode As Integer, Shift As Integer)
   'Debug.Print KeyCode
   '39:右 38:上 40:下 37:左
   If KeyCode = 38 Then
      SSTab1.SetFocus
   End If
End Sub
Private Sub SSTab3_KeyDown(KeyCode As Integer, Shift As Integer)
   'Debug.Print KeyCode
   '39:右 38:上 40:下 37:左
   If KeyCode = 38 Then
      SSTab1.SetFocus
   End If
End Sub
'2022/12/16 END

Public Sub cmdok_Click(Index As Integer)
    Select Case Index
    Case 1 '結束
         Unload Me
         
    'Add By Sindy 2010/6/21
    Case 4 '查詢
         Screen.MousePointer = vbHourglass
         Call cmdClear(True)
         '***** 一定要下,強制表單不能再呼叫非強制表單 *****
'         Check7(0).Tag = "NotShow"
         Check9.Tag = "NotShow"
         '*****
         Call QueryData
         If m_blnCallPrint = True Then
            Call SetCtrlReadOnly(True) 'Add By Sindy 2022/9/21
            Call SetCtrlReadOnly_Flow(True) 'Add By Sindy 2022/9/21
            
            cmdOK(0).Visible = False '新增/修改
            cmdOK(1).Visible = True: cmdOK(1).Enabled = True '結束
            cmdOK(2).Visible = False '清除畫面
            If InStr(UCase(Text1(6).Text), "L") > 0 And PUB_ChkLCompStaff(Text1(10).Text) = False Then
               cmdOK(5).Caption = "案源(&I)"
               cmdOK(5).Visible = True
            Else
               cmdOK(5).Visible = False
            End If
         Else
            If InStr(UCase(Text1(6).Text), "L") > 0 And PUB_ChkLCompStaff(Text1(10).Text) = False Then
               Call SetCtrlReadOnly(True)
               'Call SetCtrlReadOnly_Flow(True)
            Else
               '某些欄位應該鎖住
               Text1(6).Enabled = False
               Text1(7).Enabled = False
               Text1(8).Enabled = False
               Text1(9).Enabled = False
               Option1(0).Enabled = False: Option1(1).Enabled = False
               'cmdDel.Enabled = False
               'cmdClear2.Enabled = False
            End If
            SrcSetButton 'Added by Morgan 2020/4/17
         End If
         Screen.MousePointer = vbDefault
         
    Case 5 '查詢案源
         Screen.MousePointer = vbDefault
         Set frm090801_11.frmParent = Me
         If strLOS15 = "" Then strLOS15 = SrcGetLOS15(Text5)
         frm090801_11.strLOS15 = strLOS15
         frm090801_11.SetReadOnly
         frm090801_11.Show vbModal
         Exit Sub
    End Select
End Sub

'Modify By Sindy 2010/6/21
'Modify by Morgan 2011/3/28 +pbolPreserve
Private Sub cmdClear(Optional pbolPreserve As Boolean = False)
   Dim i As Integer 'Add by Amy 2015/10/22
   
   Call SetCtrlReadOnly(False) 'Add By Sindy 2022/12/25
   Call SetCtrlReadOnly_Flow(False) 'Add By Sindy 2022/12/25
'   m_stCRL134 = "": m_stCRL135 = ""
'   cmdCRL134.Visible = False
   txtCRL69.Text = "": txtCRL70.Text = ""
   GridCase.Tag = ""
'   cmdPic.Tag = ""
   mTQC01 = "" 'Add By Sindy 2022/10/12
   lblAPPLQ.Visible = False
   lblZip.Visible = False
   chkEnglish.Value = 0
   
   ClearAll pbolPreserve
'        If Me.Text1(10).Text <> "" Then
'            If MsgBox("是否清除員工編號???", vbExclamation + vbYesNo) = vbYes Then

   'txtEnabled1 True 'Removed by Morgan 2011/11/3
   
   If pbolPreserve = False Then 'Add by Morgan 2011/3/28 考慮輸多張接洽單只有改案號情形
      txtEnabled1 True 'Add by Morgan 2011/11/3 從外面搬進來,否則舊客戶會變可改編號 Ex.P79422
      Me.Text1(0).Text = ""
      Me.lblZone.Caption = ""
      'Modify By Sindy 2012/5/30 Mark
'      Me.Text1(10).Text = ""
'      Me.lblStaffName.Caption = ""
   End If
'            End If
'        End If
   
   'Modify By Sindy 2023/11/16
   'Me.SSTab2.Tab = 0
   If SSTab2.TabVisible(0) = True Then Me.SSTab2.Tab = 0
   '2023/11/16 END
   Me.SSTab1.Tab = 0
   'Add By Sindy 2009/08/31
   opt1(0).Value = False
   opt1(1).Value = False
   opt1(2).Value = False
   Set G_SeekPicColor.Picture = LoadPicture()
   Set tmpImg.Picture = LoadPicture()
   Set tmpPic.Picture = LoadPicture()
   Check7(0).Value = 0
   Check7(1).Value = 0
   '2009/08/31 End
   
   Option1(0).Enabled = True: Option1(1).Enabled = True 'Add By Sindy 2022/9/15
   '預設為新案
   If pbolPreserve = False Then 'Add by Morgan 2011/3/28 考慮輸多張接洽單只有改案號情形
      Me.Option1(0).Value = True
      Me.Text1(6).Enabled = True 'Add By Sindy 2022/9/15
      Me.Text1(7).Enabled = False
      Me.Text1(8).Enabled = False
      Me.Text1(9).Enabled = False
      'Modified by Lydia 2021/06/18 debug:接洽單查詢及列印使用 + And Me.Visible = True
      If Me.Text1(10).Enabled = True And Me.Visible = True Then Me.Text1(10).SetFocus
   End If
   
   'Added by Lydia 2023/11/13 DEBIT NOTE請款選項
   optDB(0).Value = 0: optDB(1).Value = 0
   Check7(2).Value = 0: Check7(3).Value = 0
   Frame33(1).BackColor = &H8000000F
   Frame33(1).Visible = False
   'end 2023/11/13
   
   'Add by Morgan 2011/4/7
   m_Note1 = "": m_Note2 = ""
   m_strGetNP01 = "" 'Add By Sindy 2015/9/17
   Text1(142).Text = "": Text1(143).Text = "" 'strYF05From = "": strYF05To = ""
   'cmdOK(3).Caption = "期限資料(&L)" 'Add By Sindy 2015/4/2
   'Add by Lydia 2014/12/22
   mPYFee = False
   Frame605.Visible = False
   
   'Add by Amy 2016/09/19
   m_strTM15 = ""
   m_strTM12 = ""
   'Add By Sindy 2014/2/6
   Check9.Value = 0: cmdCRL119.Visible = False
   m_stCRL01 = ""
   m_stCRL97 = "": m_stCRL118 = ""
   m_stCRL98 = "": m_stCRL99 = "": m_stCRL100 = "": m_stCRL101 = ""
   m_stCRL102 = "": m_stCRL103 = "": m_stCRL104 = "": m_stCRL105 = ""
   m_stCRL106 = "": m_stCRL107 = "": m_stCRL108 = "": m_stCRL109 = ""
   m_stCRL110 = "": m_stCRL111 = "": m_stCRL112 = "": m_stCRL113 = ""
   m_stCRL120 = "": m_stCRL121 = "": m_stCRL122 = "": m_stCRL123 = ""
   '2014/2/6 END
   'Add By Sindy 2015/8/28
   m_stCRL114 = "": m_stCRL115 = "": m_stCRL116 = ""
   m_stCRL117 = "": m_stCRL124 = "": m_stCRL126 = ""
   m_stCRL127 = "": m_stCRL128 = "": m_stCRL129 = ""
   m_stCRL130 = "": m_stCRL131 = "": m_stCRL132 = ""
   '2015/8/28 END
   m_lstCaseNo = "" 'add by sonia 2015/9/15 自動收文已清畫面,再輸第二筆案號又會詢問是否清除畫面,故加入此
   'Added by Lydia 2015/10/14
   cmdTMQ.Tag = ""
   GrdTMQ.Clear 'Added by Lydia 2016/04/18
   SetGrdTMQ
   TMQList = ""
   pTMQList = "" 'Added by Lydia 2016/05/05
   
   'Modified by Morgan 2020/4/28
   SrcSetButton
   bolPrintNewCase = False 'Added by Morgan 2020/6/3
   
   ClearField_Flow 'Add By Sindy 2022/9/19
End Sub

'Add By Sindy 2022/8/24 檢查是否有接洽單未送簽核資料
Private Function SrcGetCRLFlow() As Integer
   Dim stSQL As String, intQ As Integer
   Dim RsQ As ADODB.Recordset
   
On Error GoTo ErrHnd
   
   SrcGetCRLFlow = 0
   stSQL = "select CRL01 From ConsultRecordList,ConsultRecCMP,flow003" & _
           " where crl02>=" & 接洽單電子收文啟用日 & " and CRL01=f0301(+) and f0309 is null" & _
           " and CRL01=crc01(+) and crc02 is not null" & _
           " and crl78='" & strUserNum & "'" & _
           " group by CRL01"
   intQ = 1
   Set RsQ = ClsLawReadRstMsg(intQ, stSQL)
   If intQ = 1 Then
      SrcGetCRLFlow = RsQ.RecordCount
   End If
   
ErrHnd:
   Set RsQ = Nothing
End Function

Private Sub Combo1_Click(Index As Integer)
    Select Case Index
    Case 1 ', 2, 3, 4
        If Val(lblCnt.Caption) > 0 And Combo1(Index).Text = Combo1(Index).Tag Then Exit Sub 'Add by Sindy 2022/8/29 修改相同案件性質,不需清空欄位值
        Me.Text1(101).Text = "" ' + (Index - 1) * 3
        Me.Text1(102).Text = "" '+ (Index - 1) * 3
        Me.Text1(103).Text = "" ' + (Index - 1) * 3
        Me.Combo2(Index - 1).Text = "" 'Add by Sindy 2022/8/29
         'Modified by Morgan 2020/4/17
         SrcSetButton
    'Added by Morgan 2013/3/19
    Case 0
      SetNewDrug 'Added by Morgan 2021/7/20
    End Select
End Sub

'Added by Morgan 2013/4/9
Private Sub SetOpt81(pCountry As String)
   Dim ii As Integer
   Dim bolCP81old As Boolean
   
   bolCP81old = IsoptCP81 'Added by Morgan 2019/4/26
   
   'Modified by Morgan 2016/3/25
   'If pCountry = "000" Or pCountry = "101" Or pCountry = "102" Then
   If (Text1(6) = "P" Or Text1(6) = "CFP") And (pCountry = "000" Or pCountry = "101" Or pCountry = "102") Then
   'end 2016/3/25
      IsoptCP81 = True
   'Added by Morgan 2019/4/12 +日本發明案有收文實審、領證、年費
   ElseIf Text1(6) = "CFP" And (pCountry = "011" And Left(Combo6, 1) = "1") Then
      IsoptCP81 = False
      '領證、年費必須適用減免才可設定(有實審在2019/4/1以後發文)
      m_strCaseCPM = GetAllCaseCPM 'Add By Sindy 2022/8/29 取得案件性質代碼
      'For ii = 1 To 4
         'If Trim(Left(Me.Combo1(ii).Text, 4)) = "416" Then
         If InStr(m_strCaseCPM, "416") > 0 Then
            IsoptCP81 = True
            'Exit For
         'ElseIf Trim(Left(Me.Combo1(ii).Text, 4)) = "601" Then
         ElseIf InStr(m_strCaseCPM, "601") > 0 Then
            IsoptCP81 = PUB_ChkJpDiscount(Text1(6), Text1(7), Right("0" & Text1(8), 1), Right("00" & Text1(9), 2), True)
            'Exit For
         'ElseIf Trim(Left(Me.Combo1(ii).Text, 4)) = "605" Then
         ElseIf InStr(m_strCaseCPM, "605") > 0 Then
            If Val(Text1(142).Text) <= 10 Then
               IsoptCP81 = PUB_ChkJpDiscount(Text1(6), Text1(7), Right("0" & Text1(8), 1), Right("00" & Text1(9), 2), True)
            End If
            'Exit For
         End If
      'Next ii
      
      'Added by Morgan 2019/4/26
      '若狀態有改時需讀取減免身分
      If IsoptCP81 = True And bolCP81old = False Then
         Call setCP811  '設定減免身分
         Call setCP812  '設定減免身分
         Call setCP813  '設定減免身分
         Call setCP814  '設定減免身分
         Call setCP815  '設定減免身分
      End If
      'end 2019/4/26
      
   'end 2019/4/19
   Else
      IsoptCP81 = False
   End If
   
   If IsoptCP81 Then
      optCP811(0).Enabled = True
      optCP811(1).Enabled = True
      optCP812(0).Enabled = True
      optCP812(1).Enabled = True
      optCP813(0).Enabled = True
      optCP813(1).Enabled = True
      optCP814(0).Enabled = True
      optCP814(1).Enabled = True
      optCP815(0).Enabled = True
      optCP815(1).Enabled = True
   Else
      optCP811(0).Enabled = False
      optCP811(1).Enabled = False
      optCP812(0).Enabled = False
      optCP812(1).Enabled = False
      optCP813(0).Enabled = False
      optCP813(1).Enabled = False
      optCP814(0).Enabled = False
      optCP814(1).Enabled = False
      optCP815(0).Enabled = False
      optCP815(1).Enabled = False
      Erase arrAD1516 'Added by Morgan 2016/3/25
   End If
End Sub

'Add By Sindy 2019/6/24
Private Sub Combo1_DropDown(Index As Integer)
   'If Index >= 1 And Index <= 4 Then
      'Modify By Sindy 2024/2/22
      'Call SetComboCase(Index, Combo1(Index).Text)
      Call frm090801_New_SetComboCase(Index, Combo1(Index).Text, Combo1(Index), Me.Text1(6).Text, _
            Combo1(0).Text, ChkPCT, Text1(101), Text1(102), Text1(103))
      '2024/2/22 END
   'End If
End Sub

Private Sub Combo1_LostFocus(Index As Integer)
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
Dim ii As Integer
   
   'If m_blnCallPrint = True Then Exit Sub 'Add By Sindy 2014/7/25
   
   Select Case Index
      Case 0 '申請國家
          If Me.Combo1(Index).Text = "" Then
              'For ii = 1 To 4
                  Me.Combo1(1).Clear
                  GridCase.Clear: Call SetGrd 'Add By Sindy 2022/8/31
                  'Modify By Sindy 2024/2/22
                  'Call SetComboCase(1, "") '設定下拉選單 Add By Sindy 2022/10/14
                  Call frm090801_New_SetComboCase(1, "", Combo1(1), Me.Text1(6).Text, _
                        Combo1(0).Text, ChkPCT, Text1(101), Text1(102), Text1(103))
                  '2024/2/22 END
              'Next ii
          Else
             arrNation = Split(Me.Combo1(Index).Text, " ")
             
            'Added by Morgan 2013/1/15
            If Text1(6) = "P" And arrNation(0) = "000" Then
               SSTab3.TabVisible(1) = True
            Else
               SSTab3.TabVisible(1) = False
            End If
            'end 2013/1/15
               
'Modified by Morgan 2013/4/9 移到下面改呼叫函數
              SetOpt81 "" & arrNation(0)
'end 2013/4/9

              '2011/9/15 add by sonia 系統類別+申請國家有改變才要重新預設案件性質的下拉選單,否則會把由期限資料畫面點選回來的案件性質清除(連續操作多件時)
              If Text1(6).Tag = Text1(6).Text And Combo1(0).Tag = Combo1(0).Text Then
                 Exit Sub
              End If
              '2011/9/15 end
              
              'Add By Sindy 2022/10/6
              'Modify By Sindy 2019/6/24
              'For ii = 1 To 4
                 Me.Combo1(1).Clear
                 GridCase.Clear: Call SetGrd
                 'Modify By Sindy 2024/2/22
                 'Call SetComboCase(1, "") '設定下拉選單
                 Call frm090801_New_SetComboCase(1, "", Combo1(1), Me.Text1(6).Text, _
                        Combo1(0).Text, ChkPCT, Text1(101), Text1(102), Text1(103))
                 '2024/2/22 END
              'Next ii
              '2022/10/6 END
          End If
          Combo1(0).Tag = Combo1(0).Text   '2011/9/15 add by sonia
          
      Case 1 '案件性質, 2, 3, 4
          If Me.Combo1(Index).Text = "" Then
              Me.Text1(101).Text = "" ' + (Index - 1) * 3
              Me.Text1(102).Text = "" ' + (Index - 1) * 3
              Me.Text1(103).Text = "" ' + (Index - 1) * 3
          Else
               If Left(Label1(124).Caption, 4) = "專利種類" Then
                  Call SetCombo5_P 'Add by Amy 2016/06/06 +專利設計案案件性質
               End If
          End If
   End Select
End Sub

'Modify By Sindy 2022/9/9 維持寫法,因檢查時有"重算其他收文"狀況
'Added by Lydia 2021/03/29 ACS案件收文與點數及營業稅：
'案件性質為101~103Z時鎖住規費及點數欄(改案件性質>103Z時要放開)，當輸入費用後，自動依下列規則計算：
Private Sub SetACSautoFee(ByVal tInx As Integer, Optional ByRef strCheck As String)
'tInx：案件性質輸入Combo1的Index

    'Added by Lydia 2021/04/27 因為10%尚未溝通好，先上其他控制
    Dim pRate As Double
    pRate = 0
    'end 2021/04/27

   If Me.Text1(6).Text = "ACS" And tInx >= 1 And tInx <= 4 And Trim(Combo1(tInx)) <> "" Then
        If InStr("101,102,103", Trim(Left(Combo1(tInx), 3))) > 0 Then
           If Me.Text1(102 + (tInx - 1) * 3).Enabled = True Then
              Me.Text1(102 + (tInx - 1) * 3).Enabled = False
              Me.Text1(103 + (tInx - 1) * 3).Enabled = False
           End If
           If Val(Me.Text1(101 + (tInx - 1) * 3)) > 0 Then  '費用>0
               If Combo4.Text = m_CompNameJ Then
                    '甲、J公司時規費＝[費用－(費用/1.05)]+[(費用/1.05)*10%]，點數＝(費用－規費)/1000；例：費用210,000、規費＝10,000+20,000=30,000、點數180；
                    Me.Text1(102 + (tInx - 1) * 3) = (CDbl(Val(Me.Text1(101 + (tInx - 1) * 3))) - Round((CDbl(Val(Me.Text1(101 + (tInx - 1) * 3))) / 1.05), 0)) + Round((CDbl(Val(Me.Text1(101 + (tInx - 1) * 3))) / 1.05) * pRate, 0)
               Else
                    '乙、非J公司時規費=費用*10%，點數＝(費用－規費)/1000；例：費用210,000、規費＝21,000、點數189；
                    Me.Text1(102 + (tInx - 1) * 3) = Round(Val(Me.Text1(101 + (tInx - 1) * 3)) * pRate, 0)
               End If
               Me.Text1(103 + (tInx - 1) * 3) = Format((CDbl(Val(Me.Text1(101 + (tInx - 1) * 3))) - CDbl(Val(Me.Text1(102 + (tInx - 1) * 3)))) / 1000, "####0.000")
               strCheck = ""
           'Added by Lydia 2021/04/30 費用<=0,預設規費=0; ex.ACS-00064 舊案收101,因為先前已有收錢,所以無費用
           Else '包含點數
               Me.Text1(102 + (tInx - 1) * 3) = "0"
               Me.Text1(103 + (tInx - 1) * 3) = "0.000"
               strCheck = ""
           'end 2021/04/30
           End If
        ElseIf Trim(Left(Combo1(tInx), 4)) = "706" Then
            '案件性質706代收代付時，鎖住規費及點數欄，當輸入費用後，規費設定為輸入之費用，點數=0
            If Me.Text1(102 + (tInx - 1) * 3).Enabled = True Then
               Me.Text1(102 + (tInx - 1) * 3).Enabled = False
               Me.Text1(103 + (tInx - 1) * 3).Enabled = False
            End If
            If Val(Me.Text1(101 + (tInx - 1) * 3)) > 0 Then  '費用>0
                Me.Text1(102 + (tInx - 1) * 3) = Me.Text1(101 + (tInx - 1) * 3)
                Me.Text1(103 + (tInx - 1) * 3) = "0.000"
            End If
            strCheck = ""
        ElseIf Trim(Left(Combo1(tInx), 3)) > "103" Then '案件性質>103Z時
            If Combo4.Text = m_CompNameJ Then
                If Me.Text1(102 + (tInx - 1) * 3).Enabled = True Then
                   Me.Text1(102 + (tInx - 1) * 3).Enabled = False
                   Me.Text1(103 + (tInx - 1) * 3).Enabled = False
                End If
                '甲、J公司時鎖住規費及點數欄，當輸入費用後，自動計算規費＝[費用－(費用/1.05)]， 點數＝(費用－規費)/1000；例：費用420,000、規費＝20,000、點數40；
                Me.Text1(102 + (tInx - 1) * 3) = (CDbl(Val(Me.Text1(101 + (tInx - 1) * 3))) - Round((CDbl(Val(Me.Text1(101 + (tInx - 1) * 3))) / 1.05), 0))
                Me.Text1(103 + (tInx - 1) * 3) = Format((CDbl(Val(Me.Text1(101 + (tInx - 1) * 3))) - CDbl(Val(Me.Text1(102 + (tInx - 1) * 3)))) / 1000, "####0.000")
                strCheck = ""
            Else
               '乙、非J公司時如一般接洽單不鎖住規費及點數欄，由人員自行填寫費用、規費、點數欄，但檢查點數＝(費用－規費)/1000；例：費用420,000、規費＝0、點數42；
               GoTo JumpPart01
            End If
        Else
JumpPart01:
           If Me.Text1(102 + (tInx - 1) * 3).Enabled = False Then   '從鎖住規費改成非鎖住時，彈提醒
                If strCheck = "Y" Then  '列印前檢查
                    strCheck = MsgBox("[" & Combo1(tInx) & "]規費和點數請人工輸入!!" & vbCrLf & "是否繼續作業？", vbYesNo + vbDefaultButton2 + vbExclamation, "ACS案件收文")
                    If Val(strCheck) = vbYes Then
                        strCheck = ""
                    End If
                ElseIf strCheck = "A" Then '第2客戶以後有設定為不開發票,已先彈訊息
                    strCheck = ""
                Else
                    MsgBox "[" & Combo1(tInx) & "]規費和點數請人工輸入!!", vbExclamation, "ACS案件收文"
                End If
                Me.Text1(102 + (tInx - 1) * 3).Enabled = True
                Me.Text1(103 + (tInx - 1) * 3).Enabled = True
           Else
                 strCheck = ""
           End If
        End If
   Else
         strCheck = ""
   End If
   
End Sub
'end 2021/03/29

Private Sub Combo1_Validate(Index As Integer, Cancel As Boolean)
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
Dim arrNA01
Dim strMsg As String 'Add by Amy 2016/08/16
    
    Select Case Index
    Case 0 '申請國家
        If Me.Combo1(Index).Text <> "" Then
            arrNA01 = Split(Me.Combo1(Index).Text, " ")
            
'Modified by Morgan 2013/4/9
            SetOpt81 "" & arrNA01(0)
'end 2013/4/9

            'Modify By Sindy 2011/11/11 增加NA59
            StrSQLa = "Select NA01,NA03,NA59 From Nation Where Length(NA01)=3 And NA01<='9999' And substr(NA02,3,1)='0' And NA01='" & arrNA01(0) & "' Order By NA01 "
            rsA.CursorLocation = adUseClient
            rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
            If rsA.RecordCount > 0 Then
                Me.Combo1(Index).Text = "" & rsA.Fields(0).Value & " " & rsA.Fields(1).Value
                'Add by Morgan 2004/5/20
                stCountry = "" & rsA.Fields(0).Value
                
                'edit by nick 2004/08/05
                'Call setCP81  '設定減免身分
                'edit by nick 2004/10/05
                If IsoptCP81 = True Then
                    Call setCP811  '設定減免身分
                    'add by nick 2004/08/05
                    Call setCP812  '設定減免身分
                    Call setCP813  '設定減免身分
                    Call setCP814  '設定減免身分
                    Call setCP815  '設定減免身分
                End If
                'Add By Sindy 2011/11/11
                m_NA59 = "" & rsA.Fields("NA59").Value
'                Call SetFrame16
                '2011/11/11 End
'                Call setFrame21 'Add By Sindy 2012/5/8
                
               If rsA.State <> adStateClosed Then rsA.Close
               Set rsA = Nothing
            Else
               If rsA.State <> adStateClosed Then rsA.Close
               Set rsA = Nothing
               'Modify By Sindy 2019/7/2
               'Modify By Sindy 2024/2/22
               'Call SetComboCase(0, Combo1(0).Text)
               Call frm090801_New_SetComboCase(0, "", Combo1(0), Me.Text1(6).Text, _
                        Combo1(0).Text, ChkPCT, Text1(101), Text1(102), Text1(103))
               '2024/2/22 END
               If Combo1(0).Text <> "" And Combo1(0).ListCount > 1 Then
                  Cancel = True
                  SendMessage Combo1(0).hWnd, CB_SHOWDROPDOWN, True, ByVal 0&
                  Exit Sub
               Else
               '2019/7/2 END
                  If UBound(Split(Combo1(0).Text, " ")) = 0 Then
                     Cancel = True
                     MsgBox "申請國代號輸入錯誤!!!", vbExclamation + vbOKOnly
                     Exit Sub
                  End If
               End If
            End If
            
            'add by nickc 2007/05/17 加入馬德里一定要是 TF
            If arrNA01(0) = "238" And Text1(6) <> "TF" Then
                MsgBox "馬德里必須要收 TF !!!", vbExclamation + vbOKOnly
                Cancel = True
                Exit Sub
            End If
        End If
        SetNewDrug 'Added by Morgan 2021/7/20
    End Select
End Sub

'Copy From frmacc1121 By Morgan 2004/6/10
'取得此客戶所開的收據抬頭, 並預設最近一次開的收據抬頭, 若無則預設申請人
Private Sub GetReceiptTitle()
   Dim strCustNo As String
   'add by nick 2004/08/05
   Dim AllCust As Integer
   Dim strTmp As String 'Add by Amy 2016/09/01 避免選單內資料重覆ex:P-094328
   
   m_strTitleName = Me.cboTitle.Text
   Me.cboTitle.Clear
   'add by nick 2004/08/05
   For AllCust = 1 To 5
        Select Case AllCust
        Case 1
            Me.cboTitle.AddItem Text1(21).Text
            strTmp = strTmp & ";" & Text1(21).Text 'Add by Amy 2016/09/01
            strCustNo = Text1(12)
        Case 2
            'Add by Amy 2016/09/01 +if
            If InStr(strTmp, ";" & Text1(37).Text) = 0 Then
                Me.cboTitle.AddItem Text1(37).Text
                strTmp = strTmp & ";" & Text1(28).Text
            End If
            strCustNo = Text1(28)
        Case 3
            'Add by Amy 2016/09/01 +if
            If InStr(strTmp, ";" & Text1(53).Text) = 0 Then
                Me.cboTitle.AddItem Text1(53).Text
                strTmp = strTmp & ";" & Text1(53).Text
            End If
            strCustNo = Text1(44)
        Case 4
            'Add by Amy 2016/09/01 +if
            If InStr(strTmp, ";" & Text1(69).Text) = 0 Then
                Me.cboTitle.AddItem Text1(69).Text
                strTmp = strTmp & ";" & Text1(69).Text
            End If
            strCustNo = Text1(60)
        Case 5
            'Add by Amy 2016/09/01 +if
            If InStr(strTmp, ";" & Text1(85).Text) = 0 Then
                Me.cboTitle.AddItem Text1(85).Text
                strTmp = strTmp & ";" & Text1(85).Text
            End If
            strCustNo = Text1(76)
        Case Else
        End Select
            If strCustNo <> "" Then
               'Modify By Sindy 2012/12/19 +and A0K09 is null剔除已作廢收據
               strSql = "Select Distinct A0K04 From ACC0K0 Where A0K03='" & strCustNo & "' and nvl(A0K09,0)=0 Order By 1 "
               CheckOC
               
               adoRecordset.CursorLocation = adUseClient
               adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
               If adoRecordset.RecordCount > 0 Then
                  While Not adoRecordset.EOF
                     'Modify by Amy 2016/09/01 避免選單內資料重覆
                     If "" & adoRecordset.Fields(0).Value <> Me.cboTitle.List(0) And InStr(strTmp, ";" & adoRecordset.Fields(0).Value) = 0 Then
                        Me.cboTitle.AddItem "" & adoRecordset.Fields(0).Value
                         strTmp = strTmp & ";" & "" & adoRecordset.Fields(0).Value
                     End If
                     adoRecordset.MoveNext
                  Wend
               End If
               CheckOC
               'Modify By Sindy 2012/12/19 +and A0K09 is null剔除已作廢收據
               strSql = "Select A0K04 From ACC0K0 Where A0K03='" & strCustNo & "' and nvl(A0K09,0)=0 Order By A0K02 Desc "
               adoRecordset.CursorLocation = adUseClient
               adoRecordset.MaxRecords = 1
               adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
               If adoRecordset.RecordCount > 0 Then
                  'Add By Sindy 2010/4/28 增加判斷若收據抬頭沒有點選其他時才帶入值
                  If Option2(1).Value = False Then
                     cboTitle = "" & adoRecordset.Fields(0).Value
                  Else
                     cboTitle = m_strTitleName
                  End If
               End If
               CheckOC
               adoRecordset.MaxRecords = 0
               
               'Add By Sindy 2010/4/15 代出該案號的最後一次的抬頭
               'Modify By Sindy 2012/12/19 +and A0K09 is null剔除已作廢收據
               If Me.Option1(1).Value = True And Me.Text1(6).Text <> "" And Me.Text1(7).Text <> "" Then
                  If Me.Text1(8).Text = "" Then Me.Text1(8).Text = "0"
                  If Me.Text1(9).Text = "" Then Me.Text1(9).Text = "00"
                  strSql = "select A0K04 from caseprogress,ACC0K0 " & _
                              "where cp01='" & Text1(6) & "' and cp02='" & Text1(7) & "' and cp03='" & Text1(8) & "' and cp04='" & Text1(9) & "' " & _
                              "and cp60=A0K01(+) and A0K04 is not null and nvl(A0K09,0)=0 " & _
                              "Order By A0K02 Desc "
                  adoRecordset.CursorLocation = adUseClient
                  adoRecordset.MaxRecords = 1
                  adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
                  If adoRecordset.RecordCount > 0 Then
                     'Add By Sindy 2010/4/28 增加判斷若收據抬頭沒有點選其他時才帶入值
                     If Option2(1).Value = False Then
                        cboTitle = "" & adoRecordset.Fields(0).Value
                     Else
                        cboTitle = m_strTitleName
                     End If
                  End If
                  CheckOC
                  adoRecordset.MaxRecords = 0
               End If
               '2010/4/15 End
            End If
    'add by nick 2004/08/05
    Next AllCust
End Sub

'2008/9/3 add by sonia
Private Sub Combo3_Validate(Index As Integer, Cancel As Boolean)
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
Dim arrNA01

   If Trim(Me.Combo3(Index).Text) <> "" Then
      arrNA01 = Split(Me.Combo3(Index).Text, " ")
      StrSQLa = "Select NA01, NA03 From Nation Where Length(NA01)=3 And NA01<='9999' And substr(NA02,3,1)='0' And NA01='" & arrNA01(0) & "' Order By NA01 "
      rsA.CursorLocation = adUseClient
      rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
      If rsA.RecordCount > 0 Then
         If rsA.Fields(0) = "000" Then
             Me.Combo3(Index).Text = rsA.Fields(0).Value & " " & rsA.Fields(1).Value
             Me.Combo3(Index).Tag = rsA.Fields(0).Value & " " & "中華民國"
         Else
            Me.Combo3(Index).Text = rsA.Fields(0).Value & " " & rsA.Fields(1).Value
            Me.Combo3(Index).Tag = rsA.Fields(0).Value & " " & rsA.Fields(1).Value
         End If
      Else
          MsgBox "發明人國籍輸入錯誤!!!", vbExclamation + vbOKOnly
          Cancel = True
          Exit Sub
      End If
         
   Else
      If rsA.State <> adStateClosed Then rsA.Close
      Set rsA = Nothing
   End If
End Sub

Private Sub Combo5_Click()
   'Added by Morgan 2013/9/26
   If Combo5.Tag <> Combo5 Then
      SetNewDrug 'Added by Morgan 2021/7/20
   End If
   Combo5.Tag = Combo5
   'end 2013/9/26
End Sub

'Add By Sindy 2010/10/28
Private Sub Combo5_Validate(Cancel As Boolean)
   If Combo5 <> "" Then
      'Modify by Amy 2016/06/06 +專利設計案案件屬性
      If InStr(Text1(6), "P") > 0 And Left(Combo6, 1) = "3" Then
        Combo5 = Left(Combo5, 1) + "." + PUB_GetCaseAttributeName(Left(Combo5, 1), "3")
      Else
        Combo5 = Left(Combo5, 1) + "." + PUB_GetCaseAttributeName(Left(Combo5, 1))
      End If
      'end 2016/06/06
      If Combo5 = Left(Combo5, 1) + "." Then
         Combo5 = Left(Combo5, 1)
         Cancel = True
         SSTab1.Tab = 0
         'Modified by Lydia 2021/06/18 debug:接洽單查詢及列印使用  +And Me.Visible = True
         If Combo5.Enabled = True And Me.Visible = True Then Combo5.SetFocus
      End If
   End If
   Combo5.Tag = Combo5
   'end 2013/9/26
End Sub
'2010/10/28 End

'Added by Morgan 2021/7/20
Private Sub Combo6_Click()
   If Combo6.Tag <> Combo6 Then
      SetNewDrug
   End If
   'Combo6.Tag = Combo6 'Removed by Morgan 2021/9/24
End Sub

'Add By Sindy 2012/4/26
Private Sub Combo6_Validate(Cancel As Boolean)
'Dim strTemp As String
'Dim arrCaseKind As Variant
'Dim i As Integer
'
'   If Combo6 <> "" Then
'      arrCaseKind = Split(Combo6.Text, ".")
'      '判斷是否需要詢問
'      For i = 0 To UBound(arrCaseKind)
'         If ClsPDGetPatentTrademarkKind(商標, CStr(arrCaseKind(0)), strTemp, False) = 1 Then
'            Exit For
'         Else
'            Cancel = True
'            SSTab1.Tab = 0
'            Combo6.SetFocus
'         End If
'      Next i
'   End If
   
   If Combo6 <> "" Then
      If InStr(Text1(6), "P") > 0 Then
         'Add by Amy 2016/06/06 +專利設計案案件性質
         Call SetCombo5_P
         Combo6 = Left(Combo6, 1) + "." + GetPKindName(Left(Combo6, 1))
         'Combo6.Tag = Combo6 'Removed by Morgan 2021/9/24 移到下面
         'end 2016/06/06
         'Modify By Sindy 2014/7/16 Mark
         'Call SetCombo5 'Add By Sindy 2012/9/3
      ElseIf InStr(Text1(6), "T") > 0 Then
         Combo6 = Left(Combo6, 1) + "." + GetTKindName(Left(Combo6, 1))
      End If
      If Combo6 = Left(Combo6, 1) + "." Then
         Combo6 = Left(Combo6, 1)
         Cancel = True
         SSTab1.Tab = 0
         'Modified by Lydia 2021/06/18 debug:接洽單查詢及列印使用 +And Me.Visible = True
         If Combo6.Enabled = True And Me.Visible = True Then Combo6.SetFocus
      End If
   End If
   Combo6.Tag = Combo6 'Added by Morgan 2021/9/24
End Sub
'2012/4/26 End

'Add by Sindy 2012/6/6
'專利種類代碼轉中文
Private Function GetPKindName(p_Code As String) As String
   GetPKindName = ""
   strSql = "select ptm02,ptm03 from patenttrademarkmap where ptm01='1' and ptm02 in('1','2','3') and ptm02='" & p_Code & "' order by ptm02 asc"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      If RsTemp.RecordCount > 0 Then
         GetPKindName = RsTemp.Fields("ptm03")
      End If
   End If
   Select Case p_Code
'      Case "1"
'         GetPKindName = "發明"
'      Case "2"
'         GetPKindName = "新型"
'      Case "3"
'         GetPKindName = "設計"
      Case "4"
         GetPKindName = "積體電路"
   End Select
   If GetPKindName = "" And p_Code <> "" Then
      MsgBox "專利種類錯誤！", vbExclamation
   End If
End Function

'商標種類代碼轉中文
Private Function GetTKindName(p_Code As String) As String
   GetTKindName = ""
   'Modify By Sindy 2023/11/15
   'strSql = "select ptm02,ptm03 from patenttrademarkmap where ptm01='2' and ptm02 in('1','7','8','9') and ptm02='" & p_Code & "' order by ptm02 asc"
   strSql = GetCombo6_T_SQL(p_Code)
   '2023/11/15 END
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      If RsTemp.RecordCount > 0 Then
         GetTKindName = RsTemp.Fields("ptm03")
      End If
   End If
'   Select Case p_Code
'      Case "1"
'         GetTKindName = "商標"
'      Case "7"
'         GetTKindName = "證明標章"
'      Case "8"
'         GetTKindName = "團體標章"
'      Case "9"
'         GetTKindName = "團體商標"
'   End Select
   If GetTKindName = "" And p_Code <> "" Then
      MsgBox "商標種類錯誤！", vbExclamation
   End If
End Function
'2012/6/6 End

'Add By Sindy 2022/9/21
' 更新各控制項的狀態
Private Sub SetCtrlReadOnly(ByVal bEnable As Boolean)
Dim objText As Object
Dim objChk As Object
Dim objCbo As Object
Dim objOpt As Object
Dim objCom As Object
Dim objFra As Object
   
   Text5.Locked = bEnable
   
   cmdUpd.Visible = Not bEnable
   cmdDel.Visible = Not bEnable
   cmdClear2.Visible = Not bEnable
   'cmdOK(3).Visible = Not bEnable
   For Each objCom In cmdQual
      objCom.Visible = Not bEnable
   Next
   For Each objCom In cmdSerach
      objCom.Visible = Not bEnable
   Next
   For Each objCom In cmdTW
      objCom.Visible = Not bEnable
   Next
   For Each objCom In CmdSame
      objCom.Visible = Not bEnable
   Next
   For Each objCom In cmdSearchZip
      objCom.Visible = Not bEnable
   Next
   cmdPic.Visible = Not bEnable
   cmdTMQApp.Visible = Not bEnable
   cmdTMQ.Visible = Not bEnable
   
   txtCRL69.Locked = bEnable: txtCRL70.Locked = bEnable
   For Each objText In TxtC1
      objText.Locked = bEnable
   Next
   For Each objText In txtYear
      objText.Locked = bEnable
   Next
   For Each objText In txtMonth
      objText.Locked = bEnable
   Next
   For Each objText In txtDay
      objText.Locked = bEnable
   Next
   txtItemCount.Locked = bEnable
   txtItemList.Locked = bEnable
   
   For Each objText In Text1
      objText.Locked = bEnable
      objText.Enabled = bEnable
   Next
   For Each objText In Text2
      objText.Locked = bEnable
      objText.Enabled = bEnable
   Next
   For Each objText In Text3
      objText.Locked = bEnable
      objText.Enabled = bEnable
   Next
   For Each objText In Text4
      objText.Locked = bEnable
      objText.Enabled = bEnable
   Next
   Text7.Locked = bEnable
   Text7.Enabled = bEnable
   
   'Modify By Sindy 2023/6/14 原有執行過查詢(圖形附檔),但又變(其他)
'   For Each objOpt In opt1
'      objOpt.Enabled = Not bEnable
'      If objOpt.Value = True Then
'         objOpt.BackColor = &H80000005
'      Else
'         objOpt.BackColor = &H8000000F
'      End If
'   Next
   Frame15.Enabled = Not bEnable
   '2023/6/14 END
   
   For Each objOpt In Opt45
      objOpt.Enabled = Not bEnable
      If objOpt.Value = True Then
         objOpt.BackColor = &H80000005
      Else
         objOpt.BackColor = &H8000000F
      End If
   Next
   For Each objOpt In OptChoose
      objOpt.Enabled = Not bEnable
      If objOpt.Value = True Then
         objOpt.BackColor = &H80000005
      Else
         objOpt.BackColor = &H8000000F
      End If
   Next
   
   'Modify By Sindy 2023/6/14
'   For Each objOpt In optColor
'      objOpt.Enabled = Not bEnable
'      If objOpt.Value = True Then
'         objOpt.BackColor = &H80000005
'      Else
'         objOpt.BackColor = &H8000000F
'      End If
'   Next
   Frame17.Enabled = Not bEnable
   '2023/6/14 END
   
'   For Each objOpt In OptCP122
'      objOpt.Enabled = Not bEnable
'      If objOpt.Value = True Then
'         objOpt.BackColor = &H80000005
'      Else
'         objOpt.BackColor = &H8000000F
'      End If
'   Next
   For Each objOpt In OptCRL133
      objOpt.Enabled = Not bEnable
      If objOpt.Value = True Then
         objOpt.BackColor = &H80000005
      Else
         objOpt.BackColor = &H8000000F
      End If
   Next
   For Each objOpt In OptEntity
      objOpt.Enabled = Not bEnable
      If objOpt.Value = True Then
         objOpt.BackColor = &H80000005
      Else
         objOpt.BackColor = &H8000000F
      End If
   Next
   For Each objOpt In OptEP06
      objOpt.Enabled = Not bEnable
      If objOpt.Value = True Then
         objOpt.BackColor = &H80000005
      Else
         objOpt.BackColor = &H8000000F
      End If
   Next
   For Each objOpt In OptEP34
      objOpt.Enabled = Not bEnable
      If objOpt.Value = True Then
         objOpt.BackColor = &H80000005
      Else
         objOpt.BackColor = &H8000000F
      End If
   Next
   For Each objOpt In OptCP143
      objOpt.Enabled = Not bEnable
      If objOpt.Value = True Then
         objOpt.BackColor = &H80000005
      Else
         objOpt.BackColor = &H8000000F
      End If
   Next
   For Each objOpt In optCP811
      objOpt.Enabled = Not bEnable
      If objOpt.Value = True Then
         objOpt.BackColor = &H80000005
      Else
         objOpt.BackColor = &H8000000F
      End If
   Next
   For Each objOpt In optCP812
      objOpt.Enabled = Not bEnable
      If objOpt.Value = True Then
         objOpt.BackColor = &H80000005
      Else
         objOpt.BackColor = &H8000000F
      End If
   Next
   For Each objOpt In optCP813
      objOpt.Enabled = Not bEnable
      If objOpt.Value = True Then
         objOpt.BackColor = &H80000005
      Else
         objOpt.BackColor = &H8000000F
      End If
   Next
   For Each objOpt In optCP814
      objOpt.Enabled = Not bEnable
      If objOpt.Value = True Then
         objOpt.BackColor = &H80000005
      Else
         objOpt.BackColor = &H8000000F
      End If
   Next
   For Each objOpt In optCP815
      objOpt.Enabled = Not bEnable
      If objOpt.Value = True Then
         objOpt.BackColor = &H80000005
      Else
         objOpt.BackColor = &H8000000F
      End If
   Next
   For Each objOpt In Option1
      objOpt.Enabled = Not bEnable
      If objOpt.Value = True Then
         objOpt.BackColor = &H80000005
      Else
         objOpt.BackColor = &H8000000F
      End If
   Next
   For Each objOpt In Option2
      objOpt.Enabled = Not bEnable
      If objOpt.Value = True Then
         objOpt.BackColor = &H80000005
      Else
         objOpt.BackColor = &H8000000F
      End If
   Next
   For Each objOpt In Option3
      objOpt.Enabled = Not bEnable
      If objOpt.Value = True Then
         objOpt.BackColor = &H80000005
      Else
         objOpt.BackColor = &H8000000F
      End If
   Next
   For Each objOpt In Option31
      objOpt.Enabled = Not bEnable
      If objOpt.Value = True Then
         objOpt.BackColor = &H80000005
      Else
         objOpt.BackColor = &H8000000F
      End If
   Next
   For Each objOpt In Option32
      objOpt.Enabled = Not bEnable
      If objOpt.Value = True Then
         objOpt.BackColor = &H80000005
      Else
         objOpt.BackColor = &H8000000F
      End If
   Next
   For Each objOpt In Option33
      objOpt.Enabled = Not bEnable
      If objOpt.Value = True Then
         objOpt.BackColor = &H80000005
      Else
         objOpt.BackColor = &H8000000F
      End If
   Next
   For Each objOpt In Option34
      objOpt.Enabled = Not bEnable
      If objOpt.Value = True Then
         objOpt.BackColor = &H80000005
      Else
         objOpt.BackColor = &H8000000F
      End If
   Next
   For Each objOpt In Option35
      objOpt.Enabled = Not bEnable
      If objOpt.Value = True Then
         objOpt.BackColor = &H80000005
      Else
         objOpt.BackColor = &H8000000F
      End If
   Next
   For Each objOpt In Option4
      objOpt.Enabled = Not bEnable
      If objOpt.Value = True Then
         objOpt.BackColor = &H80000005
      Else
         objOpt.BackColor = &H8000000F
      End If
   Next
   For Each objOpt In Option5
      objOpt.Enabled = Not bEnable
      If objOpt.Value = True Then
         objOpt.BackColor = &H80000005
      Else
         objOpt.BackColor = &H8000000F
      End If
   Next
   For Each objOpt In Option6
      objOpt.Enabled = Not bEnable
      If objOpt.Value = True Then
         objOpt.BackColor = &H80000005
      Else
         objOpt.BackColor = &H8000000F
      End If
   Next
   For Each objOpt In Option8
      objOpt.Enabled = Not bEnable
      If objOpt.Value = True Then
         objOpt.BackColor = &H80000005
      Else
         objOpt.BackColor = &H8000000F
      End If
   Next
   For Each objOpt In Option9
      objOpt.Enabled = Not bEnable
      If objOpt.Value = True Then
         objOpt.BackColor = &H80000005
      Else
         objOpt.BackColor = &H8000000F
      End If
   Next
   For Each objOpt In OptNewDrug
      objOpt.Enabled = Not bEnable
      If objOpt.Value = True Then
         objOpt.BackColor = &H80000005
      Else
         objOpt.BackColor = &H8000000F
      End If
   Next
   
   'Modify By Sindy 2023/12/12
   Frame18.Enabled = Not bEnable
'   For Each objOpt In OptSendType
'      objOpt.Enabled = Not bEnable
'      If objOpt.Value = True Then
'         objOpt.BackColor = &H80000005
'      Else
'         objOpt.BackColor = &H8000000F
'      End If
'   Next
   '2023/12/12 END
   
   'Added by Lydia 2023/11/13
   For Each objOpt In optDB
      objOpt.Enabled = Not bEnable
      If objOpt.Value = True Then
         objOpt.BackColor = &H80000005
      Else
         objOpt.BackColor = &H8000000F
      End If
   Next
   'end 2023/11/13
   
   For Each objCbo In Combo1
      objCbo.Locked = bEnable
   Next
   For Each objCbo In Combo2
      objCbo.Locked = bEnable
   Next
   For Each objCbo In Combo3
      objCbo.Locked = bEnable
   Next
   Combo4.Locked = bEnable
   Combo5.Locked = bEnable
   Combo6.Locked = bEnable
'   Combo7.Locked = bEnable
   For Each objCbo In cboContact
      objCbo.Locked = bEnable
   Next
   cboTitle.Locked = bEnable
   
   Check1.Enabled = Not bEnable
   If Check1.Value = 1 Then
      Check1.BackColor = &H80000005
   Else
      Check1.BackColor = &H8000000F
   End If
   
   FrameCRL66.Enabled = Not bEnable
   FrameCRL90.Enabled = Not bEnable
   FrameCRL147.Enabled = Not bEnable
   Me.ChkCRL152.Enabled = Not bEnable 'Add By Sindy 2023/4/7
   Me.Check10.Enabled = Not bEnable 'Add By Sindy 2025/4/14
   
   Frame3.Enabled = Not bEnable 'Add By Sindy 2023/12/12
   
   Check2.Enabled = Not bEnable
   If Check2.Value = 1 Then
      Check2.BackColor = &H80000005
   Else
      Check2.BackColor = &H8000000F
   End If
   
   For Each objChk In Check3
      objChk.Enabled = Not bEnable
      If objChk.Value = 1 Then
         objChk.BackColor = &H80000005
      Else
         objChk.BackColor = &H8000000F
      End If
   Next
   Check4.Enabled = Not bEnable
   If Check4.Value = 1 Then
      Check4.BackColor = &H80000005
   Else
      Check4.BackColor = &H8000000F
   End If
   
   Check5.Enabled = Not bEnable
   If Check5.Value = 1 Then
      Check5.BackColor = &H80000005
   Else
      Check5.BackColor = &H8000000F
   End If
   
   For Each objChk In Check6
      objChk.Enabled = Not bEnable
      If objChk.Value = 1 Then
         objChk.BackColor = &H80000005
      Else
         objChk.BackColor = &H8000000F
      End If
   Next
   For Each objChk In Check7
      objChk.Enabled = Not bEnable
      If objChk.Value = 1 Then
         objChk.BackColor = &H80000005
      Else
         objChk.BackColor = &H8000000F
      End If
   Next
   For Each objChk In Check8
      objChk.Enabled = Not bEnable
      If objChk.Value = 1 Then
         objChk.BackColor = &H80000005
      Else
         objChk.BackColor = &H8000000F
      End If
   Next
   Check9.Enabled = Not bEnable
   If Check9.Value = 1 Then
      Check9.BackColor = &H80000005
   Else
      Check9.BackColor = &H8000000F
   End If
   
   For Each objChk In ChkAddress
      objChk.Enabled = Not bEnable
      If objChk.Value = 1 Then
         objChk.BackColor = &H80000005
      Else
         objChk.BackColor = &H8000000F
      End If
   Next
   For Each objFra In Frame30
      objFra.Enabled = Not bEnable
   Next
'   For Each objChk In ChkCRA26
'      objChk.Enabled = Not bEnable
'   Next
'   For Each objChk In ChkCRA27
'      objChk.Enabled = Not bEnable
'   Next
   
   For Each objChk In chkItem
      objChk.Enabled = Not bEnable
      If objChk.Value = 1 Then
         objChk.BackColor = &H80000005
      Else
         objChk.BackColor = &H8000000F
      End If
   Next
   ChkPCT.Enabled = Not bEnable
   If ChkPCT.Value = 1 Then
      ChkPCT.BackColor = &H80000005
   Else
      ChkPCT.BackColor = &H8000000F
   End If
End Sub

Private Sub Form_Activate()
   If Me.Enabled = True Then
      'Modify By Sindy 2019/2/15 + And cmdOK(1).Tag = ""
      '在高國碩86047電腦作業方式如下:
      '1.執行接洽單
      '2.執行申請人查詢(三晃股公司)
      '3.執行案件(P120493)----進卷宗區
      '出現附檔畫面：頁面跳動不停、當機。
'      If Text1(10).Enabled = True And cmdOK(1).Tag = "" Then Text1(10).SetFocus
      '2019/2/15 END
   End If
   
'   'Added by Morgan 2020/5/14
'   If cmdOK(1).Tag = "" Then
'      SrcLoadCheck
'   End If
'   'end 2020/5/14
   cmdOK(1).Tag = "接洽單第一次開啟" 'Add By Sindy 2019/2/15
   'Modify By Sindy 2023/5/18 + And Me.Enabled = True
   If Me.Visible = True And Me.Enabled = True Then 'Modify By Sindy 2023/4/17 +if 不然會出現"執行階段錯誤5,程序呼叫或引數不正確"
      SSTab1.SetFocus
   End If
End Sub

'Add by Morgan 2010/12/14 只要鍵盤有動作就不斷線
Private Sub Form_KeyPress(KeyAscii As Integer)
   If UCase(Forms(0).Name) = "MDIMAIN" Then Forms(0).tmrConnect.Tag = 0
End Sub

'Add By Sindy 2022/9/12
Private Sub SetListScroll(oList As ListBox)
   Dim ii As Integer
   Dim lWnow As Long, lWmax As Long
   
   lWmax = 0
   For ii = 0 To oList.ListCount - 1
      lWnow = TextWidth(oList.List(ii) & " ")
      If lWnow > lWmax Then
         lWmax = lWnow
      End If
   Next
  
   If ScaleMode = vbTwips Then lWmax = lWmax / Screen.TwipsPerPixelX  ' if twips change to pixels
   SendMessageByNum oList.hWnd, LB_SETHORIZONTALEXTENT, lWmax, 0
End Sub

Private Sub Form_Load()
Dim ii As Integer
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
Dim i As Integer
Dim objCbo As Object
   
   MoveFormToCenter Me
   Me.lblDate.BackColor = &H8000000F
   Me.lblStaffName.BackColor = &H8000000F
   Me.lblZone.BackColor = &H8000000F
   Me.SSTab1.Tab = 0
   'Modify By Sindy 2023/11/16
   'Me.SSTab2.Tab = 0
   If SSTab2.TabVisible(0) = True Then Me.SSTab2.Tab = 0
   '2023/11/16 END
   Me.SSTab3.Tab = 0 'Add By Sindy 2014/7/28
   'SSTab1.TabEnabled(4) = False 'Add By Sindy 2009/08/31
   SSTab1.TabEnabled(4) = True 'Add By Sindy 2010/6/21
   ReDim Preserve strCRL(TF_CRL) As String
       
   'Added by Lydia 2015/10/14
   If strSrvDate(1) >= TMQ電子化啟用日 Then 'Added by Lydia 2016/03/28
      Call PUB_GetTMQans("1", True) 'Added by Lydia 2016/06/02 求近似本所案
      
      If TypeName(m_PrevForm) <> "Nothing" Then
         If m_PrevForm.Name = "frm090127" Or m_PrevForm.Name = "frm090128" Then
            cmdTMQ.Visible = False
            cmdTMQApp.Visible = False 'Added by Lydia 2016/03/21
         End If
      End If
      'Added by Lydia 2016/03/21 判斷是否有查名單輸入
      If TypeName(Tmpfrm090126) = "Nothing" Then
         cmdTMQApp.Visible = False
      End If
      'end 2016/03/21
      m_UseTmqTma = "1" 'Added by Lydia 2024/11/11 預設使用原查名單
      
      SetGrdTMQ
   End If
   'end 2015/10/14
   
   Me.lblDate.Caption = Val(Mid(strSrvDate(1), 1, 4)) - 1911 & "年" & Mid(strSrvDate(1), 5, 2) & "月" & Mid(strSrvDate(1), 7, 2) & "日"
   For ii = Me.Combo1.LBound To Me.Combo1.UBound
       Me.Combo1(ii).Clear
   Next ii
   '申請國及發明人國籍
   StrSQLa = "Select NA01, NA03 From Nation Where Length(NA01)=3 And NA01<='9999' And substr(NA02,3,1)='0' Order By NA01 "
   rsA.CursorLocation = adUseClient
   rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
   While Not rsA.EOF
       Me.Combo1(0).AddItem "" & rsA.Fields(0).Value & " " & rsA.Fields(1).Value
       For i = 0 To 9
          Me.Combo3(i).AddItem "" & rsA.Fields(0).Value & " " & rsA.Fields(1).Value
       Next i
       rsA.MoveNext
   Wend
   If rsA.State <> adStateClosed Then rsA.Close
   Set rsA = Nothing
   
   'Add By Sindy 2010/5/27
   ClearChoose1Item (False)
   
   'Added by Lydia 2019/02/14 創新業務部人員收文控管
   Call PUB_ChkIsT10T20("1", strUserNum, m_Tuser, strExc(1))
   If m_Tuser <> "" Then
       Text1(10) = m_Tuser
       lblStaffName.Caption = strExc(1)
   Else
   'end 2019/02/14
        'Add By Sindy 2012/5/30
        Text1(10) = strUserNum
        lblStaffName.Caption = strUserName
        '2012/5/30 End
   End If
   'end 2019/02/14
   
'   Call Text1_Validate(10, False) 'Add By Sindy 2014/7/4

   ChkPCT.Visible = False
   Frame19.Visible = False 'Add By Sindy 2011/6/7
   
   'Add By Sindy 2013/12/13
   'Added by Lydia 2020/03/30 改用模組
   'If strSrvDate(1) >= InvoiceStartDate Then
   '   Combo4.AddItem "智權公司"
   'End If
   ''2013/12/13 END
   Call SetCombo4  '設定收據公司別(簡稱)
   
   txtPrintType = "2" ''輸出方式,印表機 Add By Sindy 2010/7/7
'   'Add By Sindy 2010/6/21
'   If Pub_StrUserSt03 = "M51" Then
'      cmdOK(4).Visible = True
'      Text5.Visible = True
'    Else
'      cmdOK(4).Visible = False
'      Text5.Visible = False
'   End If
   
   '暫時隱藏發明人
'   SSTab1.TabVisible(3) = False
   SSTab1.Tab = 0
   
   'Add By Sindy 2014/7/28 Frame設定值
   Frame1.BorderStyle = 0: Frame1.Caption = ""
   Frame10.BorderStyle = 0: Frame10.Caption = ""
   Frame11.BorderStyle = 0: Frame11.Caption = ""
   Frame12.BorderStyle = 0: Frame12.Caption = ""
   Frame13.BorderStyle = 0: Frame13.Caption = ""
   Frame14.BorderStyle = 0: Frame14.Caption = ""
   Frame15.BorderStyle = 0: Frame15.Caption = ""
   Frame16.BorderStyle = 0: Frame16.Caption = ""
   Frame17.BorderStyle = 0: Frame17.Caption = ""
   Frame18.BorderStyle = 0: Frame18.Caption = ""
   Frame19.BorderStyle = 0: Frame19.Caption = ""
   Frame2.BorderStyle = 0: Frame2.Caption = ""
   Frame20.BorderStyle = 0: Frame20.Caption = ""
   Frame21.BorderStyle = 0: Frame21.Caption = ""
   Frame22.BorderStyle = 0: Frame22.Caption = ""
   Frame23.BorderStyle = 0: Frame23.Caption = ""
'   Frame24.BorderStyle = 0: Frame24.Caption = "" '是否急件
   Frame25.BorderStyle = 0: Frame25.Caption = ""
   Frame26.BorderStyle = 0: Frame26.Caption = ""
   Frame27.BorderStyle = 0: Frame27.Caption = ""
   Frame28.BorderStyle = 0: Frame28.Caption = ""
   Frame29.BorderStyle = 0: Frame29.Caption = ""
   'Modified by Lydia 2023/11/13 預防Frame數量超過最大數,將第一頁的主Frame改為Index
   'Frame33.BorderStyle = 0: Frame33.Caption = ""
   Frame33(0).BorderStyle = 0: Frame33(0).Caption = ""
   Frame36.BorderStyle = 0: Frame36.Caption = ""
   Frame37.BorderStyle = 0: Frame37.Caption = ""
   Frame38.BorderStyle = 0: Frame38.Caption = ""
   Frame39.BorderStyle = 0: Frame39.Caption = ""
   Frame40.BorderStyle = 0: Frame40.Caption = ""
   Frame57.BorderStyle = 0: Frame57.Caption = "" 'Add By Sindy 2022/9/29
   GRD1.Clear: SetGrd2 'Add By Sindy 2022/9/19
   Frame58.BorderStyle = 0: Frame58.Caption = "" 'Add By Sindy 2022/10/14
   'Add By Sindy 2022/10/14
   FrameT.BorderStyle = 0: FrameT.Caption = "": FrameT.BackColor = &H8000000F
   FrameP.BorderStyle = 0: FrameP.Caption = "": FrameP.BackColor = &H8000000F: FrameP.Top = FrameT.Top: FrameP.Left = FrameT.Left
   FrameL.BorderStyle = 0: FrameL.Caption = "": FrameL.BackColor = &H8000000F: FrameL.Top = FrameT.Top: FrameL.Left = FrameT.Left
   Frame605.BorderStyle = 0: Frame605.Caption = ""
   Frame12.BorderStyle = 0: Frame12.Caption = ""
   FrameCase.Top = 1620: FrameCase.Left = 4170: lstCase.Clear '編輯相同案號
   '2022/10/14 END
   'Added by Lydia 2023/11/13 預防Frame數量超過最大數,將第一頁的主Frame改為Index
   Frame33(1).BackColor = &H8000000F  'DEBIT NOTE請款選項
   'end 2023/11/13
      
   cmdClear
   
   Frame4(0).BorderStyle = 0: Frame4(0).Caption = ""
   Frame4(1).BorderStyle = 0: Frame4(1).Caption = ""
   Frame4(2).BorderStyle = 0: Frame4(2).Caption = ""
   Frame4(3).BorderStyle = 0: Frame4(3).Caption = ""
   Frame4(4).BorderStyle = 0: Frame4(4).Caption = ""
   Frame6.BorderStyle = 0: Frame6.Caption = ""
   Frame7.BorderStyle = 0: Frame7.Caption = ""
   Frame8.BorderStyle = 0: Frame8.Caption = ""
   Frame9.BorderStyle = 0: Frame9.Caption = ""
   '2014/7/28 END
   Frame41.BorderStyle = 0: Frame41.Caption = "" 'Add by Amy 2015/11/13
   Frame5.Visible = False 'Add by Amy 2017/01/20
   'Added by Lydia 2018/12/10 查名是否齊備
   Frame42.BorderStyle = 0: Frame42.Caption = "": Frame42.BackColor = &H8000000F
   Frame42.Left = 60
   'Added by Lydia 2018/12/10 是否延期
   Frame43.BorderStyle = 0: Frame43.Caption = ""
   'Added by Lydia 2018/08/22 (應收帳款管控)取消預定收款日,改成付款週期
   Label1(105).Visible = False
   Text1(118).Visible = False
'   Label3.Visible = False
   
   'Added by Lydia 2019/08/06
   Frame44.Top = Frame19.Top
   Frame44.BorderStyle = 0
   
   'Added by Lydia 2021/02/24
   Frame47.BackColor = &H8000000F
   Frame47.Top = 2100
   
   'Added by Morgan 2021/7/20 大陸發明生醫案是否新藥專利設定
   Frame48.BorderStyle = 0
   Frame48.Visible = False
   
   'Frame20.Top = 2800 '2700 'Added by Lydia 2022/08/10 商標查名加註，回到原位
   
   m_ChkAmtExcept = Pub_GetSpecMan("應收帳款上限檢查排除") 'Added by Lydia 2022/09/06 改抓特殊設定
   
   'Add By Sindy 2022/8/29
   lblCnt.Caption = "": GridCase.Clear: Call SetGrd
   FrameCRC.Caption = "案件性質區（" & Val(GridCase.Rows - 1) - IIf(Trim(GridCase.TextMatrix(1, 1)) = "", 1, 0) & "）"
   Me.Height = 6920
   'Add By Sindy 2024/6/28 要先把Height紀錄下來
   Frame31.Visible = False
   PUB_SavePdfForm Me
   '2024/6/28 END
   If m_SignFlowEmp = "" Then m_SignFlowEmp = strUserNum '簽核人員
   '2022/8/29 END
   
   'Added by Lydia 2022/10/26 載入前次結束時的大小及位置
'   If UCase(TypeName(m_PrevForm)) = UCase("frm210156") Or _
'      UCase(TypeName(m_PrevForm)) = UCase("frm040101_1") Or _
'      UCase(TypeName(m_PrevForm)) = UCase("frm050101_2") Or _
'      UCase(TypeName(m_PrevForm)) = UCase("frm210148_2") Or _
'      UCase(TypeName(m_PrevForm)) = UCase("frm020101_02") Or _
'      UCase(TypeName(m_PrevForm)) = UCase("frm030201_02") Then
      PUB_SetPdfForm Me, False
'   Else
'   'end 2022/10/26
'      MoveFormToCenter Me
'   End If 'Added by Lydia 2022/10/26
   
   'Add By Sindy 2022/9/15
   For Each objCbo In Me.Combo2
      objCbo.Clear
      objCbo.AddItem ""
      objCbo.AddItem "業務失誤"
      objCbo.AddItem "專業失誤"
      objCbo.AddItem "專業支援點"
   Next
   '2022/9/15 END
   
   'Add By Sindy 2022/12/24 記錄位置
   OptEntity(0).Tag = OptEntity(0).Left
   OptEntity(1).Tag = OptEntity(1).Left
   OptEntity(2).Tag = OptEntity(2).Left
   '2022/12/24 END
End Sub

'Add By Sindy 2010/5/27
Private Sub ClearChoose1Item(blnTF As Boolean)
   '申請人國籍
   'Modify By Sindy 2023/12/19 mark:接洽單開放國籍欄，不限制外至台，國內接洽單不一定要輸
'   Label1(94).Visible = blnTF
'   Me.Text1(34).Visible = blnTF
'   Label1(110).Visible = blnTF
'   Me.Text1(50).Visible = blnTF
'   Label1(111).Visible = blnTF
'   Me.Text1(66).Visible = blnTF
'   Label1(112).Visible = blnTF
'   Me.Text1(82).Visible = blnTF
'   Label1(113).Visible = blnTF
'   Me.Text1(98).Visible = blnTF
   '2023/12/19 END
   
   '申請人地址(英)
   'Modify By Sindy 2023/2/7 國內接洽單也請顯示英文地址欄
'   Label1(92).Visible = blnTF
'   Me.Text1(125).Visible = blnTF
'   Label1(114).Visible = blnTF
'   Me.Text1(141).Visible = blnTF
'   Label1(115).Visible = blnTF
'   Me.Text1(157).Visible = blnTF
'   Label1(116).Visible = blnTF
'   Me.Text1(173).Visible = blnTF
'   Label1(117).Visible = blnTF
'   Me.Text1(189).Visible = blnTF
   '代理人
   Frame13.Visible = blnTF
   'Modified by Morgan 2020/5/7
   'Me.Text1(119).Text = "如案件內容摘要、引用條文與客戶洽談要旨等等......" & vbCrLf & _
                                       "商品類別：" & vbCrLf & _
                                       "商品名稱：" & vbCrLf & _
                                       "註冊號數：" & vbCrLf & _
                                       "優先權日：" & vbCrLf & _
                                       "優先權號：" & vbCrLf
   'If blnTF = True Then
   '   Me.Text1(119).Text = Me.Text1(119).Text & _
   '                                       "聯絡人：" & vbCrLf & _
   '                                       "彼所案號：" & vbCrLf
   'End If
   SrcSetMemo blnTF
   'end 2020/5/7
End Sub

Private Sub Form_Unload(Cancel As Integer)
   PUB_SavePdfForm Me '紀錄視窗最後的大小及位置
   
   'Add By Sindy 2024/3/25
   '特殊收據畫面存在時,要關閉
   If PUB_CheckFormExist("frm090801_7") = True Then
      Unload frm090801_7
   End If
   '2024/3/25 END
   
   'Added by Sindy 2023/1/4
   If bolIsTmp = False Then
      Set frm090801_Q = Nothing
   End If
   '2023/1/4 END
   If TypeName(m_PrevForm) <> "Nothing" Then
      If m_PrevForm.Visible = False Then
         m_PrevForm.Show
      End If
      If UCase(TypeName(m_PrevForm)) = UCase("frm12040152") Or _
         UCase(TypeName(m_PrevForm)) = UCase("frm100101_L") Then
         m_PrevForm.PubShowNextData
      End If
      Set m_PrevForm = Nothing
   End If
End Sub

'Add By Sindy 2009/08/31
Private Sub Opt1_Click(Index As Integer)
'   'Add By Sindy 2023/6/14 原有執行過查詢(圖形附檔),但又跳進來變(其他)
'   If cmdOK(1).Tag <> "" And opt1(0).Enabled = False Then
'      Exit Sub
'   End If
'   '2023/6/14 ENd
   Set G_SeekPicColor.Picture = LoadPicture()
   Set tmpPic.Picture = LoadPicture()
   Set tmpImg.Picture = LoadPicture()
   tmpImg.Visible = False
   cmdPic.Enabled = False
   Frame17.Visible = False: cmdPic.Visible = False
   'Add By Sindy 2009/11/02
   PicText.Visible = False
   PicText.Text = ""
   '2009/11/02 End
   'Add by Amy 2017/01/20 選擇同卷號時, PicText欄的位置改為本所案號欄
   Frame5.Visible = False
   TxtC1(0) = "": TxtC1(1) = "": TxtC1(2) = "": TxtC1(3) = ""
   'end 2017/01/20
   
   If opt1(0).Value = True Then
      Call LoadPic("000", "000000", "0", "03")
      'Add By Sindy 2009/11/02
      PicText.Visible = True
      'Modified by Lydia 2021/06/18 debug:接洽單查詢及列印使用 + And Me.Visible = True
      If PicText.Enabled = True And Me.Visible = True Then PicText.SetFocus
      '2009/11/02 End
   ElseIf opt1(1).Value = True Then
      tmpImg.Visible = True
      cmdPic.Enabled = True
      'If Pub_StrUserSt03 = "M51" Then
         Frame17.Visible = True: Frame17.Enabled = True: cmdPic.Visible = True
      'End If
   ElseIf opt1(2).Value = True Then
      Call LoadPic("000", "000000", "0", "02")
   'Add By Sindy 2009/11/02
   ElseIf opt1(3).Value = True Then
      'Add by Amy +系統別T/TF/CFT選擇同卷號時, PicText欄的位置改為本所案號欄
      If Trim(Text1(6)) = "T" Or Trim(Text1(6)) = "TF" Or Trim(Text1(6)) = "CFT" Then
        tmpImg.Visible = True
        Frame5.Width = 4335
        Frame5.Visible = True
        If TxtC1(0).Enabled = True And Me.Visible = True Then TxtC1(0).SetFocus
        Frame17.Visible = True 'Add By Sindy 2023/6/14
      Else
        PicText.Visible = True
        PicText.Text = "圖同"
        'Modified by Lydia 2023/01/03 程式會出錯
        'If PicText.Enabled = True Then PicText.SetFocus
        If PicText.Enabled = True And Me.Visible = True Then
          PicText.SetFocus
        End If
      End If
   ElseIf opt1(4).Value = True Then
      PicText.Visible = True
      'Modified by Lydia 2023/01/03 程式會出錯
      'If PicText.Enabled = True Then PicText.SetFocus
      If PicText.Enabled = True And Me.Visible = True Then
          PicText.SetFocus
      End If
   '2009/11/02 End
   End If
End Sub

'Add By Sindy 2009/08/31 選擇圖檔
Private Sub CmdPic_Click()
'cmdPic.Tag = "有異動"
Set G_SeekPicColor.Picture = LoadPicture()
Set tmpImg.Picture = LoadPicture()
Set frmPic001.oPic = G_SeekPicColor
Set frmPic001.oImg = tmpImg
Set frmPic001.UpForm = Me
frmPic001.oRtPic = True
frmPic001.cmdOK(4).Visible = False
frmPic001.cmdOK(5).Visible = False
frmPic001.cmdOK(6).Visible = False
frmPic001.cmdOK(7).Visible = False
frmPic001.cmdOK(2).Caption = "確定(&O)"
frmPic001.cmdOK(3).Caption = "取消(&X)"
frmPic001.Label11.Caption = "選擇圖片"
frmPic001.cmdOK(0).Left = frmPic001.cmdOK(0).Left - 250
frmPic001.cmdOK(1).Left = frmPic001.cmdOK(1).Left - 250
frmPic001.cmdOK(2).Left = frmPic001.cmdOK(2).Left - 250
frmPic001.cmdOK(3).Left = frmPic001.cmdOK(3).Left - 250
frmPic001.Width = 3800
MoveFormToCenter frmPic001
Unload frmpic002
frmPic001.SetSeekCmdok 'Add by Amy 2018/07/20
frmPic001.Show vbModal
End Sub

'Add By Sindy 2010/5/27
Private Sub optChoose_Click(Index As Integer)
   Option1(0).Enabled = True
   Option1(1).Enabled = True
   'Option1(0).Value = True
   'Call Option1_Click(0)
   Select Case Index
      Case 0 '國內
         Call ClearChoose1Item(False)
      Case 1 '大至台
         Call ClearChoose1Item(True)
   End Select
End Sub

''Add by Amy 2016/06/06 商標爭議案用
'Private Sub OptCP122_Click(Index As Integer)
'    If Index = 1 Then Exit Sub
'
'    OptCRL133(1).Value = True
'End Sub
'
'Private Sub OptCRL133_Click(Index As Integer)
'    If Index = 0 And OptCRL133(Index).Value = True Then
'        OptCP122(1).Value = True
'    End If
'End Sub

Private Sub Option1_Click(Index As Integer)
   'If m_blnCallPrint = True Then Exit Sub 'Add By Sindy 2014/7/25
    Select Case Index
    Case 0 '新案
        cmdAddAtt.Caption = "文件匯入" 'Add By Sindy 2022/9/6
        'Added by Lydia 2020/11/19 CFP和CFT英國脫歐案管制：帶入期限後，新案第一次先不清除畫面
        If bolCase201 = True Then
            bolCase201 = False
            Exit Sub
        End If
        'end 2020/11/19
        
        'Added by Morgan 2020/2/5 舊案點選新案時詢問是否清除，否則會有舊案期限資料殘留問題(Ex:P)--經理
        If Text1(7) <> "" And strLOS18 = "" And _
            Text5.Visible = False And cmdOK(0).Visible = True Then 'Modified by Morgan 2020/5/27 案源輸入不用問
            
            If MsgBox("是否清除原畫面資料？", vbYesNo + vbDefaultButton2) = vbYes Then
               cmdClear True
            End If
        End If
        'end 2020/2/5
        
        txtCRL69.Text = "": txtCRL70.Text = "" 'Add By Sindy 2022/9/15
        m_Note1 = "": m_Note2 = "" 'Add By Sindy 2010/4/23
        m_strGetNP01 = "" 'Add By Sindy 2015/9/17
        Text1(142).Text = "": Text1(143).Text = "" 'strYF05From = "": strYF05To = "" 'Add By Sindy 2010/7/9
        Me.Text1(7).Text = ""
        Me.Text1(8).Text = ""
        Me.Text1(9).Text = ""
        Me.Text1(7).Enabled = False
        Me.Text1(8).Enabled = False
        Me.Text1(9).Enabled = False
        Me.Text1(18).Enabled = True
        txtEnabled1 True
        'cmdOK(3).Enabled = False
         
         ClearCustTxt 1
         ClearCustTxt 2
         ClearCustTxt 3
         ClearCustTxt 4
         ClearCustTxt 5
         ClearFagentTxt 'Add By Sindy 2010/5/27
         cboTitle.Clear
         'Add By Sindy 2013/12/13
         'Modified by Lydia 2020/03/30 改用模組
         'If strSrvDate(1) >= InvoiceStartDate Then
         '   Combo4.Enabled = True
         '   If Combo4.ListCount < 4 Then
         '      Combo4.AddItem "智權公司"
         '   End If
         'End If
         ''2013/12/13 END
         Call SetCombo4 '設定收據公司別(簡稱)
         'end 2020/03/30
         
         Call Combo1_LostFocus(0) 'Add By Sindy 2013/12/14 執行預設動作
         
         'Modified by Lydia 2021/06/18 debug:接洽單查詢及列印使用 + And Me.Visible = True
         If Me.Text1(6).Enabled = True And Me.Visible = True Then Me.Text1(6).SetFocus
         
         'Added by Morgan 2020/5/27
         '案源輸入法務案時預設國家及申請人
         If (strLOS18 <> "" Or strLOS15 <> "" And strLCaseNo(1) = "") And m_blnCallPrint = False Then
            'Modify By Sindy 2024/2/22
            'SetComboCase 0, "000"
            'SetComboCase 0, Combo1(0).Text
            Call frm090801_New_SetComboCase(0, "000", Combo1(0), Me.Text1(6).Text, _
                  Combo1(0).Text, ChkPCT, Text1(101), Text1(102), Text1(103))
            Call frm090801_New_SetComboCase(0, Combo1(0).Text, Combo1(0), Me.Text1(6).Text, _
                  Combo1(0).Text, ChkPCT, Text1(101), Text1(102), Text1(103))
            '2024/2/22 END
            If strLOS18 <> "" Then
               SrcLoadCustbyCRL01 strLOS18
            Else
               SrcSetCustByVar
            End If
         End If
         'end 2020/5/27
         
    Case 1 '舊案
        cmdAddAtt.Caption = "回覆單匯入" 'Add By Sindy 2022/9/6
        m_Note1 = "": m_Note2 = "" 'Add By Sindy 2010/4/23
        m_strGetNP01 = "" 'Add By Sindy 2015/9/17
        Text1(142).Text = "": Text1(143).Text = "" 'strYF05From = "": strYF05To = "" 'Add By Sindy 2010/7/9
        Me.Text1(7).Enabled = True
        Me.Text1(8).Enabled = True
        Me.Text1(9).Enabled = True
        Me.Text1(18).Enabled = False
        txtEnabled1 False
        Me.Text1(18).Text = ""
        'edit by nickc 2005/04/06
        'Me.Option3(1).Value = True
        Option31(1).Value = True
        Option32(1).Value = True
        Option33(1).Value = True
        Option34(1).Value = True
        Option35(1).Value = True

        'Modified by Lydia 2021/06/18 debug:接洽單查詢及列印使用 + And Me.Visible = True
        If Me.Text1(6).Enabled = True And Me.Visible = True Then Me.Text1(6).SetFocus

         'Add By Sindy 2013/12/13
         If strSrvDate(1) >= InvoiceStartDate Then
            'Modified by Lydia 2020/03/30
            'If Combo4.ListCount = 4 Then
            '   Combo4.RemoveItem 3
            'Memo by Lydia 2022/09/22 移除選項是因為舊案不可更換收據公司別；
            '  若確定要更換公司別，需先請財務修改之前收據後會自動發Email通知「程式管理人員」，
            '  經過人員修改資料後接洽單才可以更換收據公司別。
            If Combo4.ListCount = m_CompJforIdx + 1 Then
               Combo4.RemoveItem m_CompJforIdx
            'end 2020/03/30
            End If
         End If
         '2013/12/13 END
    End Select
    
    'Add by Morgan 2004/5/20
    'edit by nick 2004/08/05
    'Call setCP81
    Call setCP811
    'add by nick 2004/08/05
    Call setCP812
    Call setCP813
    Call setCP814
    Call setCP815
    LockContact 'Add by Morgan 2008/9/16
    SetNewDrug 'Addd by Morgan 2021/7/20
    Call SetCombo4 'Added by Lydia 2020/03/30 設定收據公司別(簡稱)
End Sub

'Add by Amy 2016/03/25
Private Sub Option2_Click(Index As Integer)
    Frame33(1).Visible = False  'Added by Lydia 2023/11/13 DEBIT NOTE請款選項;預設不顯示
    'Add by Amy 2016/06/24 +勾選Debit Note不可點選特殊收據
    If Index = 2 Then
        If Option2(2).Value = True Then
            If Check9.Value = 1 Then Check9.Value = 0
            Call Check9_Click
            Check9.Enabled = False
            If OptChoose(0).Value = True Then 'Added by Lydia 2023/11/14 排除外至台案件(MCTF)
               Frame33(1).Visible = True  'Added by Lydia 2023/11/13 DEBIT NOTE請款選項
            End If
        End If
    Else
        'Add by Amy 2016/12/23 +舊案多申請人勾同申請人,自動勾特殊收據(P-091969)
        If Index = 0 Then
            If Option1(1).Value = True And (Me.Text1(21) <> "" Or Me.Text1(22) <> "") And _
              (Me.Text1(37) <> "" Or Me.Text1(38) <> "") And Check9.Value = 0 Then
                Check9.Value = 1
                Call Check9_Click
            End If
        End If
        'end 2016/12/23
        Check9.Enabled = True
    End If
End Sub

Private Sub Option31_Click(Index As Integer)
    Select Case Index
    Case 0 '新客戶
        txtEnabled2 True, 1
        ClearCustTxt 1
        If IsoptCP81 = True Then
            optCP811(0).Enabled = True
            optCP811(1).Enabled = True
        End If
    Case 1 '舊客戶
        txtEnabled2 False, 1
        If IsoptCP81 = True Then
            optCP811(0).Enabled = False
            optCP811(1).Enabled = False
        End If
        Call Combo1_Validate(0, False)
    End Select
    LockContact 1
    
    SetQualVisible 'Added by Morgan 2013/4/11
End Sub

Private Sub Option32_Click(Index As Integer)
    Select Case Index
    Case 0 '新客戶
        txtEnabled2 True, 2
        ClearCustTxt 2
        If IsoptCP81 = True Then
            optCP812(0).Enabled = True
            optCP812(1).Enabled = True
        End If
    Case 1 '舊客戶
        txtEnabled2 False, 2
        If IsoptCP81 = True Then
            optCP812(0).Enabled = False
            optCP812(1).Enabled = False
        End If
        Call Combo1_Validate(0, False)
    End Select
    LockContact 2
    
    SetQualVisible 'Added by Morgan 2013/4/11
End Sub

Private Sub Option33_Click(Index As Integer)
    Select Case Index
    Case 0 '新客戶
        txtEnabled2 True, 3
        ClearCustTxt 3
        If IsoptCP81 = True Then
            optCP813(0).Enabled = True
            optCP813(1).Enabled = True
        End If
    Case 1 '舊客戶
        txtEnabled2 False, 3
        If IsoptCP81 = True Then
            optCP813(0).Enabled = False
            optCP813(1).Enabled = False
        End If
        Call Combo1_Validate(0, False)
    End Select
    LockContact 3
    
    SetQualVisible 'Added by Morgan 2013/4/11
End Sub

Private Sub Option34_Click(Index As Integer)
    Select Case Index
    Case 0 '新客戶
        txtEnabled2 True, 4
        ClearCustTxt 4
        If IsoptCP81 = True Then
            optCP814(0).Enabled = True
            optCP814(1).Enabled = True
        End If
    Case 1 '舊客戶
        txtEnabled2 False, 4
        If IsoptCP81 = True Then
            optCP814(0).Enabled = False
            optCP814(1).Enabled = False
        End If
        Call Combo1_Validate(0, False)
    End Select
    LockContact 4
    
    SetQualVisible 'Added by Morgan 2013/4/11
End Sub

Private Sub Option35_Click(Index As Integer)
    Select Case Index
    Case 0 '新客戶
        txtEnabled2 True, 5
        ClearCustTxt 5
        If IsoptCP81 = True Then
            optCP815(0).Enabled = True
            optCP815(1).Enabled = True
        End If
    Case 1 '舊客戶
        txtEnabled2 False, 5
        If IsoptCP81 = True Then
            optCP815(0).Enabled = False
            optCP815(1).Enabled = False
        End If
        Call Combo1_Validate(0, False)
    End Select
    LockContact 5
    
    SetQualVisible 'Added by Morgan 2013/4/11
End Sub

Private Sub Option4_Click(Index As Integer)
    Select Case Index
    Case 0
        Me.Option4(1).Value = Not Me.Option4(0).Value
        Me.Option4(2).Value = Not Me.Option4(0).Value
    Case 1
        Me.Option4(0).Value = Not Me.Option4(1).Value
        Me.Option4(2).Value = Not Me.Option4(1).Value
    Case 2
        Me.Option4(0).Value = Not Me.Option4(2).Value
        Me.Option4(1).Value = Not Me.Option4(2).Value
        'Modified by Lydia 2021/06/18 debug:接洽單查詢及列印使用 + And Me.Visible = True
        If Me.Text1(97).Enabled = True And Me.Visible = True Then Me.Text1(97).SetFocus
    End Select
End Sub

'Add by Morgan 2011/3/8
Private Sub Text1_Change(Index As Integer)
   If Index = 6 Then
      SetNewDrug 'Added by Morgan 2021/7/20
      SrcSetMemo OptChoose(1).Value, True  'Added by Morgan 2020/5/7
      
      'Add By Sindy 2009/08/31
      '商標
      If Text1(6) = "T" Or Text1(6) = "TF" Or Text1(6) = "FCT" Or _
         Text1(6) = "TS" Or Text1(6) = "S" Or Text1(6) = "CFT" Then
         SSTab1.TabVisible(4) = True
      Else
         SSTab1.TabVisible(4) = False
      End If
      
      'Add By Sindy 2022/12/15
      '專利
      If Text1(6) = "P" Or Text1(6) = "CFP" Then
         SSTab1.TabVisible(3) = True
      Else
         SSTab1.TabVisible(3) = False
      End If
   End If
   
   'Added by Morgan 2016/7/20
   'Modified by Morgan 2016/12/27 減免身分資料也要清除
   Select Case Index
   Case 12: stCustNo1 = "": If IsoptCP81 = True Then Call setCP811
   Case 28: stCustNo2 = "": If IsoptCP81 = True Then Call setCP812
   Case 44: stCustNo3 = "": If IsoptCP81 = True Then Call setCP813
   Case 60: stCustNo4 = "": If IsoptCP81 = True Then Call setCP814
   Case 76: stCustNo5 = "": If IsoptCP81 = True Then Call setCP815
   End Select
   'end 2016/12/27
   'end 2016/7/20
   
   'Added by Morgan 2022/3/14
   If Index = 119 Then
      PUB_RefreshText Text1(119)
   End If
   'end 2022/3/14

End Sub

Private Sub Text1_GotFocus(Index As Integer)
   
'   'Added by Morgan 2022/3/17
'   '修正Tab鍵無法正常跳至下一駐點問題(因TextBox的Visible設True會自動取得駐點)
'   'Modify By Sindy 2022/11/29
'   If Not Me.ActiveControl Is Nothing Then
'   '2022/11/29 END
'      If Me.ActiveControl.Name = "Text1" Then
'         If Me.ActiveControl.Index = Index Then
'            If m_bolText1SetFocus Then
'               If m_objControl.Enabled And m_objControl.Visible Then
'                  'Modified by Morgan 2022/3/25 非目前欄位才設駐點
'                  If Index <> m_objControl.Index Then
'                     m_objControl.SetFocus
'                  End If
'               End If
'               m_bolText1SetFocus = False
'            End If
'         End If
'      End If
'   End If
'   'end 2022/3/17

    Select Case Index
    Case 6
        'Add By Sindy 2010/5/27
        If OptChoose(0).Visible = True Then
            If OptChoose(0).Value = False And OptChoose(1).Value = False Then
               MsgBox "請點選接洽單種類!!!", vbExclamation + vbOKOnly
               SSTab1.Tab = 0
               Exit Sub
            End If
        End If
         TextInverse Me.Text1(Index) 'Add by Lydia 2014/12/39 案號反白
    Case 119
        Me.Text1(Index).SelStart = Len(Me.Text1(Index).Text)
    Case Else
'        TextInverse Me.Text1(Index)
    End Select
    'add by nickc 2007/07/13 將輸入法改成使用API
    If Index = 598 Or Index = 498 Or Index = 398 Or Index = 298 Or Index = 97 Or _
       Index = 198 Or Index = 21 Or Index = 23 Or Index = 26 Or Index = 27 Or _
       Index = 37 Or Index = 39 Or Index = 42 Or Index = 43 Or Index = 53 Or _
       Index = 55 Or Index = 58 Or Index = 59 Or Index = 69 Or Index = 71 Or _
       Index = 74 Or Index = 75 Or Index = 85 Or Index = 87 Or Index = 90 Or Index = 91 Then
        OpenIme
    Else
        CloseIme
    End If
        
End Sub

'Add By Sindy 2022/12/26
Private Sub SetFrmCol()
   m_strCaseCPM = GetAllCaseCPM() '取得案件性質代碼
   Call SetFrame41 'Add by Amy 2015/11/13
   Call SetFrame16 'Add By Sindy 2011/11/11
   Call SetFrame27 'Add By Sindy 2013/2/25
   Call SetFrame20 'Add By Sindy 2012/3/6
   Call setFrame21 'Add By Sindy 2012/5/8
   Call SetFrameChg 'Add By Sindy 2022/12/16
   Call SetFrame28 'Add By Sindy 2022/12/24
End Sub

'Modify By Sindy 2014/5/23
'Private Sub Text1_LostFocus(Index As Integer)
Public Sub Text1_LostFocus(Index As Integer)
'2014/5/23 END
   'Added by Morgan 2022/3/17
   '修正Tab鍵無法正常跳至下一駐點問題(因TextBox的Visible設True會自動取得駐點)
   'Modify By Sindy 2022/11/29
   If Not Me.ActiveControl Is Nothing Then
   '2022/11/29 END
      If Not m_bolText1SetFocus Then
         If Screen.ActiveForm.Name = Me.Name Then 'Added by Morgan 2022/3/25 視窗切換的觸發不必管
            If Me.ActiveControl.Name = "Text1" Then
               Set m_objControl = Me.ActiveControl
               m_bolText1SetFocus = True
            End If
         End If
      End If
   End If
   'end 2022/3/17
   
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
'Add By Sindy 2009/08/31
Dim strText119 As String
Dim i As Integer, j As Integer
Dim strCP10(1 To 4) As String
Dim blnCancel As Boolean
Dim bolNoAsk As Boolean 'Add by Morgan 2011/3/28
'Add by Amy 2015/10/22 +新客戶確認臺灣地址格式
Dim strZipCode As String, strAddr As String, strCityN As String, strIndArea As String, strNewArea As String, strROC As String
Dim bolTWSame As Boolean, bolMany As Boolean, intArea As Integer
Dim intTW As Integer 'Add by Amy 2017/06/07 for 判斷台灣地址回傳用
   
   If Index = 6 Then
      m_strSys = CheckSys(Text1(6)) 'Add By Sindy 2022/8/30
      
      'Add By Sindy 2022/10/14
      FrameP.Visible = False: FrameT.Visible = False: FrameL.Visible = False
      If m_strSys = "1" Or m_strSys = "5" Then
         FrameP.Visible = True
      ElseIf m_strSys = "2" Or m_strSys = "6" Then
         FrameT.Visible = True
      Else
         FrameL.Visible = True
      End If
      '2022/10/14 END
   End If
   
   If m_blnCallPrint = True Then Exit Sub 'Add By Sindy 2014/7/25
   
    Select Case Index
    Case 6 '系統類別
        Combo1_LostFocus 0
        'add by nickc 2007/03/23
        If Index = 6 Then
            ChkPCT.Visible = False
            If Text1(6) = "P" Or Text1(6) = "CFP" Then
               ChkPCT.Visible = True
            End If
            'Add By Sindy 2011/6/7 系統類別為L,CFL且為新案時,可輸入案件屬性
            Frame19.Visible = False
            'Modified by Morgan 2020/6/22 +FCL
            If (Text1(6) = "L" Or Text1(6) = "CFL" Or Text1(6) = "FCL") And Option1(0).Value = True Then
               Frame19.Visible = True
            End If
            '2011/6/7 End
            
            Call SetFrmCol 'Add By Sindy 2022/12/26
            'Add By Sindy 2014/7/16
            If Me.Option1(0).Value = True Then
            '2014/7/16 END
               Call SetCombo6 'Add By Sindy 2012/6/6
               Call SetCombo5 'Add By Sindy 2012/9/3
            End If
            
            'Added by Lydia 2020/03/30
            If Text1(6).Text <> Text1(6).Tag Then
                Call SetCombo4 '設定收據公司別(簡稱)
                'Modified by Lydia 2020/04/08 改成共用模組
                'If ChkSalesL(Text1(6).Text, Text1(10).Text) = False Then  '法務案(L、CFL)及顧問案LA之智權人員只能是法律所人員
                'Modified by Morgan 2020/4/17 非法務人員收法務案改自動收文介紹案源
                'If PUB_ChkSalesL(Text1(6).Text, Text1(10).Text) = False Then
                'End If
                SrcSetButton
                'end 2020/4/17
                'Added by Lydia 2021/02/24  CFT申請案：欄位「商標英文大寫」
                If Text7.Text <> "" Then
                    Text7.Text = UCase(Text7.Text)
                End If
                'end 2021/02/24
            End If
            'end 2020/03/30
        End If
        
        'Modify By Sindy 2014/2/21 Mark
'        'Add By Sindy 2010/4/29
'        '商標 or 專利
'        If Text1(6) = "T" Or Text1(6) = "TF" Or Text1(6) = "FCT" Or _
'            Text1(6) = "TS" Or Text1(6) = "S" Or Text1(6) = "CFT" Or _
'            Text1(6) = "P" Or Text1(6) = "CFP" Or Text1(6) = "PS" Or Text1(6) = "CPS" Then
            Label1(90).Visible = True
            Combo4.Visible = True
'        Else
'            Label1(90).Visible = False
'            Combo4.Visible = False
'            Combo4.ListIndex = 0
'        End If
        '2010/4/29 End
        
        'Add By Sindy 2009/08/31
        '商標
        If Text1(6) = "T" Or Text1(6) = "TF" Or Text1(6) = "FCT" Or _
            Text1(6) = "TS" Or Text1(6) = "S" Or Text1(6) = "CFT" Then
            SSTab1.TabEnabled(4) = True
         Else
            SSTab1.TabEnabled(4) = False
        End If
                 
        '專利
        Dim bolMsgQ As Boolean, bolAnsYN As Boolean
        bolMsgQ = False   '是否有詢問過
        bolAnsYN = False '是否要保留
        If Text1(6) = "P" Or Text1(6) = "CFP" Or Text1(6) = "PS" Or Text1(6) = "CPS" Then
            If Text1(119) <> "" Then
               strText119 = ""
               arrCaseProperty = Split(Text1(119).Text, vbCrLf)
               '判斷是否需要詢問
               For iiiii = 0 To UBound(arrCaseProperty)
                  If (InStr(1, arrCaseProperty(iiiii), "商品類別：") <> 0 And _
                        Len("商品類別：") <> Len(arrCaseProperty(iiiii))) Or _
                     (InStr(1, arrCaseProperty(iiiii), "商品名稱：") <> 0 And _
                        Len("商品名稱：") <> Len(arrCaseProperty(iiiii))) Or _
                     (InStr(1, arrCaseProperty(iiiii), "註冊號數：") <> 0 And _
                        Len("註冊號數：") <> Len(arrCaseProperty(iiiii))) Then
                     If bolMsgQ = False Then
                        If MsgBox("商品資料已輸入,是否要保留?", vbExclamation + vbYesNo) = vbYes Then
                           bolAnsYN = True
                        End If
                        bolMsgQ = True
                     End If
                  End If
               Next iiiii
               '開始截取資料
               For iiiii = 0 To UBound(arrCaseProperty)
                  If InStr(1, arrCaseProperty(iiiii), "商品類別：") <> 0 Or _
                     InStr(1, arrCaseProperty(iiiii), "商品名稱：") <> 0 Or _
                     InStr(1, arrCaseProperty(iiiii), "註冊號數：") <> 0 Then
                     If bolAnsYN = True Then
                        strText119 = strText119 & arrCaseProperty(iiiii) & vbCrLf
                     End If
                  Else
                     If arrCaseProperty(iiiii) <> "" Then
                        strText119 = strText119 & arrCaseProperty(iiiii) & vbCrLf
                     End If
                  End If
               Next iiiii
               If Text1(119).Tag = Text1(119) Then Text1(119).Tag = strText119 'Added by Morgan 2020/5/7
               Me.Text1(119).Text = strText119
            End If
        End If
        '2009/08/31 End
'        '2010/1/6 ADD BY SONIA 抓台灣發明案實審規費
'        m_416Fee = 0
'        '2010/8/17 MODIFY BY SONIA
'        'm_416Fee = GetPatentOfficialFee(Text1(6), "416", "", m_strPA08, stCountry, "", , Text1(7), IIf(Text1(8) = "", "0", Text1(8)), IIf(Text1(9) = "", "00", Text1(9)))
'        m_416Fee = GetPatentOfficialFee(Text1(6), "416", "", m_strPA08, stCountry, m_strPA16, m_strPA14, Text1(7), IIf(Text1(8) = "", "0", Text1(8)), IIf(Text1(9) = "", "00", Text1(9)))
'        '2010/1/6 END
        
        'Added by Lydia 2021/03/29 ACS案件收文與點數及營業稅：不管新舊案，系統類別ACS跳離時都先彈訊息提醒
        If Text1(6).Tag <> Text1(6).Text And Text1(6) = "ACS" And strSrvDate(1) >= strACSdate1 Then
            MsgBox "ACS案件官方規費請與服務費分開填案件性質，" & vbCrLf & "官方規費請填案件性質706代收代付！", vbInformation, "ACS案件收文"
        End If
        'end 2021/03/29
        
        Text1(6).Tag = Text1(6).Text   '2011/9/15 add by sonia
                
         'Added by Morgan 2013/1/15
         If Text1(6) = "P" And Left(Trim(Combo1(0)), 3) = "000" Then
            SSTab3.TabVisible(1) = True
         Else
            SSTab3.TabVisible(1) = False
         End If
         'end 2013/1/15

    Case 7, 9 '本所案號
        'Add By Sindy 2010/6/10
        If Me.ActiveControl.Name = "Text1" Then
            If (Me.ActiveControl.Index = 8 Or Me.ActiveControl.Index = 9) Then
               '正常步驟，繼續操作...
               Exit Sub
            End If
        End If
        '2010/6/10 End
        
        If Me.Text1(7).Text = "" Then
            Me.Combo1(0).Text = ""
            Me.Combo1(1).Text = ""
'            Me.Combo1(2).Text = ""
'            Me.Combo1(3).Text = ""
'            Me.Combo1(4).Text = ""
            Me.Text1(11).Text = ""
            'Add by Lydia 2014/12/22 清空費用
            For i = 101 To 112
               Text1(i).Text = ""
            Next i
                
            Combo5.Text = "" 'Add By Sindy 2010/10/28
            ClearCustTxt 1
            ClearCustTxt 2
            ClearCustTxt 3
            ClearCustTxt 4
            ClearCustTxt 5
            m_strPA08 = "1"    '2010/3/1 MODIFY BY SONIA 改預設"1",原為""
            m_strPA16 = ""
            m_strTM15 = ""
            m_strTM12 = ""
            m_strPA14 = ""     '2010/8/17 add by sonia
            Call GetReceiptTitle 'Add by Morgan 2004/6/11
            Exit Sub
        Else
            'Me.Option3(1).Value = True
            Me.Option31(1).Value = True
        End If
        '若為舊案
        If Me.Option1(1).Value = True And Me.Text1(6).Text <> "" And Me.Text1(7).Text <> "" Then
            If Me.Text1(8).Text = "" Then Me.Text1(8).Text = "0"
            If Me.Text1(9).Text = "" Then Me.Text1(9).Text = "00"
        End If
        
        m_strCaseCPM = GetAllCaseCPM 'Add By Sindy 2022/9/6 取得案件性質代碼
        'Add By Sindy 2022/9/28 L-888888案號,設定對造為本所客戶為是。
        If Text1(6) = "L" And Text1(7) = "888888" Then
            Check7(0).Value = 1 '是
            Check7(0).Enabled = False
            Check7(1).Enabled = False
        Else
            Check7(0).Enabled = True
            Check7(1).Enabled = True
        End If
        '2022/9/38 END
        
        'Added by Lydia 2021/05/07 ACS智財顧問專業分配比例管制
        If Text1(6).Text & Text1(7).Text & Text1(8).Text & Text1(9).Text <> Text1(6).Tag & Text1(7).Tag & Text1(8).Tag & Text1(9).Tag Then
             Call SetACS112data
        End If
        'end 2021/05/07
        
        'Added by Lydia 2020/02/05
        Call setFrame21  '如果連續收同一種案件,需要重設文件齊備、是否急件、查名齊備
        '暫存本所案號
        Me.Text1(7).Tag = Me.Text1(7).Text
        Me.Text1(8).Tag = Me.Text1(8).Text
        Me.Text1(9).Tag = Me.Text1(9).Text
        'end 2020/02/05
        
        '若有輸入本所案號
        If Me.Text1(6).Text <> "" And Me.Text1(7).Text <> "" Then
            
            If Text1(6) = "LA" Then Me.Text1(11).Enabled = True 'Add By Sindy 2011/6/17
            
            '2008/3/21 modify by sonia 商標案加抓tm15,tm12
            'Modify by Morgan 2008/8/1 加抓個案申請人聯絡人編號
            'Modify by Sindy 2009/05/21 加抓商標檔申請人2,3,4,5及服務檔申請人4,5
            '2010/3/31 MODIFY BY SONIA 專利案加申請日
            '2010/8/17 modify by sonia 專利案加公告日pa14
            'Modify By Sindy 2011/2/21 增加LC43,LC44,LC45,LC46,HC24,HC25,HC26,HC27
            'Modify By Sindy 2013/12/16 +pa161
            StrSQLa = "Select NA01||' '||NA03, Nvl(PA05, Nvl(PA06, PA07)), PA26, PA27, PA28, PA29, PA30, PA08, PA16, PA48, '', '',PA149,PA10,PA75,PA14,PA158,pa161 From Patent, Nation,Customer Where PA09=NA01(+) And PA01='" & Me.Text1(6).Text & "' And PA02='" & Me.Text1(7).Text & "' And PA03='" & Me.Text1(8).Text & "' And PA04='" & Me.Text1(9).Text & "' and cu01(+)=substr(pa26,1,8) and cu02(+)=substr(pa26,9,1) "
            StrSQLa = StrSQLa & " Union Select NA01||' '||NA03, Nvl(TM05, Nvl(TM06, TM07)), TM23, TM78, TM79, TM80, TM81, '', '', TM35,TM15,TM12,TM123,0,TM44,0,'',tm130 as pa161 From Trademark, Nation,Customer Where TM10=NA01(+) And TM01='" & Me.Text1(6).Text & "' And TM02='" & Me.Text1(7).Text & "' And TM03='" & Me.Text1(8).Text & "' And TM04='" & Me.Text1(9).Text & "' and cu01(+)=substr(TM23,1,8) and cu02(+)=substr(TM23,9,1)  "
            StrSQLa = StrSQLa & " Union Select NA01||' '||NA03, Nvl(LC05, Nvl(LC06, LC07)), LC11, LC43,LC44,LC45,LC46, '', '', LC17, '', '',LC42,0,LC22,0,'',lc48 as pa161 From Lawcase, Nation,Customer Where LC15=NA01(+) And LC01='" & Me.Text1(6).Text & "' And LC02='" & Me.Text1(7).Text & "' And LC03='" & Me.Text1(8).Text & "' And LC04='" & Me.Text1(9).Text & "' and cu01(+)=substr(LC11,1,8) and cu02(+)=substr(LC11,9,1) "
            StrSQLa = StrSQLa & " Union Select NA01||' '||NA03, HC06, HC05, HC24,HC25,HC26,HC27, '', '', '', '', '',HC23,0,'',0,'','' as pa161 From Hirecase, Nation,Customer Where '000'=NA01(+) And HC01='" & Me.Text1(6).Text & "' And HC02='" & Me.Text1(7).Text & "' And HC03='" & Me.Text1(8).Text & "' And HC04='" & Me.Text1(9).Text & "' and cu01(+)=substr(HC05,1,8) and cu02(+)=substr(HC05,9,1) "
            StrSQLa = StrSQLa & " Union Select NA01||' '||NA03, Nvl(SP05, Nvl(SP06, SP07)), SP08, SP58, SP59, SP65, SP66, '', '', SP29, '', '',SP78,0,SP26,0,'',sp85 as pa161 From Servicepractice, Nation,Customer Where SP09=NA01(+) And SP01='" & Me.Text1(6).Text & "' And SP02='" & Me.Text1(7).Text & "' And SP03='" & Me.Text1(8).Text & "' And SP04='" & Me.Text1(9).Text & "' and cu01(+)=substr(SP08,1,8) and cu02(+)=substr(SP08,9,1) "
            rsA.CursorLocation = adUseClient
            rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
            If rsA.RecordCount > 0 Then
                
               'Add by Amy 2016/09/01 當601異議/603評定/605廢止原新商標案改收舊案時不彈
               If bolNotShow = False Then
                    'Add by Morgan 2011/3/28 考慮輸多張接洽單只有改案號情形
                    If m_lstCaseNo <> "" And m_lstCaseNo <> Text1(6).Text & "-" & Text1(7).Text & "-" & Text1(8).Text & "-" & Text1(9).Text Then
                       'Add By Sindy 2014/5/23 外部呼叫時一定要先清除再使用
                       'Modified by Morgan 2020/6/29
                       'If bolExternalCall = True Then
                       If bolExternalCall = True Or strLCaseNo(1) = "L" Then
                       'end 2020/6/29
                          cmdClear True
                       Else
                       '2014/5/23 END
                          If MsgBox("是否清除原畫面資料？", vbYesNo + vbDefaultButton2) = vbYes Then
                             cmdClear True
                          '2013/10/7 ADD BY SONIA 改案號無論如何都清本所期限及法定期限,否則操作人員若不清畫面會有不同案號不同期限的問題(先收T-086187延展,再收T-134151延展)
                          Else
                            ' Text1(1) = "": Text1(3) = ""
                            'Add by Lydia 2015/01/06 清除費用資料
                            Call FreeClear
                          '2013/10/7 END
                          End If
                       End If
                       bolNoAsk = True
                    Else
                       bolNoAsk = False
                    End If
               End If
               'end 2016/09/01
               'Add By Sindy 2013/12/16 舊案若為智權公司則收據公司別設定為智權公司則鎖住
               If strSrvDate(1) >= InvoiceStartDate Then
                  Combo4.Enabled = True
                  If "" & rsA.Fields("pa161").Value = "J" Then
                     'Modified by Lydia 2020/03/30
                     'If Combo4.ListCount < 4 Then
                     '   Combo4.AddItem "智權公司"
                     'End If
                     'Combo4.ListIndex = 3
                     If Combo4.ListCount < m_CompJforIdx + 1 Then
                        Combo4.AddItem m_CompNameJ
                     End If
                     Combo4.ListIndex = m_CompJforIdx
                     'end 2020/03/30
                     Combo4.Enabled = False
                  Else
                     'Modified by Lydia 2020/03/30
                     'If Combo4.ListCount = 4 Then
                     '   Combo4.RemoveItem 3
                     If Combo4.ListCount = m_CompJforIdx + 1 Then
                        Combo4.RemoveItem m_CompJforIdx
                     'end 2020/03/30
                     End If
                  End If
               End If
               '2013/12/16 END
               
               '2011/10/19 cancel by sonia移到下面去
               'm_lstCaseNo = Text1(6).Text & "-" & Text1(7).Text & "-" & Text1(8).Text & "-" & Text1(9).Text
                '專利種類
                m_strPA08 = "" & rsA.Fields(7).Value
                'Add By Sindy 2014/7/15
                Call SetCombo6 'Add By Sindy 2014/7/16
                If m_strPA08 <> "" Then
                  If Combo6.Visible = True Then 'Modify By Sindy 2014/7/18 +if
                     Combo6.ListIndex = m_strPA08 - 1
                  End If
                End If
                '案件屬性
                Call SetCombo5
                If "" & rsA.Fields("pa158").Value <> "" Then
                  If Combo5.Visible = True Then 'Modify By Sindy 2014/7/18 +if
                     Combo5.ListIndex = "" & rsA.Fields("pa158").Value - 1
                  End If
                End If
                '2014/7/15 END
                m_strPA16 = "" & rsA.Fields(8).Value
                '2010/8/17 add by sonia
                m_strPA14 = "" & rsA.Fields(15).Value
                If m_strPA14 = "0" Then m_strPA14 = ""
                '2010/8/17 END
                m_strPA10 = "" & rsA.Fields(13).Value  '2010/3/31 ADD BY SONIA
                '2008/3/21 ADD BY SONIA
                m_strTM15 = "" & rsA.Fields(10).Value
                m_strTM12 = "" & rsA.Fields(11).Value
                '2008/3/21 END
                'Add By Sindy 2010/7/9
                If m_strPA08 <> "" Then '代表有pa資料
                  ReDim pa(1 To TF_PA) As String
                  Call PUB_ReadPatentData(pa(), Text1(6), Text1(7), Text1(8), Text1(9))
                End If
                '舊案
                Me.Option1(0).Value = False
                Me.Option1(1).Value = True
                '申請國
                Me.Combo1(0).Text = "" & rsA.Fields(0).Value
                stCountry = Trim(Mid(Combo1(0).Text, 1, 4))
                
'                'Modified by Lydia 2014/12/25 電子化(2015/1/1)後可減少列印(大陸案例外)
'                Call settxtPCnt
                
                SetOpt81 stCountry 'Added by Morgan 2013/4/9
                
                '主題
                Me.Text1(11).Text = "" & rsA.Fields(1).Value
                
                'Add By Sindy 2010/10/28 案件屬性
                'Modified by Lydia 2019/08/12 +""
                'If IsNull(rsA.Fields(16).Value) Then
                If "" & rsA.Fields(16).Value = "" Then
                  Combo5 = ""
                Else
                  'Modified by Lydia 2024/04/24 +專利種類IIf(Combo6.Text <> "", Left(Combo6, 1), "")
                  Combo5 = rsA.Fields(16).Value + "." + PUB_GetCaseAttributeName(rsA.Fields(16).Value, IIf(Combo6.Text <> "", Left(Combo6, 1), ""))
                End If
                '2010/10/28 End
                
                'Add By Sindy 2010/5/27
                '代理人
                ClearFagentTxt
                If "" & rsA.Fields(14).Value <> "" Then
                    SetFagentTxt "" & rsA.Fields(14).Value
                    'Added by Lydia 2019/01/30 大至台案件彈出代理人D/N備註
                    If OptChoose(1).Value = True Then
                        If Left(Text1(6), 1) = "T" And m_FA110 <> "" Then '商標D/N備註
                            MsgBox "代理人D/N備註：" & vbCrLf & m_FA110, vbExclamation, "代理人D/N備註"
                        ElseIf m_FA45 <> "" Then  '專利D/N備註
                            MsgBox "代理人D/N備註：" & vbCrLf & m_FA45, vbExclamation, "代理人D/N備註"
                        End If
                    End If
                    'end 2019/01/30
                End If
                
                '申請人1
                ClearCustTxt 1
                If "" & rsA.Fields(2).Value <> "" Then
                    SetCustTxt 1, "" & rsA.Fields(2).Value
                    'Add by Morgan 2008/8/1
                    'Modified by Morgan 2022/1/20
                    PUB_AddContact rsA.Fields(2), cboContact(1), "" & rsA.Fields("PA149"), True, True, m_strContactList(1)
                    'Added by Lydia 2019/08/15 提醒及列印顧問服務件數,放在最後會出現未知錯誤(P.S.陳建宏輸入LA-003308-0-00在輸入完４欄位用滑鼠跳回第２欄，有程式錯誤messag之後才彈"聘任期間xxxxx")
                    Call SetLAdata
                End If
                '申請人2
                ClearCustTxt 2
                If "" & rsA.Fields(3).Value <> "" Then
                    SetCustTxt 2, "" & rsA.Fields(3).Value
                    'Add by Morgan 2008/8/1
                    'Modified by Morgan 2022/1/20 改2.0
                    PUB_AddContact rsA.Fields(3), cboContact(2), , True, True, m_strContactList(2)
                End If
                '申請人3
                ClearCustTxt 3
                If "" & rsA.Fields(4).Value <> "" Then
                    SetCustTxt 3, "" & rsA.Fields(4).Value
                    'Add by Morgan 2008/8/1
                    'Modified by Morgan 2022/1/20 改2.0
                    PUB_AddContact rsA.Fields(4), cboContact(3), , True, True, m_strContactList(3)
                End If
                '申請人4
                ClearCustTxt 4
                If "" & rsA.Fields(5).Value <> "" Then
                    SetCustTxt 4, "" & rsA.Fields(5).Value
                    'Add by Morgan 2008/8/1
                    'Modified by Morgan 2022/1/20 改2.0
                    PUB_AddContact rsA.Fields(5), cboContact(4), , True, True, m_strContactList(4)
                End If
                '申請人5
                ClearCustTxt 5
                If "" & rsA.Fields(6).Value <> "" Then
                    SetCustTxt 5, "" & rsA.Fields(6).Value
                    'Add by Morgan 2008/8/1
                    'Modified by Morgan 2022/1/20 改2.0
                    PUB_AddContact rsA.Fields(6), cboContact(5), , True, True, m_strContactList(5)
                End If
                Combo1_LostFocus 0
                
                'Add By Sindy 2015/7/27 原在按下列印鍵時提醒,改先提醒
                IsSpecCu = False
                SpecCUName = ""
                SpecMemo = ""
                IsCuMemo = False
                CuMemoName = ""
                CuMemo = ""
                SetCuData IsSpecCu, SpecCUName, SpecMemo, IsCuMemo, CuMemoName, CuMemo
                If m_blnCallPrint = False Then
                   If IsCuMemo Then
                      MsgBox CuMemoName & "此客戶有業務備註！！", vbInformation, "業務備註！"
                   End If
                   'Added by Morgan 2020/3/16
                   'P台灣申請中案件(無公告日,專用期間)，最後發文之A或B類(不續辦、閉卷、取消收文除外)之出名代理人有76012桂所長的案件，檢查若該申請人無總委案件則顯示訊息
                   If Text1(6) = "P" And stCountry = "000" Then
                     strExc(1) = ""
                     If PUB_ChkIsGuiCase(pa(1), pa(2), pa(3), pa(4), , strExc(1)) Then
                        If strExc(1) <> "" Then
                           MsgBox "本案須辦理變更代理人為閻+林,請協助簽署下列申請人總(個)委任書！" & vbCrLf & vbCrLf & strExc(1), vbExclamation
                        End If
                     End If
                   End If
                   'end 2020/3/16
                End If
                '2015/7/27 END
                
                '客戶案號(型號)
                Me.Text1(18).Text = "" & rsA.Fields(9).Value
                If Me.Text1(18).Text = "" Then Me.Text1(18).Enabled = True
                
                '2011/10/19 add by sonia舊案檢查申請地址與客戶中文地址是否相同,若不同則提醒
                If m_lstCaseNo = "" Or (m_lstCaseNo <> "" And m_lstCaseNo <> Text1(6).Text & "-" & Text1(7).Text & "-" & Text1(8).Text & "-" & Text1(9).Text) Then
                   '申請人地址
                   For i = 1 To 5
                      If Text1(27 + (i - 1) * 16) <> "" Then
                         Call CheckCustAddr(Text1(27 + (i - 1) * 16), Text1(12 + (i - 1) * 16), i)
                      End If
                   Next i
                End If
                
                m_lstCaseNo = Text1(6).Text & "-" & Text1(7).Text & "-" & Text1(8).Text & "-" & Text1(9).Text  '2011/10/19 從上面移下來
                '2011/10/19 end
                
            Else
                m_strPA08 = "1"    '2010/3/1 MODIFY BY SONIA 改預設"1",原為""
                m_strPA16 = ""
                m_strPA10 = ""     '2010/3/31 ADD BY SONIA
                m_strTM15 = ""
                m_strTM12 = ""
                m_strPA14 = ""     '2010/8/17 ADD BY SONIA
                MsgBox "查無此本所案號資料!!!", vbExclamation + vbOKOnly
                Me.Combo1(0).Text = ""
                Me.Combo1(1).Text = ""
'                Me.Combo1(2).Text = ""
'                Me.Combo1(3).Text = ""
'                Me.Combo1(4).Text = ""
                Me.Combo1(1).Clear
'                Me.Combo1(2).Clear
'                Me.Combo1(3).Clear
'                Me.Combo1(4).Clear
                Me.Text1(11).Text = ""
                Combo5.Text = "" 'Add By Sindy 2010/10/28
                ClearCustTxt 1
                ClearCustTxt 2
                ClearCustTxt 3
                ClearCustTxt 4
                ClearCustTxt 5
                'Add by Lydia 2014/12/22 清空費用
                For i = 101 To 112
                   Text1(i).Text = ""
                Next i
                Frame605.Visible = False
                
                Me.Text1(18).Text = ""
                'Modified by Lydia 2021/06/18 debug:接洽單查詢及列印使用 + And Me.Visible = True
                If Me.Text1(6).Enabled = True And Me.Visible = True Then Me.Text1(6).SetFocus

                Text1_GotFocus 6
            End If
            If rsA.State <> adStateClosed Then rsA.Close
            Set rsA = Nothing
        End If
        Call GetReceiptTitle 'Add by Morgan 2004/6/11
'        '2010/1/6 ADD BY SONIA 抓台灣發明案實審規費
'        m_416Fee = 0
'        '2010/8/17 MODIFY BY SONIA
'        'm_416Fee = GetPatentOfficialFee(Text1(6), "416", "", m_strPA08, stCountry, "", , Text1(7), IIf(Text1(8) = "", "0", Text1(8)), IIf(Text1(9) = "", "00", Text1(9)))
'        m_416Fee = GetPatentOfficialFee(Text1(6), "416", "", m_strPA08, stCountry, m_strPA16, m_strPA14, Text1(7), IIf(Text1(8) = "", "0", Text1(8)), IIf(Text1(9) = "", "00", Text1(9)))
'        '2010/1/6 END
    'Add by Amy 2015/11/02 +新客戶確認臺灣地址格式(舊客戶欄位鎖住)
    Case 26, 42, 58, 74, 90 '聯絡地址
    Case 27, 43, 59, 75, 91 '申請地址
    Case 101 '費用, 104, 107, 110
        If Me.Combo1((Index - 101) / 3 + 1).Text <> "" Then
            If Me.Text1(Index).Text = "" Then Me.Text1(Index).Text = "0"
            'edit by nickc 2005/03/02 控制小數點
            Me.Text1(Index + 2).Text = Format((CDbl(Val(Me.Text1(Index).Text)) - CDbl(Val(Me.Text1(Index + 1).Text))) / 1000, "####0.000")
            Call Check613612 'Add By Sindy 2012/8/31
            'Added by Lydia 2021/03/29 ACS案件收文與點數及營業稅
            If strSrvDate(1) >= strACSdate1 Then
               'Modify By Sindy 2022/8/31
               'Call SetACSautoFee(IIf(Index = 101, 1, IIf(Index = 104, 2, IIf(Index = 107, 3, 4))))
               Call SetACSautoFee(1)
            End If
            'end 2021/03/29
        Else
            Me.Text1(Index).Text = ""
            Me.Text1(Index + 1).Text = ""
            Me.Text1(Index + 2).Text = ""
        End If
    Case 102 '規費, 105, 108, 111
        If Me.Combo1((Index - 102) / 3 + 1).Text <> "" Then
            If Me.Text1(Index).Text = "" Then Me.Text1(Index).Text = "0"
            'edit by nickc 2005/03/02 控制小數點
            Me.Text1(Index + 1).Text = Format((CDbl(Val(Me.Text1(Index - 1).Text)) - CDbl(Val(Me.Text1(Index).Text))) / 1000, "####0.000")
            Call Check613612 'Add By Sindy 2012/8/31
        Else
            Me.Text1(Index - 1).Text = ""
            Me.Text1(Index).Text = ""
            Me.Text1(Index + 1).Text = ""
        End If
    Case 103 '點數, 106, 109, 112
        Text1(Index - 1).Tag = Text1(Index - 1) 'Add by Morgan 2011/3/7
        If Me.Combo1((Index - 103) / 3 + 1).Text <> "" Then
            If Me.Text1(Index).Text = "" Then Me.Text1(Index).Text = "0"
            'edit by nickc 2005/03/02 控制小數點
            Me.Text1(Index - 1).Text = CDbl(Val(Me.Text1(Index - 2).Text)) - CDbl(Val(Me.Text1(Index).Text) * 1000)
            Call Check613612 'Add By Sindy 2012/8/31
        Else
            Me.Text1(Index - 2).Text = ""
            Me.Text1(Index - 1).Text = ""
            Me.Text1(Index).Text = ""
        End If
        
'         If Text1(Index - 1) <> Text1(Index - 1).Tag Then 'Add by Morgan 2011/3/7 規費有變才做,否則會重複提醒
'            '2010/11/19 add by sonia 重新檢查規費
'            Text1_Validate Index - 1, blnCancel
'            If blnCancel = True Then
'               If Me.Text1(Index - 1).Enabled = True Then Me.Text1(Index - 1).SetFocus
'            End If
'            '2010/11/19 end
'         End If
    End Select
End Sub

'Add By Sindy 2012/8/31 案件性質為613補充答辯或612補充理由時，若點數>=8時預設為會稿，反之預設為不會稿且鎖住欄位
'Modify By Sindy 2013/3/12 改為點數>=5時預設為會稿，反之預設為不會稿且鎖住欄位
Private Sub Check613612()
Dim ii As Integer

   'Modified by Lydia 2018/12/10 判斷會稿才預設
   'If Frame21.Visible = True Then
   If Frame21.Visible = True And Frame23.Visible = True Then
      OptEP34(0).Enabled = True
      OptEP34(1).Enabled = True
      'Added by Lydia 2022/07/15 TC案之文件齊備日管控; T大陸案之齊備日管控
      If Trim(Text1(6)) = "TC" Then
          '台灣TC案不會稿(不顯示),但大陸TC案要會稿 => setFrame21有設定
      Else
      'end 2022/07/15
        If Trim(Left(Trim(Combo1(1)), 4)) = "613" Or Trim(Left(Trim(Combo1(1)), 4)) = "612" Then
           If Text1(103) >= 5 Then
              OptEP34(0).Value = True '會稿
           Else
              OptEP34(1).Value = True '不會稿
              OptEP34(0).Enabled = False
              OptEP34(1).Enabled = False
           End If
        Else
            For ii = 1 To GridCase.Rows - 1
                If Trim(GridCase.TextMatrix(ii, 1)) <> "" Then
                   If Trim(Left(Trim(GridCase.TextMatrix(ii, 1)), 4)) = "613" Or Trim(Left(Trim(GridCase.TextMatrix(ii, 1)), 4)) = "612" Then
                      If Val(Trim(GridCase.TextMatrix(ii, 4))) >= 5 Then
                         OptEP34(0).Value = True '會稿
                      Else
                         OptEP34(1).Value = True '不會稿
                         OptEP34(0).Enabled = False
                         OptEP34(1).Enabled = False
                      End If
                      Exit For
                   End If
                Else
                   Exit For
                End If
            Next ii
        End If
        'Modify By Sindy 2022/8/29 Mark
'        If Trim(Left(Trim(Combo1(2)), 4)) = "613" Or Trim(Left(Trim(Combo1(2)), 4)) = "612" Then
'           If Text1(106) >= 5 Then
'              OptEP34(0).Value = True '會稿
'           Else
'              OptEP34(1).Value = True '不會稿
'              OptEP34(0).Enabled = False
'              OptEP34(1).Enabled = False
'           End If
'        End If
'        If Trim(Left(Trim(Combo1(3)), 4)) = "613" Or Trim(Left(Trim(Combo1(3)), 4)) = "612" Then
'           If Text1(109) >= 5 Then
'              OptEP34(0).Value = True '會稿
'           Else
'              OptEP34(1).Value = True '不會稿
'              OptEP34(0).Enabled = False
'              OptEP34(1).Enabled = False
'           End If
'        End If
'        If Trim(Left(Trim(Combo1(4)), 4)) = "613" Or Trim(Left(Trim(Combo1(4)), 4)) = "612" Then
'           If Text1(112) >= 5 Then
'              OptEP34(0).Value = True '會稿
'           Else
'              OptEP34(1).Value = True '不會稿
'              OptEP34(0).Enabled = False
'              OptEP34(1).Enabled = False
'           End If
'        End If
      End If 'Added by Lydia 2022/07/15
   End If
End Sub

'Add By Sindy 2011/11/11
Private Sub SetFrame16()
Dim bolNewCase As Boolean
   
   m_strCaseCPM = GetAllCaseCPM(bolNewCase) 'Add By Sindy 2022/8/29 取得案件性質代碼
   
   'CFP且非英語系國家且為新申請案案件性質時,開放Frame16欄位
   Frame16.Visible = False
   If Text1(6) = "CFP" And m_NA59 = "N" And ( _
      (InStr(NewCasePtyList, Trim(Left(Trim(Combo1(1)), 4))) > 0 And Trim(Left(Trim(Combo1(1)), 4)) <> "") Or _
      bolNewCase = True _
      ) Then
'      Or _
'      (InStr(NewCasePtyList, Trim(Left(Trim(Combo1(2)), 4))) > 0 And Trim(Left(Trim(Combo1(2)), 4)) <> "") Or _
'      (InStr(NewCasePtyList, Trim(Left(Trim(Combo1(3)), 4))) > 0 And Trim(Left(Trim(Combo1(3)), 4)) <> "") Or _
'      (InStr(NewCasePtyList, Trim(Left(Trim(Combo1(4)), 4))) > 0 And Trim(Left(Trim(Combo1(4)), 4)) <> "")
      Frame16.Visible = True
      '但若已收938超頁費,939超項費,則不顯示Frame16欄位
'      If Trim(Left(Trim(Combo1(1)), 4)) = "938" Or _
'         Trim(Left(Trim(Combo1(2)), 4)) = "938" Or _
'         Trim(Left(Trim(Combo1(3)), 4)) = "938" Or _
'         Trim(Left(Trim(Combo1(4)), 4)) = "938" Or _
'         Trim(Left(Trim(Combo1(1)), 4)) = "939" Or _
'         Trim(Left(Trim(Combo1(2)), 4)) = "939" Or _
'         Trim(Left(Trim(Combo1(3)), 4)) = "939" Or _
'         Trim(Left(Trim(Combo1(4)), 4)) = "939" Then
      If InStr(m_strCaseCPM, "938") > 0 Or Trim(Left(Trim(Combo1(1)), 4)) = "938" _
         Or InStr(m_strCaseCPM, "939") > 0 Or Trim(Left(Trim(Combo1(1)), 4)) = "939" Then
         Frame16.Visible = False
      End If
   Else
      Option3(0).Value = False
      Option3(1).Value = False
   End If
End Sub

'Add By Sindy 2013/2/25
Private Sub SetFrame27()
   
   'm_strCaseCPM = GetAllCaseCPM() 'Add By Sindy 2022/8/29 取得案件性質代碼
   
   '填寫CFT移轉案之接洽記錄單時，不管新舊案一定要勾選才可列印接洽單
   Frame27.Visible = False
   If Text1(6) = "CFT" And ( _
      (InStr("501", Trim(Left(Trim(Combo1(1)), 4))) > 0 And Trim(Left(Trim(Combo1(1)), 4)) <> "") Or _
      (InStr(m_strCaseCPM, "501") > 0 And m_strCaseCPM <> "") _
      ) Then
'      Or _
'      (InStr("501", Trim(Left(Trim(Combo1(2)), 4))) > 0 And Trim(Left(Trim(Combo1(2)), 4)) <> "") Or _
'      (InStr("501", Trim(Left(Trim(Combo1(3)), 4))) > 0 And Trim(Left(Trim(Combo1(3)), 4)) <> "") Or _
'      (InStr("501", Trim(Left(Trim(Combo1(4)), 4))) > 0 And Trim(Left(Trim(Combo1(4)), 4)) <> "")
      Frame27.Visible = True
   Else
      Option8(0).Value = False
      Option8(1).Value = False
   End If
End Sub

'Add By Sindy 2012/3/6
Private Sub SetFrame20()
Dim bolNewCase As Boolean
   
   m_strCaseCPM = GetAllCaseCPM(bolNewCase) 'Add By Sindy 2022/8/29 取得案件性質代碼
   
   'Modified by Lydia 2016/09/22
   ''台灣商標T並且案件性質為申請時,開放Frame20欄位
   '商標案並且案件性質為申請時,開放Frame20欄位
   Frame20.Visible = False
   Option6(4).Visible = False  'Added by Lydia 2016/09/22
   
   If (Text1(6) = "T" Or Text1(6) = "CFT" Or Text1(6) = "TF") And ( _
      (Trim(Left(Trim(Combo1(1)), 4)) = "101" And Trim(Left(Trim(Combo1(1)), 4)) <> "") Or _
      bolNewCase = True _
      ) Then
'      Or _
'      (Trim(Left(Trim(Combo1(2)), 4)) = "101" And Trim(Left(Trim(Combo1(2)), 4)) <> "") Or _
'      (Trim(Left(Trim(Combo1(3)), 4)) = "101" And Trim(Left(Trim(Combo1(3)), 4)) <> "") Or _
'      (Trim(Left(Trim(Combo1(4)), 4)) = "101" And Trim(Left(Trim(Combo1(4)), 4)) <> "")
      Frame20.Visible = True
      'Added by Lydia 2016/09/22 非台灣T,TF才有不查名
      If Trim(Left(Combo1(0).Text, 4)) <> "000" And Combo1(0).Text <> "" And (Text1(6) = "T" Or Text1(6) = "TF") Then
         Option6(4).Visible = True
      End If
   Else
      Option6(0).Value = False
      Option6(1).Value = False
      Option6(2).Value = False
      Option6(3).Value = False
      Option6(4).Value = False 'Added by Lydia 2016/09/22
      Option6(5).Value = False 'Added by Lydia 2018/11/13
   End If
End Sub

'Add By Sindy 2012/5/8
Private Sub setFrame21()
Dim intP As Integer 'Added by Lydia 2019/11/05
Dim bolNewCase As Boolean, bolTMdebate As Boolean
   
   m_strCaseCPM = GetAllCaseCPM(bolNewCase, bolTMdebate) 'Add By Sindy 2022/8/29 取得案件性質代碼
   
   'Add By Sindy 2022/10/14 LA顧問案之0顧問聘任,ACS案之112智財顧問
   Label1(140).Visible = False: Text1(139).Visible = False: Text1(140).Visible = False
   If (Text1(6) = "LA" And (m_strCaseCPM = "0" Or Trim(Left(Trim(Combo1(1)), 2)) = "0")) Or _
      (Text1(6) = "ACS" And (InStr(m_strCaseCPM, "112") > 0 Or Trim(Left(Trim(Combo1(1)), 4)) = "112")) Then
       Label1(140).Visible = True: Text1(139).Visible = True: Text1(140).Visible = True
   End If
   '2022/10/14 END
   
   'Added by Lydia 2018/12/10 非專利案隱藏 申請優先權證明書 和 申請技術報告項數
   Label1(88).Visible = False: Text1(4).Visible = False
   Label1(125).Visible = False: Text1(114).Visible = False
   If Text1(6) = "P" Or Text1(6) = "CFP" Then
      'Modify By Sindy 2023/1/18
      If InStr(m_strCaseCPM, "405") > 0 Then '405.申請優先權證明書
      '2023/1/18 END
         Label1(88).Visible = True: Text1(4).Visible = True
      'Modify By Sindy 2023/1/18 421申請技術報告
      'Modify By Sindy 2023/5/12 + 807 第三人申請技術報告
      ElseIf InStr(m_strCaseCPM, "421") > 0 Or _
             InStr(m_strCaseCPM, "807") > 0 Then
      '2023/1/18 END
         Label1(125).Visible = True: Text1(114).Visible = True
      End If
'   Else
'       Label1(88).Visible = False: Text1(4).Visible = False
'       Label1(125).Visible = False: Text1(114).Visible = False
   End If
   
   Frame23.Visible = True  '預設會稿顯示
   Frame42.Visible = False  '查名是否齊備
   Frame43.Visible = True  '預設延期顯示
   Label23.Caption = "資料是否齊備："
   'Modified by Lydia 2020/02/05 判斷完整本所案號; ex.T-222561延期自動收文,但是文件齊備預設為N
   'If Text1(6).Text <> Text1(6).Tag Then '預設清空
   If Text1(6).Text & Text1(7).Text & Text1(8).Text & Text1(9).Text <> Text1(6).Tag & Text1(7).Tag & Text1(8).Tag & Text1(9).Tag Then
        OptEP06(0).Value = False: OptEP06(1).Value = False
        OptEP34(0).Value = False: OptEP34(1).Value = False
'        OptCP122(0).Value = False: OptCP122(1).Value = False
        OptCP143(0).Value = False: OptCP143(1).Value = False
   End If
   'end 2018/12/10
   
   If Val(strSrvDate(1)) >= Val(TMdebateStarDT) Then
      '台灣商標Ｔ案若收文爭議案件性質時,開放Frame21欄位
      Frame21.Visible = False
      'Modified by Lydia 2022/07/15 T大陸案之齊備日管控
      'If Text1(6) = "T" And Left(Trim(Combo1(0)), 3) = "000" And ( _
         (InStr(TMdebate, Trim(Left(Trim(Combo1(1)), 4))) > 0 And Trim(Left(Trim(Combo1(1)), 4)) <> "") Or _
         (InStr(TMdebate, Trim(Left(Trim(Combo1(2)), 4))) > 0 And Trim(Left(Trim(Combo1(2)), 4)) <> "") Or _
         (InStr(TMdebate, Trim(Left(Trim(Combo1(3)), 4))) > 0 And Trim(Left(Trim(Combo1(3)), 4)) <> "") Or _
         (InStr(TMdebate, Trim(Left(Trim(Combo1(4)), 4))) > 0 And Trim(Left(Trim(Combo1(4)), 4)) <> "") _
         ) Then
      If Text1(6) = "T" And InStr("000,020", Left(Trim(Combo1(0)), 3)) > 0 And ( _
         (InStr(TMdebate, Trim(Left(Trim(Combo1(1)), 4))) > 0 And Trim(Left(Trim(Combo1(1)), 4)) <> "") Or _
         bolTMdebate = True _
         ) Then
'         Or _
'         (InStr(TMdebate, Trim(Left(Trim(Combo1(2)), 4))) > 0 And Trim(Left(Trim(Combo1(2)), 4)) <> "") Or _
'         (InStr(TMdebate, Trim(Left(Trim(Combo1(3)), 4))) > 0 And Trim(Left(Trim(Combo1(3)), 4)) <> "") Or _
'         (InStr(TMdebate, Trim(Left(Trim(Combo1(4)), 4))) > 0 And Trim(Left(Trim(Combo1(4)), 4)) <> "")
         Frame21.Visible = True
         '案件性質為613補充答辯或612補充理由時，則只可不會稿
'         If Trim(Left(Trim(Combo1(1)), 4)) = "613" Or _
'            Trim(Left(Trim(Combo1(2)), 4)) = "613" Or _
'            Trim(Left(Trim(Combo1(3)), 4)) = "613" Or _
'            Trim(Left(Trim(Combo1(4)), 4)) = "613" Or _
'            Trim(Left(Trim(Combo1(1)), 4)) = "612" Or _
'            Trim(Left(Trim(Combo1(2)), 4)) = "612" Or _
'            Trim(Left(Trim(Combo1(3)), 4)) = "612" Or _
'            Trim(Left(Trim(Combo1(4)), 4)) = "612" Then
         If Trim(Left(Trim(Combo1(1)), 4)) = "613" Or Trim(Left(Trim(Combo1(1)), 4)) = "612" Or _
            InStr(m_strCaseCPM, "613") > 0 Or InStr(m_strCaseCPM, "612") > 0 Then
            OptEP34(1).Value = True
         End If
      'Added by Lydia 2018/12/10 T台灣案填寫接洽單管控文件及查名是否齊備
      'Modified by Lydia 2022/07/15 T大陸案之齊備日管控
      'ElseIf Text1(6) = "T" And Left(Trim(Combo1(0)), 3) = "000" Then
      ElseIf Text1(6) = "T" Then
            Frame21.Visible = True
            Frame43.Visible = False '非商爭案，無延期
            Label23.Caption = "文件是否齊備："  '原本是「資料是否齊備」
'            OptCP122(1).Value = True '急件：預設N
            '查名限101申請
            Frame23.Visible = False
            'If InStr(Trim(Left(Combo1(1).Text, 4)) & "," & Trim(Left(Combo1(2).Text, 4)) & "," & Trim(Left(Combo1(3).Text, 4)) & "," & Trim(Left(Combo1(4).Text, 4)), "101") > 0 Then
            If InStr(Trim(Left(Combo1(1).Text, 4)), "101") > 0 Or InStr(m_strCaseCPM, "101") Then
               Frame42.Visible = True '查名是否齊備
            End If
      'end 2018/12/10
      'Mark by Lydia 2020/02/05 前面以本所案號判斷是否清空
      'Else
      '   '預設清空
      '   OptEP06(0).Value = False: OptEP06(1).Value = False
      '   OptEP34(0).Value = False: OptEP34(1).Value = False
      '   OptCP122(0).Value = False: OptCP122(1).Value = False
      '   OptCP143(0).Value = False: OptCP143(1).Value = False
      'end 2020/02/05
      'Added by Lydia 2022/07/15 TC案之文件齊備日管控
      ElseIf Text1(6) = "TC" Then
            Frame21.Visible = True
            Frame43.Visible = False '非商爭案，無延期
            Label23.Caption = "文件是否齊備："  '原本是「資料是否齊備」
'            OptCP122(1).Value = True '急件：預設N
            Frame23.Visible = False   '會稿：不顯示
'            Frame24.Visible = False   '急件：不顯示
            If Left(Trim(Combo1(0)), 3) <> "000" Then '台灣TC案不會稿
               Frame23.Visible = True   '會稿：顯示
            End If
      'end 2022/07/15
      End If
     
      'Added by Lydia 2019/11/05 特定申請人會稿案件: 業務收文新申請案(發明、實用新型、外觀設計)，且申請國為中國大陸時，希望能於收文時跳出【請確認此申請人之中國大陸案是否要會稿】
      Frame45.Visible = False
      'P新案,申請國為大陸
      If Option1(0).Value = True And Text1(6).Text = "P" And Trim(Left(Combo1(0), 4)) = "020" Then
'           For intP = 1 To 4
'              If Trim(Left(Trim(Combo1(1)), 4)) <> "" And InStr(NewCasePtyList, Trim(Left(Trim(Combo1(1)), 4))) > 0 Then
              If bolNewCase = True Then
                  Frame45.Visible = True
                  Opt45(0).Value = False: Opt45(1).Value = False  'Added by Lydia 2020/02/26
              End If
'           Next intP
      End If
      
   End If
End Sub

'Add By Sindy 2022/12/24
Private Sub SetFrame28()
Dim arrText As Variant, intCnt As Integer
Dim arrCF209() As String

   'm_strCaseCPM = GetAllCaseCPM()
   chkEnglish.Visible = False
   Frame28.Visible = False
   OptEntity(0).Enabled = False
   OptEntity(1).Enabled = False
   OptEntity(2).Enabled = False
   If Me.Combo1(1).Text <> "" Then
      m_strCaseCPM = m_strCaseCPM & IIf(m_strCaseCPM <> "", ",", "") & Trim(Left(Combo1(1).Text, 4))
   End If
   'CFP美國,P大陸以外之各國身份種類(frm090801之Frame28)
   'CFP美國案用現有檢查
   'P大陸案Frame28改放(1)無(2)台灣專利說明書
   If Text1(6).Text <> "" And Trim(Combo1(0)) <> "" And ((Option1(0).Value = True And m_strCaseCPM <> "") Or Option1(1).Value = True) Then
      If m_strCaseCPM = "" Then m_strCaseCPM = " "
      arrText = Split(m_strCaseCPM, ",")
'      strSql = "select distinct cf205,cf209 from casefee2" & _
'               " where cf201='" & Text1(6).Text & "' and cf202='" & Left(Trim(Combo1(0)), 3) & "' and cf203='" & arrText(0) & "'" & _
'               " and cf210=(select max(cf210) from casefee2 where cf201='" & Text1(6).Text & "' and cf202='" & Left(Trim(Combo1(0)), 3) & "' and cf203='" & arrText(0) & "'" & _
'               " and cf210<=" & strSrvDate(1) & ")" & IIf(Text1(6).Text = "P" And Left(Trim(Combo1(0)), 3) = "020", "", " and cf209 is not null") & _
'               " order by cf205"
'      intI = 1
'      Set RsTemp = ClsLawReadRstMsg(intI, strSql)
'      If intI = 1 Then
      'Modify By Sindy 2025/3/10
      If Option1(1).Value = True And m_strSys = "1" Then '專利資料
         If pa(179) <> "" Then '有個體別
            arrText(0) = ""
         End If
      End If
      '2025/3/10 END
      If Pub_GetCF209(Text1(6).Text, Left(Trim(Combo1(0)), 3), arrText(0), arrCF209) = True Then
         If Text1(6).Text = "P" And Left(Trim(Combo1(0)), 3) = "020" Then '大陸案
            Frame28.Visible = True
            OptEntity(0).Left = OptEntity(0).Tag
            OptEntity(0).Width = 2000
            OptEntity(0).Caption = arrCF209(0) '"無台灣專利說明書" Modify By Sindy 2025/2/3
            OptEntity(0).Visible = True
            OptEntity(0).Enabled = True
            OptEntity(1).Left = OptEntity(1).Tag + 1000
            OptEntity(1).Width = 2000
            OptEntity(1).Caption = arrCF209(1) '"有台灣專利說明書" Modify By Sindy 2025/2/3
            OptEntity(1).Visible = True
            OptEntity(1).Enabled = True
            OptEntity(2).Visible = False
         Else 'If "" & RsTemp.Fields("cf209") <> "" Then
            Frame28.Visible = True
            OptEntity(0).Left = OptEntity(0).Tag
            OptEntity(0).Width = 1100 '870
            OptEntity(0).Visible = False
            OptEntity(1).Left = OptEntity(1).Tag
            OptEntity(1).Width = 1315
            OptEntity(1).Visible = False
            OptEntity(2).Visible = False
            'OptEntity(2).Left = OptEntity(2).Tag
'            RsTemp.MoveFirst
'            intCnt = 0
'            Do While Not RsTemp.EOF
            For intCnt = 0 To 2
               'Add By Sindy 2023/3/22
               If arrCF209(intCnt) <> "" Then
               '2023/3/22 END
                  OptEntity(intCnt).Caption = arrCF209(intCnt) 'RsTemp.Fields("cf209")
                  OptEntity(intCnt).Visible = True
                  OptEntity(intCnt).Enabled = True
               End If
'               intCnt = intCnt + 1
'               RsTemp.MoveNext
'            Loop
            Next intCnt
         End If
      End If
      
      '同時申請三國(含)以上之美日德可多5點
      If Text1(6).Text = "CFP" Then
         If (Left(Trim(Combo1(0)), 3) = "011" Or Left(Trim(Combo1(0)), 3) = "101" Or Left(Trim(Combo1(0)), 3) = "231") And _
            InStr(CaseMapIn, arrText(0)) > 0 Then
            chkEnglish.Visible = True
         Else
            chkEnglish.Value = 0
         End If
      Else
         chkEnglish.Value = 0
      End If
   End If
End Sub

'Add By Sindy 2012/9/3 案件屬性
Private Sub SetCombo5()
   'Add By Sindy 2010/10/28
   Label1(168).Visible = False
   Combo5.Visible = False
   'Combo5.Text = "" 'Modify By Sindy 2023/4/28 Mark
   'Modify By Sindy 2012/9/3 專利種類為3.設計或4.積體電路時,不可填寫案件屬性; 案件性質103,105,117,125時,案件屬性欄清空並鎖住
   'Modify By Sindy 2014/7/14
'   If (Text1(6) = "P" Or Text1(6) = "CFP") And _
'      Trim(Left(Trim(Combo1(1)), 4)) <> "103" And _
'      Trim(Left(Trim(Combo1(2)), 4)) <> "103" And _
'      Trim(Left(Trim(Combo1(3)), 4)) <> "103" And _
'      Trim(Left(Trim(Combo1(4)), 4)) <> "103" And _
'      Trim(Left(Trim(Combo1(1)), 4)) <> "105" And _
'      Trim(Left(Trim(Combo1(2)), 4)) <> "105" And _
'      Trim(Left(Trim(Combo1(3)), 4)) <> "105" And _
'      Trim(Left(Trim(Combo1(4)), 4)) <> "105" And _
'      Trim(Left(Trim(Combo1(1)), 4)) <> "117" And _
'      Trim(Left(Trim(Combo1(2)), 4)) <> "117" And _
'      Trim(Left(Trim(Combo1(3)), 4)) <> "117" And _
'      Trim(Left(Trim(Combo1(4)), 4)) <> "117" And _
'      Trim(Left(Trim(Combo1(1)), 4)) <> "125" And _
'      Trim(Left(Trim(Combo1(2)), 4)) <> "125" And _
'      Trim(Left(Trim(Combo1(3)), 4)) <> "125" And _
'      Trim(Left(Trim(Combo1(4)), 4)) <> "125" And _
'      (Trim(Left(Trim(Combo6), 1)) <> "3" And Trim(Left(Trim(Combo6), 1)) <> "4") Then
   'Modify By Sindy 2023/5/11 專利種類只有4.積體電路,不需填寫案件屬性
   If (Text1(6) = "P" Or Text1(6) = "CFP") And _
      ((Option1(0).Value = True And Trim(Left(Trim(Combo6), 1)) <> "4") Or _
       (Option1(1).Value = True And m_strPA08 <> "4")) Then
   '2014/7/14 END
      Label1(168).Visible = True
      Combo5.Visible = True
      'Modified by Lydia 2021/06/18 debug:接洽單查詢及列印使用
      'If Combo5.Enabled = True Then Combo5.SetFocus  '2010/11/19 ADD BY SONIA
      If Combo5.Enabled = True And Me.Visible = True Then
         If Combo5.Enabled = True And Me.Visible = True Then Combo5.SetFocus
      End If
      'end 2021/06/18
   'Add By Sindy 2023/5/11
   Else
      Combo5.Text = ""
      '2023/5/11 END
   End If
End Sub

'Add By Sindy 2023/11/15 讀取商標種類的語法
Private Function GetCombo6_T_SQL(Optional p_Code As String) As String
Dim strCon As String
   
   strSql = ""
   If p_Code <> "" Then strCon = " and ptm02='" & p_Code & "'"
   'Modify By Sindy 2023/2/15
   'T台灣案
   If Text1(6) = "T" And Left(Combo1(0), 3) = "000" Then
      'Modify By Sindy 2023/11/15
      'strSql = "Select Ptm02,Ptm03 From Patenttrademarkmap Where Ptm01='2' AND (PTM02='1' OR PTM02>='7') order by ptm02 asc"
      strSql = "Select Ptm02,Ptm03 From Patenttrademarkmap Where Ptm01='2' AND (PTM02='1' OR PTM02>='7')" & strCon & _
               " union Select spt02,spt03 From SpecialPatenttrademark Where spt01='2'" & Replace(strCon, "ptm02", "spt02") & _
               " order by ptm02 asc"
      '2023/11/15 END
   'T大陸案
   ElseIf Text1(6) = "T" And Left(Combo1(0), 3) = "020" Then
      'Modify By Sindy 2023/11/15
      'strSql = "Select Ptm02,Ptm04 as Ptm03 From Patenttrademarkmap Where Ptm01='2' And (Ptm02='1' Or Ptm02>='7') and instr(ptm04,'（無）')=0 order by ptm02 asc"
      strSql = "select * from (" & _
               " Select Ptm02,Ptm04 as ptm03 From Patenttrademarkmap Where Ptm01='2' AND (PTM02='1' OR PTM02>='7')" & _
               " union Select spt02,spt04 From SpecialPatenttrademark Where spt01='2')" & _
               " where ptm03<>'（無）'" & strCon & _
               " order by ptm02 asc"
      '2023/11/15 END
   'TF
   Else
   '2023/2/15 END
      strSql = "select ptm02,ptm03 from patenttrademarkmap where ptm01='2' and ptm02 in('1','7','8','9')" & strCon & _
               " order by ptm02 asc"
   End If
   GetCombo6_T_SQL = strSql
End Function

'Add By Sindy 2012/6/6 商標,專利種類
Private Sub SetCombo6()
Dim strCombo6 As String, ii As Integer 'Add By Sindy 2023/4/28

   Label1(124).Visible = False
   Combo6.Visible = False
   Combo6.Enabled = False
   strCombo6 = Trim(Combo6.Text) 'Add By Sindy 2023/4/28
   Combo6.Clear
   'Add By Sindy 2012/4/26
   'CFT維持不顯示商標種類欄
   If (Text1(6) = "T" Or Text1(6) = "TF") And Option1(0).Value = True Then
      Label1(124).Visible = True
      Label1(124).Caption = "商標種類:"
      Combo6.Visible = True
      Combo6.Enabled = True
      strSql = GetCombo6_T_SQL 'Modify By Sindy 2023/11/15
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strSql)
      If intI = 1 Then
         RsTemp.MoveFirst
         Do While Not RsTemp.EOF
            Combo6.AddItem RsTemp.Fields("ptm02") & "." & RsTemp.Fields("ptm03")
            RsTemp.MoveNext
         Loop
      End If
'      Combo6.AddItem "1.商標"
'      Combo6.AddItem "7.證明標章"
'      Combo6.AddItem "8.團體標章"
'      Combo6.AddItem "9.團體商標"
   End If
   '2012/4/26 End
   'Add By Sindy 2012/6/6
   'If (Text1(6) = "P" Or Text1(6) = "CFP") And Option1(0).Value = True Then
   If (Text1(6) = "P" Or Text1(6) = "CFP") Then
      Label1(124).Visible = True
      Label1(124).Caption = "專利種類:"
      'Modify By Sindy 2014/7/14
      If Option1(0).Value = True Then
         Combo6.Visible = True
         Combo6.Enabled = True
      Else
         Combo6.Visible = True
         Combo6.Enabled = False
      End If
      '2014/7/14 END
      strSql = "select ptm02,ptm03 from patenttrademarkmap where ptm01='1' and ptm02 in('1','2','3') order by ptm02 asc"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strSql)
      If intI = 1 Then
         RsTemp.MoveFirst
         Do While Not RsTemp.EOF
            Combo6.AddItem RsTemp.Fields("ptm02") & "." & RsTemp.Fields("ptm03")
            RsTemp.MoveNext
         Loop
      End If
'      Combo6.AddItem "1.發明"
'      Combo6.AddItem "2.新型"
'      Combo6.AddItem "3.設計"
      Combo6.AddItem "4.積體電路"
   End If
   '2012/6/6 End
   
   'Add By Sindy 2023/4/28
   If strCombo6 <> "" Then
      For ii = 0 To Combo6.ListCount - 1
         If Combo6.List(ii) = strCombo6 Then
            Combo6.ListIndex = ii
            Exit For
         End If
      Next ii
   End If
   '2023/4/28 END
End Sub

'Add By Sindy 2022/12/16 系統別,申請國家,案件性質的欄位變化
Private Sub SetFrameChg()
   
   'm_strCaseCPM = GetAllCaseCPM() '取得案件性質代碼
   
   'Modified by Lydia 2016/04/25 顯示判斷 + TS案
   'Modified by Lydia 2021/11/23 因為T案增加737智財協作,調整判斷
   'strExc(1) = IIf(Trim(Left(Trim(Combo1(1)), 4)) <> "", Trim(Left(Trim(Combo1(1)), 4)), "") & "," & IIf(Trim(Left(Trim(Combo1(2)), 4)) <> "", Trim(Left(Trim(Combo1(2)), 4)), "") & _
          "," & IIf(Trim(Left(Trim(Combo1(3)), 4)) <> "", Trim(Left(Trim(Combo1(3)), 4)), "") & "," & IIf(Trim(Left(Trim(Combo1(4)), 4)) <> "", Trim(Left(Trim(Combo1(4)), 4)), "")
   'If strSrvDate(1) >= TMQ電子化啟用日 And Left(Trim(Combo1(0)), 3) = "000" And Option1(0).Value = True And ((Text1(6).Text = "T" And Text1(7) & Text1(8) = "" And InStr(strExc(1), TMQ_T案) > 0) _
       Or (Text1(6).Text = "TS" And Text1(7) & Text1(8) = "" And InStr(strExc(1), TMQ_TS案) > 0)) Then
   If GetTMQArea = True Then
   'end 2021/11/23
      FRTMQ.Visible = True
   Else
      FRTMQ.Visible = False
   End If
'end 2015/10/14
   'Added by Lydia 2021/02/24 大陸商標申請案及CFT申請案，均強迫填入相關資訊。
   Frame47.Visible = False
   'Added by Lydia 2022/04/28 debug: 補上抓所有案件性質
'           strExc(1) = IIf(Trim(Left(Trim(Combo1(1)), 4)) <> "", Trim(Left(Trim(Combo1(1)), 4)), "") & "," & IIf(Trim(Left(Trim(Combo1(2)), 4)) <> "", Trim(Left(Trim(Combo1(2)), 4)), "") & _
'                  "," & IIf(Trim(Left(Trim(Combo1(3)), 4)) <> "", Trim(Left(Trim(Combo1(3)), 4)), "") & "," & IIf(Trim(Left(Trim(Combo1(4)), 4)) <> "", Trim(Left(Trim(Combo1(4)), 4)), "")
   'end 2022/04/28
   If (Left(Trim(Combo1(0)), 3) = "020" And Text1(6).Text = "T" And InStr(m_strCaseCPM, "101") > 0) Or Text1(6).Text = "CFT" And InStr(m_strCaseCPM, "101") > 0 Then
      Frame47.Visible = True
      If Text1(6).Text = "T" Then
         fra47Title.Caption = "商標說明："
      Else
         fra47Title.Caption = "商標英文大寫："
      End If
   End If
   'end 2021/02/24
End Sub

Private Sub Text1_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
   'Added by Morgan 2022/1/22
   If Index = 11 Or Index = 119 Or Index = 130 _
      Or Index = 21 Or Index = 22 Or Index = 23 Or Index = 24 Or Index = 26 Or Index = 27 Or Index = 125 _
      Or Index = 37 Or Index = 38 Or Index = 39 Or Index = 40 Or Index = 42 Or Index = 43 Or Index = 141 _
      Or Index = 53 Or Index = 54 Or Index = 55 Or Index = 56 Or Index = 58 Or Index = 59 Or Index = 157 _
      Or Index = 69 Or Index = 70 Or Index = 71 Or Index = 72 Or Index = 74 Or Index = 75 Or Index = 173 _
      Or Index = 85 Or Index = 86 Or Index = 87 Or Index = 88 Or Index = 90 Or Index = 91 Or Index = 189 Then
      If Button = 2 Then Forms(0).PopupMenu2 Text1(Index)
   End If
   'end 2022/1/22
End Sub

'Add by Morgan 2004/1/9
Private Function CheckCaseProperty() As Boolean
   Dim bolReturn As Boolean, ii As Integer
   
   m_strCaseCPM = GetAllCaseCPM 'Add By Sindy 2022/9/8 取得案件性質代碼
   'Modify by Morgan 2011/1/14 移轉,讓與才傳true
   'bolReturn = True
   If (Text1(6) = "P" Or Text1(6) = "FCP" Or Text1(6) = "CFP") Then
'      For ii = 1 To 4
      'Modified by Lydia 2018/07/16 +繼承(703)
'         If Left(Combo1(ii), 3) = "701" Or Left(Combo1(ii), 3) = "703" Or Left(Combo1(ii), 3) = "708" Then
         If InStr(m_strCaseCPM, "701") > 0 Or InStr(m_strCaseCPM, "703") > 0 Or InStr(m_strCaseCPM, "708") > 0 Then
            bolReturn = True
'            Exit For
         End If
'      Next
   ElseIf (Text1(6) = "T" Or Text1(6) = "FCT" Or Text1(6) = "CFT") Then
'      For ii = 1 To 4
'         If Left(Combo1(ii), 3) = "501" Then
         If InStr(m_strCaseCPM, "501") > 0 Then
            bolReturn = True
'            Exit For
         End If
'      Next
   End If
   'end 2011/1/14
   CheckCaseProperty = bolReturn
End Function

'Add By Sindy 2010/5/27
Private Sub ClearFagentTxt()
   Me.Text1(5).Text = ""
   Me.Text1(130).Text = ""
   'Added by Lydia 2019/01/30
   m_FA45 = ""
   m_FA110 = ""
End Sub

Private Sub ClearCustTxt(intTab As Integer)
Dim ii As Integer
   
   m_strCaseCPM = GetAllCaseCPM 'Add By Sindy 2022/8/31 取得案件性質代碼
   'Add By Sindy 2022/11/2
   If Trim(Combo1(1).Text) <> "" Then
      arrCaseProperty = Split(Me.Combo1(1).Text, " ")
      m_strCaseCPM = arrCaseProperty(0) & IIf(m_strCaseCPM <> "", "," & m_strCaseCPM, "")
   End If
   '2022/11/2 END
   
   'Add By Sindy 2017/2/18
   Select Case intTab
   Case 1 '申請人1
      'Add By Sindy 2016/12/8 新案才能點選新客戶
      If Option1(0).Value = True Then '新案
         Option31(0).Enabled = True
         Option31(1).Enabled = True
      Else
         'Modify by Sindy 2017/2/18
         'Modified by Lydia 2018/07/16 +繼承(703)
         If Option1(1).Value = True And _
            ((Trim(Text1(6).Text) = "P" Or Trim(Text1(6).Text) = "FCP" Or Trim(Text1(6).Text) = "CFP") And _
             (InStr(m_strCaseCPM, "701") > 0 Or InStr(m_strCaseCPM, "703") > 0 Or InStr(m_strCaseCPM, "708") > 0) _
            ) Or _
            ((Trim(Text1(6).Text) = "CFT" Or Trim(Text1(6).Text) = "CFC" Or Trim(Text1(6).Text) = "FCT" Or _
              Trim(Text1(6).Text) = "T" Or Trim(Text1(6).Text) = "TB" Or Trim(Text1(6).Text) = "TC" Or _
              Trim(Text1(6).Text) = "TF") And _
             (InStr(m_strCaseCPM, "501") > 0)) Then
            Option31(0).Enabled = True
            Option31(1).Enabled = True
            Option31(0).Value = True 'Add by Amy 2020/04/28 預設新客戶才可輸入編號/名稱 查
         Else
         '2017/2/18 END
            Option31(0).Enabled = False
            Option31(1).Enabled = False
         End If
      End If
      '2016/12/8 END
      ChkCRA26(0).Value = 0 'Add By Sindy 2022/11/9
      ChkCRA27(0).Value = 0 'Add By Sindy 2022/11/9
      For ii = 12 To 17
          Me.Text1(ii).Text = ""
      Next ii
      For ii = 19 To 27
          Me.Text1(ii).Text = ""
      Next ii
      Me.Text1(92).Text = ""
      Me.Text1(92).Enabled = True 'Add By Sindy 2023/3/17 還原為可以輸入狀態
      'add by nickc 2007/05/23
      Text1(120).Text = ""
      'Add By Sindy 2010/5/27
      Text1(34).Text = ""
      Text1(125).Text = ""
      m_CU144(1) = "" 'Add By Sindy 2013/12/16 不可開立發票
   Case 2 '申請人2
      'Add By Sindy 2016/12/8 新案才能點選新客戶
      If Option1(0).Value = True Then '新案
         Option32(0).Enabled = True
         Option32(1).Enabled = True
      Else
         'Modify by Sindy 2017/2/18
         'Modified by Lydia 2018/07/16 +繼承(703)
         If Option1(1).Value = True And _
            ((Trim(Text1(6).Text) = "P" Or Trim(Text1(6).Text) = "FCP" Or Trim(Text1(6).Text) = "CFP") And _
             (InStr(m_strCaseCPM, "701") > 0 Or InStr(m_strCaseCPM, "703") > 0 Or InStr(m_strCaseCPM, "708") > 0) _
            ) Or _
            ((Trim(Text1(6).Text) = "CFT" Or Trim(Text1(6).Text) = "CFC" Or Trim(Text1(6).Text) = "FCT" Or _
              Trim(Text1(6).Text) = "T" Or Trim(Text1(6).Text) = "TB" Or Trim(Text1(6).Text) = "TC" Or _
              Trim(Text1(6).Text) = "TF") And _
             (InStr(m_strCaseCPM, "501") > 0)) Then
            Option32(0).Enabled = True
            Option32(1).Enabled = True
            Option32(0).Value = True 'Add by Amy 2020/04/28 預設新客戶才可輸入編號/名稱 查
         Else
         '2017/2/18 END
            Option32(0).Enabled = False
            Option32(1).Enabled = False
         End If
      End If
      '2016/12/8 END
      ChkCRA26(1).Value = 0 'Add By Sindy 2022/11/9
      ChkCRA27(1).Value = 0 'Add By Sindy 2022/11/9
      For ii = 28 To 33
          Me.Text1(ii).Text = ""
      Next ii
      For ii = 35 To 43
          Me.Text1(ii).Text = ""
      Next ii
      Me.Text1(93).Text = ""
      Me.Text1(93).Enabled = True 'Add By Sindy 2023/3/17 還原為可以輸入狀態
      'add by nickc 2007/05/23
      Text1(121).Text = ""
      'Add By Sindy 2010/5/27
      Text1(50).Text = ""
      Text1(141).Text = ""
      m_CU144(2) = "" 'Add By Sindy 2013/12/16 不可開立發票
   Case 3 '申請人3
      'Add By Sindy 2016/12/8 新案才能點選新客戶
      If Option1(0).Value = True Then '新案
         Option33(0).Enabled = True
         Option33(1).Enabled = True
      Else
         'Modify by Sindy 2017/2/18
         'Modified by Lydia 2018/07/16 +繼承(703)
         If Option1(1).Value = True And _
            ((Trim(Text1(6).Text) = "P" Or Trim(Text1(6).Text) = "FCP" Or Trim(Text1(6).Text) = "CFP") And _
             (InStr(m_strCaseCPM, "701") > 0 Or InStr(m_strCaseCPM, "703") > 0 Or InStr(m_strCaseCPM, "708") > 0) _
            ) Or _
            ((Trim(Text1(6).Text) = "CFT" Or Trim(Text1(6).Text) = "CFC" Or Trim(Text1(6).Text) = "FCT" Or _
              Trim(Text1(6).Text) = "T" Or Trim(Text1(6).Text) = "TB" Or Trim(Text1(6).Text) = "TC" Or _
              Trim(Text1(6).Text) = "TF") And _
             (InStr(m_strCaseCPM, "501") > 0)) Then
            Option33(0).Enabled = True
            Option33(1).Enabled = True
            Option33(0).Value = True 'Add by Amy 2020/04/28 預設新客戶才可輸入編號/名稱 查
         Else
         '2017/2/18 END
            Option33(0).Enabled = False
            Option33(1).Enabled = False
         End If
      End If
      '2016/12/8 END
      ChkCRA26(2).Value = 0 'Add By Sindy 2022/11/9
      ChkCRA27(2).Value = 0 'Add By Sindy 2022/11/9
      For ii = 44 To 49
          Me.Text1(ii).Text = ""
      Next ii
      For ii = 51 To 59
          Me.Text1(ii).Text = ""
      Next ii
      Me.Text1(94).Text = ""
      Me.Text1(94).Enabled = True 'Add By Sindy 2023/3/17 還原為可以輸入狀態
      'add by nickc 2007/05/23
      Text1(122).Text = ""
      'Add By Sindy 2010/5/27
      Text1(66).Text = ""
      Text1(157).Text = ""
      m_CU144(3) = "" 'Add By Sindy 2013/12/16 不可開立發票
   Case 4 '申請人4
      'Add By Sindy 2016/12/8 新案才能點選新客戶
      If Option1(0).Value = True Then '新案
         Option34(0).Enabled = True
         Option34(1).Enabled = True
      Else
         'Modify by Sindy 2017/2/18
         'Modified by Lydia 2018/07/16 +繼承(703)
         If Option1(1).Value = True And _
            ((Trim(Text1(6).Text) = "P" Or Trim(Text1(6).Text) = "FCP" Or Trim(Text1(6).Text) = "CFP") And _
             (InStr(m_strCaseCPM, "701") > 0 Or InStr(m_strCaseCPM, "703") > 0 Or InStr(m_strCaseCPM, "708") > 0) _
            ) Or _
            ((Trim(Text1(6).Text) = "CFT" Or Trim(Text1(6).Text) = "CFC" Or Trim(Text1(6).Text) = "FCT" Or _
              Trim(Text1(6).Text) = "T" Or Trim(Text1(6).Text) = "TB" Or Trim(Text1(6).Text) = "TC" Or _
              Trim(Text1(6).Text) = "TF") And _
             (InStr(m_strCaseCPM, "501") > 0)) Then
            Option34(0).Enabled = True
            Option34(1).Enabled = True
            Option34(0).Value = True 'Add by Amy 2020/04/28 預設新客戶才可輸入編號/名稱 查
         Else
         '2017/2/18 END
            Option34(0).Enabled = False
            Option34(1).Enabled = False
         End If
      End If
      '2016/12/8 END
      ChkCRA26(3).Value = 0 'Add By Sindy 2022/11/9
      ChkCRA27(3).Value = 0 'Add By Sindy 2022/11/9
      For ii = 60 To 65
          Me.Text1(ii).Text = ""
      Next ii
      For ii = 67 To 75
          Me.Text1(ii).Text = ""
      Next ii
      Me.Text1(95).Text = ""
      Me.Text1(95).Enabled = True 'Add By Sindy 2023/3/17 還原為可以輸入狀態
      'add by nickc 2007/05/23
      Text1(123).Text = ""
      'Add By Sindy 2010/5/27
      Text1(82).Text = ""
      Text1(173).Text = ""
      m_CU144(4) = "" 'Add By Sindy 2013/12/16 不可開立發票
   Case 5 '申請人5
      'Add By Sindy 2016/12/8 新案才能點選新客戶
      If Option1(0).Value = True Then '新案
         Option35(0).Enabled = True
         Option35(1).Enabled = True
      Else
         'Modify by Sindy 2017/2/18
         'Modified by Lydia 2018/07/16 +繼承(703)
         If Option1(1).Value = True And _
            ((Trim(Text1(6).Text) = "P" Or Trim(Text1(6).Text) = "FCP" Or Trim(Text1(6).Text) = "CFP") And _
             (InStr(m_strCaseCPM, "701") > 0 Or InStr(m_strCaseCPM, "703") > 0 Or InStr(m_strCaseCPM, "708") > 0) _
            ) Or _
            ((Trim(Text1(6).Text) = "CFT" Or Trim(Text1(6).Text) = "CFC" Or Trim(Text1(6).Text) = "FCT" Or _
              Trim(Text1(6).Text) = "T" Or Trim(Text1(6).Text) = "TB" Or Trim(Text1(6).Text) = "TC" Or _
              Trim(Text1(6).Text) = "TF") And _
             (InStr(m_strCaseCPM, "501") > 0)) Then
            Option35(0).Enabled = True
            Option35(1).Enabled = True
            Option35(0).Value = True 'Add by Amy 2020/04/28 預設新客戶才可輸入編號/名稱 查
         Else
         '2017/2/18 END
            Option35(0).Enabled = False
            Option35(1).Enabled = False
         End If
      End If
      '2016/12/8 END
      ChkCRA26(4).Value = 0 'Add By Sindy 2022/11/9
      ChkCRA27(4).Value = 0 'Add By Sindy 2022/11/9
      For ii = 76 To 81
          Me.Text1(ii).Text = ""
      Next ii
      For ii = 83 To 91
          Me.Text1(ii).Text = ""
      Next ii
      Me.Text1(96).Text = ""
      Me.Text1(96).Enabled = True 'Add By Sindy 2023/3/17 還原為可以輸入狀態
      'add by nickc 2007/05/23
      Text1(124).Text = ""
      'Add By Sindy 2010/5/27
      Text1(98).Text = ""
      Text1(189).Text = ""
      m_CU144(5) = "" 'Add By Sindy 2013/12/16 不可開立發票
   End Select
   'Add by Morgan 2008/8/1
   cboContact(intTab).Clear
   m_strContactList(intTab) = ""
End Sub

'Add By Sindy 2010/5/27
Private Function SetFagentTxt(strCUCode As String) As Boolean
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
   
   SetFagentTxt = False
   strCUCode = Left(strCUCode & "000000000", 9)
   StrSQLa = "Select * From Fagent Where FA01='" & Mid(strCUCode, 1, 8) & "' And FA02='" & Mid(strCUCode, 9, 1) & "' "
   rsA.CursorLocation = adUseClient
   rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
   If rsA.RecordCount > 0 Then
       SetFagentTxt = True
       '代理人
       Me.Text1(5).Text = "" & rsA("FA01").Value & rsA("FA02").Value
       '代理人中文名稱
       'Modified by Morgan 2020/7/2 +判斷查詢語文
       'Me.Text1(130).Text = "" & rsA("FA04").Value
       If strLang = "英" Then
         Me.Text1(130).Text = RTrim(rsA("FA05").Value & " " & rsA("FA63").Value & " " & rsA("FA64").Value & " " & rsA("FA65").Value)
       ElseIf strLang = "日" Then
         Me.Text1(130).Text = "" & rsA("FA06").Value
       ElseIf Not IsNull(rsA("FA04").Value) Then
         Me.Text1(130).Text = "" & rsA("FA04").Value
       ElseIf Not IsNull(rsA("FA05").Value) Then
         Me.Text1(130).Text = RTrim(rsA("FA05").Value & " " & rsA("FA63").Value & " " & rsA("FA64").Value & " " & rsA("FA65").Value)
       Else
         Me.Text1(130).Text = "" & rsA("FA06").Value
       End If
       'end 2020/7/2
       
'       Text1(5).Enabled = False
'       Text1(130).Enabled = False
       'Added by Lydia 2019/01/30 記錄代理人D/N備註
       m_FA45 = "" & rsA.Fields("FA45")
       m_FA110 = "" & rsA.Fields("FA110")
       'end 2019/01/30
   End If
   If rsA.State <> adStateClosed Then rsA.Close
   Set rsA = Nothing
End Function

Private Function SetCustTxt(intTab As Integer, strCUCode As String) As Boolean
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset, i As Integer

SetCustTxt = False
bolCusCAddr(intTab) = False 'Add by Amy 2017/07/12 先輸P-092481接洽單(案件無申請地址)再輸CFT-011714,印出接洽單會有 案件無中文地址但客戶檔有，請通知專業部補輸「案件中文地址」字樣
strCUCode = Left(strCUCode & "000000000", 9)
StrSQLa = "Select * From Customer Where CU01='" & Mid(strCUCode, 1, 8) & "' And CU02='" & Mid(strCUCode, 9, 1) & "' "
rsA.CursorLocation = adUseClient
rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
If intTab = 1 Then m_CU143 = "" 'Add By Sindy 2013/11/20
If rsA.RecordCount > 0 Then
    'Modify by Amy 2020/04/28 依不同頁籤設定,ex: 申請人2查詢 張復龍,會將原來設申請人1-新客戶,改成舊客戶
    Select Case intTab
        Case 1
            Option31(1).Value = True 'Add by Amy 2016/12/23 點新客戶發現為舊客戶,輸入客戶編號後自行點選舊客戶,導致intTWAdd被清空-雅娟
        Case 2
            Option32(1).Value = True
        Case 3
            Option33(1).Value = True
        Case 4
            Option34(1).Value = True
        Case 5
            Option35(1).Value = True
    End Select
    'end 2020/04/28
    SetCustTxt = True
    
    If intTab = 1 Then m_CU143 = "" & rsA("CU143").Value 'Add By Sindy 2013/11/20 預定收款日放寬月數
    m_CU144(intTab) = "" & rsA("CU144").Value 'Add By Sindy 2013/12/16 不可開立發票
    'Added by Lydia 2021/03/29 ACS案件收文與點數及營業稅：客戶設定為不開發票CU144='N'者也自動取消收據公司；
    If strSrvDate(1) >= strACSdate1 Then
        If Option1(0).Value = True And Text1(6) = "ACS" And m_CU144(intTab) = "N" And Combo4.Text <> "" Then
            Combo4.ListIndex = 0
            '填寫到第二頁客戶欄時，若此客戶有設定為不開發票者，因為第一頁之費用已填寫完畢，所以彈訊息提醒「此客戶設定為不開發票，將自動取消收據之智權公司，收費項目之費用、規費、點數欄請再自行調整！」。
            MsgBox "申請人" & intTab & "設定為不開發票，將自動取消收據之智權公司，" & vbCrLf & "收費項目之費用、規費、點數欄請再自行調整！", vbExclamation, "ACS案件收文"
            
            '重新計算規費和點數
            For intI = 1 To GridCase.Rows - 1 '4
               Call GetGridCaseData(intI)
               Call SetACSautoFee(1, IIf(intTab > 1, "A", "")) 'intI
            Next
        End If
    End If
    'end 2021/03/29
    
    '申請人
    Me.Text1(12 + (intTab - 1) * 16).Text = "" & rsA("CU01").Value & rsA("CU02").Value
    '接洽人
    Me.Text1(13 + (intTab - 1) * 16).Text = "" & rsA("CU08").Value
    '電話1
    Me.Text1(14 + (intTab - 1) * 16).Text = "" & rsA("CU16").Value
    'Modified by Lydia 2021/08/26 電話2=> LINE ID
    ''電話2
    'Me.Text1(15 + (intTab - 1) * 16).Text = "" & rsA("CU17").Value
    If Len("" & rsA("CU21").Value) > 20 Then ' 因為LINE ID可存50字, 但是列印最多20字,多出來的用...省略
      Me.Text1(15 + (intTab - 1) * 16).Text = Mid("" & rsA("CU21").Value, 1, 16) & " ..."
    Else
        Me.Text1(15 + (intTab - 1) * 16).Text = "" & rsA("CU21").Value
    End If
    'end 2021/08/26
    '傳真1
    Me.Text1(16 + (intTab - 1) * 16).Text = "" & rsA("CU18").Value
    '傳真2
    Me.Text1(17 + (intTab - 1) * 16).Text = "" & rsA("CU19").Value
    'E-Mail
    Me.Text1(19 + (intTab - 1) * 16).Text = "" & rsA("CU20").Value
    '手機
    Me.Text1(20 + (intTab - 1) * 16).Text = "" & rsA("CU22").Value
    '申請人中文
    'Modified by Lydia 2017/06/19 長度控制
    'Me.Text1(21 + (intTab - 1) * 16).Text = "" & rsA("CU04").Value
    'Modified by Lydia 2017/06/28 + trim 清空白
    'Modified by Lydia 2017/11/20 trim 會造成造字錯誤,拿掉
    'Me.Text1(21 + (intTab - 1) * 16).Text = Trim(IIf(Me.Text1(21 + (intTab - 1) * 16).MaxLength = 0, "" & rsA("CU04").Value, convForm("" & rsA("CU04").Value, Me.Text1(21 + (intTab - 1) * 16).MaxLength)))
    Me.Text1(21 + (intTab - 1) * 16).Text = IIf(Me.Text1(21 + (intTab - 1) * 16).MaxLength = 0, "" & rsA("CU04").Value, PUB_StrToStr("" & rsA("CU04").Value, Me.Text1(21 + (intTab - 1) * 16).MaxLength))
    'Added by Lydia 2017/06/19 記錄客戶檔的名稱
    Me.Text1(21 + (intTab - 1) * 16).Tag = Me.Text1(21 + (intTab - 1) * 16).Text
    '申請人英文
    'Modified by Lydia 2017/06/19 長度控制
    'Me.Text1(22 + (intTab - 1) * 16).Text = "" & rsA("CU05").Value & " " & rsA("CU88").Value & " " & rsA("CU89").Value & " " & rsA("CU90").Value
    'Modified by Lydia 2017/06/28 + trim 清空白
    'Modified by Lydia 2017/11/20 trim 會造成造字錯誤,拿掉
    'Me.Text1(22 + (intTab - 1) * 16).Text = Trim(IIf(Me.Text1(22 + (intTab - 1) * 16).MaxLength = 0, "" & rsA("CU05").Value & " " & rsA("CU88").Value & " " & rsA("CU89").Value & " " & rsA("CU90").Value, convForm("" & rsA("CU05").Value & " " & rsA("CU88").Value & " " & rsA("CU89").Value & " " & rsA("CU90").Value, Me.Text1(22 + (intTab - 1) * 16).MaxLength)))
    Me.Text1(22 + (intTab - 1) * 16).Text = IIf(Me.Text1(22 + (intTab - 1) * 16).MaxLength = 0, "" & rsA("CU05").Value & " " & rsA("CU88").Value & " " & rsA("CU89").Value & " " & rsA("CU90").Value, PUB_StrToStr("" & rsA("CU05").Value & " " & rsA("CU88").Value & " " & rsA("CU89").Value & " " & rsA("CU90").Value, Me.Text1(22 + (intTab - 1) * 16).MaxLength))
    'Added by Lydia 2017/06/19 記錄客戶檔的名稱
    Me.Text1(22 + (intTab - 1) * 16).Tag = Me.Text1(22 + (intTab - 1) * 16).Text

    '代表人1中文
'    If OptChoose(1).Value = True Then
'      Me.Text1(23 + (intTab - 1) * 16).Text = "" & rsA("CU39").Value '代表人1(中)
'    Else
      Me.Text1(23 + (intTab - 1) * 16).Text = "" & rsA("CU07").Value '公司負責人
'    End If
    '代表人1英文
    'Modify by Morgan 2004/4/27
    'Me.Text1(24 + (intTab - 1) * 16).Text = ""
    Me.Text1(24 + (intTab - 1) * 16).Text = "" & rsA("CU103").Value
    '2011/5/18 ADD BY SONIA 若帶出之長度超過16碼X29285(接洽單只能印16碼),則最後二碼改為..
    If GetTextLength("" & rsA("CU103").Value) > 16 Then
      Me.Text1(24 + (intTab - 1) * 16).Text = Left(Me.Text1(24 + (intTab - 1) * 16).Text, 14) & ".."
    End If
    '2011/5/18 END
    ''聯絡地址郵遞區號
    Me.Text1(25 + (intTab - 1) * 16).Text = "" & rsA("CU30").Value
    '聯絡地址
    Me.Text1(26 + (intTab - 1) * 16).Text = "" & rsA("CU31").Value
    '申請地址
    Me.Text1(27 + (intTab - 1) * 16).Text = "" & rsA("CU23").Value
    'ID No.
    Me.Text1(92 + (intTab - 1)).Text = "" & rsA("CU11").Value
    'Add By Sindy 2014/9/11
    Me.Text1(92 + (intTab - 1)).Tag = "" & rsA("CU15").Value '個人或公司
    '2014/9/11 END
    'Add By Sindy 2014/2/6 舊客戶為公司且無統一編號台灣客戶者開放使用者輸入
    Me.Text1(92 + (intTab - 1)).Enabled = False
    If "" & rsA("CU15").Value = "1" And "" & rsA("CU11").Value = "" And "" & rsA("CU10").Value < "010" Then
      Me.Text1(92 + (intTab - 1)).Enabled = True
    End If
    '2014/2/6 END
    
    'add by nickc 2006/02/08
    Me.Text1(120 + (intTab - 1)).Text = "" & rsA("CU112").Value
   
    'Add by Morgan 2008/8/1
    'Modified by Morgan 2022/1/20 改2.0
    PUB_AddContact rsA("CU01"), cboContact(intTab), "" & rsA("CU127"), , True, m_strContactList(intTab) '設定接洽人選單
    
    'Add By Sindy 2010/5/27
    '國籍
    Me.Text1(34 + (intTab - 1) * 16).Text = GetPrjNationName("" & rsA("CU10").Value)
    '申請地址英文
    'Modify By Sindy 2023/2/8 + & "" & rsA("CU102").Value
    Me.Text1(125 + (intTab - 1) * 16).Text = "" & rsA("CU24").Value & "" & rsA("CU25").Value & "" & rsA("CU26").Value & "" & rsA("CU27").Value & "" & rsA("CU28").Value & "" & rsA("CU102").Value
    '2010/5/27 End
End If
If rsA.State <> adStateClosed Then rsA.Close
Set rsA = Nothing

'Add By Sindy 2011/1/21 專利商標的舊案抓案件基本檔的申請地址
'Dim strCAddr As String, strEAddr As String  '2011/10/21 cancel by sonia
Dim bolSetAddr As Boolean

If (Me.Option1(1).Value = True And Me.Text1(6).Text <> "" And Me.Text1(7).Text <> "") Then
   Select Case Me.Text1(6).Text
      '專利
      Case "P", "FCP", "CFP"
         'Add By Sindy 2011/3/3
         bolSetAddr = True
'         For i = 1 To 4
'            If Me.Combo1(i).Text <> "" Then
'               arrCaseProperty = Split(Me.Combo1(i).Text, " ") '案件性質
               'Modified by Lydia 2018/07/16 +繼承(703)
               If InStr(m_strCaseCPM, "701") > 0 Or InStr(m_strCaseCPM, "703") > 0 Or InStr(m_strCaseCPM, "708") > 0 Then
                  bolSetAddr = False
'                  Exit For
               End If
'            End If
'         Next i
         If bolSetAddr = True Then
         '2011/3/3 End
            StrSQLa = "Select * From Patent Where pa01='" & Text1(6) & "' And pa02='" & Text1(7) & "' And pa03='" & Text1(8) & "' And pa04='" & Text1(9) & "' "
            rsA.CursorLocation = adUseClient
            rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
            If rsA.RecordCount > 0 Then
               If intTab = 1 Then
                  strCAddr = Trim("" & rsA("pa31").Value)
                  strEAddr = Trim("" & rsA("pa36").Value)
               ElseIf intTab = 2 Then
                  strCAddr = Trim("" & rsA("pa32").Value)
                  strEAddr = Trim("" & rsA("pa37").Value)
               ElseIf intTab = 3 Then
                  strCAddr = Trim("" & rsA("pa33").Value)
                  strEAddr = Trim("" & rsA("pa38").Value)
               ElseIf intTab = 4 Then
                  strCAddr = Trim("" & rsA("pa34").Value)
                  strEAddr = Trim("" & rsA("pa39").Value)
               ElseIf intTab = 5 Then
                  strCAddr = Trim("" & rsA("pa35").Value)
                  strEAddr = Trim("" & rsA("pa40").Value)
               End If
               AddrToZipAddr    '2011/10/21 add by sonia舊案申請地址拆成郵遞區號及地址
               '申請地址
               If Me.Text1(27 + (intTab - 1) * 16) <> MsgText(601) And strCAddr = MsgText(601) Then bolCusCAddr(intTab) = True 'Add by Amy 2016/12/23
               Me.Text1(27 + (intTab - 1) * 16).Text = strCAddr
               '申請地址英文
               Me.Text1(125 + (intTab - 1) * 16).Text = strEAddr
               '中文地址郵遞區號
               '2011/10/21 modify by sonia
               'Me.Text1(120 + (intTab - 1)).Text = "" 'Add By Sindy 2011/3/3
               Me.Text1(120 + (intTab - 1)).Text = strZip
               '2011/10/21 end
            End If
            If rsA.State <> adStateClosed Then rsA.Close
            Set rsA = Nothing
         End If
      '商標
      Case "T", "FCT", "CFT", "TF"
         'Add By Sindy 2011/3/3
         bolSetAddr = True
'         For i = 1 To 4
'            If Me.Combo1(i).Text <> "" Then
'               arrCaseProperty = Split(Me.Combo1(i).Text, " ") '案件性質
               If InStr(m_strCaseCPM, "501") > 0 Then
                  bolSetAddr = False
'                  Exit For
               End If
'            End If
'         Next i
         If bolSetAddr = True Then
         '2011/3/3 End
            StrSQLa = "Select * From Trademark Where tm01='" & Text1(6) & "' And tm02='" & Text1(7) & "' And tm03='" & Text1(8) & "' And tm04='" & Text1(9) & "' "
            rsA.CursorLocation = adUseClient
            rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
            If rsA.RecordCount > 0 Then
               If intTab = 1 Then
                  strCAddr = Trim("" & rsA("tm24").Value)
                  strEAddr = Trim("" & rsA("tm25").Value)
               ElseIf intTab = 2 Then
                  strCAddr = Trim("" & rsA("tm82").Value)
                  strEAddr = Trim("" & rsA("tm86").Value)
               ElseIf intTab = 3 Then
                  strCAddr = Trim("" & rsA("tm83").Value)
                  strEAddr = Trim("" & rsA("tm87").Value)
               ElseIf intTab = 4 Then
                  strCAddr = Trim("" & rsA("tm84").Value)
                  strEAddr = Trim("" & rsA("tm88").Value)
               ElseIf intTab = 5 Then
                  strCAddr = Trim("" & rsA("tm85").Value)
                  strEAddr = Trim("" & rsA("tm89").Value)
               End If
               AddrToZipAddr    '2011/10/21 add by sonia舊案申請地址拆成郵遞區號及地址
               '申請地址
               If Me.Text1(27 + (intTab - 1) * 16) <> MsgText(601) And strCAddr = MsgText(601) Then bolCusCAddr(intTab) = True 'Add by Amy 2016/12/23
               Me.Text1(27 + (intTab - 1) * 16).Text = strCAddr
               '申請地址英文
               Me.Text1(125 + (intTab - 1) * 16).Text = strEAddr
               '中文地址郵遞區號
               '2011/10/21 modify by sonia
               'Me.Text1(120 + (intTab - 1)).Text = "" 'Add By Sindy 2011/3/3
               Me.Text1(120 + (intTab - 1)).Text = strZip
               '2011/10/21 end
            End If
            If rsA.State <> adStateClosed Then rsA.Close
            Set rsA = Nothing
         End If
   End Select
End If
'2011/1/21 End

Select Case intTab
Case 1
    stCustNo1 = Me.Text1(12 + (intTab - 1) * 16).Text
    'add by nick 2004/10/05
    If IsoptCP81 = True Then
        Call setCP811
    End If
Case 2
    stCustNo2 = Me.Text1(12 + (intTab - 1) * 16).Text
    'add by nick 2004/10/05
    If IsoptCP81 = True Then
        Call setCP812
    End If
Case 3
    stCustNo3 = Me.Text1(12 + (intTab - 1) * 16).Text
    'add by nick 2004/10/05
    If IsoptCP81 = True Then
        Call setCP813
    End If
Case 4
    stCustNo4 = Me.Text1(12 + (intTab - 1) * 16).Text
    'add by nick 2004/10/05
    If IsoptCP81 = True Then
        Call setCP814
    End If
Case 5
    stCustNo5 = Me.Text1(12 + (intTab - 1) * 16).Text
    'add by nick 2004/10/05
    If IsoptCP81 = True Then
        Call setCP815
    End If
Case Else
End Select
End Function

'Modify by Morgan 2011/3/28 +pbolPreserve
Private Sub ClearAll(Optional pbolPreserve As Boolean = False)
Dim objText As Object
Dim objChk As Object
Dim objCbo As Object
Dim objOpt As Object
Dim objLbl As Object 'Add By Sindy 2023/1/18

For Each objText In Me.Text1
    If objText.Index <> 0 And objText.Index <> 10 Then
         If Not (pbolPreserve = True And (objText.Index = 6 Or objText.Index = 7 Or objText.Index = 8 Or objText.Index = 9)) Then 'Add by Morgan 2011/3/28 考慮輸多張接洽單只有改案號情形
            objText.Text = ""
            objText.Tag = "" 'Add By Sindy 2014/9/25
         End If
    End If
Next
Text1(1).Enabled = True 'Add By Sindy 2014/12/10
Text1(3).Enabled = True 'Add By Sindy 2014/12/10
m_strSaveFiles = "": m_strSaveFiles2 = "" 'Add By Sindy 2015/1/7
Me.cmdAddAtt.BackColor = &H808080 '灰色 Add By Sindy 2022/9/30
Frame57.Visible = False 'Add By Sindy 2022/10/4

If pbolPreserve = False Then 'Add by Morgan 2011/3/28 考慮輸多張接洽單只有改案號情形
   For Each objOpt In Me.Option1
       objOpt.Value = False
   Next
End If

For Each objOpt In Me.Option2
    objOpt.Value = False
Next

'是否新客戶
For Each objOpt In Me.Option31
    objOpt.Value = False
Next
For Each objOpt In Me.Option32
    objOpt.Value = False
Next
For Each objOpt In Me.Option33
    objOpt.Value = False
Next
For Each objOpt In Me.Option34
    objOpt.Value = False
Next
For Each objOpt In Me.Option35
    objOpt.Value = False
Next

Me.OptSendType(1).Value = True 'Add By Sindy 2010/11/22

'Add By Sindy 2010/11/2
Me.PicText.Text = ""
For Each objOpt In Me.optCP811
    objOpt.Value = False
Next
For Each objOpt In Me.optCP812
    objOpt.Value = False
Next
For Each objOpt In Me.optCP813
    objOpt.Value = False
Next
For Each objOpt In Me.optCP814
    objOpt.Value = False
Next
For Each objOpt In Me.optCP815
    objOpt.Value = False
Next
For Each objOpt In Me.optColor
    objOpt.Value = False
Next
'Added by Lydia 2019/11/05
For Each objOpt In Me.Opt45
    objOpt.Value = False
Next

Me.ChkPCT.Value = vbUnchecked
Me.Check1.Value = vbUnchecked
Me.Check2.Value = vbUnchecked
For Each objChk In Me.Check7
    objChk.Value = vbUnchecked
Next

'Add By Sindy 2022/9/15
For Each objChk In Me.ChkCRA26
    objChk.Value = vbUnchecked
    objChk.Enabled = True
Next
For Each objChk In Me.ChkCRA27
    objChk.Value = vbUnchecked
    objChk.Enabled = True
Next
ChkCRL66.Value = vbUnchecked
ChkCRL66.BackColor = &H8000000F '灰色
'Add By Sindy 2022/11/4 急件
Check11.Value = vbUnchecked
Check11.BackColor = &H8000000F '灰色
'Add By Sindy 2022/12/13 費用已核准
Check12.Value = vbUnchecked
Check12.BackColor = &H8000000F '灰色
ChkCRL152.Value = vbUnchecked 'Add By Sindy 2023/4/7
Check10.Enabled = 0 'Add By Sindy 2025/4/14

Me.cboTitle.Text = ""
Me.Combo4.ListIndex = 0
Me.Combo4.Tag = "" 'Add by Amy 2015/10/22
Me.Combo5.Text = ""
For Each objCbo In Me.cboContact
    objCbo.Text = ""
Next
For Each objOpt In Me.opt1
    objOpt.Value = False
Next
'2010/11/2 End

Me.Combo6.Text = "" 'Add By Sindy 2012/4/26
Me.Combo5.Text = "" 'Add By Sindy 2014/7/15

'Add By Sindy 2011/6/7
Text1(127).Text = ""
For Each objChk In Me.Check3
    objChk.Value = vbUnchecked
Next
'2011/6/7 End

For Each objOpt In Me.Option4
    objOpt.Value = False
Next
For Each objOpt In Me.Option5
    objOpt.Value = False
Next
For Each objChk In Me.Check6
    objChk.Value = vbUnchecked
Next

'申請國/案件性質
'Add By Sindy 2022/8/29
lblCnt.Caption = ""
GridCase.Clear: Call SetGrd
FrameCRC.Caption = "案件性質區（" & Val(GridCase.Rows - 1) - IIf(Trim(GridCase.TextMatrix(1, 1)) = "", 1, 0) & "）"
'2022/8/29 END
For Each objCbo In Me.Combo1
    objCbo.Tag = ""  '2011/9/20 add by sonia
    objCbo.Text = ""
    If objCbo.Index <> 0 Then
        objCbo.Clear
    End If
Next
'案件性質備註
For Each objCbo In Me.Combo2
    objCbo.Text = ""
Next

'2010/3/4 add by sonia
If pbolPreserve = False Then 'Add by Morgan 2011/3/28 考慮輸多張接洽單只有改案號情形
   Me.Option1(0).Value = True
End If
Me.Option31(1).Value = True
Me.Option32(1).Value = True
Me.Option33(1).Value = True
Me.Option34(1).Value = True
Me.Option35(1).Value = True
'2010/3/4 end

'2008/8/25 add for Toni use in 發明人
For Each objText In Text2
   objText.Text = ""
Next
For Each objText In Text3
   objText.Text = ""
Next
For Each objCbo In Me.Combo3
   objCbo.Text = ""
Next
For Each objText In Text4
   objText.Text = ""
   objText.Tag = ""
Next
For Each objChk In Me.ChkAddress
   objChk.Value = vbUnchecked
Next
'Add By Sindy 2023/1/18
For Each objLbl In Label5
   objLbl.Tag = ""
Next
'2023/1/18 END

m_Note1 = "": m_Note2 = "" 'Add By Sindy 2010/4/23
m_strGetNP01 = "" 'Add By Sindy 2015/9/17
Text1(142).Text = "": Text1(143).Text = "" 'strYF05From = "": strYF05To = "" 'Add By Sindy 2010/7/9

Me.SSTab1.Tab = 0
'Modify By Sindy 2023/11/16
'Me.SSTab2.Tab = 0
If SSTab2.TabVisible(0) = True Then Me.SSTab2.Tab = 0
'2023/11/16 END
'Modified by Morgan 2020/5/7
'Me.Text1(119).Text = "如案件內容摘要、引用條文與客戶洽談要旨等等......" & vbCrLf & _
'                                    "商品類別：" & vbCrLf & _
'                                    "商品名稱：" & vbCrLf & _
'                                    "註冊號數：" & vbCrLf & _
'                                    "優先權日：" & vbCrLf & _
'                                    "優先權號：" & vbCrLf
''Add By Sindy 2010/5/27
'If OptChoose(1).Value = True Then
'   Me.Text1(119).Text = Me.Text1(119).Text & "聯絡人：" & vbCrLf & _
'                                                                           "彼所案號：" & vbCrLf
'End If
SrcSetMemo OptChoose(1).Value
'end 2020/5/7

'發明人暫存資料清除  20080902 add by Toni
strInventorNo = ""
strPetition = ""
strInventorName = "" 'Add By Sindy 2011/1/31 +strInventorName
m_CRL02 = "" 'Add By Sindy 2011/10/18
'Modify By Sindy 2012/6/11 Mark
''Add By Sindy 2012/5/30
'Text1(10) = strUserNum
'lblStaffName.Caption = strUserName
''2012/5/30 End

'Add By Sindy 2012/11/12
'Modified by Morgan 2013/11/26
'For Each objOpt In Me.Check8
'    objOpt.Value = 0
For Each objChk In Me.Check8
   objChk.Value = vbUnchecked
'end 2013/11/26
Next
'2012/11/12 End
'Add by Amy 2017/01/20
For Each objText In Me.TxtC1
   objText.Text = ""
Next

m_LAmsg = "" 'Added by Lydia 2019/04/10

'Me.Combo7.Text = "" 'Added by Lydia 2019/08/06

'Added by Lydia 2020/11/19
strCaseNA239 = ""
bolCase201 = False
strCaseNA239 = "" 'Added by Lydia 2021/03/05
strGetNp15 = "" 'Added by Lydia 2021/04/15

strCaseNo1 = "": strCaseNo4 = "": StrCaseNo3 = "": strCaseNo4 = "": strTM28 = "" 'Added by Lydia 2020/12/15

'Added by Lydia 2021/02/24
Frame47.Visible = False
Text7.Text = ""

'Added by Lydia 2021/03/29 開放編輯
If strSrvDate(1) >= strACSdate1 Then
    Combo4.Enabled = True
'    For intI = 1 To 4
'        Text1(102 + (intI - 1) * 3).Enabled = True
'        Text1(103 + (intI - 1) * 3).Enabled = True
'    Next intI
    Text1(102).Enabled = True
    Text1(103).Enabled = True
End If

m_ACS112msg = "" 'Added by Lydia 2021/05/07

'Added by Lydia 2023/11/13 DEBIT NOTE請款選項
optDB(0).Value = 0: optDB(1).Value = 0
Check7(2).Value = 0: Check7(3).Value = 0
Frame33(1).BackColor = &H8000000F
Frame33(1).Visible = False
'end 2023/11/13

End Sub

Private Function GetCopies(strStaffNo As String) As Integer
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset

StrSQLa = "Select ST06 From Staff Where ST01='" & strStaffNo & "'"
rsA.CursorLocation = adUseClient
rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
If rsA.RecordCount > 0 Then
    If "" & rsA.Fields(0).Value = "1" Then
        GetCopies = 2
    Else
        GetCopies = 3
    End If
Else
    GetCopies = 0
End If
If rsA.State <> adStateClosed Then rsA.Close
Set rsA = Nothing
End Function

Private Sub txtEnabled1(blnTF As Boolean)
Dim ii As Integer
    
    Me.Combo1(0).Enabled = blnTF
    'Modify By Sindy 2011/6/7
    If Text1(6) = "LA" Then
      Me.Text1(11).Enabled = True
    Else
    '2011/6/7 End
      Me.Text1(11).Enabled = blnTF
    End If
    Me.Combo5.Enabled = blnTF 'Add By Sindy 2010/10/28
    Me.Combo6.Enabled = blnTF 'Add By Sindy 2012/4/26
    
    'Add By Sindy 2010/11/22
    Text1(1).Enabled = True
    Text1(3).Enabled = True
    '2010/11/22 End
    
'    'Add By Sindy 2010/5/27
'    If optChoose(1).Value = True Then
'       Me.Text1(5).Enabled = blnTF
'       Me.Text1(130).Enabled = blnTF
'    End If
'    '2010/5/27 End
    
    For ii = 1 To 5
        Me.Text1(12 + (ii - 1) * 16).Enabled = blnTF
        Me.Text1(13 + (ii - 1) * 16).Enabled = blnTF
        Me.Text1(14 + (ii - 1) * 16).Enabled = blnTF
        Me.Text1(15 + (ii - 1) * 16).Enabled = blnTF
        Me.Text1(16 + (ii - 1) * 16).Enabled = blnTF
        Me.Text1(17 + (ii - 1) * 16).Enabled = blnTF
    
        Me.Text1(19 + (ii - 1) * 16).Enabled = blnTF
        Me.Text1(20 + (ii - 1) * 16).Enabled = blnTF
        Me.Text1(21 + (ii - 1) * 16).Enabled = blnTF
        Me.Text1(22 + (ii - 1) * 16).Enabled = blnTF
        Me.Text1(23 + (ii - 1) * 16).Enabled = blnTF
        Me.Text1(24 + (ii - 1) * 16).Enabled = blnTF
        Me.Text1(25 + (ii - 1) * 16).Enabled = blnTF
        Me.Text1(26 + (ii - 1) * 16).Enabled = blnTF
        Me.Text1(27 + (ii - 1) * 16).Enabled = blnTF
        
        Me.Text1(92 + (ii - 1)).Enabled = blnTF
        'Add by Amy 2016/01/04
        Me.Text1(120 + ii).Enabled = blnTF
        
        'Add By Sindy 2010/5/27
        If OptChoose(1).Value = True Then
            Me.Text1(34 + (ii - 1) * 16).Enabled = blnTF
            'Modify By Sindy 2023/2/7 國內接洽單也請顯示英文地址欄
            'Me.Text1(125 + (ii - 1) * 16).Enabled = blnTF
        End If
        '2010/5/27 End
        Me.Text1(125 + (ii - 1) * 16).Enabled = blnTF 'Modify By Sindy 2023/2/7 國內接洽單也請顯示英文地址欄
        Me.cmdSerach(ii - 1).Enabled = blnTF
    Next ii
End Sub

Private Sub txtEnabled2(blnTF As Boolean, ii As Integer)
      'blnTF = True 'Add By Sindy 查詢不要灰色字體
      Select Case ii
      Case 1
            Frame30(0).Enabled = blnTF 'Add By Sindy 2022/11/9
            If Me.Option31(0).Value = True Then
                Me.Text1(12 + (ii - 1) * 16).Enabled = True
                Me.Text1(21 + (ii - 1) * 16).Enabled = True
                Me.Text1(22 + (ii - 1) * 16).Enabled = True
                Me.cboContact(ii).Enabled = True              '2008/9/4  add by sonia
                Me.cmdSerach(ii - 1).Enabled = True
                'Add by Amy 2015/11/02 +顯示同上鈕
                CmdSame(ii).Visible = True
                Me.Text1(27 + (ii - 1) * 16).Width = 5730 '5700
                Me.Text1(27 + (ii - 1) * 16).Left = 1580
            Else
                CmdSame(ii).Visible = False
                Me.Text1(27 + (ii - 1) * 16).Width = 6300
                Me.Text1(27 + (ii - 1) * 16).Left = 1005
                'end 2015/11/02
                ChkCRA26(0).Value = 0 'Add By Sindy 2022/11/9
                ChkCRA27(0).Value = 0 'Add By Sindy 2022/11/9
            End If
      Case 2
            Frame30(1).Enabled = blnTF 'Add By Sindy 2022/11/9
            If Me.Option32(0).Value = True Then
                Me.Text1(12 + (ii - 1) * 16).Enabled = True
                Me.Text1(21 + (ii - 1) * 16).Enabled = True
                Me.Text1(22 + (ii - 1) * 16).Enabled = True
                Me.cboContact(ii).Enabled = True              '2008/9/4  add by sonia
                Me.cmdSerach(ii - 1).Enabled = True
                'Add by Amy 2015/11/02 +顯示同上鈕
                CmdSame(ii).Visible = True
                Me.Text1(27 + (ii - 1) * 16).Width = 5680 '5650
                Me.Text1(27 + (ii - 1) * 16).Left = 1610
            Else
                CmdSame(ii).Visible = False
                Me.Text1(27 + (ii - 1) * 16).Width = 6255
                Me.Text1(27 + (ii - 1) * 16).Left = 1030
                'end 2015/11/02
                ChkCRA26(1).Value = 0 'Add By Sindy 2022/11/9
                ChkCRA27(1).Value = 0 'Add By Sindy 2022/11/9
            End If
      Case 3
            Frame30(2).Enabled = blnTF 'Add By Sindy 2022/11/9
            If Me.Option33(0).Value = True Then
                Me.Text1(12 + (ii - 1) * 16).Enabled = True
                Me.Text1(21 + (ii - 1) * 16).Enabled = True
                Me.Text1(22 + (ii - 1) * 16).Enabled = True
                Me.cboContact(ii).Enabled = True              '2008/9/4  add by sonia
                Me.cmdSerach(ii - 1).Enabled = True
                'Add by Amy 2015/11/02 +顯示同上鈕
                CmdSame(ii).Visible = True
                Me.Text1(27 + (ii - 1) * 16).Width = 5735
                Me.Text1(27 + (ii - 1) * 16).Left = 1555
             Else
                CmdSame(ii).Visible = False
                Me.Text1(27 + (ii - 1) * 16).Width = 6285
                Me.Text1(27 + (ii - 1) * 16).Left = 1005
                'end 2015/11/02
                ChkCRA26(2).Value = 0 'Add By Sindy 2022/11/9
                ChkCRA27(2).Value = 0 'Add By Sindy 2022/11/9
            End If
      Case 4
            Frame30(3).Enabled = blnTF 'Add By Sindy 2022/11/9
            If Me.Option34(0).Value = True Then
                Me.Text1(12 + (ii - 1) * 16).Enabled = True
                Me.Text1(21 + (ii - 1) * 16).Enabled = True
                Me.Text1(22 + (ii - 1) * 16).Enabled = True
                Me.cboContact(ii).Enabled = True              '2008/9/4  add by sonia
                Me.cmdSerach(ii - 1).Enabled = True
                'Add by Amy 2015/11/02 +顯示同上鈕
                CmdSame(ii).Visible = True
                Me.Text1(27 + (ii - 1) * 16).Width = 5740
                Me.Text1(27 + (ii - 1) * 16).Left = 1510
             Else
                CmdSame(ii).Visible = False
                Me.Text1(27 + (ii - 1) * 16).Width = 6285
                Me.Text1(27 + (ii - 1) * 16).Left = 975
                'end 2015/11/02
                ChkCRA26(3).Value = 0 'Add By Sindy 2022/11/9
                ChkCRA27(3).Value = 0 'Add By Sindy 2022/11/9
            End If
      Case 5
            Frame30(4).Enabled = blnTF 'Add By Sindy 2022/11/9
            If Me.Option35(0).Value = True Then
                Me.Text1(12 + (ii - 1) * 16).Enabled = True
                Me.Text1(21 + (ii - 1) * 16).Enabled = True
                Me.Text1(22 + (ii - 1) * 16).Enabled = True
                Me.cboContact(ii).Enabled = True              '2008/9/4  add by sonia
                Me.cmdSerach(ii - 1).Enabled = True
                'Add by Amy 2015/11/02 +顯示同上鈕
                CmdSame(ii).Visible = True
                Me.Text1(27 + (ii - 1) * 16).Width = 5750
                Me.Text1(27 + (ii - 1) * 16).Left = 1550
             Else
                CmdSame(ii).Visible = False
                Me.Text1(27 + (ii - 1) * 16).Width = 6285
                Me.Text1(27 + (ii - 1) * 16).Left = 1005
                'end 2015/11/02
                ChkCRA26(4).Value = 0 'Add By Sindy 2022/11/9
                ChkCRA27(4).Value = 0 'Add By Sindy 2022/11/9
            End If
      Case Else
      End Select
      Me.Text1(13 + (ii - 1) * 16).Enabled = blnTF
      Me.Text1(14 + (ii - 1) * 16).Enabled = blnTF
      Me.Text1(15 + (ii - 1) * 16).Enabled = blnTF
      Me.Text1(16 + (ii - 1) * 16).Enabled = blnTF
      Me.Text1(17 + (ii - 1) * 16).Enabled = blnTF
      Me.Text1(19 + (ii - 1) * 16).Enabled = blnTF
      Me.Text1(20 + (ii - 1) * 16).Enabled = blnTF
      Me.Text1(22 + (ii - 1) * 16).Enabled = blnTF
      Me.Text1(23 + (ii - 1) * 16).Enabled = blnTF
      Me.Text1(24 + (ii - 1) * 16).Enabled = blnTF
      'Modify by Amy 2016/01/04 有值鎖住
      If blnTF = False Then
        If Me.Text1(25 + (ii - 1) * 16) <> MsgText(601) Then Me.Text1(25 + (ii - 1) * 16).Enabled = blnTF: Me.cmdSearchZip(ii * 2 - 2).Enabled = blnTF
      Else
        Me.Text1(25 + (ii - 1) * 16).Enabled = blnTF: Me.cmdSearchZip(ii * 2 - 2).Enabled = blnTF
      End If
      Me.Text1(26 + (ii - 1) * 16).Enabled = blnTF
      Me.Text1(27 + (ii - 1) * 16).Enabled = blnTF
      Me.Text1(92 + (ii - 1)).Enabled = blnTF
      'Add by Amy 2016/01/04 有值鎖住
      If blnTF = False Then
        If Me.Text1(119 + ii) <> MsgText(601) Then Me.Text1(119 + ii).Enabled = blnTF: Me.cmdSearchZip(ii * 2 - 1).Enabled = blnTF
      Else
        Me.Text1(119 + ii).Enabled = blnTF: Me.cmdSearchZip(ii * 2 - 1).Enabled = blnTF
      End If
      'Add By Sindy 2010/5/27
      If OptChoose(1).Value = True Then
         Me.Text1(34 + (ii - 1) * 16).Enabled = blnTF
         'Modify By Sindy 2023/2/7 國內接洽單也請顯示英文地址欄
         'Me.Text1(125 + (ii - 1) * 16).Enabled = blnTF
      End If
      '2010/5/27 End
      Me.Text1(125 + (ii - 1) * 16).Enabled = blnTF 'Modify By Sindy 2023/2/7 國內接洽單也請顯示英文地址欄
End Sub

Private Sub Text2_GotFocus(Index As Integer)
   TextInverse Me.Text2(Index)
   OpenIme
End Sub

Private Sub Text2_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
   If Button = 2 Then Forms(0).PopupMenu2 Text2(Index)
End Sub

'Private Sub Text2_Validate(Index As Integer, Cancel As Boolean)
'Dim StrSQLa As String
'Dim rsA As New ADODB.Recordset
'Dim arrNA01
'Dim tmpNa As String
'
'   If Me.Text2(Index).Text <> "" Then
'      tmpNa = "000"
'
'      StrSQLa = "Select NA01,NA03 from Nation where NA01='" & tmpNa & "'"
'      rsA.CursorLocation = adUseClient
'      rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
'      If rsA.RecordCount > 0 Then
'         Me.Combo3(Index).Text = rsA.Fields(0).Value & " " & rsA.Fields(1).Value
'         'Add By Sindy 2012/10/4
'         Me.Combo3(Index).Tag = rsA.Fields(0).Value & " " & "中華民國"
'         '2012/10/4 End
'      End If
'
'      If rsA.State <> adStateClosed Then rsA.Close
'      Set rsA = Nothing
'   End If
'End Sub

Private Sub Text3_GotFocus(Index As Integer)
   TextInverse Me.Text3(Index)
   CloseIme
End Sub

Private Sub Text3_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

'Private Sub Text3_Validate(Index As Integer, Cancel As Boolean)
'Dim strTmp As String, i As Integer, j As Integer
'
'   If Text3(Index).Text <> "" Then
'      'Modify  by Amy 2020/09/01 +if 國籍為大陸於列印前才檢查身份證號
'      If Left(Combo3(Index), 3) <> "020" Then
'            'Modify By Sindy 2023/5/16
'            'If GetTextLength(Text3(Index).Text) <> 10 Then
'            If Len(Text3(Index).Text) <> 10 Then
'            '2023/5/16 END
'                  Call Text3_GotFocus(Index)
'                  strTmp = "發明人ID必須是10碼 !"
'                  If MsgBox(strTmp, vbOKOnly) = vbOK Then
'                     Cancel = True
'                     Exit Sub
'                  '2008/9/3 ADD BY SONIA 不修改也不檢查ID
'                  Else
'                     Exit Sub
'                  End If
'            End If
'
'            If CheckID(i, Text3(Index).Text) = False Then
'               strTmp = "發明人ID錯誤，是否修正 ?"
'               If MsgBox(strTmp, vbYesNo + vbCritical) = vbYes Then
'                  Cancel = True
'                  Call Text3_GotFocus(Index)
'               End If
'            End If
'      End If
'   End If
'
'End Sub

Private Sub Text4_GotFocus(Index As Integer)
   TextInverse Me.Text4(Index)
   OpenIme
End Sub

Private Sub Text4_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
   If Button = 2 Then Forms(0).PopupMenu2 Text4(Index)
End Sub

Private Sub Text4_Validate(Index As Integer, Cancel As Boolean)
   If CheckLengthIsOK(Text4(Index), 70) = False Then
       Cancel = True
   End If
End Sub

Private Sub Text7_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   If Button = 2 Then Forms(0).PopupMenu2 Text7
End Sub

'Add by Amy 2017/01/20
Private Sub TxtC1_GotFocus(Index As Integer)
    TextInverse Me.TxtC1(Index)
End Sub

Private Sub TxtC1_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index = 0 Then
        KeyAscii = UpperCase(KeyAscii)
    End If
End Sub

'Private Sub TxtC1_LostFocus(Index As Integer)
'    Dim strMsg As String
'
'    If Trim(TxtC1(0)) = MsgText(601) Or Trim(TxtC1(1)) = MsgText(601) Then Exit Sub
'    If Index = 2 And TxtC1(2) = MsgText(601) Then TxtC1(2) = "0": Exit Sub
'    If Index <> 3 Then Exit Sub
'
'    If TxtC1(3) = MsgText(601) Then TxtC1(3) = "00"
'    Set G_SeekPicColor.Picture = LoadPicture()
'    Set tmpImg.Picture = LoadPicture()
'    If ExistCheck("TradeMark", "Tm01||Tm02||Tm03||Tm04", TxtC1(0) & TxtC1(1) & TxtC1(2) & TxtC1(3), strMsg) = False Then
''        TxtC1(0).SetFocus
'        Exit Sub
'    End If
'    'Modify by Amy 2018/07/31 ChkIsExistImg不使用
'    'If ChkIsExistImg(TxtC1(0), TxtC1(1), TxtC1(2), TxtC1(3)) = False Then
'    If ChkImgByteFile(TxtC1(0), TxtC1(1), TxtC1(2), TxtC1(3)) = False Then
'        MsgBox "此案號無代表圖！"
'        Exit Sub
'    End If
'
'    frmPic001.bolNoMsg = True
'    frmPic001.oCP01 = Me.TxtC1(0)
'    frmPic001.oCP02 = Me.TxtC1(1)
'    frmPic001.oCP03 = Me.TxtC1(2)
'    frmPic001.oCP04 = Me.TxtC1(3)
'    frmPic001.StrMenu
'    MoveFormToCenter frmPic001
'    frmPic001.cmdOK_Click (7) '複制
'    Set G_SeekPicColor.Picture = LoadPicture()
'    Set tmpImg.Picture = LoadPicture()
'    Set frmPic001.oPic = G_SeekPicColor
'    Set frmPic001.oImg = tmpImg
'    Set frmPic001.UpForm = Me
'    frmPic001.oRtPic = True
'    frmPic001.cmdOK_Click (0) '貼上
'    Unload frmpic002
'    frmPic001.cmdOK_Click (2) '確定
'    frmPic001.bolNoMsg = False
'End Sub

Private Sub TxtC1_Validate(Index As Integer, Cancel As Boolean)
    If Index <> 0 Then Exit Sub
    If Trim(TxtC1(0)) = MsgText(601) Then Exit Sub
    
    If Not (Trim(TxtC1(0)) = "T" Or Trim(TxtC1(0)) = "TF" Or Trim(TxtC1(0)) = "CFT") Then
        MsgBox "系統別只可輸入T/TF/CFT！"
        Cancel = True
'        TxtC1(0).SetFocus
        Exit Sub
    End If
End Sub
'end 2017/01/20

Private Sub txtItemCount_GotFocus()
   TextInverse txtItemCount
   CloseIme
End Sub

Private Sub txtItemCount_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 8 And Not IsNumeric(Chr(KeyAscii)) Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub txtItemList_GotFocus()
   CloseIme
End Sub

'Add by Sindy 2009/08/31
Private Sub LoadPic(strIBF01 As String, strIBF02 As String, strIBF03 As String, strIBF04 As String)
   Set PicRs = New ADODB.Recordset
   If PicRs.State = 1 Then PicRs.Close
   PicRs.CursorLocation = adUseClient
   'Modify By Sindy 2018/10/31 TF馬德里商標圖檔-子案的圖同母案
   '固定都以IBF01=tm01 AND IBF02=substr(tm02,1,5)||'0' AND IBF03='0' AND IBF04='00' 去抓代表圖
   If strIBF01 = "TF" Then
      PicRs.Open "select * from ImgByteFile where ibf05='1' and ibf01='" & strIBF01 & "' and ibf02='" & Mid(strIBF02, 1, 5) & "0" & "' and ibf03='0' And ibf04='00' ", cnnConnection, adOpenStatic, adLockOptimistic
   Else
   '2018/10/31 END
      PicRs.Open "select * from ImgByteFile where ibf05='1' and ibf01='" & strIBF01 & "' and ibf02='" & strIBF02 & "' and ibf03='" & strIBF03 & "' and ibf04='" & strIBF04 & "' ", cnnConnection, adOpenStatic, adLockOptimistic
   End If
   If PicRs.RecordCount <> 0 Then
      PicRs.MoveFirst
      '加入無圖式的格式
      If CheckStr(PicRs.Fields("ibf06")) = "3" Or CheckStr(PicRs.Fields("ibf06")) = "4" Or CheckStr(PicRs.Fields("ibf06")) = "6" Then
         IsWmf = True
      Else
         IsWmf = False
      End If
      'Add By Sindy 2017/8/10
'      If "" & PicRs.Fields("IBF15") <> "" Then
         If IsWmf = False Then
            Call PUB_GetFtpFile(PicRs.Fields("IBF15"), App.path & "\NowPic.jpg", UCase("ImgByteFile"))
         Else
            Call PUB_GetFtpFile(PicRs.Fields("IBF15"), App.path & "\NowPic.wmf", UCase("ImgByteFile"))
         End If
'      Else
'      '2017/8/10 END
'         ReDim bytes(Val(PicRs.Fields("ibf13").Value))
'         bytes() = PicRs.Fields("ibf14").GetChunk(Val(PicRs.Fields("ibf13").Value))
'         file_num = FreeFile
'         If IsWmf = False Then
'             Open App.path & "\NowPic.jpg" For Binary Access Write As #file_num
'         Else
'             Open App.path & "\NowPic.wmf" For Binary Access Write As #file_num
'         End If
'         Put #file_num, , bytes()
'         Close #file_num
'      End If
      
      If IsWmf = False Then
          tmpPic.Picture = LoadPicture(Trim(App.path & "\NowPic.jpg"))
          tmpImg.Picture = LoadPicture(Trim(App.path & "\NowPic.jpg"))
      Else
          tmpPic.Picture = LoadPicture(Trim(App.path & "\NowPic.wmf"))
          tmpImg.Picture = LoadPicture(Trim(App.path & "\NowPic.wmf"))
      End If
      If Dir(App.path & "\NowPic.jpg") <> "" Then
          Kill App.path & "\NowPic.jpg"
      End If
      If Dir(App.path & "\NowPic.wmf") <> "" Then
          Kill App.path & "\NowPic.wmf"
      End If
   End If
End Sub

'Add by Morgan 2008/9/16
'接洽人控制
Private Sub LockContact(Optional Index As Integer = -1)
   Dim ii As Integer
   Dim bNewCase As Boolean, bNewCust As Boolean
   Dim bChange As Boolean 'Added by Lydia 2018/12/03 國內接洽單收讓與案時，開放可選擇接洽單
   
   m_strCaseCPM = GetAllCaseCPM 'Add By Sindy 2022/9/8 取得案件性質代碼
   bNewCase = Option1(0).Value
   'Added by Lydia 2018/12/03 Added by Lydia 2018/12/03 國內接洽單收讓與案時，開放可選擇接洽單
    If Option1(1).Value = True And ((Trim(Text1(6).Text) = "P" Or _
        Trim(Text1(6).Text) = "FCP" Or _
        Trim(Text1(6).Text) = "CFP") And _
        (InStr(m_strCaseCPM, "701") > 0 Or InStr(m_strCaseCPM, "703") > 0 Or InStr(m_strCaseCPM, "708") > 0)) Or _
        ((Trim(Text1(6).Text) = "CFT" Or _
        Trim(Text1(6).Text) = "CFC" Or _
        Trim(Text1(6).Text) = "FCT" Or _
        Trim(Text1(6).Text) = "T" Or _
        Trim(Text1(6).Text) = "TB" Or _
        Trim(Text1(6).Text) = "TC" Or _
        Trim(Text1(6).Text) = "TF") And _
        (InStr(m_strCaseCPM, "501") > 0)) Then
            bChange = True
   End If
   'end 2018/12/03
   
   For ii = 1 To 5
      If Index = -1 Or Index = ii Then
         Select Case ii
            Case 1
               bNewCust = Option31(0).Value
            Case 2
               bNewCust = Option32(0).Value
            Case 3
               bNewCust = Option33(0).Value
            Case 4
               bNewCust = Option34(0).Value
            Case 5
               bNewCust = Option35(0).Value
         End Select
         If bNewCust = True Then
            cboContact(ii).Clear
            cboContact(ii).Enabled = True
            cboContact(ii).Tag = "1"
            m_strContactList(ii) = ""
         'Modified by Lydia 2018/12/03
         'ElseIf bNewCase = True Then
         ElseIf bNewCase = True Or bChange = True Then
            cboContact(ii).Enabled = True
            cboContact(ii).Tag = "0"
         Else
             cboContact(ii).Enabled = False
            cboContact(ii).Tag = "0"
         End If
         If cboContact(ii).ListIndex = -1 Then
            cboContact(ii) = ""
         End If
      End If
      If Index = ii Then Exit For
   Next
End Sub

'Add By Sindy 2010/6/21
Sub InitAll()
'    DIB.Destroy
'    DIBPal.Clear
'    PicFrame.Clear
    Set tmpPic.Picture = LoadPicture()
    Set tmpImg.Picture = LoadPicture()
    Set G_SeekPicColor.Picture = LoadPicture()
'    Set G_SeekPicBW.Picture = LoadPicture()
'    Set Me.CropPic.Picture = LoadPicture()
'    Set PicMain.Picture = LoadPicture()
    Set pic1.Picture = LoadPicture()
End Sub

'Add By Sindy 2010/6/21
Sub PicToObj(oFileNameAndPath As String)
    Dim objImg As StdPicture
    Dim nSrcWidth, nSrcHeight, nWidth, nHeight
    Dim tBI      As BITMAP
    InitAll
    On Error GoTo BE
    If UCase(pvGetExt(oFileNameAndPath)) = "WMF" Or UCase(pvGetExt(oFileNameAndPath)) = "EMF" Then
        Set G_SeekPicColor.Picture = LoadPicture(App.path & "\NowPic.wmf")
        'tmpImg.BackColor = &H80000009
    Else
        'tmpPic.BackColor = &H8000000A
        Set objImg = pvGetStdPicture(oFileNameAndPath)
        Set m_Image = New cImage
        Set m_Jpeg = New cJpeg
        Set G_SeekPicColor.Picture = objImg
        If G_SeekPicColor.Picture <> 0 Then
            Call GetObject(objImg.handle, Len(tBI), tBI)
            If tBI.bmWidth = 0 Or tBI.bmHeight = 0 Then
                MsgBox "發生錯誤！", vbExclamation, "圖檔格式錯誤"
'                'edit by nickc 2007/11/19
'                If oRtPic = False Then
'                    StrMenu
'                End If
                Exit Sub
            End If
            If tBI.bmWidth > 2000 Or tBI.bmHeight > 2000 Then
                nSrcWidth = tBI.bmWidth
                nSrcHeight = tBI.bmHeight
                If nSrcWidth > nSrcHeight Then
                    nWidth = 1200
                    nHeight = nSrcHeight / (nSrcWidth / nWidth)
                ElseIf nSrcWidth < nSrcHeight Then
                    nHeight = 1200
                    nWidth = nSrcWidth / (nSrcHeight / nHeight)
                Else
                    nHeight = 1200
                    nWidth = 1200
                End If
                pic1.Width = nWidth
                pic1.Height = nHeight
                pic1.BackColor = &H8000000A
                '重新定義大小
                pic1.Scale (0, 0)-(nWidth, nHeight)
                '縮小
                pic1.PaintPicture objImg, 0, 0, nWidth, nHeight, , , , , vbSrcCopy
                '存檔
                SavePicture pic1.Image, App.path & "\NowPic.bmp"
                Set objImg = pvGetStdPicture(App.path & "\NowPic.bmp")
            End If
            m_Image.CopyStdPicture objImg
            ' m_Jpeg.SetSamplingFrequencies 1, 1, 0, 0, 0, 0 'Removed by Morgan 2023/7/27 不必轉灰階
            m_Jpeg.Quality = 75
            m_Jpeg.SampleHDC m_Image.hDC, m_Image.Width, m_Image.Height
            RidFile App.path & "\NowPic.jpg"
            m_Jpeg.SaveFile App.path & "\NowPic.jpg"
            m_Image.CopyStdPicture pvGetStdPicture(App.path & "\NowPic.jpg")
            Set G_SeekPicColor.Picture = LoadPicture(App.path & "\NowPic.jpg")

        End If
    End If
    Dim t_hd As Double
    Dim t_wd As Double
    t_hd = G_SeekPicColor.ScaleHeight / tmpPic.ScaleHeight
    t_wd = G_SeekPicColor.ScaleWidth / tmpPic.ScaleWidth
    If t_hd > t_wd Then
        t_wd = G_SeekPicColor.ScaleWidth / t_hd
        t_hd = G_SeekPicColor.ScaleHeight / t_hd
    Else
        t_hd = G_SeekPicColor.ScaleHeight / t_wd
        t_wd = G_SeekPicColor.ScaleWidth / t_wd
    End If
        tmpImg.Width = t_wd
        tmpImg.Height = t_hd
        tmpImg.Move (tmpPic.ScaleWidth - tmpImg.Width) / 2, (tmpPic.ScaleHeight - tmpImg.Height) / 2, t_wd, t_hd
    
    Set tmpImg.Picture = G_SeekPicColor.Picture
    Set objImg = Nothing

    Exit Sub
BE:
    Resume Next
End Sub

'Add By Sindy 2010/6/21
Public Function QueryData() As Boolean
Dim intIndex As Integer
Dim Cancel  As Boolean
Dim strTemp As String
Dim rsA As New ADODB.Recordset
Dim rsD As New ADODB.Recordset
Dim stAttPath As String, stFullName As String
Dim bolNotTransCase As Boolean '非轉案
   
   QueryData = True
   If m_blnCallPrint = True Then cmdCRL55.Visible = False
   
   '*****************
   '接洽記錄單主檔
   '*****************
   'Modify by Morgan 2010/12/20 +抓進度檔
   'strExc(0) = "Select * From consultrecordlist Where crl01 ='" & Trim(Text5) & "' "
   strExc(0) = "Select * From consultrecordlist,caseprogress Where crl01 ='" & Trim(Text5) & "' and cp140(+)=crl01 order by cp01 asc,cp02 asc,cp03 asc,cp04 asc"
   intI = 1
   Set rsA = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      rsA.MoveFirst
'      Do While Not rsA.EOF
         m_CP05 = "" & rsA.Fields("cp05") 'Add by Morgan 2010/12/20
         m_CP09 = "" & rsA.Fields("cp09") 'Add by Morgan 2010/12/20
         
         '案件資料/收費項目
         Me.SSTab1.Tab = 0
         m_CRL02 = "" & rsA.Fields("CRL02") '填表日期 Add By Sindy 2011/10/18
         lblDate.Caption = Val(Mid(rsA.Fields("CRL02"), 1, 4)) - 1911 & "年" & Mid(rsA.Fields("CRL02"), 5, 2) & "月" & Mid(rsA.Fields("CRL02"), 7, 2) & "日"
         Text1(10) = rsA.Fields("CRL03")
         lblStaffName = GetStaffName(Me.Text1(10).Text)
         Text1(0) = "" & rsA.Fields("CRL04")
         'Added by Lydia 2023/12/27
         If m_CRL02 > 新部門啟用日 Then
            lblZone = GetDeptNameA0922(Text1(10))
         Else
         'end 2023/12/27
            lblZone = GetDepartmentName(Me.Text1(0).Text)
         End If
         If rsA.Fields("CRL05") = "1" Then
            OptChoose(0).Value = True
         ElseIf rsA.Fields("CRL05") = "2" Then
            OptChoose(1).Value = True
         End If
         If "" & rsA.Fields("CRL06") = "Y" Then
            Option1(0).Value = True
         Else
            Option1(1).Value = True
         End If
         Text1(6) = "" & rsA.Fields("CRL07")
         Text1(6).Tag = Text1(6).Text 'Added by Morgan 2020/5/28
         Call Text1_LostFocus(6) 'Add By Sindy 2022/10/14
         Text1(7) = "" & rsA.Fields("CRL08")
         Text1(8) = "" & rsA.Fields("CRL09")
         Text1(9) = "" & rsA.Fields("CRL10")
         Call GetMainData '抓取主檔資料
         If "" & rsA.Fields("CRL11") = "Y" Then ChkPCT.Value = 1 Else ChkPCT.Value = 0
         If Not IsNull(rsA.Fields("CRL12")) Then Text1(3) = ChangeWStringToTString(rsA.Fields("CRL12"))
         If Not IsNull(rsA.Fields("CRL13")) Then Text1(1) = ChangeWStringToTString(rsA.Fields("CRL13"))
         '申請國家
         If Not IsNull(rsA.Fields("CRL15")) Then Combo1(0) = rsA.Fields("CRL15"): Call Combo1_Validate(0, False): Call Combo1_LostFocus(0)
         
         'Add By Sindy 2022/8/29
'*****************
'案件性質
'*****************
         '先檢查是否已收文
         strExc(0) = "Select crc01,CRC08 " & _
                       "From ConsultRecCMP " & _
                     "Where crc01 ='" & Trim(Text5) & "' and CRC08 is not null "
         intI = 1
         Set rsD = ClsLawReadRstMsg(intI, strExc(0))
         LblRecved.Visible = False
         If intI = 1 Then
            LblRecved.Visible = True
         End If
         GridCase.Clear
         Call SetGrd
         'Modify By Sindy 2023/7/18
'         strExc(0) = "Select CRC02 順序,CRC03||' '||decode(crl15,'000',CPM03,CPM04) 案件性質,decode(CRC04,'有修改',CRC04,'刪收文',CRC04,to_char(CRC04,'99,999,999')) 費用" & _
'                     ",decode(CRC05,'有修改',CRC05,'刪收文',CRC05,to_char(CRC05,'99,999,999')) 規費,decode(CRC06,'有修改',CRC06,'刪收文',CRC06,to_char(CRC06,'99,999.000')) 點數,CRC07 備註,cp01||cp02||cp03||cp04 案號,CRC08 總收文號" & _
'                     ",CRC10 算案件數,CRC11 計件值,CRC12 加乘註記,cp01,cp02,cp03,cp04,decode(crl15,'000',CPM03,CPM04) CPMn " & _
'                     "From ConsultRecCMP,consultrecordlist,CasePropertyMap,caseprogress " & _
'                     "Where crc01 ='" & Trim(Text5) & "' and crl01(+)=crc01 and CPM01(+)=crl07 And CPM02(+)=crc03 And CRC08=CP09(+) " & _
'                     "order by crc02 asc "
         'Modify By Sindy 2025/4/14 +,CRC13 規費調整
         strExc(0) = "Select CRC02 順序,CRC03||' '||decode(crl15,'000',CPM03,CPM04) 案件性質" & _
                     ",decode(translate(CRC04,'/0123456789','/'),null,to_char(CRC04,'99,999,999'),CRC04) 費用" & _
                     ",decode(translate(CRC05,'/0123456789','/'),null,to_char(CRC05,'99,999,999'),CRC05) 規費" & _
                     ",decode(translate(CRC06,'/0123456789','/'),null,to_char(CRC06,'99,999.000'),CRC06) 點數" & _
                     ",CRC07 備註,cp01||cp02||cp03||cp04 案號,CRC08 總收文號" & _
                     ",CRC10 算案件數,CRC11 計件值,CRC12 加乘註記,cp01,cp02,cp03,cp04,decode(crl15,'000',CPM03,CPM04) CPMn" & _
                     ",decode(translate(CRC05,'/0123456789','/'),null,CRC13,'') 規費調整" & _
                     " From ConsultRecCMP,consultrecordlist,CasePropertyMap,caseprogress " & _
                     "Where crc01 ='" & Trim(Text5) & "' and crl01(+)=crc01 and CPM01(+)=crl07 And CPM02(+)=crc03 And CRC08=CP09(+) " & _
                     "order by crc02 asc "
         '2023/7/18 END
         intI = 1
         Set rsD = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            If rsD.RecordCount > 0 Then
               Set GridCase.Recordset = rsD
               FrameCRC.Caption = "案件性質區（" & rsD.RecordCount & "）"
               rsD.MoveFirst
               Do While Not rsD.EOF
                  'Add By Sindy 2023/6/19
                  If rsD.Fields("cp01") = "TF" And rsD.Fields("案件性質") = "104 領土延伸" Then
                     bolNotTransCase = True '非轉案
                  End If
                  '2023/6/19 END
                  
                  If "" & rsD.Fields("計件值") <> "" Or "" & rsD.Fields("加乘註記") <> "" Then
                     Text1(136) = Text1(136) & rsD.Fields("CPMn") & "："
                     If "" & rsD.Fields("計件值") <> "" Then
                        Text1(136) = Text1(136) & "計件值(" & rsD.Fields("計件值") & ");"
                     End If
                     If "" & rsD.Fields("加乘註記") <> "" Then
                        Text1(136) = Text1(136) & "加乘註記(" & rsD.Fields("加乘註記") & ");"
                     End If
                  End If
                  rsD.MoveNext
               Loop
               rsD.MoveFirst
'               '填入案號
'               If Option1(0).Value = True And "" & rsD.Fields("cp02") <> "" Then
'                  Text1(7).Text = "" & rsD.Fields("cp02")
'                  Text1(8).Text = "" & rsD.Fields("cp03")
'                  Text1(9).Text = "" & rsD.Fields("cp04")
'               End If
               'Modify By Sindy 2023/3/24
               '有轉案填入案號
               m_strTransCase = ""
               If "" & rsD.Fields("cp02") <> "" And bolNotTransCase = False Then 'Modify By Sindy 2023/6/19 排除非轉案
                  If Text1(6).Text & Text1(7).Text <> rsD.Fields("cp01") & rsD.Fields("cp02") Then
                     m_strTransCase = "（原為" & Text1(6).Text & Text1(7).Text & Text1(8).Text & Text1(9).Text & "案）"
                     FrameCRC.Caption = FrameCRC.Caption & "　" & m_strTransCase
                     Text1(6).Text = "" & rsD.Fields("cp01")
                     Text1(7).Text = "" & rsD.Fields("cp02")
                     Text1(8).Text = "" & rsD.Fields("cp03")
                     Text1(9).Text = "" & rsD.Fields("cp04")
                     Call GetMainData 'Added by Morgan 2025/9/8 要重抓取主檔資料,否則原案號若已不存在時會沒有設定到基本資料
                     Text1(6).Tag = Text1(6).Text 'Add by Sindy 2025/9/9
                     Call Text1_LostFocus(6) 'Add by Sindy 2025/9/9
                  End If
               End If
               '2023/3/24 END
            End If
         End If
         rsD.Close
         GridCase.col = 1
         GridCase.row = 1
         m_strCaseCPM = GetAllCaseCPM
         Call SetFrmCol 'Add By Sindy 2022/12/26
         '2022/8/29 END
         
         If Not IsNull(rsA.Fields("CRL16")) Then Text1(18) = rsA.Fields("CRL16")
         If Not IsNull(rsA.Fields("CRL17")) Then Text1(11) = rsA.Fields("CRL17")
         If "" & rsA.Fields("CRL18") = "Y" Then
            Check7(0).Value = 1
         ElseIf "" & rsA.Fields("CRL18") = "N" Then
            Check7(1).Value = 1
         Else
            Check7(0).Value = 0
            Check7(1).Value = 0
         End If
'         'Add By Sindy 2022/9/13 +CRL134(對造陳報主管)/135(原因)
'         m_stCRL134 = ""
'         If Not IsNull(rsA.Fields("CRL134")) Then m_stCRL134 = rsA.Fields("CRL134")
'         m_stCRL135 = ""
'         If Not IsNull(rsA.Fields("CRL135")) Then m_stCRL135 = rsA.Fields("CRL135")
'         '2022/9/13 END
         
         If Not IsNull(rsA.Fields("CRL39")) Then Text1(113) = rsA.Fields("CRL39") '後金
                  
         If Not IsNull(rsA.Fields("CRL58")) Then Text1(4) = rsA.Fields("CRL58")
         Text1(114) = "" & rsA.Fields("CRL40") 'Added by Morgan 2012/12/18 改放申請技術報告項數
         
         If "" & rsA.Fields("CRL41") = "1" Then
            Option2(0).Value = True
         ElseIf "" & rsA.Fields("CRL41") = "2" Then
            Option2(2).Value = True
         ElseIf "" & rsA.Fields("CRL41") = "3" Then
            Option2(1).Value = True
            If Not IsNull(rsA.Fields("CRL42")) Then cboTitle = rsA.Fields("CRL42")
         End If
         'Added by Lydia 2023/11/13 DEBIT NOTE請款選項
         'Modified by Lydia 2023/11/14 排除外至台案件(MCTF) => rsA.Fields("CRL05") = "1"
         If "" & rsA.Fields("CRL41") = "2" And "" & rsA.Fields("CRL05") = "1" Then
            Frame33(1).Visible = True
            If "" & rsA.Fields("CRL153") = "1" Then
               optDB(0).Value = 1
            Else
               optDB(1).Value = 1
               If "" & rsA.Fields("CRL153") = "2" Then
                  Check7(2).Value = 1
               Else
                  Check7(3).Value = 1
               End If
            End If
         Else
            Frame33(1).Visible = False
         End If
         'end 2023/11/13
         
         If "" & rsA.Fields("CRL43") = "Y" Then
            Check6(0).Value = 1
         Else
            Check6(0).Value = 0
         End If
         If Not IsNull(rsA.Fields("CRL44")) Then Text1(116) = rsA.Fields("CRL44")
         If "" & rsA.Fields("CRL45") = "Y" Then
            Check6(1).Value = 1
         Else
            Check6(1).Value = 0
         End If
         If Not IsNull(rsA.Fields("CRL46")) Then Text1(117) = ChangeWStringToTString(rsA.Fields("CRL46"))
         If "" & rsA.Fields("CRL47") = "Y" Then
            Check1.Value = 1
         Else
            Check1.Value = 0
         End If
         If Not IsNull(rsA.Fields("CRL48")) Then Text1(118) = ChangeWStringToTString(rsA.Fields("CRL48"))
         'Modify By Sindy 2016/3/18
         'If Not IsNull(rsA.Fields("CRL49")) Then Combo4.ListIndex = rsA.Fields("CRL49")
         If Not IsNull(rsA.Fields("CRL49")) Then
            Call SetCombo4(m_CRL02) '用舊公司名稱預設 Modify By Sindy 2025/4/16 從下面程式往上移
            'Added by Lydia 2020/03/30
            If m_CRL02 >= 智慧所更名日 Then
                Select Case "" & rsA.Fields("CRL49")
                     Case "1": Combo4.ListIndex = m_Comp1forIdx
                     Case "2": Combo4.ListIndex = m_Comp2forIdx
                     Case "J": Combo4.ListIndex = m_CompJforIdx
                     Case "L": Combo4.ListIndex = m_CompLforIdx
                End Select
            Else
            'end 2020/03/30
                'Modified by Lydia 2020/03/30
                'If rsA.Fields("CRL49") = 3 Then
                '   Combo4.AddItem "智權公司"
                'End If
                m_CompName2 = "" '清空目前預設
                'Call SetCombo4(m_CRL02) '用舊公司名稱預設 Modify By Sindy 2025/4/16 mark
                'end 2020/03/30
                Combo4.ListIndex = rsA.Fields("CRL49")
            End If
         End If
         '2016/3/18 END

         If "" & rsA.Fields("CRL50") = "Y" Then
            Check2.Value = 1
         Else
            Check2.Value = 0
         End If
         '案件來源及申請人
         'Me.SSTab1.Tab = 1
         If "" & rsA.Fields("CRL51") = "01" Then
            Option4(0).Value = True
         ElseIf "" & rsA.Fields("CRL51") = "02" Then
            Option4(3).Value = True
         End If
         If "" & rsA.Fields("CRL52") = "1" Then
            Option4(1).Value = True
         ElseIf "" & rsA.Fields("CRL52") = "2" Then
            Option4(2).Value = True
         End If
         If Not IsNull(rsA.Fields("CRL53")) Then
            Text1(97) = rsA.Fields("CRL53")
            'Add By Sindy 2022/12/6 財務處要顯示員工編號
            '依員工姓名抓取員工編號
            strExc(10) = GetPrjSalesNM_2(rsA.Fields("CRL53"), , , , , False)
            If strExc(10) <> "" Then
               Label1(142).Visible = True
               Label1(142).Caption = "( " & strExc(10) & " )"
            End If
         End If
         If Not IsNull(rsA.Fields("CRL54")) Then Text1(99) = rsA.Fields("CRL54")
         If "" & rsA.Fields("CRL56") = "Y" Then
            Option5(0).Value = True
         ElseIf "" & rsA.Fields("CRL56") = "N" Then
            Option5(1).Value = True
         End If
         
         If Not IsNull(rsA.Fields("CRL60")) Then Text1(5) = rsA.Fields("CRL60") & rsA.Fields("CRL61")
         If Not IsNull(rsA.Fields("CRL62")) Then Text1(130) = rsA.Fields("CRL62")
         If Not IsNull(rsA.Fields("CRL77")) Then Text1(137) = rsA.Fields("CRL77") 'Add By Sindy 2022/10/14 代理人彼所案號
         
         '商標圖
         'Me.SSTab1.Tab = 4
         If "" & rsA.Fields("CRL63") = "1" Then
            opt1(0).Value = True
         ElseIf "" & rsA.Fields("CRL63") = "2" Then
            opt1(1).Value = True
         ElseIf "" & rsA.Fields("CRL63") = "3" Then
            opt1(2).Value = True
         ElseIf "" & rsA.Fields("CRL63") = "4" Then
            opt1(3).Value = True
         ElseIf "" & rsA.Fields("CRL63") = "5" Then
            opt1(4).Value = True
         End If
         If Not IsNull(rsA.Fields("CRL64")) Then PicText = rsA.Fields("CRL64")
         'Add By Sindy 2022/9/20 商標圖同卷號
         If Not IsNull(rsA.Fields("CRL75")) Then
            Frame5.Visible = True
            TxtC1(0) = Mid(rsA.Fields("CRL75"), 1, Len(rsA.Fields("CRL75")) - 9)
            TxtC1(1) = Mid(rsA.Fields("CRL75"), Len(rsA.Fields("CRL75")) - 8, 6)
            TxtC1(2) = Mid(rsA.Fields("CRL75"), Len(rsA.Fields("CRL75")) - 2, 1)
            TxtC1(3) = Mid(rsA.Fields("CRL75"), Len(rsA.Fields("CRL75")) - 1, 2)
         End If
         
         'Add By Sindy 2022/9/20 7501判決分析 提供書面分析 / 請律師向當事人說明
         If "" & rsA.Fields("CRL72") <> "" Then
            Frame41.Visible = True
            If rsA.Fields("CRL72") = "1" Then
               Option9(0).Value = True '提供書面分析
            ElseIf rsA.Fields("CRL72") = "2" Then
               Option9(1).Value = True '請律師向當事人說明
            End If
         End If
                  
         'Add By Sindy 2022/9/23 分割成幾案
         Label1(123).Visible = False
         If Not IsNull(rsA.Fields("CRL76")) Then
            Is307or308 = True
            CountBy307308 = rsA.Fields("CRL76")
            Label1(123).Visible = True
            Label1(123).Caption = "分割成 " & CountBy307308 & " 案"
         End If
         
         'Add By Sindy 2022/9/15
         '證書形式
         Label1(141).Visible = False
         Text1(145).Visible = False
         Label28(1).Visible = False
         If Not IsNull(rsA.Fields("CRL59")) Then
            Label1(141).Visible = True
            Text1(145).Visible = True
            Label28(1).Visible = True
            Text1(145) = rsA.Fields("CRL59")
         End If
         '關連表單編號
         If Not IsNull(rsA.Fields("CRL65")) Then
            Text1(144) = rsA.Fields("CRL65")
         End If
         '對造已簽准
         If "" & rsA.Fields("CRL66") = "Y" Then
            ChkCRL66.Value = 1
         Else
            ChkCRL66.Value = 0
         End If
         'Add By Sindy 2023/4/7 自行送簽核
         If "" & rsA.Fields("CRL152") = "Y" Then
            ChkCRL152.Value = 1
         Else
            ChkCRL152.Value = 0
         End If
         '2023/4/7 END
         
         '案件說明處理事項
         'Me.SSTab1.Tab = 2
         If Not IsNull(rsA.Fields("CRL57")) Then
            Text1(119) = rsA.Fields("CRL57")
         'Add By Sindy 2024/6/12
         Else
            Text1(119) = ""
            '2024/6/12 END
         End If
         '系統加註
         If Not IsNull(rsA.Fields("CRL70")) Then txtCRL70 = rsA.Fields("CRL70")
         'Add By Sindy 2025/7/1 人員有勾規費調整
         If InStr(txtCRL70, "有勾規費調整") > 0 Then
            Me.Check10.Value = 1
         End If
         '2025/7/1 END
         '呈主管簽核
         If Not IsNull(rsA.Fields("CRL69")) Then txtCRL69 = rsA.Fields("CRL69")
         
         If Text1(8) & Text1(9) = "000" Then
            strExc(1) = Text1(6) & "-" & Text1(7)
         Else
            strExc(1) = Text1(6) & "-" & Text1(7) & "-" & Text1(8) & "-" & Text1(9)
         End If
         '一案兩請案號
         If Not IsNull(rsA.Fields("CRL67")) Then
            If InStr(rsA.Fields("CRL67"), strExc(1)) = 0 Then
               Text1(134) = rsA.Fields("CRL67")
            End If
         End If
         '擬制喪失新穎性案號
         If Not IsNull(rsA.Fields("CRL68")) Then
            If InStr(rsA.Fields("CRL68"), strExc(1)) = 0 Then
               Text1(135) = rsA.Fields("CRL68")
            End If
         End If
         '相關案號
         If Not IsNull(rsA.Fields("CRL55")) Then
            'Modify By Sindy 2023/4/20 Mark:P-131300收125衍生設計申請,不用判斷此if
            'If InStr(rsA.Fields("CRL55"), strExc(1)) = 0 Then
               Text1(100) = rsA.Fields("CRL55") '本案與總號
            'End If
         End If
         'Add By Sindy 2022/9/20 相關案號的類別
         If Not IsNull(rsA.Fields("CRL74")) Then
            cmdCRL55.Visible = False
            If Len(rsA.Fields("CRL74")) = 2 Then
               Label1(84).Caption = "案源單號："
               strLSourceType = rsA.Fields("CRL74")
               strLOS15 = SrcGetLOS15(Text5)
               
            ElseIf rsA.Fields("CRL74") = "1" And Not IsNull(rsA.Fields("CRL55")) Then '1.查名代號
               Label1(84).Caption = "查名代號："
               mTQC01 = rsA.Fields("CRL55")
               cmdTMQ.Tag = mTQC01
               'Added by Lydia 2024/11/11 查名單(網中)
               m_UseTmqTma = "1" '預設使用原查名單
               If strSrvDate(1) >= 查名單網中系統平行測試 Then
                  strExc(0) = "SELECT TMA01,TMA08 From TMQCASEMAP,TMQAPPFORM WHERE TQC01='" & mTQC01 & "' and tqc03<>'" & cntTQC自動記錄 & "' AND TQC03=TMA01(+) AND NVL(TMA01,'N') <> 'N' "
                  intI = 1
                  Set rsD = ClsLawReadRstMsg(intI, strExc(0))
                  If intI = 1 Then
                     m_UseTmqTma = "2"
                  End If
               End If
               'end 2024/11/11
               Call QueryTMQ(mTQC01)
            ElseIf rsA.Fields("CRL74") = "2" And Not IsNull(rsA.Fields("CRL55")) Then '2.英國脫歐案
               Label1(84).Caption = "英國脫歐案："
               strCaseNA239 = rsA.Fields("CRL55")
            ElseIf rsA.Fields("CRL74") = "3" And Not IsNull(rsA.Fields("CRL55")) Then '3.CFT緬甸重新申請案
               Label1(84).Caption = "CFT緬甸重新申請案："
               strExc(10) = rsA.Fields("CRL55")
               strCaseNo1 = Left(strExc(10), 3)
               strCaseNo2 = Mid(strExc(10), 4, 6)
               StrCaseNo3 = Mid(strExc(10), 10, 1)
               strCaseNo4 = Mid(strExc(10), 11, 2)
            'Add By Sindy 2025/8/4
            ElseIf rsA.Fields("CRL74") = "4" And Not IsNull(rsA.Fields("CRL55")) Then '4.C類來函
               Label1(84).Caption = "C類來函："
               '2025/8/4 END
            End If
         End If
         '2022/9/15 END
         
         'Add By Sindy 2022/10/14
         Frame58.Visible = False
         If Not IsNull(rsA.Fields("CRL73")) Then
            Text1(138) = rsA.Fields("CRL73") '商品類別
            Frame58.Visible = True
         End If
         '2022/10/14 END
         
         'Added by Morgan 2020/5/28
         SetCombo6
         If IsNull(rsA.Fields("CRL87")) Then
            Combo6 = ""
         Else
            If InStr(Text1(6), "P") > 0 Then
               Combo6 = rsA.Fields("CRL87") + "." + GetPKindName(rsA.Fields("CRL87"))
            ElseIf InStr(Text1(6), "T") > 0 Then
               Combo6 = rsA.Fields("CRL87") + "." + GetTKindName(rsA.Fields("CRL87"))
            End If
         End If
         If Left(Combo6, 1) = "3" Then
            chkItem(0).Enabled = False
            chkItem(0).Value = vbUnchecked
            txtItemCount.Enabled = False
            txtItemCount = ""
            chkItem(1).Enabled = False
            chkItem(1).Value = vbUnchecked
            txtItemList = "第項"
            chkItem(2).Enabled = True
         Else
            chkItem(0).Enabled = True
            chkItem(1).Enabled = True
            chkItem(2).Value = vbUnchecked
            chkItem(2).Enabled = False
         End If
         'end 2020/5/28
         
         'Modify By Sindy 2022/11/9 Mark
         SetCombo5 'Added by Morgan 2020/5/28
         '2022/11/9 END
         'Add By Sindy 2010/10/28 案件屬性
         If IsNull(rsA.Fields("CRL81")) Then
            Combo5 = ""
         Else
            'Modified by Lydaia 2024/04/24 +專利種類CRL87
            Combo5 = "" & rsA.Fields("CRL81") + "." + PUB_GetCaseAttributeName("" & rsA.Fields("CRL81"), "" & rsA.Fields("CRL87"))
         End If
         
         'Add By Sindy 2023/12/12
         If m_CRL02 < 指定日期啟用日 Then
            Frame3.Visible = False
         Else
            Frame3.Visible = True
         End If
         '2023/12/12 END
         'Add By Sindy 2010/11/22 送件方式
         Frame18.Visible = True 'Add By Sindy 2024/3/18
         If Not IsNull(rsA.Fields("CRL82")) Then
            Frame18.Visible = True
            'Modify By Sindy 2024/1/22
            '送件方式選項1的中文
            OptSendType(0).Caption = PUB_GetCP114Opt1Desc(Trim(Text1(6)), m_strCaseCPM)
            '2024/1/22 END
            If "" & rsA.Fields("CRL82") = "1" Then
               OptSendType(0).Value = True
            ElseIf "" & rsA.Fields("CRL82") = "2" Then
               OptSendType(1).Value = True
            ElseIf "" & rsA.Fields("CRL82") = "3" Then
               OptSendType(2).Value = True
            End If
            '指定日期
            If Not IsNull(rsA.Fields("CRL83")) Then Text1(126) = ChangeWStringToTString(rsA.Fields("CRL83"))
            'Add By Sindy 2023/12/12 指定日期方式
            If Not IsNull(rsA.Fields("CRL155")) Then
               If rsA.Fields("CRL155") = "1" Then
                  OptCP164(0).Value = True
               ElseIf rsA.Fields("CRL155") = "2" Then
                  OptCP164(1).Value = True
               ElseIf rsA.Fields("CRL155") = "3" Then
                  OptCP164(2).Value = True
               End If
            End If
            '2023/12/12 END
         Else
            OptSendType(0).Value = False
            OptSendType(1).Value = False
            OptSendType(2).Value = False
            Frame18.Visible = False
         End If
         
         'Add By Sindy 2011/6/7 法務案件屬性
         'Modified by Lydia 2019/08/06 +ACS
         'If Not IsNull(rsA.Fields("CRL84")) Then Frame19.Visible = True: Text1(127) = rsA.Fields("CRL84")
         If Not IsNull(rsA.Fields("CRL84")) Then
'             If Text1(6) = "ACS" Then
'                 Frame44.Visible = True
'                 Me.Combo7.Text = "" & rsA.Fields("CRL84")
'             Else
                 Frame19.Visible = True
                 Text1(127) = "" & rsA.Fields("CRL84")
'             End If
         End If
         
         'Add By Sindy 2012/3/6
         '超頁超項備註
         If Not IsNull(rsA.Fields("CRL85")) Then
            Frame16.Visible = True
            If "" & rsA.Fields("CRL85") = "1" Then
               Option3(0).Value = True
            ElseIf "" & rsA.Fields("CRL85") = "2" Then
               Option3(1).Value = True
            End If
         End If
         '加註查名結果
         If Not IsNull(rsA.Fields("CRL86")) Then
            Frame20.Visible = True
            If "" & rsA.Fields("CRL86") = "1" Then
               Option6(0).Value = True
            ElseIf "" & rsA.Fields("CRL86") = "2" Then
               Option6(1).Value = True
            ElseIf "" & rsA.Fields("CRL86") = "3" Then
               Option6(2).Value = True
            ElseIf "" & rsA.Fields("CRL86") = "4" Then
               Option6(3).Value = True
            'Added by Lydia 2016/09/22
            ElseIf "" & rsA.Fields("CRL86") = "5" Then
               Option6(4).Value = True
            'Added by Lydia 2018/11/13
            ElseIf "" & rsA.Fields("CRL86") = "6" Then
               Option6(5).Value = True
            End If
            'Modify by Sindy 2022/9/16 + 相同文字／圖形，申請人已註冊之審定號
            If Not IsNull(rsA.Fields("CRL136")) Then Text1(133) = rsA.Fields("CRL136")
            '2022/9/16 END
         End If
         '2012/3/6 End
         
         'Add By Sindy 2012/5/8
         '資料是否齊備
         If Not IsNull(rsA.Fields("CRL88")) Then
            Frame21.Visible = True
            Frame22.Visible = True
            If rsA.Fields("CRL88") = "Y" Then
               OptEP06(0).Value = True
            ElseIf rsA.Fields("CRL88") = "N" Then
               OptEP06(1).Value = True
            End If
         Else
            OptEP06(0).Value = False
            OptEP06(1).Value = False
         End If
         '是否會稿
         If Not IsNull(rsA.Fields("CRL89")) Then
            Frame21.Visible = True
            Frame23.Visible = True
            If rsA.Fields("CRL89") = "Y" Then
               OptEP34(0).Value = True
            ElseIf rsA.Fields("CRL89") = "N" Then
               OptEP34(1).Value = True
            End If
         Else
            OptEP34(0).Value = False
            OptEP34(1).Value = False
         End If
         '是否急件
         'Modify By Sindy 2022/11/8
'         If Not IsNull(rsA.Fields("CRL90")) Then
'            Frame21.Visible = True
'            Frame24.Visible = True
'            If rsA.Fields("CRL90") = "Y" Then
'               OptCP122(0).Value = True
'            ElseIf rsA.Fields("CRL90") = "N" Then
'               OptCP122(1).Value = True
'            End If
'         Else
'            OptCP122(0).Value = False
'            OptCP122(1).Value = False
'         End If
         '急件
         If "" & rsA.Fields("CRL90") = "Y" Then
            Check11.Value = 1
         Else
            Check11.Value = 0
         End If
         '費用已核准
         If "" & rsA.Fields("CRL147") = "Y" Then
            Check12.Value = 1
         Else
            Check12.Value = 0
         End If
         '同時申請三國(含)以上之美日德
         chkEnglish.Visible = False: chkEnglish.Value = 0
         If "" & rsA.Fields("CRL148") <> "" Then
            chkEnglish.Visible = True
            If "" & rsA.Fields("CRL148") = "Y" Then chkEnglish.Value = 1
         End If
         '2012/5/8 End
         'Add by Amy 2016/06/06 +可否延期
         If Not IsNull(rsA.Fields("CRL133")) Then
            Frame21.Visible = True
            Frame43.Visible = True
            If rsA.Fields("CRL133") = "Y" Then
               OptCRL133(0).Value = True
            ElseIf rsA.Fields("CRL133") = "N" Then
               OptCRL133(1).Value = True
            End If
         Else
            OptCRL133(0).Value = False
            OptCRL133(1).Value = False
         End If
         'end 2016/06/06
         'Added by Lydia 2018/12/10 查名是否齊備
         If Not IsNull(rsA.Fields("CRL137")) Then
            Frame21.Visible = True
            Frame42.Visible = True
            If rsA.Fields("CRL137") = "Y" Then
               OptCP143(0).Value = True
            ElseIf rsA.Fields("CRL137") = "N" Then
               OptCP143(1).Value = True
            End If
         End If
         'end 2018/12/10
         
         'Add By Sindy 2012/8/17 優惠期日
         If Not IsNull(rsA.Fields("CRL91")) Then
            Frame25.Visible = True
            Text1(128) = ChangeWStringToTString(rsA.Fields("CRL91"))
         Else
            Frame25.Visible = False
         End If
         
         'Add By Sindy 2012/11/12 收據自動列印時間點
         If "" & rsA.Fields("CRL92") = "1" Then
            Check8(0).Value = 1
         ElseIf "" & rsA.Fields("CRL92") = "2" Then
            Check8(1).Value = 1
         ElseIf "" & rsA.Fields("CRL92") = "3" Then
            Check8(2).Value = 1
         End If
         '2012/11/12 End
         
         'Add By Sindy 2013/2/25
         '是否在該國有無近似案件
         If Not IsNull(rsA.Fields("CRL93")) Then
            Frame27.Visible = True
            If "" & rsA.Fields("CRL93") = "Y" Then
               Option8(0).Value = True
            ElseIf "" & rsA.Fields("CRL93") = "N" Then
               Option8(1).Value = True
            End If
         End If
         '2013/2/25 End
         
         'Added by Morgan 2013/4/9
         If Not IsNull(rsA.Fields("CRL94")) Then
            Select Case rsA.Fields("CRL94")
            Case "1"
               OptEntity(0).Value = True
            Case "2"
               OptEntity(1).Value = True
            Case "3"
               OptEntity(2).Value = True
            End Select
         End If
         'end 2013/4/9
         
         'Added by Morgan 2013/9/23
         If rsA.Fields("CRL95") = "Y" Then
            Check4.Visible = True 'Add By Sindy 2023/2/15
            Check4.Value = 1
         Else
            Check4.Visible = False 'Add By Sindy 2023/2/15
            Check4.Value = 0
         End If
         'end 2013/9/23
         'Added by Morgan 2013/9/23
         If rsA.Fields("CRL96") = "Y" Then
            Check5.Value = 1
            Check5.Visible = True 'Add Sindy 2023/8/16
         Else
            Check5.Value = 0
            Check5.Visible = False 'Add Sindy 2023/8/16
         End If
         'end 2013/9/23
         
         'Add By Sindy 2014/2/6 特殊收據
         If "" & rsA.Fields("CRL119") = "Y" Then
'            '外部呼叫且不列印特殊收據頁
'            If m_blnCallPrint = True And m_blnCallPrint_CRL119 = False Then
'               Check9.Value = 0
'               cmdCRL119.Visible = False
'            Else
               Check9.Value = 1
               'cmdCRL119.Visible = True
'            End If
            m_stCRL01 = "" & rsA.Fields("CRL01")
            m_stCRL97 = "" & rsA.Fields("CRL97")
            m_stCRL98 = "" & rsA.Fields("CRL98")
            m_stCRL99 = "" & rsA.Fields("CRL99")
            m_stCRL100 = "" & rsA.Fields("CRL100")
            m_stCRL101 = "" & rsA.Fields("CRL101")
            m_stCRL102 = "" & rsA.Fields("CRL102")
            m_stCRL103 = "" & rsA.Fields("CRL103")
            m_stCRL104 = "" & rsA.Fields("CRL104")
            m_stCRL105 = "" & rsA.Fields("CRL105")
            m_stCRL106 = "" & rsA.Fields("CRL106")
            m_stCRL107 = "" & rsA.Fields("CRL107")
            m_stCRL108 = "" & rsA.Fields("CRL108")
            m_stCRL109 = "" & rsA.Fields("CRL109")
            m_stCRL110 = "" & rsA.Fields("CRL110")
            m_stCRL111 = "" & rsA.Fields("CRL111")
            m_stCRL112 = "" & rsA.Fields("CRL112")
            m_stCRL113 = "" & rsA.Fields("CRL113")
            m_stCRL114 = "" & rsA.Fields("CRL114")
            m_stCRL115 = "" & rsA.Fields("CRL115")
            m_stCRL116 = "" & rsA.Fields("CRL116")
            m_stCRL117 = "" & rsA.Fields("CRL117")
            m_stCRL118 = "" & rsA.Fields("CRL118")
            m_stCRL120 = "" & rsA.Fields("CRL120")
            m_stCRL121 = "" & rsA.Fields("CRL121")
            m_stCRL122 = "" & rsA.Fields("CRL122")
            m_stCRL123 = "" & rsA.Fields("CRL123")
            m_stCRL124 = "" & rsA.Fields("CRL124")
            'Add By Sindy 2015/8/28
            m_stCRL126 = "" & rsA.Fields("CRL126")
            m_stCRL127 = "" & rsA.Fields("CRL127")
            m_stCRL128 = "" & rsA.Fields("CRL128")
            m_stCRL129 = "" & rsA.Fields("CRL129")
            m_stCRL130 = "" & rsA.Fields("CRL130")
            m_stCRL131 = "" & rsA.Fields("CRL131")
            m_stCRL132 = "" & rsA.Fields("CRL132")
            '2015/8/28 END
         Else
            Check9.Value = 0
            'cmdCRL119.Visible = False
         End If
         '2014/2/6 END
         
'Add By Sindy 2024/5/9 此2個常變數存入資料庫內, 因操作子畫面組出來的
'                      記錄下來, 以利人員新增後再回頭修改時才能保留住程式原狀況, 部分程式判斷上才不會出錯
         If Trim("" & rsA.Fields("CRL161")) <> "" Then
            m_Note1 = Trim("" & rsA.Fields("CRL161"))
         End If
         If Trim("" & rsA.Fields("CRL162")) <> "" Then
            m_Note2 = Trim("" & rsA.Fields("CRL162"))
         End If
'2024/5/9 End
         
         'Add By Sindy 2023/9/27 記錄延期案的下一程序文號
         'Modify By Sindy 2025/8/5 + Or InStr(m_strCaseCPM & ",", "727,") > 0)
         If Left(Text1(6), 1) = "T" And _
            (InStr(m_strCaseCPM & ",", "303,") > 0 Or InStr(m_strCaseCPM & ",", "727,") > 0) Then
            m_strGetNP01 = "" & rsA.Fields("CRL71")
            If "" & rsA.Fields("CRL70") <> "" And InStr("" & rsA.Fields("CRL70"), "延期案件性質") > 0 Then
               strExc(9) = InStr(rsA.Fields("CRL70"), "延期案件性質")
               strExc(10) = InStr(Mid(rsA.Fields("CRL70"), strExc(9)), vbCrLf)
               'm_Note2 = Mid(rsA.Fields("CRL70"), strExc(9), strExc(10) - 1)
               Frame605.Visible = True
               labelYF.Caption = m_Note2
               Text1(142).Visible = False
               Text1(143).Visible = False
            End If
            '2023/9/27 END
         'Add By Sindy 2014/2/19 審定號
         'Modify By Sindy 2022/9/20 + 申請案號
         ElseIf Trim("" & rsA.Fields("CRL125")) <> "" Or Trim("" & rsA.Fields("CRL71")) <> "" Then
            Frame29.Visible = True
            Text1(129).Text = "" & rsA.Fields("CRL125")
            Text1(132).Text = "" & rsA.Fields("CRL71")
         Else
            Frame29.Visible = False
            Text1(129).Text = ""
            Text1(132).Text = ""
         End If
         '2014/2/19 END
         
         'Added by Lydia 2021/02/24
         '補上：特定申請人會稿案件CRL138
         If "" & rsA.Fields("CRL138") <> "" Then
             Frame45.Visible = True
             If "" & rsA.Fields("CRL138") = "Y" Then
                 Opt45(0).Value = True
             Else
                 Opt45(1).Value = True
             End If
         End If
         '大陸商標申請案及CFT申請案，均強迫填入相關資訊。
         If "" & rsA.Fields("CRL139") <> "" Then
             Frame47.Visible = True
             Text7.Text = "" & rsA.Fields("CRL139")
         End If
         'end 2021/02/24
         
         'Added by Morgan 2021/7/20 大陸發明生醫案是否新藥專利設定
         If "" & rsA.Fields("CRL140") <> "" Then
            Frame48.Visible = True
            If "" & rsA.Fields("CRL140") = "Y" Then
               OptNewDrug(1).Value = True
            ElseIf "" & rsA.Fields("CRL140") = "N" Then
               OptNewDrug(0).Value = True
            End If
         End If
         'end 2021/7/20
         
         'Add By Sindy 2022/9/20
         SSTab3.TabVisible(1) = False
'         If Text1(6) = "P" And Left(Combo1(0).Text, 3) = "000" Then
'            SSTab3.TabVisible(1) = True
'         Else
'            SSTab3.TabVisible(1) = False
'         End If
         '聘任期間
         If (Text1(6) = "LA" Or Text1(6) = "ACS") And Not IsNull(rsA.Fields("CRL144")) Then
            Label1(140).Visible = True: Text1(139).Visible = True: Text1(140).Visible = True
            Text1(139) = rsA.Fields("CRL144") - 19110000
            Text1(140) = rsA.Fields("CRL145") - 19110000
         '年費期間
         ElseIf ((Text1(6) = "P" Or Text1(6) = "CFP") And _
                 (InStr(m_strCaseCPM, "605") > 0 Or InStr(m_strCaseCPM, "606") > 0 Or InStr(m_strCaseCPM, "607") > 0)) Or _
            (Text1(6) = "P" And InStr(m_strCaseCPM, "601") > 0) Or _
            (Text1(6) = "CFP" And InStr(m_strCaseCPM, "613") > 0) Then
            Frame605.Visible = True
            If InStr(m_strCaseCPM, "605") > 0 Or (Text1(6) = "P" And InStr(m_strCaseCPM, "601") > 0) Then
               labelYF.Caption = "繳費年度：         ~"
            Else
               labelYF.Caption = "繳費次數：         ~"
            End If
            If Not IsNull(rsA.Fields("CRL144")) Then
               Text1(142) = rsA.Fields("CRL144")
               Text1(143) = rsA.Fields("CRL145")
            End If
         '舉發聲明
         ElseIf Not IsNull(rsA.Fields("CRL141")) Or _
            Not IsNull(rsA.Fields("CRL143")) Or _
            Not IsNull(rsA.Fields("CRL144")) Then
            SSTab3.TabVisible(1) = True
            
            If "" & rsA.Fields("CRL141") = "1" Then
               chkItem(0).Value = vbChecked
               txtItemCount = "" & rsA.Fields("CRL142")
            ElseIf "" & rsA.Fields("CRL141") = "2" Then
               chkItem(1).Value = vbChecked
               txtItemList = "" & rsA.Fields("CRL142")
            End If
            If Val("" & rsA.Fields("CRL144")) > 0 Then
               txtYear(0) = Left(rsA.Fields("CRL144"), 4) - 1911
               txtMonth(0) = Mid(rsA.Fields("CRL144"), 5, 2)
               txtDay(0) = Mid(rsA.Fields("CRL144"), 7, 2)
               If Val("" & rsA.Fields("CRL145")) > 0 Then
                  txtYear(1) = Left(rsA.Fields("CRL145"), 4) - 1911
                  txtMonth(1) = Mid(rsA.Fields("CRL145"), 5, 2)
                  txtDay(1) = Mid(rsA.Fields("CRL145"), 7, 2)
               End If
            End If
            If "" & rsA.Fields("CRL143") <> "" Then
               If InStr(rsA.Fields("CRL143"), "1,") > 0 Then chkItem(2).Value = vbChecked
               If InStr(rsA.Fields("CRL143"), "2,") > 0 Then chkItem(3).Value = vbChecked
               If InStr(rsA.Fields("CRL143"), "3,") > 0 Then chkItem(4).Value = vbChecked
               If InStr(rsA.Fields("CRL143"), "4,") > 0 Then chkItem(5).Value = vbChecked
            End If
         End If
         '2022/9/20 END
         
'         rsA.MoveNext
'      Loop
   End If
   rsA.Close
   
   'Add By Sindy 2022/9/9
   '*****************
   '電子檔: 不抓已收文的卷宗區附件
   '*****************
   Me.m_strSaveFiles = "": Me.m_strSaveFiles2 = ""
   Me.cmdAddAtt.BackColor = &H808080 '灰色
   strExc(0) = "Select * " & _
                 "From CasePaperPDF " & _
               "Where cpp11 ='" & Trim(Text5) & "' and cpp10<>'D' " & _
               "and substr(upper(cpp02),-4)<>upper('.del') and substr(upper(cpp02),-6)<>upper('.order') " & _
               "order by cpp02 asc "
   intI = 1
   Set rsA = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      Me.cmdAddAtt.BackColor = &HC0C0FF '粉紅色
   End If
   If LblRecved.Visible = True Then '已收文
      Me.cmdAddAtt.Visible = False
   End If
   rsA.Close
   '(有文件)
   strExc(0) = "Select * " & _
                 "From CasePaperPDF " & _
               "Where cpp11 ='" & Trim(Text5) & "' and length(cpp01)=9 and cpp10<>'D' " & _
               "and substr(upper(cpp02),-4)<>upper('.del') and substr(upper(cpp02),-6)<>upper('.order') " & _
               "order by cpp02 asc "
   intI = 1
   Set rsA = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      Label38.Visible = True
   Else
      Label38.Visible = False
   End If
   rsA.Close
   '2022/9/9 END
   '*****************
   '申請人
   '*****************
   SSTab2.TabVisible(0) = False
   SSTab2.TabVisible(1) = False
   SSTab2.TabVisible(2) = False
   SSTab2.TabVisible(3) = False
   SSTab2.TabVisible(4) = False
   strExc(0) = "Select * " & _
                 "From consultrecapp " & _
               "Where cra01 ='" & Trim(Text5) & "' " & _
               "order by cra02 asc "
   intI = 1
   Set rsA = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      rsA.MoveFirst
      Do While Not rsA.EOF
         '申請人1
         If rsA.Fields("CRA02") = "1" Then
            SSTab2.TabVisible(0) = True
            Me.SSTab2.Tab = 0
            If "" & rsA.Fields("CRA03") = "Y" Then
               Option31(0).Value = True
               '有客戶編號或特例不可以再異動
               If Not IsNull(rsA.Fields("CRA05")) Or "" & rsA.Fields("CRA26") = "Y" Then
                  Option31(0).Enabled = False
                  Option31(1).Enabled = False
               End If
            Else
               Option31(1).Value = True
            End If
         '申請人2
         ElseIf rsA.Fields("CRA02") = "2" Then
            SSTab2.TabVisible(1) = True
            Me.SSTab2.Tab = 1
            If "" & rsA.Fields("CRA03") = "Y" Then
               Option32(0).Value = True
               '有客戶編號或特例不可以再異動
               If Not IsNull(rsA.Fields("CRA05")) Or "" & rsA.Fields("CRA26") = "Y" Then
                  Option32(0).Enabled = False
                  Option32(1).Enabled = False
               End If
            Else
               Option32(1).Value = True
            End If
         '申請人3
         ElseIf rsA.Fields("CRA02") = "3" Then
            SSTab2.TabVisible(2) = True
            Me.SSTab2.Tab = 2
            If "" & rsA.Fields("CRA03") = "Y" Then
               Option33(0).Value = True
               '有客戶編號或特例不可以再異動
               If Not IsNull(rsA.Fields("CRA05")) Or "" & rsA.Fields("CRA26") = "Y" Then
                  Option33(0).Enabled = False
                  Option33(1).Enabled = False
               End If
            Else
               Option33(1).Value = True
            End If
         '申請人4
         ElseIf rsA.Fields("CRA02") = "4" Then
            SSTab2.TabVisible(3) = True
            Me.SSTab2.Tab = 3
            If "" & rsA.Fields("CRA03") = "Y" Then
               Option34(0).Value = True
               '有客戶編號或特例不可以再異動
               If Not IsNull(rsA.Fields("CRA05")) Or "" & rsA.Fields("CRA26") = "Y" Then
                  Option34(0).Enabled = False
                  Option34(1).Enabled = False
               End If
            Else
               Option34(1).Value = True
            End If
         '申請人5
         ElseIf rsA.Fields("CRA02") = "5" Then
            SSTab2.TabVisible(4) = True
            Me.SSTab2.Tab = 4
            If "" & rsA.Fields("CRA03") = "Y" Then
               Option35(0).Value = True
               '有客戶編號或特例不可以再異動
               If Not IsNull(rsA.Fields("CRA05")) Or "" & rsA.Fields("CRA26") = "Y" Then
                  Option35(0).Enabled = False
                  Option35(1).Enabled = False
               End If
            Else
               Option35(1).Value = True
            End If
         End If
         intIndex = rsA.Fields("CRA02")
         If Not IsNull(rsA.Fields("CRA04")) Then
            Text1(IIf(intIndex = 1, 198, IIf(intIndex = 2, 298, IIf(intIndex = 3, 398, IIf(intIndex = 4, 498, 598))))) = rsA.Fields("CRA04")
            Text1(IIf(intIndex = 1, 198, IIf(intIndex = 2, 298, IIf(intIndex = 3, 398, IIf(intIndex = 4, 498, 598))))).BackColor = &HC0FFC0 '底色變淺綠色
         End If
         If Not IsNull(rsA.Fields("CRA05")) Then
            Text1(IIf(intIndex = 1, 12, IIf(intIndex = 2, 28, IIf(intIndex = 3, 44, IIf(intIndex = 4, 60, 76))))) = rsA.Fields("CRA05") & rsA.Fields("CRA06")
            'Modify By Sindy 2022/10/12 新客戶欄位鎖住
            If "" & rsA.Fields("CRA03") = "Y" Then
               Text1(IIf(intIndex = 1, 12, IIf(intIndex = 2, 28, IIf(intIndex = 3, 44, IIf(intIndex = 4, 60, 76))))).Enabled = False
            End If
            '2022/10/12 END
         End If
         If m_blnCallPrint = False Then
            '舊客戶需 Run SetCustTxt 不然會易出現郵遞區號問題
            If Trim("" & rsA.Fields("CRA05") & rsA.Fields("CRA06")) <> "" And "" & rsA.Fields("CRA03") <> "Y" Then
               SetCustTxt Val(rsA.Fields("CRA02")), rsA.Fields("CRA05") & rsA.Fields("CRA06")
               PUB_AddContact rsA.Fields("CRA05") & rsA.Fields("CRA06"), cboContact(Val(rsA.Fields("CRA02"))), , True, True, m_strContactList(Val(rsA.Fields("CRA02")))
            End If
         End If
         If Not IsNull(rsA.Fields("CRA07")) Then '中文客戶名稱
            Text1(IIf(intIndex = 1, 21, IIf(intIndex = 2, 37, IIf(intIndex = 3, 53, IIf(intIndex = 4, 69, 85))))) = rsA.Fields("CRA07")
            'Add By Sindy 2022/10/12 + 記錄Tag,按下存檔檢查資料會用到
            Text1(IIf(intIndex = 1, 21, IIf(intIndex = 2, 37, IIf(intIndex = 3, 53, IIf(intIndex = 4, 69, 85))))).Tag = rsA.Fields("CRA07")
         End If
         If Not IsNull(rsA.Fields("CRA08")) Then '英文客戶名稱
            Text1(IIf(intIndex = 1, 22, IIf(intIndex = 2, 38, IIf(intIndex = 3, 54, IIf(intIndex = 4, 70, 86))))) = rsA.Fields("CRA08")
            'Add By Sindy 2022/10/12 + 記錄Tag,按下存檔檢查資料會用到
            Text1(IIf(intIndex = 1, 22, IIf(intIndex = 2, 38, IIf(intIndex = 3, 54, IIf(intIndex = 4, 70, 86))))).Tag = rsA.Fields("CRA08")
         End If
         If Not IsNull(rsA.Fields("CRA09")) Then Text1(IIf(intIndex = 1, 23, IIf(intIndex = 2, 39, IIf(intIndex = 3, 55, IIf(intIndex = 4, 71, 87))))) = rsA.Fields("CRA09")
         If Not IsNull(rsA.Fields("CRA10")) Then cboContact(intIndex) = rsA.Fields("CRA10")
         If Not IsNull(rsA.Fields("CRA11")) Then Text1(IIf(intIndex = 1, 34, IIf(intIndex = 2, 50, IIf(intIndex = 3, 66, IIf(intIndex = 4, 82, 98))))) = rsA.Fields("CRA11")
         If Not IsNull(rsA.Fields("CRA12")) Then Text1(IIf(intIndex = 1, 92, IIf(intIndex = 2, 93, IIf(intIndex = 3, 94, IIf(intIndex = 4, 95, 96))))) = rsA.Fields("CRA12")
         If Not IsNull(rsA.Fields("CRA13")) Then Text1(IIf(intIndex = 1, 14, IIf(intIndex = 2, 30, IIf(intIndex = 3, 46, IIf(intIndex = 4, 62, 78))))) = rsA.Fields("CRA13")
         If Not IsNull(rsA.Fields("CRA14")) Then Text1(IIf(intIndex = 1, 15, IIf(intIndex = 2, 31, IIf(intIndex = 3, 47, IIf(intIndex = 4, 63, 79))))) = rsA.Fields("CRA14")
         If Not IsNull(rsA.Fields("CRA15")) Then Text1(IIf(intIndex = 1, 16, IIf(intIndex = 2, 32, IIf(intIndex = 3, 48, IIf(intIndex = 4, 64, 80))))) = rsA.Fields("CRA15")
         If Not IsNull(rsA.Fields("CRA16")) Then Text1(IIf(intIndex = 1, 17, IIf(intIndex = 2, 33, IIf(intIndex = 3, 49, IIf(intIndex = 4, 65, 81))))) = rsA.Fields("CRA16")
         If Not IsNull(rsA.Fields("CRA17")) Then Text1(IIf(intIndex = 1, 19, IIf(intIndex = 2, 35, IIf(intIndex = 3, 51, IIf(intIndex = 4, 67, 83))))) = rsA.Fields("CRA17")
         If Not IsNull(rsA.Fields("CRA18")) Then Text1(IIf(intIndex = 1, 20, IIf(intIndex = 2, 36, IIf(intIndex = 3, 52, IIf(intIndex = 4, 68, 84))))) = rsA.Fields("CRA18")
         If Not IsNull(rsA.Fields("CRA19")) Then '中文地址郵遞區號
            Text1(IIf(intIndex = 1, 120, IIf(intIndex = 2, 121, IIf(intIndex = 3, 122, IIf(intIndex = 4, 123, 124))))) = rsA.Fields("CRA19")
            'Modify By Sindy 2022/10/12 舊客戶欄位鎖住
            If "" & rsA.Fields("CRA03") = "" Then
               Text1(IIf(intIndex = 1, 120, IIf(intIndex = 2, 121, IIf(intIndex = 3, 122, IIf(intIndex = 4, 123, 124))))).Enabled = False
            End If
            '2022/10/12 END
         End If
         If Not IsNull(rsA.Fields("CRA20")) Then Text1(IIf(intIndex = 1, 27, IIf(intIndex = 2, 43, IIf(intIndex = 3, 59, IIf(intIndex = 4, 75, 91))))) = rsA.Fields("CRA20")
         If Not IsNull(rsA.Fields("CRA21")) Then Text1(IIf(intIndex = 1, 125, IIf(intIndex = 2, 141, IIf(intIndex = 3, 157, IIf(intIndex = 4, 173, 189))))) = rsA.Fields("CRA21")
         If Not IsNull(rsA.Fields("CRA22")) Then '聯絡地址郵遞區號
            Text1(IIf(intIndex = 1, 25, IIf(intIndex = 2, 41, IIf(intIndex = 3, 57, IIf(intIndex = 4, 73, 89))))) = rsA.Fields("CRA22")
            'Modify By Sindy 2022/10/12 舊客戶欄位鎖住
            If "" & rsA.Fields("CRA03") = "" Then
               Text1(IIf(intIndex = 1, 25, IIf(intIndex = 2, 41, IIf(intIndex = 3, 57, IIf(intIndex = 4, 73, 89))))).Enabled = False
            End If
            '2022/10/12 END
         End If
         If Not IsNull(rsA.Fields("CRA23")) Then Text1(IIf(intIndex = 1, 26, IIf(intIndex = 2, 42, IIf(intIndex = 3, 58, IIf(intIndex = 4, 74, 90))))) = rsA.Fields("CRA23")
         If Not IsNull(rsA.Fields("CRA24")) Then Text1(IIf(intIndex = 1, 24, IIf(intIndex = 2, 40, IIf(intIndex = 3, 56, IIf(intIndex = 4, 72, 88))))) = rsA.Fields("CRA24")
         
         'Add By Sindy 2022/9/15 有對造
         If "" & rsA.Fields("CRA26") = "Y" Then
            ChkCRA26(Val(rsA.Fields("CRA02")) - 1).Value = 1
            lblAPPLQ.Visible = True
            'ChkCRA26(Val(rsA.Fields("CRA02")) - 1).Enabled = False
         End If
         '2022/9/15 END
         'Add By Sindy 2022/11/8 有跨所
         If "" & rsA.Fields("CRA27") = "Y" Then
            ChkCRA27(Val(rsA.Fields("CRA02")) - 1).Value = 1
            lblZip.Visible = True
            'ChkCRA27(Val(rsA.Fields("CRA02")) - 1).Enabled = False
         End If
         '2022/11/8 END
         
         '申請人1
         If rsA.Fields("CRA02") = "1" Then
            If "" & rsA.Fields("CRA25") = "Y" Then
               optCP811(0).Value = True
            ElseIf "" & rsA.Fields("CRA25") = "N" Then
               optCP811(1).Value = True
            End If
         '申請人2
         ElseIf rsA.Fields("CRA02") = "2" Then
            If "" & rsA.Fields("CRA25") = "Y" Then
               optCP812(0).Value = True
            ElseIf "" & rsA.Fields("CRA25") = "N" Then
               optCP812(1).Value = True
            End If
         '申請人3
         ElseIf rsA.Fields("CRA02") = "3" Then
            If "" & rsA.Fields("CRA25") = "Y" Then
               optCP813(0).Value = True
            ElseIf "" & rsA.Fields("CRA25") = "N" Then
               optCP813(1).Value = True
            End If
         '申請人4
         ElseIf rsA.Fields("CRA02") = "4" Then
            If "" & rsA.Fields("CRA25") = "Y" Then
               optCP814(0).Value = True
            ElseIf "" & rsA.Fields("CRA25") = "N" Then
               optCP814(1).Value = True
            End If
         '申請人5
         ElseIf rsA.Fields("CRA02") = "5" Then
            If "" & rsA.Fields("CRA25") = "Y" Then
               optCP815(0).Value = True
            ElseIf "" & rsA.Fields("CRA25") = "N" Then
               optCP815(1).Value = True
            End If
         End If
         rsA.MoveNext
      Loop
   End If
   rsA.Close
   '*****************
   '發明人
   '*****************
   SSTab1.TabVisible(3) = False
   strExc(0) = "Select * " & _
                 "From consultrecinv " & _
               "Where cri01 ='" & Trim(Text5) & "' " & _
               "order by cri02 asc "
   intI = 1
   Set rsA = ClsLawReadRstMsg(intI, strExc(0))
   strInventorNo = "": strInventorName = "" 'Add By Sindy 2011/1/31 +strInventorName
   If intI = 1 Then
      SSTab1.TabVisible(3) = True
      rsA.MoveFirst
      Do While Not rsA.EOF
         intIndex = rsA.Fields("CRI02") - 1
         If Not IsNull(rsA.Fields("CRI03")) Then
            If strInventorNo <> "" Then strInventorNo = strInventorNo & ","
            strInventorNo = strInventorNo & rsA.Fields("CRI03") & rsA.Fields("CRI04")
            'Add By Sindy 2011/1/31 +strInventorName
            If strInventorName <> "" Then strInventorName = strInventorName & ","
            strInventorName = strInventorName & rsA.Fields("CRI06")
            '2011/1/31 End
            Label5(intIndex).Tag = rsA.Fields("CRI03") & rsA.Fields("CRI04") 'Add By Sindy 2023/1/18
         End If
         If Not IsNull(rsA.Fields("CRI05")) Then Text3(intIndex) = rsA.Fields("CRI05")
         If Not IsNull(rsA.Fields("CRI06")) Then Text2(intIndex) = rsA.Fields("CRI06")
         If Not IsNull(rsA.Fields("CRI07")) Then Text4(intIndex) = rsA.Fields("CRI07")
         If Not IsNull(rsA.Fields("CRI08")) Then Combo3(intIndex) = rsA.Fields("CRI08")
         Call Combo3_Validate(intIndex, Cancel)
         If "" & rsA.Fields("CRI09") = "Y" Then ChkAddress(intIndex).Value = 1
         rsA.MoveNext
      Loop
   End If
   Call SetColColor 'Add By Sindy 2023/1/18
   rsA.Close
   '*****************
   '商標圖檔
   '*****************
   strExc(0) = "Select * " & _
                 "From consultrecimagef " & _
               "Where crif01 ='" & Trim(Text5) & "' "
   intI = 1
   Set rsA = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      '顯示圖檔
      rsA.MoveFirst
      If CheckStr(rsA.Fields("crif02")) = "1" Or CheckStr(rsA.Fields("crif02")) = "3" Then optColor(0).Value = True Else optColor(1).Value = True
      '加入無圖式的格式
      If CheckStr(rsA.Fields("crif02")) = "3" Or CheckStr(rsA.Fields("crif02")) = "4" Or CheckStr(rsA.Fields("crif02")) = "6" Then IsWmf = True Else IsWmf = False
      If IsWmf = False Then
         stAttPath = App.path & "\NowPic.jpg"
      Else
         stAttPath = App.path & "\NowPic.wmf"
      End If
      
      'Add By Sindy 2017/6/1
      If "" & rsA.Fields("crif05") <> "" Then
         PUB_GetFtpFile rsA.Fields("crif05"), stAttPath, UCase("consultrecimagef")
      Else
      '2017/6/1 END
         ReDim bytes(Val(rsA.Fields("crif03").Value))
         bytes() = rsA.Fields("crif04").GetChunk(Val(rsA.Fields("crif03").Value))
         file_num = FreeFile
'         If IsWmf = False Then
'             Open App.path & "\NowPic.jpg" For Binary Access Write As #file_num
'         Else
'             Open App.path & "\NowPic.wmf" For Binary Access Write As #file_num
'         End If
         Open stAttPath For Binary Access Write As #file_num
         Put #file_num, , bytes()
         Close #file_num
      End If
'      If IsWmf = False Then
'          PicToObj Trim(App.path & "\NowPic.jpg")
'      Else
'          PicToObj Trim(App.path & "\NowPic.wmf")
'      End If
      PicToObj Trim(stAttPath)
'      If Dir(App.path & "\NowPic.jpg") <> "" Then
'          Kill App.path & "\NowPic.jpg"
'      End If
'      If Dir(App.path & "\NowPic.wmf") <> "" Then
'          Kill App.path & "\NowPic.wmf"
'      End If
      If Dir(stAttPath) <> "" Then
         cmdSavePic.Visible = True 'Added by Morgan 2023/7/6
         Kill stAttPath
      End If
      If Dir(App.path & "\tmp.tif") <> "" Then
          Kill App.path & "\tmp.tif"
      End If
   End If
   rsA.Close
   If SSTab2.TabVisible(0) = True Then Me.SSTab2.Tab = 0
   Me.SSTab1.Tab = 0
   
   Call QueryData_Flow '查詢簽核狀況 Add By Sindy 2022/9/19
   
   Set rsA = Nothing
   Set rsD = Nothing
End Function

'Add By Sindy 2023/1/18
Private Sub SetColColor()
Dim i As Integer
   
   For i = 0 To 9
      If Label5(i).Tag = "" And Trim(Text2(i).Text) <> "" Then
         Label5(i).BackColor = &H80FFFF   '黃色
      Else
         Label5(i).BackColor = &H8000000F '底色:按鈕表面
      End If
   Next i
End Sub

'Add By Sindy 2010/6/21
Private Sub GetMainData()
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
   
   m_strPA08 = ""
   m_strPA16 = ""
   m_strPA10 = ""
   m_strTM15 = ""
   m_strTM12 = ""
   m_strPA14 = ""   '2010/8/17 ADD BY SONIA
   
   '2008/3/21 modify by sonia 商標案加抓tm15,tm12
   'Modify by Morgan 2008/8/1 加抓個案申請人聯絡人編號
   'Modify by Sindy 2009/05/21 加抓商標檔申請人2,3,4,5及服務檔申請人4,5
   '2010/3/31 MODIFY BY SONIA 專利案加申請日
   '2010/8/17 MODIFY BY SONIA 專利案加公告日pa14
   'Modify By Sindy 2011/2/21 增加LC43,LC44,LC45,LC46,HC24,HC25,HC26,HC27
   StrSQLa = "Select NA01||' '||NA03, Nvl(PA05, Nvl(PA06, PA07)), PA26, PA27, PA28, PA29, PA30, PA08, PA16, PA48, '', '',PA149,PA10,PA75,PA14 From Patent, Nation,Customer Where PA09=NA01(+) And PA01='" & Me.Text1(6).Text & "' And PA02='" & Me.Text1(7).Text & "' And PA03='" & Me.Text1(8).Text & "' And PA04='" & Me.Text1(9).Text & "' and cu01(+)=substr(pa26,1,8) and cu02(+)=substr(pa26,9,1) "
   StrSQLa = StrSQLa & " Union Select NA01||' '||NA03, Nvl(TM05, Nvl(TM06, TM07)), TM23, TM78, TM79, TM80, TM81, '', '', TM35,TM15,TM12,TM123,0,TM44,0 From Trademark, Nation,Customer Where TM10=NA01(+) And TM01='" & Me.Text1(6).Text & "' And TM02='" & Me.Text1(7).Text & "' And TM03='" & Me.Text1(8).Text & "' And TM04='" & Me.Text1(9).Text & "' and cu01(+)=substr(TM23,1,8) and cu02(+)=substr(TM23,9,1)  "
   StrSQLa = StrSQLa & " Union Select NA01||' '||NA03, Nvl(LC05, Nvl(LC06, LC07)), LC11, LC43,LC44,LC45,LC46, '', '', LC17, '', '',LC42,0,LC22,0 From Lawcase, Nation,Customer Where LC15=NA01(+) And LC01='" & Me.Text1(6).Text & "' And LC02='" & Me.Text1(7).Text & "' And LC03='" & Me.Text1(8).Text & "' And LC04='" & Me.Text1(9).Text & "' and cu01(+)=substr(LC11,1,8) and cu02(+)=substr(LC11,9,1) "
   StrSQLa = StrSQLa & " Union Select NA01||' '||NA03, HC06, HC05, HC24,HC25,HC26,HC27, '', '', '', '', '',HC23,0,'',0 From Hirecase, Nation,Customer Where '000'=NA01(+) And HC01='" & Me.Text1(6).Text & "' And HC02='" & Me.Text1(7).Text & "' And HC03='" & Me.Text1(8).Text & "' And HC04='" & Me.Text1(9).Text & "' and cu01(+)=substr(HC05,1,8) and cu02(+)=substr(HC05,9,1) "
   StrSQLa = StrSQLa & " Union Select NA01||' '||NA03, Nvl(SP05, Nvl(SP06, SP07)), SP08, SP58, SP59, SP65, SP66, '', '', SP29, '', '',SP78,0,SP26,0 From Servicepractice, Nation,Customer Where SP09=NA01(+) And SP01='" & Me.Text1(6).Text & "' And SP02='" & Me.Text1(7).Text & "' And SP03='" & Me.Text1(8).Text & "' And SP04='" & Me.Text1(9).Text & "' and cu01(+)=substr(SP08,1,8) and cu02(+)=substr(SP08,9,1) "
   rsA.CursorLocation = adUseClient
   rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
   If rsA.RecordCount > 0 Then
      m_strPA08 = "" & rsA.Fields(7).Value
      m_strPA16 = "" & rsA.Fields(8).Value
      '2010/8/17 ADD BY SONIA
      m_strPA14 = "" & rsA.Fields(15).Value
      If m_strPA14 = "0" Then m_strPA14 = ""
      '2010/8/17 END
      m_strPA10 = "" & rsA.Fields(13).Value  '2010/3/31 ADD BY SONIA
      '2008/3/21 ADD BY SONIA
      m_strTM15 = "" & rsA.Fields(10).Value
      m_strTM12 = "" & rsA.Fields(11).Value
      'Add By Sindy 2010/7/9
      If m_strPA08 <> "" Then '代表有pa資料
         ReDim pa(1 To TF_PA) As String
         Call PUB_ReadPatentData(pa(), Text1(6), Text1(7), Text1(8), Text1(9))
      End If
   End If
   If rsA.State <> adStateClosed Then rsA.Close
   Set rsA = Nothing
End Sub

Private Sub DelPic()
   Dim strPicFileName As String
   strPicFileName = App.path & "\$tmp_*.tmp"
   If Dir(strPicFileName) <> "" Then
      Kill strPicFileName
   End If
   m_Device.Line (0, 0)-(m_Device.Width, m_Device.Height), QBColor(15), BF
End Sub

Private Sub SetPic(idx As Integer)

   Dim strPicFileName As String
   strPicFileName = App.path & "\$tmp_" & idx & ".tmp"
   
   SavePicture Picture1.Image, strPicFileName
   '要用覆蓋的否則會錯誤--VB Bug
   'Picture1.Cls
   m_Device.Line (0, 0)-(m_Device.Width, m_Device.Height), QBColor(15), BF
   
End Sub

'Add by Morgan 2010/12/10
'檢查案件是否已閉卷
Private Function IfCaseClosed() As Boolean
   If Option1(1).Value = True Then
      strExc(0) = "select pa57 from patent where pa01='" & Text1(6) & "' and pa02='" & Text1(7) & "' and pa03='" & Text1(8) & "' and pa04='" & Text1(9) & "' and pa57 is not null" & _
         " union select tm29 from trademark where tm01='" & Text1(6) & "' and tm02='" & Text1(7) & "' and tm03='" & Text1(8) & "' and tm04='" & Text1(9) & "' and tm29 is not null" & _
         " union select sp15 from servicepractice where sp01='" & Text1(6) & "' and sp02='" & Text1(7) & "' and sp03='" & Text1(8) & "' and sp04='" & Text1(9) & "' and sp15 is not null" & _
         " union select lc08 from lawcase where lc01='" & Text1(6) & "' and lc02='" & Text1(7) & "' and lc03='" & Text1(8) & "' and lc04='" & Text1(9) & "' and lc08 is not null" & _
         " union select hc09 from hirecase where hc01='" & Text1(6) & "' and hc02='" & Text1(7) & "' and hc03='" & Text1(8) & "' and hc04='" & Text1(9) & "' and hc09 is not null"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         IfCaseClosed = True
      End If
   End If
End Function
'Add by Morgan 2010/12/10
'是否有減免退費未辦理
Private Function IfDiscFeeNotPaid(pCaseNo() As String, pCurDiscStat As String) As Boolean
   If pCurDiscStat = "Y" And Left(Combo1(0), 3) = "000" Then
      IfDiscFeeNotPaid = PUB_CheckYearFeeReturn(pCaseNo, False)
   End If
End Function

'Add by Morgan 2011/1/10 從 cmdok_Click 搬來以便共用
'檢查是否特殊客戶
'Modify By Sindy 2015/7/27 +Optional ByVal pIndex As Integer = 0 pIndex:傳入要檢查第x個申請人資料,若0代表全部檢查一次
Private Sub SetCuData(ByRef pIsSpecCu As Boolean, Optional ByRef pSpecCUName As String, Optional ByRef pSpecMemo As String, _
                      Optional ByRef pIsCuMemo As Boolean, Optional ByRef pCuMemoName As String, Optional ByRef pCuMemo As String, _
                      Optional ByVal pIndex As Integer = 0)
   
   pIsSpecCu = False 'Add By Sindy 2012/12/12 預設變數值
   For iiiii = IIf(pIndex = 0, 1, pIndex) To IIf(pIndex = 0, 5, pIndex)
       'edit by nickc 2008/01/18 加入客戶業務備註
       'Modified by Lydia 2022/10/26 固定用CU02=0判斷業務備註; ex.P-124312 目前客戶編號為X70604001
       'strSql = "select cu01,cu02,cu121,nvl(cu04,nvl(cu05||cu88||cu89||cu90,cu06)) cu04,cl04,cu125 from customer, (select * from custspeciallog  where (cl01,cl02,cl03) in (select cl01,cl02,max(cl03) from custspeciallog where cl01='" & Me.Text1(12 + (iiiii - 1) * 16).Text & "' " & _
                      " and cl02 in (select max(cl02) from custspeciallog where cl01='" & Me.Text1(12 + (iiiii - 1) * 16).Text & "' ) group by cl01,cl02)) Ncl " & _
                      " Where CU01='" & Mid(Me.Text1(12 + (iiiii - 1) * 16).Text, 1, 8) & "' And CU02='" & Mid(Me.Text1(12 + (iiiii - 1) * 16).Text, 9, 1) & "' and cu01||cu02=Ncl.cl01(+) "
       strSql = "select cu01,cu02,cu121,nvl(cu04,nvl(cu05||cu88||cu89||cu90,cu06)) cu04,cl04,cu125 from customer, (select * from custspeciallog  where (cl01,cl02,cl03) in (select cl01,cl02,max(cl03) from custspeciallog where cl01='" & Left(Me.Text1(12 + (iiiii - 1) * 16).Text, 8) & "0" & "' " & _
                      " and cl02 in (select max(cl02) from custspeciallog where cl01='" & Left(Me.Text1(12 + (iiiii - 1) * 16).Text, 8) & "0" & "' ) group by cl01,cl02)) Ncl " & _
                      " Where CU01='" & Mid(Me.Text1(12 + (iiiii - 1) * 16).Text, 1, 8) & "' And CU02='0' and cu01||cu02=Ncl.cl01(+) "
       CheckOC3
       AdoRecordSet3.CursorLocation = adUseClient
       AdoRecordSet3.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
       If AdoRecordSet3.RecordCount <> 0 Then
           If CheckStr(AdoRecordSet3.Fields("cu121")) = "Y" Then
               pIsSpecCu = True
               pSpecCUName = pSpecCUName & CheckStr(AdoRecordSet3.Fields("cu04")) & "==>" & CheckStr(AdoRecordSet3.Fields("cl04")) & vbCrLf
               If pSpecMemo = "" Then    '秀玲說只要記錄第一筆
                   pSpecMemo = CheckStr(AdoRecordSet3.Fields("cl04"))
               End If
           End If
           'add by nickc 2008/01/18 加入客戶業務備註
           If CheckStr(AdoRecordSet3.Fields("cu125")) <> "" Then
               pIsCuMemo = True
               pCuMemoName = pCuMemoName & CheckStr(AdoRecordSet3.Fields("cu04")) & "==>" & CheckStr(AdoRecordSet3.Fields("cu125")) & vbCrLf
               'Modify by Amy 2016/03/14 申請人1~5有業務備註都要顯示
               'If pCuMemo = "" Then   '記錄第一筆
                   'Modify by Amy 2016/08/15 顯示於Text1(119),但列印於最後
                   pCuMemo = pCuMemo & "申請人" & iiiii & "：" & CheckStr(AdoRecordSet3.Fields("cu125")) & "@;@"
               'End If
           End If
       End If
   Next iiiii
   If pCuMemo <> MsgText(601) Then pCuMemo = Left(pCuMemo, Len(pCuMemo) - 3) 'Modify by Amy 2016/08/15
End Sub

'2011/10/19 add by sonia 舊案檢查申請地址與客戶中文地址是否相同,若不同則提醒
Private Sub CheckCustAddr(ByVal strCaseAddr As String, ByVal strCustNo As String, ByVal strI As String)
   If strCustNo <> "" Then
      strExc(0) = "select cu23 from customer where cu01='" & Mid(strCustNo, 1, 8) & "' and cu02='" & Mid(strCustNo, 9, 1) & "'"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         If RTrim(strCaseAddr) <> "" & RTrim(RsTemp.Fields(0)) Then
            ShowMsg "此案件申請人" & Chr(Asc(strI) - 23937) & "的地址(" & RTrim(strCaseAddr) & ")" & vbCrLf & _
                    "　　　與客戶現行地址(" & "" & RTrim(RsTemp.Fields(0)) & ") 不同！" & vbCrLf & _
                    "若要變更地址請於接洽單加註！"
         End If
         'Added by Lydia 2016/12/02 若客戶更名,提醒收據抬頭是否要修改
         If Mid(strCustNo, 9, 1) <> "0" And Mid(strCustNo, 9, 1) <> "" Then
            ShowMsg "此案件申請人" & Chr(Asc(strI) - 23937) & "已更名，請注意收據抬頭是否要修改！"
         End If
         'end 2016/12/02
      End If
   End If
End Sub
'2011/10/19 end

'2011/10/21 add by sonia 專利商標舊案申請地址拆成郵遞區號及地址
Private Sub AddrToZipAddr()
   strZip = ""
   Do While Left(strCAddr, 1) >= "０" And Left(strCAddr, 1) <= "９"
      strZip = strZip & Left(strCAddr, 1)
      strCAddr = Mid(strCAddr, 2)
   Loop
End Sub
'2011/10/21 end

Private Sub txtDay_GotFocus(Index As Integer)
   TextInverse txtDay(Index)
   CloseIme
End Sub

Private Sub txtDay_KeyPress(Index As Integer, KeyAscii As Integer)
   If KeyAscii <> 8 And (KeyAscii > Asc("9") Or KeyAscii < Asc("0")) Then
      KeyAscii = 0
      Beep
   End If
End Sub
Private Sub txtMonth_GotFocus(Index As Integer)
   TextInverse txtMonth(Index)
   CloseIme
End Sub

Private Sub txtMonth_KeyPress(Index As Integer, KeyAscii As Integer)
   If KeyAscii <> 8 And (KeyAscii > Asc("9") Or KeyAscii < Asc("0")) Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub txtYear_GotFocus(Index As Integer)
   TextInverse txtYear(Index)
   CloseIme
End Sub

Private Sub txtYear_KeyPress(Index As Integer, KeyAscii As Integer)
   If KeyAscii <> 8 And Not IsNumeric(Chr(KeyAscii)) Then
      KeyAscii = 0
      Beep
   End If
End Sub

'Added by Morgan 2013/4/9
'案件是否可減免改抓畫面上的設定
'Y:符合,N:不符合,empty:未設定
Private Function GetCP81() As String
   If optCP811(0).Value = True _
      And (Trim(Text1(37) & Text1(38)) = "" Or optCP812(0).Value = True) _
      And (Trim(Text1(53) & Text1(54)) = "" Or optCP813(0).Value = True) _
      And (Trim(Text1(69) & Text1(70)) = "" Or optCP814(0).Value = True) _
      And (Trim(Text1(85) & Text1(86)) = "" Or optCP815(0).Value = True) Then
      GetCP81 = "Y"
   ElseIf optCP811(1).Value = True _
      Or (Trim(Text1(37) & Text1(38)) <> "" And optCP812(1).Value = True) _
      Or (Trim(Text1(53) & Text1(54)) <> "" And optCP813(1).Value = True) _
      Or (Trim(Text1(69) & Text1(70)) <> "" And optCP814(1).Value = True) _
      Or (Trim(Text1(85) & Text1(86)) <> "" And optCP815(1).Value = True) Then
      GetCP81 = "N"
   Else
      GetCP81 = Empty
   End If
End Function
'Added by Morgan 2013/4/10
'檢查是否可減免有無變更
Private Function DiscStatusIsChanged() As Boolean
   '舊客戶且減免狀態有變
   
   If (Option31(1).Value = True And ((optCP811(0).Value = True And arrAD1516(1, 0) <> "Y") Or (optCP811(1).Value = True And arrAD1516(1, 0) <> "N"))) _
      Or (Option32(1).Value = True And ((optCP812(0).Value = True And arrAD1516(2, 0) <> "Y") Or (optCP812(1).Value = True And arrAD1516(2, 0) <> "N"))) _
      Or (Option33(1).Value = True And ((optCP813(0).Value = True And arrAD1516(3, 0) <> "Y") Or (optCP813(1).Value = True And arrAD1516(3, 0) <> "N"))) _
      Or (Option34(1).Value = True And ((optCP814(0).Value = True And arrAD1516(4, 0) <> "Y") Or (optCP814(1).Value = True And arrAD1516(4, 0) <> "N"))) _
      Or (Option35(1).Value = True And ((optCP815(0).Value = True And arrAD1516(5, 0) <> "Y") Or (optCP815(1).Value = True And arrAD1516(5, 0) <> "N"))) Then
      DiscStatusIsChanged = True
   'Added by Morgan 2019/4/15
   '日本案減免身分或資格變更也算(減免金額可能會不同)
   ElseIf stCountry = "011" And ((Option31(1).Value = True And optCP811(0).Value = True And (arrAD1516(1, 3) <> arrAD1516(1, 5) Or arrAD1516(1, 1) <> arrAD1516(1, 6))) _
      Or (Option32(1).Value = True And optCP812(0).Value = True And (arrAD1516(2, 3) <> arrAD1516(2, 5) Or arrAD1516(2, 1) <> arrAD1516(2, 6))) _
      Or (Option33(1).Value = True And optCP813(0).Value = True And (arrAD1516(3, 3) <> arrAD1516(3, 5) Or arrAD1516(3, 1) <> arrAD1516(3, 6))) _
      Or (Option34(1).Value = True And optCP814(0).Value = True And (arrAD1516(4, 3) <> arrAD1516(4, 5) Or arrAD1516(4, 1) <> arrAD1516(4, 6))) _
      Or (Option35(1).Value = True And optCP815(0).Value = True And (arrAD1516(5, 3) <> arrAD1516(5, 5) Or arrAD1516(5, 1) <> arrAD1516(5, 6)))) Then
      DiscStatusIsChanged = True
   Else
      DiscStatusIsChanged = False
   End If
End Function

'Added by Morgan 2013/4/10
'控制台灣專利領證或繳年費才可設定中小企業減免資格
Private Sub SetQualVisible(Optional iTab As Integer = -1, Optional bolCheck As Boolean)
   Dim ii As Integer, bolEnabled As Boolean
   
   m_strCaseCPM = GetAllCaseCPM 'Add By Sindy 2022/8/29 取得案件性質代碼
   
   'Added by Morgan 2019/4/12
   Label1(5) = "符合年費減免：": cmdQual(1).Caption = "中小企業減免資格"
   Label1(27) = Label1(5): cmdQual(2).Caption = cmdQual(1).Caption
   Label1(40) = Label1(5): cmdQual(3).Caption = cmdQual(1).Caption
   Label1(53) = Label1(5): cmdQual(4).Caption = cmdQual(1).Caption
   Label1(66) = Label1(5): cmdQual(5).Caption = cmdQual(1).Caption
  'end 2019/4/12
  
   bolCheck = False
   If iTab = -1 Then iTab = SSTab2.Tab
   
   cmdQual(iTab + 1).Visible = False
   
   If Text1(6) = "P" And Trim(Left(Combo1(0), 4)) = "000" Then
      ii = 5
      'For ii = 1 To 4
         'If Trim(Left(Me.Combo1(ii).Text, 4)) = "601" Then
         If InStr(m_strCaseCPM, "601") > 0 Then
            ii = 1
            If DBDATE(Text1(1)) >= "20130718" Then bolCheck = True
            'Exit For
         'ElseIf Trim(Left(Me.Combo1(ii).Text, 4)) = "605" Then
         ElseIf InStr(m_strCaseCPM, "605") > 0 Then
            ii = 1
            If DBDATE(Text1(1)) >= "20130801" Then bolCheck = True
            'Exit For
         End If
      'Next
      If ii < 5 Then
         Select Case iTab
         Case 0 '申請人1
            '可減免
            If optCP811(0).Value = True Then
               '舊客戶
               If Option31(1).Value = True Then
                  If Text1(12) <> "" Then
                     If GetCU15(Text1(12)) = "1" Then
                        cmdQual(1).Visible = True
                     End If
                  End If
               '新客戶
               Else
                  cmdQual(1).Visible = True
               End If
            End If
            
         Case 1 '申請人2
            '可減免
            If optCP812(0).Value = True Then
               '舊客戶
               If Option32(1).Value = True Then
                  If Text1(28) <> "" Then
                     If GetCU15(Text1(28)) = "1" Then
                        cmdQual(2).Visible = True
                     End If
                  End If
               '新客戶
               Else
                  cmdQual(2).Visible = True
               End If
            End If
         
         Case 2 '申請人3
            '可減免
            If optCP813(0).Value = True Then
               '舊客戶
               If Option33(1).Value = True Then
                  If Text1(44) <> "" Then
                     If GetCU15(Text1(44)) = "1" Then
                        cmdQual(3).Visible = True
                     End If
                  End If
               '新客戶
               Else
                  cmdQual(3).Visible = True
               End If
            End If
         Case 3 '申請人4
            '可減免
            If optCP814(0).Value = True Then
               '舊客戶
               If Option34(1).Value = True Then
                  If Text1(60) <> "" Then
                     If GetCU15(Text1(60)) = "1" Then
                        cmdQual(4).Visible = True
                     End If
                  End If
               '新客戶
               Else
                  cmdQual(4).Visible = True
               End If
            End If
         
         Case 4 '申請人5
            '可減免
            If optCP815(0).Value = True Then
               '舊客戶
               If Option35(1).Value = True Then
                  If Text1(76) <> "" Then
                     If GetCU15(Text1(76)) = "1" Then
                        cmdQual(5).Visible = True
                     End If
                  End If
               '新客戶
               Else
                  cmdQual(5).Visible = True
               End If
            End If
         End Select
      End If

   'Added by Morgan 2019/4/12 +CFP日本發明專利案
   ElseIf Text1(6) = "CFP" And Trim(Left(Combo1(0), 4)) = "011" Then
   
      Label1(5) = "符合減免資格：": cmdQual(1).Caption = "減免資格"
      Label1(27) = Label1(5): cmdQual(2).Caption = cmdQual(1).Caption
      Label1(40) = Label1(5): cmdQual(3).Caption = cmdQual(1).Caption
      Label1(53) = Label1(5): cmdQual(4).Caption = cmdQual(1).Caption
      Label1(66) = Label1(5): cmdQual(5).Caption = cmdQual(1).Caption
         
      SetOpt81 stCountry '重新設定一次(因需考慮專利種類或收文性質)
      
      'Modify By Sindy 2022/11/9 + And Not Me.ActiveControl Is Nothing
      If IsoptCP81 And Not Me.ActiveControl Is Nothing Then
         Select Case iTab
         Case 0 '申請人1
            cmdQual(1).Visible = optCP811(0).Value
            If cmdQual(1).Visible And Me.ActiveControl.Name = optCP811(0).Name Then cmdQual(1).Value = True
         Case 1 '申請人2
            cmdQual(2).Visible = optCP812(0).Value
            If cmdQual(2).Visible And Me.ActiveControl.Name = optCP812(0).Name Then cmdQual(2).Value = True
         Case 2 '申請人3
            cmdQual(3).Visible = optCP813(0).Value
            If cmdQual(3).Visible And Me.ActiveControl.Name = optCP813(0).Name Then cmdQual(3).Value = True
         Case 3 '申請人4
            cmdQual(4).Visible = optCP814(0).Value
            If cmdQual(4).Visible And Me.ActiveControl.Name = optCP814(0).Name Then cmdQual(4).Value = True
         Case 4 '申請人5
            cmdQual(5).Visible = optCP815(0).Value
            If cmdQual(5).Visible And Me.ActiveControl.Name = optCP815(0) Then cmdQual(5).Value = True
         End Select
      End If
   End If
End Sub

Private Function GetCU15(pCuNo As String) As String
   Dim stSQL As String, intR As Integer
   Dim rsQuery As ADODB.Recordset
   stSQL = "select cu15 from customer where cu01='" & Left(pCuNo & "000", 8) & "' and cu02='" & Mid(pCuNo & "000", 9, 1) & "'"
   intR = 1
   Set rsQuery = ClsLawReadRstMsg(intR, stSQL)
   If intR = 1 Then
      GetCU15 = "" & rsQuery(0)
   End If
   Set rsQuery = Nothing
End Function

'Added by Morgan 2021/7/20
'大陸發明生醫案是否新藥專利設定
Private Sub SetNewDrug()
   If Option1(0).Value = True And Text1(6) = "P" And Left(Combo1(0), 3) = "020" And Left(Combo6, 1) = "1" And Left(Combo5, 1) = "3" Then
      OptNewDrug(0).Value = False
      OptNewDrug(1).Value = False
      Frame48.Visible = True
   Else
      Frame48.Visible = False
   End If
End Sub

'Added by Morgan 2013/4/10
'更新中小企業減免資格
Private Function UpdateDiscountQual(Optional pbolInTrans As Boolean = True) As Boolean
   Dim stSQL As String
   
   If DiscStatusIsChanged() = False Then

      If pbolInTrans = False Then
         cnnConnection.BeginTrans
         On Error GoTo ErrHnd
      End If
      'Modified by Morgan 2019/4/15 +改判斷有修改才執行並加記錄修改人員(減免身分只有日本案可設定且目前不會直接更新)
      If stCustNo1 <> "" And optCP811(0).Value = True And cmdQual(1).Visible And (arrAD1516(1, 1) <> arrAD1516(1, 6) Or arrAD1516(1, 2) <> arrAD1516(1, 7) Or arrAD1516(1, 3) <> arrAD1516(1, 5)) Then
         stSQL = "update applicantdiscount set ad10='" & arrAD1516(1, 3) & "',ad15='" & arrAD1516(1, 1) & "',ad16=" & Val(arrAD1516(1, 2)) & ",ad07='" & strUserNum & "',ad08=to_char(sysdate,'yyyymmdd'),ad09=to_char(sysdate,'hh24mi') where ad01='" & Left(stCustNo1 & "00", 8) & "' and ad02='" & stCountry & "'"
         Pub_SeekTbLog stSQL
         cnnConnection.Execute stSQL, intI
      End If
      If stCustNo2 <> "" And optCP812(0).Value = True And cmdQual(2).Visible And (arrAD1516(2, 1) <> arrAD1516(2, 6) Or arrAD1516(2, 2) <> arrAD1516(2, 7) Or arrAD1516(2, 3) <> arrAD1516(2, 5)) Then
         stSQL = "update applicantdiscount set ad10='" & arrAD1516(2, 3) & "',ad15='" & arrAD1516(2, 1) & "',ad16=" & Val(arrAD1516(2, 2)) & ",ad07='" & strUserNum & "',ad08=to_char(sysdate,'yyyymmdd'),ad09=to_char(sysdate,'hh24mi') where ad01='" & Left(stCustNo2 & "00", 8) & "' and ad02='" & stCountry & "'"
         Pub_SeekTbLog stSQL
         cnnConnection.Execute stSQL, intI
      End If
      If stCustNo3 <> "" And optCP813(0).Value = True And cmdQual(3).Visible And (arrAD1516(3, 1) <> arrAD1516(3, 6) Or arrAD1516(3, 2) <> arrAD1516(3, 7) Or arrAD1516(3, 3) <> arrAD1516(3, 5)) Then
         stSQL = "update applicantdiscount set ad10='" & arrAD1516(3, 3) & "',ad15='" & arrAD1516(3, 1) & "',ad16=" & Val(arrAD1516(3, 2)) & ",ad07='" & strUserNum & "',ad08=to_char(sysdate,'yyyymmdd'),ad09=to_char(sysdate,'hh24mi') where ad01='" & Left(stCustNo3 & "00", 8) & "' and ad02='" & stCountry & "'"
         Pub_SeekTbLog stSQL
         cnnConnection.Execute stSQL, intI
      End If
      If stCustNo4 <> "" And optCP814(0).Value = True And cmdQual(4).Visible And (arrAD1516(4, 1) <> arrAD1516(4, 6) Or arrAD1516(4, 2) <> arrAD1516(4, 7) Or arrAD1516(4, 3) <> arrAD1516(4, 5)) Then
         stSQL = "update applicantdiscount set ad10='" & arrAD1516(4, 3) & "',ad15='" & arrAD1516(4, 1) & "',ad16=" & Val(arrAD1516(4, 2)) & ",ad07='" & strUserNum & "',ad08=to_char(sysdate,'yyyymmdd'),ad09=to_char(sysdate,'hh24mi') where ad01='" & Left(stCustNo4 & "00", 8) & "' and ad02='" & stCountry & "'"
         Pub_SeekTbLog stSQL
         cnnConnection.Execute stSQL, intI
      End If
      If stCustNo5 <> "" And optCP815(0).Value = True And cmdQual(5).Visible And (arrAD1516(5, 1) <> arrAD1516(5, 6) Or arrAD1516(5, 2) <> arrAD1516(5, 7) Or arrAD1516(5, 3) <> arrAD1516(5, 5)) Then
         stSQL = "update applicantdiscount set ad10='" & arrAD1516(5, 3) & "',ad15='" & arrAD1516(5, 1) & "',ad16=" & Val(arrAD1516(5, 2)) & ",ad07='" & strUserNum & "',ad08=to_char(sysdate,'yyyymmdd'),ad09=to_char(sysdate,'hh24mi') where ad01='" & Left(stCustNo5 & "00", 8) & "' and ad02='" & stCountry & "'"
         Pub_SeekTbLog stSQL
         cnnConnection.Execute stSQL, intI
      End If
      If pbolInTrans = False Then cnnConnection.CommitTrans
   End If
   UpdateDiscountQual = True
   Exit Function
   
ErrHnd:
   If pbolInTrans = False Then
      cnnConnection.RollbackTrans
      MsgBox "中小企業減免資格更新失敗！", vbCritical
   End If
End Function

'Add By Sindy 2014/2/6
Private Function SaveCustData(Optional pbolInTrans As Boolean = True) As Boolean
Dim i As Integer
Dim bolOldCust As Boolean
Dim intRec As Integer 'Add by Amy 2016/01/04
'Add by Amy 2016/05/20
Dim stField As String
'Add by Amy 2016/12/23
Dim strUpd As String '原用strsql 因為共用變數其他function也有用到會有問題
Dim stCU13 As String
'Added by Lydia 2020/03/30
Dim rsA1 As New ADODB.Recordset
Dim strA1 As String

On Error GoTo ErrHnd
   
   If pbolInTrans = False Then
      cnnConnection.BeginTrans
   End If
   
   For i = 0 To 4
      If i = 0 Then bolOldCust = Option31(1).Value
      If i = 1 Then bolOldCust = Option32(1).Value
      If i = 2 Then bolOldCust = Option33(1).Value
      If i = 3 Then bolOldCust = Option34(1).Value
      If i = 4 Then bolOldCust = Option35(1).Value
      '舊客戶有客戶編號且統一編號欄位開放可以輸入時
      If bolOldCust = True And Trim(Me.Text1(12 + (16 * i)).Text) <> "" And Me.Text1(92 + i).Enabled = True Then
         If Me.Text1(92 + i).Text <> "" Then
            strUpd = "update customer" & _
                     " set cu11='" & Me.Text1(92 + i).Text & "'" & _
                     " where cu01='" & Left(Trim(Me.Text1(12 + (16 * i)).Text), 8) & "'" & _
                       " and cu02='" & Right(Trim(Me.Text1(12 + (16 * i)).Text), 1) & "'"
            cnnConnection.Execute strUpd
         End If
      End If
      
      'Add by Amy 2016/01/04 新案舊客戶有客戶編號且申請地址/聯絡地址郵遞區號有改過時(原本為null也回寫)
      If Option1(0).Value = True Then
        If bolOldCust = True Then
            strUpd = ""
            'Add by Amy 2016/05/20 畫面收據公司別與申請人1預設公司別不同時更新相對應欄位
            'Modify by Amy 2016/12/23 +同業務區或為MCTF同組人員才可回寫收據公司別
            If i = 0 Then
                If ChkSameCuArea(Trim(Me.Text1(12 + (16 * i))), Trim(Me.Text1(10))) = True Then
                    'Modify by Amy 2019/12/04 +GetReceiptVal
                    If Me.Combo4.Tag <> GetReceiptVal(Me.Combo4, False) Then
                        Select Case Text1(6)
                            Case "P", "CFP", "FCP"
                                If Left(Combo1(0), 3) = 台灣國家代號 Then
                                    stField = ",cu160="
                                Else
                                    stField = ",cu161="
                                End If
                            Case "T", "TF", "CFT", "FCT"
                                If Left(Combo1(0), 3) = 台灣國家代號 Then
                                    stField = ",cu162="
                                Else
                                    stField = ",cu163="
                                End If
                            Case Else
                                If Left(Combo1(0), 3) = 台灣國家代號 Then
                                    stField = ",cu164="
                                Else
                                    stField = ",cu165="
                                End If
                        End Select
                        'Modify by Amy 2019/12/04 改寫至GetReceiptVal
                        stField = stField & CNULL(GetReceiptVal(Me.Combo4, True))
                        If stField <> MsgText(601) Then strUpd = strUpd & stField
                    End If
                End If
                'end 2016/12/23
            End If
            'Added by Lydia 2020/03/30 新案以系統類別抓系統種類對照表，若SK02為3、4(法務案)及7、8(顧問案)的收據公司別鎖住，且不更新回客戶檔
            If i = 0 And Combo4.Enabled = False Then
                'Modified by Lydia 2020/04/10 請改為instr(系統類別,'L')>0的收據公司別鎖住，且不更新回客戶檔
                'strA1 = "select sk02 from systemkind where sk01='" & Text1(6) & "' "
                'i = 1
                'Set rsA1 = ClsLawReadRstMsg(i, strA1)
                'If i = 1 Then
                '    If InStr("3,4,7,8", "" & rsA1.Fields("sk02")) > 0 Then
                '        strUpd = ""
                '    End If
                'End If
                'Set rsA1 = Nothing
                If InStr(UCase(Text1(6)), "L") > 0 Then
                    strUpd = ""
                End If
                'end 2020/04/10
            End If
            'end 2020/03/30
            
            If strUpd <> MsgText(601) Then
                '客戶基本檔
                strUpd = "Update Customer" & _
                            " Set " & Mid(strUpd, 2) & _
                            " Where cu01='" & Left(Trim(Me.Text1(12 + (16 * i)).Text), 8) & "'" & _
                            " and cu02='" & Right(Trim(Me.Text1(12 + (16 * i)).Text), 1) & "'"
                Pub_SeekTbLog strUpd
                cnnConnection.Execute strUpd, intRec
                If intRec > 0 Then
                    'Modify by Amy 2019/12/04 +GetReceiptVal
                    If i = 0 And Me.Combo4.Tag <> GetReceiptVal(Me.Combo4, False) Then Me.Combo4.Tag = GetReceiptVal(Me.Combo4, False)
                    Me.Text1(120 + i).Tag = Me.Text1(120 + i)
                    Me.Text1(25 + i).Tag = Me.Text1(25 + i)
                End If
            End If
            'end 2015/05/20
        End If
      End If
      'end 2016/01/04
   Next i
   
   If pbolInTrans = False Then cnnConnection.CommitTrans
   SaveCustData = True
   Exit Function
   
ErrHnd:
   If pbolInTrans = False Then
      cnnConnection.RollbackTrans
      'Modify by Amy 2016/07/27
      'MsgBox "客戶統一編號更新失敗！", vbCritical
      'Modified by Lydia 2019/08/12 +Titile
      'MsgBox Err.Description, vbCritical
      MsgBox Err.Description, vbCritical, "SaveCustData"
   End If
End Function

'Add by Amy 2014/07/18 依系統別抓取案件性質名稱
Private Function GetCasePropertyName(ByVal strSystemKind As String, ByVal strCPM02 As String, ByVal strField As String) As String
    Dim strSqlq As String, intQ As Integer
    Dim RsQ As New ADODB.Recordset
    
    GetCasePropertyName = ""
    strSqlq = "Select " & strField & " From CasePropertyMap Where CPM01='" & strSystemKind & "' And CPM02='" & strCPM02 & "' "
    intQ = 1
    Set RsQ = ClsLawReadRstMsg(intQ, strSqlq)
    If intQ = 1 Then
        GetCasePropertyName = "" & RsQ.Fields(0)
    End If
End Function
'end 2014/07/18

'Add by Lydia 2014/12/25 台灣P案未設定減免身份提示訊息
Private Function MsgDiscSet() As Boolean
If Trim(Left(Me.Combo1(0).Text, 4)) = "000" And Trim(Text1(6)) = "P" Then
  MsgDiscSet = False
  
    If optCP811(0).Value = 0 And optCP811(1).Value = 0 Then
       MsgBox "申請人1尚未設定減免身份!!": MsgDiscSet = True
    ElseIf (Trim(Text1(37) & Text1(38))) <> "" And optCP812(0).Value = 0 And optCP812(1).Value = 0 Then
        MsgBox "申請人2尚未設定減免身份!!": MsgDiscSet = True
    ElseIf (Trim(Text1(53) & Text1(54))) <> "" And optCP813(0).Value = 0 And optCP813(1).Value = 0 Then
        MsgBox "申請人3尚未設定減免身份!!": MsgDiscSet = True
    ElseIf (Trim(Text1(69) & Text1(70))) <> "" And optCP814(0).Value = 0 And optCP814(1).Value = 0 Then
        MsgBox "申請人4尚未設定減免身份!!": MsgDiscSet = True
    ElseIf (Trim(Text1(85) & Text1(86))) <> "" And optCP815(0).Value = 0 And optCP815(1).Value = 0 Then
        MsgBox "申請人5尚未設定減免身份!!": MsgDiscSet = True
    End If
End If
End Function

'Add by Lydia 2015/01/06 使用者會有保留前案畫面,進行下一案例操作的習慣,所以案號一改需清空費用相關計算變數和欄位
Private Sub FreeClear()
    txtCRL69.Text = "": txtCRL70.Text = "" 'Add By Sindy 2022/9/15
    'cmdOK(3).Caption = "期限資料(&L)" 'Add By Sindy 2015/4/2
    m_Note1 = "": m_Note2 = ""
    m_strGetNP01 = "" 'Add By Sindy 2015/9/17
    Text1(142).Text = "": Text1(143).Text = ""
    Text1(1) = "": Text1(3) = ""
    Check11.Value = 0 'Add By Sindy 2022/11/4
    Check12.Value = 0 'Add By Sindy 2022/12/13
    lblCnt.Caption = "": Combo1(1) = "": Text1(101) = "": Text1(102) = "": Text1(103) = "": Combo2(0) = ""
    GridCase.Clear
    Call SetGrd
    mPYFee = False
    Frame605.Visible = False
End Sub

'Add By Sindy 2022/8/29 案件性質
Private Sub SetGrd()
   Dim arrGridHeadText, arrGridHeadWidth
   Dim iRow As Integer
   'Modify By Sindy 2025/4/14 +,CRC13 規費調整
   '                        0       1           2       3       4       5       6       7           8           9         10          11      12      13      14      15      16
   arrGridHeadText = Array("順序", "案件性質", "費用", "規費", "點數", "備註", "案號", "總收文號", "算案件數", "計件值", "加乘註記", "cp01", "cp02", "cp03", "cp04", "CPMn", "規費調整")
   If LblRecved.Visible = True Then
      If Text1(6) = "P" Or Text1(6) = "PS" Or Text1(6) = "CFP" Or Text1(6) = "CPS" Then
         arrGridHeadWidth = Array(400, 1700, 1000, 1000, 1000, 1500, 1200, 1000, 800, 800, 800, 0, 0, 0, 0, 0, 0)
      Else
         arrGridHeadWidth = Array(400, 1700, 1000, 1000, 1000, 1500, 1200, 1000, 0, 0, 0, 0, 0, 0, 0, 0, 0)
      End If
   Else
      arrGridHeadWidth = Array(400, 1700, 1000, 1000, 1000, 1500, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)
   End If
   GridCase.Visible = False
   GridCase.Cols = UBound(arrGridHeadText) + 1
   GridCase.Rows = 2
   For iRow = 0 To GridCase.Cols - 1
      GridCase.row = 0
      GridCase.col = iRow
      GridCase.Text = arrGridHeadText(iRow)
      GridCase.ColWidth(iRow) = arrGridHeadWidth(iRow)
      If iRow = 1 Then
         GridCase.CellAlignment = flexAlignLeftCenter '儲存格內容中間靠左對齊。這是對字串的預設值。
      Else
         GridCase.CellAlignment = flexAlignLeftCenter 'flexAlignCenterCenter 'flexAlignRightCenter
      End If
   Next
   GridCase.Visible = True
End Sub

'Add By Sindy 2022/8/29
'案件性質編輯區
'清除:欄位清空
Private Sub cmdClear2_Click()
   lblCnt.Caption = "" '順序
   Combo1(1).Text = "": Combo1(1).Tag = "" '案件性質
   Text1(101).Text = "": Text1(101).Tag = "" '費用
   Text1(102).Text = "": Text1(102).Tag = "" '規費
   Text1(103).Text = "": Text1(103).Tag = "" '點數
   Combo2(0).Text = "": Combo2(0).Tag = "" '備註
End Sub
'刪除:移除勾選的資料列
Private Sub cmdDel_Click()
Dim j As Integer
   
   If Val(lblCnt.Caption) > 0 Then
      GridCase.Tag = "有異動"
      If GridCase.Rows - 1 = 1 Then
         GridCase.Clear: Call SetGrd
      Else
         GridCase.RemoveItem Val(lblCnt.Caption)
      End If
      '重新整理順序
      For j = 1 To GridCase.Rows - 1
         If Trim(GridCase.TextMatrix(j, 1)) <> "" Then
            GridCase.TextMatrix(j, 0) = j
         End If
      Next j
      FrameCRC.Caption = "案件性質編輯區（" & Val(GridCase.Rows - 1) - IIf(Trim(GridCase.TextMatrix(1, 1)) = "", 1, 0) & "）"
      Call cmdClear2_Click '清除
   End If
End Sub
'加入:新增或修改
Private Sub cmdUpd_Click()
Dim intRow As Integer, j As Integer, Cancel As Boolean
   
   If Me.Combo1(1).Text = "" Then Exit Sub
   If Me.Combo1(1).Text = "" Then
      MsgBox "案件性質不可空白！", vbExclamation + vbOKOnly
      SSTab1.Tab = 0
'      Combo1(1).SetFocus
      Exit Sub
   Else
      Cancel = False
      Combo1_Validate 1, Cancel
      If Cancel = True Then
         SSTab1.Tab = 0
'         Combo1(1).SetFocus
         Exit Sub
      End If
   End If
   
   GridCase.Tag = "有異動"
   '修改
   If Val(lblCnt.Caption) > 0 Then
      intRow = Val(lblCnt.Caption)
   '新增
   Else
      intRow = GridCase.Rows - 1
      
      If InStr(Text1(6), "L") > 0 And intRow >= 4 Then
         MsgBox "L的案件性質輸入不可超過 4 筆！", vbExclamation + vbOKOnly
         SSTab1.Tab = 0
         Exit Sub
      End If
      
      If Trim(GridCase.TextMatrix(intRow, 1)) <> "" Then
         GridCase.AddItem ""
         intRow = intRow + 1
      End If
   End If
   
   GridCase.TextMatrix(intRow, 0) = intRow '順序
   GridCase.TextMatrix(intRow, 1) = " " & Trim(Combo1(1).Text) '案件性質
   GridCase.TextMatrix(intRow, 2) = Format(Text1(101).Text, "#,##0")  '費用
   GridCase.TextMatrix(intRow, 3) = Format(Text1(102).Text, "#,##0") '規費
   GridCase.TextMatrix(intRow, 4) = Format(Text1(103).Text, "#,##0.000") '點數
   GridCase.TextMatrix(intRow, 5) = Combo2(0).Text '備註
   GridCase.col = 1
   GridCase.row = intRow
   FrameCRC.Caption = "案件性質編輯區（" & Val(GridCase.Rows - 1) - IIf(Trim(GridCase.TextMatrix(1, 1)) = "", 1, 0) & "）"
   Call cmdClear2_Click '清除
   m_strCaseCPM = GetAllCaseCPM 'Add By Sindy 2022/9/12 取得案件性質代碼
End Sub

'Add By Sindy 2022/8/29 取得目前輸入的案件性質代碼
'Optional ByRef bolNewCase As Boolean = False : 回傳是否為新申請案案件性質
'Optional ByRef bolTMdebate As Boolean = False : 回傳是否為商爭案件性質
'Optional ByRef strNewCP10 As String = "" : 回傳新申請案的案件性質
'Optional ByRef bolHad10Point As Boolean = False : 檢查是否有點數超過或等於10點的
'Optional ByRef dblTotOFee As Double = 0 : 回傳總規費
'Optional ByRef dblTotRvFee As Double = 0 : 回傳總費用
'Optional ByRef dblTotPoint As Double = 0 : 回傳總點數
Private Function GetAllCaseCPM(Optional ByRef bolNewCase As Boolean = False, _
   Optional ByRef bolTMdebate As Boolean = False, Optional ByRef strNewCP10 As String = "", _
   Optional ByRef bolHad10Point As Boolean = False, Optional ByRef dblTotOFee As Double = 0, _
   Optional ByRef dblTotRvFee As Double = 0, Optional ByRef dblTotPoint As Double = 0) As String
Dim j As Integer
Dim arrData As Variant
      
   GetAllCaseCPM = ""
   bolNewCase = False: bolTMdebate = False: strNewCP10 = "": bolHad10Point = False: dblTotOFee = 0: dblTotRvFee = 0: dblTotPoint = 0
   For j = 1 To GridCase.Rows - 1
      If Trim(GridCase.TextMatrix(j, 1)) <> "" Then
         arrData = Split(Trim(GridCase.TextMatrix(j, 1)), " ")
         'Add By Sindy 2025/2/4
         If InStr(arrData(0), "修改") > 0 And Trim(GridCase.TextMatrix(j, 7)) <> "" Then
            strExc(0) = "Select cp09,cp10 " & _
                        "From caseprogress " & _
                        "Where cp09 ='" & Trim(GridCase.TextMatrix(j, 7)) & "'"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               GetAllCaseCPM = GetAllCaseCPM & "," & RsTemp.Fields("cp10")
            End If
         Else
         '2025/2/4 END
            GetAllCaseCPM = GetAllCaseCPM & "," & arrData(0)
         End If
         '專利
         If InStr("1", m_strSys) > 0 And InStr(NewCasePtyList, arrData(0)) > 0 Then
            strNewCP10 = arrData(0)
            bolNewCase = True
         '商標
         ElseIf InStr("2", m_strSys) > 0 Then
            If arrData(0) = "101" Then
               strNewCP10 = arrData(0)
               bolNewCase = True
            End If
            'Modify By Sindy 2025/7/30 FCT案727分析不屬於爭議案 + FCT_NotTMdebate
            If InStr(TMdebate, arrData(0)) > 0 And Not (Text1(6) = "FCT" And InStr(FCT_NotTMdebate, arrData(0)) > 0) Then
               bolTMdebate = True
            End If
         End If
         '檢查點數是否超過或等於10點
         If Format(Val(GridCase.TextMatrix(j, 4)), "###0.000") >= 10 Then
            bolHad10Point = True
         End If
         '總規費:dblTotOFee
         If Val(Format(GridCase.TextMatrix(j, 3), "###0")) > 0 Then
            dblTotOFee = dblTotOFee + Val(Format(GridCase.TextMatrix(j, 3), "###0"))
         End If
         '總費用:dblTotRvFee
         If Val(Format(GridCase.TextMatrix(j, 2), "###0")) > 0 Then
            dblTotRvFee = dblTotRvFee + Val(Format(GridCase.TextMatrix(j, 2), "###0"))
         End If
         '總點數:dblTotPoint
         If Val(Format(GridCase.TextMatrix(j, 4), "###0")) > 0 Then
            dblTotPoint = dblTotPoint + Val(Format(GridCase.TextMatrix(j, 4), "###0"))
         End If
         'Add By Sindy 2025/4/14 規費有異常時,欄位變色
         If GridCase.TextMatrix(j, 16) <> "" Then
            GridCase.col = 3
            GridCase.row = j
            GridCase.CellBackColor = QBColor(14) '淡黃色
            Me.Check10.Value = 1
            Me.Check10.BackColor = QBColor(14) '淡黃色
         Else
            GridCase.col = 2
            GridCase.row = j
            strExc(10) = GridCase.CellBackColor
            GridCase.col = 3
            GridCase.CellBackColor = strExc(10)
         End If
         '2025/4/14 END
      End If
   Next j
   If GetAllCaseCPM <> "" Then GetAllCaseCPM = Mid(GetAllCaseCPM, 2)
End Function

'Add By Sindy 2022/8/29
Private Sub GridCase_Click()
Dim m_intRow As Integer
Dim m_intCol As Integer

   m_intRow = GridCase.MouseRow
   m_intCol = GridCase.MouseCol
   If m_intRow <> 0 Then
      If GridCase.TextMatrix(m_intRow, 1) <> "" Then
         If ((Val(lblCnt) = 0 And Trim(Combo1(1).Text) <> "")) Then 'Or ChkIsModify = True
            If MsgBox("案件性質有異動，要放棄嗎？", vbYesNo + vbCritical) = vbNo Then
               Exit Sub
            Else
               Call cmdClear2_Click '清除
            End If
         End If
         Call GetGridCaseData(m_intRow)
      End If
   End If
End Sub
'取得資料視窗中的資料列欄位值
Private Sub GetGridCaseData(intRow As Integer)
Dim j As Integer, i As Integer

   '檢查目前那一列為反白列
   dblPrevRow = 0
   For j = 1 To GridCase.Rows - 1
      GridCase.col = 1
      GridCase.row = j
      If GridCase.CellBackColor <> QBColor(15) Then
         dblPrevRow = j
         Exit For
      End If
   Next j
   If dblPrevRow <> intRow Then
      '上一筆資料列清除反白
      If dblPrevRow > 0 And dblPrevRow <= (GridCase.Rows - 1) Then
         For i = 0 To GridCase.Cols - 1
            GridCase.col = i
            If GridCase.CellBackColor <> &H8080FF And GridCase.CellBackColor <> QBColor(14) Then
               GridCase.CellBackColor = QBColor(15)
            End If
         Next i
      End If
   End If
   If intRow <> 0 Then
      '目前資料列反白
      GridCase.col = 1
      GridCase.row = intRow
      dblPrevRow = GridCase.row
      'If GridCase.CellBackColor = QBColor(15) Then
         For i = 0 To GridCase.Cols - 1
            GridCase.col = i
            If GridCase.CellBackColor <> &H8080FF And GridCase.CellBackColor <> QBColor(14) Then
               GridCase.CellBackColor = &HFFC0C0
            End If
         Next i
      'End If
   End If
   If GridCase.TextMatrix(intRow, 1) <> "" Then
      lblCnt.Caption = GridCase.TextMatrix(intRow, 0) '順序
      Combo1(1).Text = Trim(GridCase.TextMatrix(intRow, 1)) '案件性質
      Combo1(1).Tag = Combo1(1).Text
      Text1(101).Text = Format(GridCase.TextMatrix(intRow, 2), "###0")  '費用
      Text1(101).Tag = Text1(101).Text
      Text1(102).Text = Format(GridCase.TextMatrix(intRow, 3), "###0") '規費
      Text1(102).Tag = Text1(102).Text
      Text1(103).Text = Format(GridCase.TextMatrix(intRow, 4), "###0.000") '點數
      Text1(103).Tag = Text1(103).Text
      Combo2(0).Text = GridCase.TextMatrix(intRow, 5) '備註
      Combo2(0).Tag = Combo2(0).Text
   End If
End Sub

'Add By Sindy 2022/8/29
Private Sub GridCase_SelChange()
Dim j As Integer, i As Integer

GridCase.Visible = False
If GridCase.MouseRow = 0 Then
   '已選取的資料列清除反白
   For j = 1 To GridCase.Rows - 1
      GridCase.col = 1
      GridCase.row = j
      If GridCase.CellBackColor <> QBColor(15) Then
         For i = 0 To GridCase.Cols - 1
            GridCase.col = i
            GridCase.CellBackColor = QBColor(15)
         Next i
         Exit For
      End If
   Next j
Else
   Call GetGridCaseData(GridCase.MouseRow)
'   '檢查目前那一列為反白列
'   dblPrevRow = 0
'   For j = 1 To GridCase.Rows - 1
'      GridCase.col = 1
'      GridCase.row = j
'      If GridCase.CellBackColor <> QBColor(15) Then
'         dblPrevRow = j
'         Exit For
'      End If
'   Next j
'   If dblPrevRow <> GridCase.MouseRow Then
'      '上一筆資料列清除反白
'      If dblPrevRow > 0 And dblPrevRow <= (GridCase.Rows - 1) Then
'         For i = 0 To GridCase.Cols - 1
'            GridCase.col = i
'            If GridCase.CellBackColor <> &H8080FF Then
'               GridCase.CellBackColor = QBColor(15)
'            End If
'         Next i
'      End If
'      '目前資料列反白
'      GridCase.col = 1
'      GridCase.row = GridCase.MouseRow
'      dblPrevRow = GridCase.row
'      If GridCase.CellBackColor = QBColor(15) Then
'         For i = 0 To GridCase.Cols - 1
'            GridCase.col = i
'            If GridCase.CellBackColor <> &H8080FF Then
'               GridCase.CellBackColor = &HFFC0C0
'            End If
'         Next i
'      End If
'   End If
End If
GridCase.Visible = True
End Sub

Private Sub GrdTMQ_Click()
Dim TmpRow As Integer
Dim inR As Integer
TmpRow = GrdTMQ.MouseRow

If TmpRow > 0 Then
   If GrdTMQ.TextMatrix(TmpRow, 0) = "V" Then
      '清空資料
      GrdTMQ.col = 0
      GrdTMQ.row = TmpRow
      GrdTMQ.Text = ""
      For inR = 0 To GrdTMQ.Cols - 1
         GrdTMQ.col = inR
         GrdTMQ.CellBackColor = QBColor(15)
      Next inR
   Else
      '目前資料列反白
      GrdTMQ.col = 0
      GrdTMQ.row = TmpRow
      'Modified by Lydia 2016/05/05
      'If PUB_MGridGetValue(tmpRow, "TMQ11", GrdTMQ) <> "" Then
      If TMQ_CtrRead = False Then
         GrdTMQ.Text = "V"
      'Modified by Lydia 2016/03/28 是否控制查覆完畢
      'Else
      Else
         If m_UseTmqTma = "1" Then 'Added by Lydia 2024/11/11
            If TMQ_CtrRead And Trim(PUB_MGridGetValue(TmpRow, "TMQ11", GrdTMQ)) = "" Then
               MsgBox "委查單" & PUB_MGridGetValue(TmpRow, "委查單號", GrdTMQ) & " 尚未查覆完畢,請洽查名人員 " & GetStaffName(PUB_MGridGetValue(TmpRow, "TMQ10", GrdTMQ)), vbCritical, "委查結果"
            End If
         End If  'Added by Lydia 2024/11/11
      End If
   End If
End If
End Sub
'Modify By Sindy 2022/12/23 + Optional strTQC01 As String = "" : 查名代號
Public Sub QueryTMQ(Optional strTQC01 As String = "")
Dim rsAD As New ADODB.Recordset
Dim inX As Integer
Dim inA As Integer
Dim strA1 As String

    'Added by Lydia 2016/04/18
    GrdTMQ.FixedCols = 0
    SetGrdTMQ
    'end 2016/04/18
    If cmdTMQ.Tag <> "" Then
       strA1 = GetAddStr(cmdTMQ.Tag)
       'Added by Lydia 2024/11/11 查名單(網中)
       If m_UseTmqTma = "2" Then
          strSql = "select 'V' v,tma26 as 文字檢索,decode(tma27,null,null,'(圖形檢索)') as 圖形檢索,tma18 as 客戶名稱," & _
                  " decode(tma25,'2',null,decode(tma19," & PUB_GetTMQans("3", True) & ", tma39)) as 文字檢索結果,decode(tma25,'1',null,decode(tma19," & PUB_GetTMQans("3", True) & ",tma41)) as 圖形檢索結果," & _
                  " decode(tma20,null,decode(tma22,null,decode(tma23,null,null,tma23||',')||decode(tma24,null,null,tma24||','),tma22||','),tma20) as 類別組群,tma33 as 智權備註,tma01" & _
                  " from tmqappform WHERE "
          If strTQC01 <> "" Then
             strSql = strSql & " tma01 in (select tqc03 from tmqcasemap where tqc01='" & strTQC01 & "' and tqc03<>'" & cntTQC自動記錄 & "') "
          Else
             strSql = strSql & " tma01 in (" & strA1 & ") "
          End If
          strSql = strSql & " ORDER BY TMA01 desc "
       Else
       'end 2024/11/11
          'Modified by Lydia 2016/01/27 排除整張不查的委查單
          'Modified by Lydia 2016/03/28 +控制TMQ_CtrRead
          'Modified by Lydia 2016/04/06 +控制TMQ_ReApp
          'Modified by Lydia 2016/06/02 覆核結果取代查名結果MIN(TQD06)=>MIN(NVL(TQD09,TQD06)),TMQ_結果查詢改成模組PUB_GetTMQans
          'Modified by Lydia 2018/03/20 用TMQ20判斷是否已刪除明細(+And nvl(tmq20,'N') = 'N') 'Remove by Lydia 2018/03/21 影響速度
          strSql = "select 'V' V,TMQ19 讀,DECODE(TQA06,'1',DECODE(TQA13,NULL,'(非正常字)',TQA13),'2','(圖形查詢)') 文字1,DECODE(TQA06,'1',DECODE(TQA14,NULL,DECODE(TQA08,NULL,'','(非正常字)'),TQA14),'2','') 文字2," & _
                  "TQA04 客戶名稱,decode(v1c2," & PUB_GetTMQans("3", True) & ") 結果, TMQ03 組群,TMQ01,TMQ11,TMQ10,TMQ20,TMQ21 " & _
                   "FROM TMQAPP,trademarkquery,(select tqd02 v1c1, MIN(NVL(TQD09,TQD06)) v1c2 from tmqdetail group by tqd02) VT1 " & _
                   "WHERE TQA01=TMQ18(+) and tmq01=v1c1(+) AND NOT(TMQ03 IS NULL) " & _
                   IIf(TMQ_ReApp = False, "and tmq21 is null ", "") & _
                   IIf(TMQ_CtrRead = True, "AND V1C2<" & CNULL(TMQ_不查), "")
          'Modify By Sindy 2022/12/23
          If strTQC01 <> "" Then
            'Modified by Lydia 2023/08/18 +排除記錄
            strSql = strSql & "and tmq01 in (select tqc03 from tmqcasemap where tqc01='" & strTQC01 & "' and tqc03<>'" & cntTQC自動記錄 & "') "
          Else
            strSql = strSql & "and tqa01 in (" & strA1 & ") "
          End If
          '2022/12/23 END
          'Modified by lydia 2016/04/18 改依日期降冪(最新的再上面)
          strSql = strSql & " ORDER BY TMQ18 desc,TMQ01 desc "
       End If 'Added by Lydia 2024/11/11
       inA = 1
       Set rsAD = ClsLawReadRstMsg(inA, strSql)
       If inA = 1 Then
          Set GrdTMQ.Recordset = rsAD
          SetGrdTMQ (rsAD.RecordCount + 1)
          GrdTMQ.FixedCols = 4
       End If
    End If
End Sub
Private Sub SetGrdTMQ(Optional ByVal iR As Integer = 2)
   Dim arrGridHeadText, arrGridHeadWidth
   Dim iRow As Integer
   
   'Added by Lydia 2024/11/11 查名單(網中)
   If m_UseTmqTma = "2" Then
      arrGridHeadText = Array("V", "文字檢索", "圖形檢索", "客戶名稱", "文字結果", "圖形結果", "類別組群", "智權備註", "查名單號")
      arrGridHeadWidth = Array(200, 860, 860, 860, 860, 860, 1000, 1000, 1000)
   Else
   'end 2024/11/11
      arrGridHeadText = Array("V", "讀", "文字1", "文字2", "客戶名稱", "結果", "組群", "委查單號", "TMQ11", "TMQ10", "TMQ20", "TMQ21")
      arrGridHeadWidth = Array(200, 300, 860, 860, 800, 800, 860, 1000, 0, 0, 0, 0)
   End If 'Added by Lydia 2024/11/11
   GrdTMQ.Visible = False
   GrdTMQ.Cols = UBound(arrGridHeadText) + 1
   GrdTMQ.Rows = iR
   With GrdTMQ
        For iRow = 0 To .Cols - 1
           .row = 0
           .col = iRow
           .Text = arrGridHeadText(iRow)
           .ColWidth(iRow) = arrGridHeadWidth(iRow)
           .CellAlignment = flexAlignCenterCenter
        Next
        For intI = 1 To iR - 1
          .row = intI
          For iRow = 0 To 3
            .col = iRow
            .CellBackColor = QBColor(15)
          Next iRow
        Next intI
   End With
   GrdTMQ.Visible = True

End Sub
'end 2015/10/14

'Add by Amy 2015/11/13 +L及CFL 且案件性質選7501判決分析,開放Frame41欄位
Private Sub SetFrame41()
   'm_strCaseCPM = GetAllCaseCPM() 'Add By Sindy 2022/8/29 取得案件性質代碼
   
   Frame41.Visible = False
   If (Text1(6) = "L" Or Text1(6) = "CFL") And ( _
       Left(Trim(Combo1(1)), 4) = "7501" Or InStr(m_strCaseCPM, "7501") > 0 _
      ) Then
'        (Left(Trim(Combo1(2)), 4) = "7501" And Trim(Combo1(2)) <> "") Or _
'       (Left(Trim(Combo1(3)), 4) = "7501" And Trim(Combo1(3)) <> "") Or (Left(Trim(Combo1(4)), 4) = "7501" And Trim(Combo1(4)) <> "")
      Frame41.Visible = True
   Else
      Option9(0).Value = False
      Option9(1).Value = False
   End If
End Sub

'Add by Amy 2016/06/06 +專利設計案案件性質
Private Sub SetCombo5_P()
    If InStr(Text1(6), "P") = 0 Then Exit Sub
    If bolNotChk = True Then Exit Sub '按下列印時不執行
    
    '當專利種類由1、2變3 或3變1、2 時重抓案件屬性
    If ((Combo6.Tag = MsgText(601) Or Left(Combo6.Tag, 1) = "1" Or Left(Combo6.Tag, 1) = "2") And Left(Combo6, 1) = "3") _
      Or (Left(Combo6.Tag, 1) = "3" And (Left(Combo6, 1) = "1" Or Left(Combo6, 1) = "2")) Then
        Call PUB_AddCaseAttributeCombo(Combo5, Left(Combo6, 1)) '專利案件屬性選單 Modify By Sindy 2020/3/10
'        Combo5.Clear
'        If Left(Combo6, 1) = "3" Then
'            Combo5.AddItem ""
'            Combo5.AddItem "1.整體"
'            Combo5.AddItem "2.部分"
'            Combo5.AddItem "3.圖像"
'            Combo5.AddItem "4.成組"
'        Else
'            Combo5.AddItem ""
'            Combo5.AddItem "1.機械"
'            Combo5.AddItem "2.電子電機"
'            Combo5.AddItem "3.化學生醫"
'        End If
    End If
End Sub

Private Function ChkIsNumeric(ByVal stZip As String) As Boolean
    Dim j As Integer
    
    ChkIsNumeric = False
    If Trim(stZip) = MsgText(601) Then Exit Function
    
    For j = 1 To Len(stZip)
        If Not IsNumeric(Mid(stZip, j, 1)) Then Exit Function
    Next j
    ChkIsNumeric = True
End Function

'Added by Morgan 2016/7/20
Private Sub Combo1_Change(Index As Integer)
   If Index = 0 Then
      stCountry = ""
   End If
End Sub

'Add by Amy 2016/09/01 卷宗性質
Private Function GetTM28() As String
    Dim RsQ As New ADODB.Recordset
    Dim strQ As String, intQ As Integer
    
    GetTM28 = ""
    strQ = "Select TM28 From TradeMark Where tm01='" & Text1(6) & "' And tm02='" & Text1(7) & "' And tm03='" & Text1(8) & "' And tm04='" & Text1(9) & "' "
    intQ = 1
    Set RsQ = ClsLawReadRstMsg(intQ, strQ)
    If intQ = 1 Then
       GetTM28 = "" & RsQ.Fields("TM28")
    End If
    Set RsQ = Nothing
End Function

'Add by Amy 2018/08/09 以承辦人抓其主管 st52且在職,抓到有值或 st52=null 為止
Private Function GetAllST52(ByVal stCP14 As String, ByRef m_CP14) As Boolean
    Dim RsQ As New ADODB.Recordset
    Dim strQ As String, strST52 As String
    Dim intQ As Integer
    
    GetAllST52 = False
    strQ = "Select st01,st52,st04 From Staff Where st01="
    Do While stCP14 <> MsgText(601)
        intQ = 1
        Set RsQ = ClsLawReadRstMsg(intQ, strQ & "'" & stCP14 & "'")
        If intQ = 1 Then
            '離職
            If "" & RsQ.Fields("st04") = "2" Then
                'st52 為null 會離開回圈
                stCP14 = "" & RsQ.Fields("st52")
            Else
                stCP14 = "": m_CP14 = "" & RsQ.Fields("st01")
                GetAllST52 = True
            End If
        End If
    Loop
    Set RsQ = Nothing
End Function

'Added by Lydia 2018/12/10 檢查匯入的查名單是否齊備
Private Function ChkTMQList(ByVal mList As String) As Boolean
Dim stCon As String
Dim stMid As String

    If Frame42.Visible = True Then
       ChkTMQList = True
       If Option6(3).Value = True And OptCP143(0).Value = True Then '尚待查名+查名已齊備
           MsgBox "尚待查名和查名已齊備兩者結果為相反 !"
           GoTo ExitProc
       ElseIf mList <> "" And (OptCP143(0).Value = True Or Option6(0).Value = True Or Option6(1).Value = True Or Option6(2).Value = True) Then   '匯入查名單+查名已齊備/已查名...
            stMid = PUB_TMQchkCP143(mList)
            If stMid = "N" Then
               If OptCP143(0).Value = True Then
                   stCon = "查名已齊備"
                   OptCP143(0).Value = False
               ElseIf Option6(0).Value = True Then
                   stCon = Option6(0).Caption
               ElseIf Option6(1).Value = True Then
                   stCon = Option6(1).Caption
               ElseIf Option6(2).Value = True Then
                   stCon = Option6(2).Caption
               End If
               If stCon <> "" Then
                     MsgBox "勾選的查名單尚未查覆完畢，請勿勾選: " & stCon & " !"
                     GoTo ExitProc
               End If
            End If
       End If
    Else
       ChkTMQList = True
    End If
    
    Exit Function
    
ExitProc:
SSTab1.Tab = 2
SSTab3.Tab = 0
ChkTMQList = False

End Function

'Added by Lydia 2019/04/10 提醒及列印顧問服務件數
Private Sub SetLAdata()
'填寫接洽單時，若輸入為LA之舊案時，
'在本所案號欄跳離時讀取下列資訊顯示訊息給操作者，
'同時列印在接洽單上。
'抓該案號最大收文日且未取消收文的顧問聘任(CP10='0')
'依其聘任期間(CP53~CP54)，再抓該案號於聘任期間A類且未收費之收文，
'累計次數及工作時數(CP113)顯示訊息：
'聘任期間：民國年/月/日~民國年/月/日, 已服務  次,  工時(PS: 工時為0時不顯示)
Dim rsB As New ADODB.Recordset
Dim intB As Integer
Dim strB1 As String

On Error GoTo ErrHandle

m_LAmsg = ""

If Option1(1).Value = True And Text1(6).Text = "LA" And Len(Trim(Text1(7).Text)) = 6 Then
     strB1 = "select c1.cp09,substr(sqldatet(c1.cp53),1,9) cp53t,substr(sqldatet(c1.cp54),1,9) cp54t,sum(x02) x02,sum(x03) x03 from caseprogress c1" & _
                  ",(select cp09 as x00,cp05 as x01, 1 as x02,nvl(cp113,0) as x03 from caseprogress where cp01='" & Text1(6) & "' and cp02='" & Text1(7) & "' and cp03='" & Left(Text1(8) & "0", 1) & "' and cp04='" & Left(Text1(9) & "00", 2) & "' and substr(cp09,1,1)='A' and nvl(cp18,0)=0) x1 " & _
                  "where c1.cp09 in ( " & _
                  "select substr(max(cp05||cp09),9,9) mno from caseprogress where cp01='" & Text1(6) & "' and cp02='" & Text1(7) & "' and cp03='" & Left(Text1(8) & "0", 1) & "' and cp04='" & Left(Text1(9) & "00", 2) & "' and cp10='0' and cp158=0 and cp159=0 " & _
                   ") and x01>=c1.cp53 and x01<=c1.cp54 group by c1.cp09,substr(sqldatet(c1.cp53),1,9) ,substr(sqldatet(c1.cp54),1,9) "
    intB = 1
    Set rsB = ClsLawReadRstMsg(intB, strB1)
    If intB = 1 Then
        'Modified by Lydia 2019/08/12 + ""
        If "" & rsB.Fields("cp53t") <> "" And "" & rsB.Fields("cp54t") <> "" Then
            m_LAmsg = "聘任期間：" & rsB.Fields("cp53t") & "~" & rsB.Fields("cp54t")
            m_LAmsg = m_LAmsg & ", 已服務 " & rsB.Fields("x02") & " 次"
            If Val("" & rsB.Fields("x03")) > 0 Then m_LAmsg = m_LAmsg & ", 工時： "
            MsgBox m_LAmsg, vbInformation, "顧問服務件數"
            'Added by Lydia 2021/02/22 加入備註; ex. LA-003219於109/1/9已有顧問聘任期間(輸錯109/2/1~111/1/31)，又於110/1/20重複顧問聘任期間
            If InStr(Text1(119), m_LAmsg) = 0 Then
                Text1(119) = Text1(119) & IIf(Text1(119) <> "", vbCrLf & vbCrLf, "") & m_LAmsg & vbCrLf
            End If
            'end 2021/02/22
        End If
    End If
    
    Set rsB = Nothing
End If

'Added by Lydia 2019/08/13
    Exit Sub
    
ErrHandle:
    If Err.Number <> 0 Then
        MsgBox Err.Description, "提醒顧問服務件數" '及列印
        Resume Next
    End If
'end 2019/08/13
End Sub

'Added by Lydia 2019/08/06 ACS案之接洽單案件屬性
Private Sub SetFrame44()

   Frame44.Visible = False
   If Text1(6) = "ACS" And Option1(0).Value = True Then
       Frame44.Visible = True
       Frame19.Visible = False
'       If Me.Combo7.ListCount = 0 Then
'          PUB_SetCasePTM "4", , Me.Combo7
'       End If
   End If
End Sub

'Add by Amy 2019/12/04 傳入收據公司別名稱,回傳對應之編號
Private Function GetReceiptVal(ByVal stReceiptName As String, ByVal bolSql As Boolean) As String
    'Added by Lydia 2020/03/30
    If strSrvDate(1) >= 智慧所更名日 Then
        Select Case stReceiptName
            Case m_CompName1
                GetReceiptVal = "1"
            Case m_CompName2
                'Modified by Lydia 2022/09/05 因為預設2公司,所以不用回寫=空白
                'GetReceiptVal = "2"
                GetReceiptVal = ""
            Case m_CompNameJ
               GetReceiptVal = "J"
            Case m_CompNameL
               GetReceiptVal = "L"
            Case Else
                GetReceiptVal = ""
        End Select
    Else
    'end 2020/03/30
        Select Case stReceiptName
            Case "專利商標"
                GetReceiptVal = "1"
            Case "專利法律"
                GetReceiptVal = "2"
            'Modify by Amy 2016/07/27
            Case "智權公司"
               GetReceiptVal = "J"
            Case Else
                GetReceiptVal = ""
        End Select
    End If
    
    'Removed by Morgan 2020/7/1 不管有無編號，回傳形式應該要相同，否則無編號時會更新成單引號 Ex:X25140
    'If bolSql = True And GetReceiptVal = MsgText(601) Then
    '    GetReceiptVal = "'" & GetReceiptVal & "'"
    'End If
    'end 2020/7/1
End Function

'Added by Lydia 2020/03/30 設定收據公司別(簡稱)
Private Sub SetCombo4(Optional ByVal pStartDay As String)
'pStartDay  : 可傳入接洽單日期CRL02
Dim strB As String, strB2 As String
Dim intB As Integer
Dim rsB As New ADODB.Recordset
Dim m_Comp  As String
     
     If pStartDay = "" Then pStartDay = strSrvDate(1) '預設系統日
     
     '記錄收據公司(簡稱)
     If m_CompName2 = "" Then
        If pStartDay >= 智慧所更名日 Then
            m_CompName1 = CompNameQuery("1", "4")
            m_CompName2 = CompNameQuery("2", "4")
            m_CompNameJ = CompNameQuery("J", "4")
            m_CompNameL = CompNameQuery("L", "4")
        Else
            m_CompName1 = "專利商標"
            m_CompName2 = "專利法律"
            m_CompNameJ = "智權公司"
            m_CompNameL = "L"
        End If
     End If

    '收據公司欄自智慧所更名日起，下拉選單預設僅以2抓公司檔簡稱，智權公司的預設也以J抓公司檔簡稱
    If pStartDay < 智慧所更名日 Then  '1+2公司
          strB = "1,2,J"
    Else
          '其他:
          strB = "2,J"  '抓2+J公司
         '新案以系統類別抓系統種類對照表，若SK02為3、4(法務案)及7、8(顧問案)的收據公司別鎖住，且不更新回客戶檔
        If Text1(6).Text <> "" Then
            'Modified by Lydia 2020/04/10 '請改為instr(系統類別,'L')>0的收據公司別鎖住，且不更新回客戶檔；instr(系統類別,'L')=0不可選L公司；
            'strB2 = "select sk02 from systemkind where sk01='" & Text1(6).Text & "' "
            'intB = 1
            'Set rsB = ClsLawReadRstMsg(intB, strB2)
            'If intB = 1 Then
            '    If InStr("3,4,7,8", "" & rsB.Fields("sk02")) > 0 Then
                If InStr(UCase(Text1(6)), "L") > 0 Then
            'end 2020/04/10
                    If Option1(0).Value = True Then '新案限制為L公司
                        m_Comp = "L"
                    Else   '舊案=L
                        strB = "L"
                    End If
                End If
            'End If 'Mark by Lydia 2020/04/10
        End If
    End If

    strB = "select a0801,a0820 from acc080 where a0801 in (" & GetAddStr(IIf(m_Comp <> "", m_Comp, strB)) & ") order by a0801 "
    intB = 1
    Set rsB = ClsLawReadRstMsg(intB, strB)
    If intB = 1 Then
      Combo4.Clear
      intB = 0
      Combo4.AddItem "", intB
      rsB.MoveFirst
      Do While Not rsB.EOF
          intB = intB + 1
          'Combo4.AddItem "" & rsB.Fields("a0820"), intB
          Select Case "" & rsB.Fields("a0801")
               Case "1"
                   Combo4.AddItem m_CompName1, intB
                   m_Comp1forIdx = intB
               Case "2"
                   Combo4.AddItem m_CompName2, intB
                   m_Comp2forIdx = intB
               Case "J"
                   Combo4.AddItem m_CompNameJ, intB
                   m_CompJforIdx = intB
               Case "L"
                   Combo4.AddItem m_CompNameL, intB
                   m_CompLforIdx = intB
          End Select
          rsB.MoveNext
      Loop
    End If
    '新案以系統類別抓系統種類對照表，若SK02為3、4(法務案)及7、8(顧問案)的收據公司別鎖住，且不更新回客戶檔
    'Memo by Lydia 2020/04/10 請改為instr(系統類別,'L')>0的收據公司別鎖住，且不更新回客戶檔；instr(系統類別,'L')=0不可選L公司；
    If m_Comp = "L" Then
         Combo4.ListIndex = m_CompLforIdx
         Combo4.Enabled = False
    Else
         Combo4.Enabled = True
    End If
    
    'Added by Lydia 2021/03/29 ACS案件收文與點數及營業稅：勾選新案時，將收據公司欄鎖住並預設為J公司(智權公司)
    If strSrvDate(1) >= strACSdate1 Then
        If Text1(6) = "ACS" Then
            If Option1(0).Value = True Then
                Combo4.ListIndex = m_CompJforIdx
                Combo4.Enabled = False
            Else
                Combo4.Enabled = True
            End If
        End If
    End If
    'end 2021/03/29
End Sub

'Added by Morgan 2020/4/17
'Modify By Sindy 2022/10/6
Private Sub SrcSetButton()
Dim intRow As Integer
Dim bolLosCase As Boolean 'Add By Sindy 2022/10/6

   '預設
   If Text5 = "" Then
      LblText5.Visible = False
      Text5.Visible = False
   Else
      LblText5.Visible = True
      Text5.Visible = True
   End If
   
   If m_blnCallPrint = True Then '僅查詢
      cmdSend.Visible = False
      Exit Sub
   End If
   
   Me.Enabled = False
   cmdSend.Caption = "待收文區"
   cmdSend.Visible = False
   cmdOK(1).Enabled = True '可操作結束
   
   cmdOK(5).Visible = False
   cmdOK(0).Visible = True
   
   bolLosCase = False
   If InStr(UCase(Text1(6).Text), "L") > 0 And PUB_ChkLCompStaff(Text1(10).Text) = False Then
      bolLosCase = True
   End If
   
   '外部呼叫的收文,直接送出
   If bolExternalCall = True Or cmdOK(0).Caption = "送出" Then
      cmdOK(0).Caption = "送出"
      cmdOK(2).Visible = False
   Else
      intRow = SrcGetCRLFlow
      If intRow > 0 Then
         cmdSend.Caption = "待收文區(" & intRow & ")"
         cmdSend.Visible = True
         cmdOK(1).Enabled = False '不可操作結束
      End If
      If Text5 = "" Then
         cmdOK(0).Caption = "新增"
         cmdOK(2).Visible = True
      Else
         If UCase(TypeName(m_PrevForm)) = UCase("frm210148") Then
            cmdOK(0).Visible = False
         ElseIf UCase(TypeName(m_PrevForm)) = UCase("frm210147") And m_F0309 = Flow_退回 Then
            If bolLosCase = True Then m_blnCallPrint = True '案源不能修改資料
            cmdOK(0).Caption = "重送"
         Else
            If bolLosCase = True Then
               m_blnCallPrint = True '案源不能修改資料
               MsgBox "案源資料已寫入，不可修改！", vbInformation
'            ElseIf mTQC01 <> "" Then '有查名代號了,不可修改存檔
'               m_blnCallPrint = True
'               Call SetCtrlReadOnly(True)
'               MsgBox "有查名代號了，不可修改！", vbInformation
            End If
            
            cmdOK(0).Caption = "存檔"
         End If
         cmdOK(2).Visible = False
      End If
      If m_blnCallPrint = True Then If cmdOK(0).Caption <> "重送" Then cmdOK(0).Visible = False
   End If
   
   If bolLosCase = True Then
      If m_blnCallPrint = True Then
         cmdOK(5).Caption = "案源(&I)"
      Else
         cmdOK(5).Caption = "案源輸入(&I)"
      End If
      cmdOK(5).Visible = True
      If cmdOK(0).Caption <> "重送" Then cmdOK(0).Visible = False
   End If
   
   If m_PrevForm Is Nothing = False Then
      '查詢狀況
      cmdSend.Visible = False
      cmdOK(1).Enabled = True '可操作結束
   End If
   Me.Enabled = True
End Sub

'Added by Morgan 2020/4/22
'Modified by Morgan 2020/7/23 pIsLCase2:是否為法律案接洽單2
Public Sub Load4Print(pLOS15 As String, pCRL01 As String, Optional pIsLCase As Boolean = True, Optional pIsLCase2 As Boolean = False)
   strLOS15 = pLOS15
   
   '列印
   If pCRL01 <> "" Then
      bolIsTmp = True 'Added by Morgan 2020/8/3
      Text5 = pCRL01
      cmdOK(4).Value = True
      m_CP05 = ""
      m_CP09 = ""
      
      If pIsLCase = True Then
         Text1(10) = strUserNum
         cmdOK(5).Visible = True
      'Added by Morgan 2022/4/25
      'P/T案接洽單無費用不印特殊收據
      ElseIf Check9.Value = 1 Then
         Check9.Value = 0
      'end 2022/4/25
      End If
            
      cmdOK(0).Caption = "列印(&P)　　份"
      cmdOK(0).Visible = True
      cmdOK(2).Visible = False
   '自行填法務案接洽單
   Else
      SrcLoadLCasebyLOS15 strLOS15
   End If
   
   'Modified by Morgan 2020/7/23
   'bolLawOfficeCase = True
   bolLawOfficeCase = pIsLCase
   bolLawOfficeCase2 = pIsLCase2
   'end 2020/7/23
End Sub

'Added by Morgan 2020/4/23
'更新接洽單列印時間
Private Sub SrcUpdPrintTime()
On Error GoTo ErrHnd
   'Modified by Morgan 2020/7/23 +los24列印人員
   If bolLawOfficeCase2 Then
      strSql = "update lawofficesource set los25=sysdate,los26='" & strUserNum & "' where los15='" & strLOS15 & "'"
   Else
      strSql = "update lawofficesource set los14=sysdate,los24='" & strUserNum & "' where los15='" & strLOS15 & "'"
   End If
   cnnConnection.Execute strSql, intI
ErrHnd:
   If Err.Number <> 0 Then MsgBox Err.Description, vbCritical
End Sub

'Added by Morgan 2020/5/22
'檢查是否0費用
Private Function SrcZoroFeeChk() As Boolean
   Dim ii As Integer, bCheck As Boolean
   bCheck = True
   For ii = 1 To GridCase.Rows - 1 '4
      'If Val(Text1(101 + 3 * (ii - 1))) <> 0 Then
      If Val(Format(Trim(GridCase.TextMatrix(ii, 2)), "###0")) <> 0 Then
         SSTab1.Tab = 0
         'Text1(101 + 3 * (ii - 1)).SetFocus
         bCheck = False
         Exit For
      'ElseIf Val(Text1(102 + 3 * (ii - 1))) <> 0 Then
      ElseIf Val(Format(Trim(GridCase.TextMatrix(ii, 3)), "###0")) <> 0 Then
         SSTab1.Tab = 0
         'Text1(102 + 3 * (ii - 1)).SetFocus
         bCheck = False
         Exit For
      'ElseIf Val(Text1(103 + 3 * (ii - 1))) <> 0 Then
      ElseIf Val(Format(Trim(GridCase.TextMatrix(ii, 4)), "###0.000")) <> 0 Then
         SSTab1.Tab = 0
         'Text1(103 + 3 * (ii - 1)).SetFocus
         bCheck = False
         Exit For
      End If
   Next
   If bCheck = False Then
      If Left(strLSourceType, 1) = "B" Then
         MsgBox "【B類】案源接洽單之費用、規費、點數都必須為 0 !!" & vbCrLf & "若欲收文有費用的案件性質，請分不同接洽單填寫!!", vbExclamation
      'Removed by Morgan 2021/7/19 取消 C類案源
      'ElseIf strLSourceType = "C" Then
      '   MsgBox "【C類】案源法務案之費用、規費、點數都必須為 0 !!", vbExclamation
      'end 2021/7/19
      ElseIf strLCaseNo(1) = "L" Then
         MsgBox "【B類】案源補收文之費用、規費、點數都必須為 0 !!", vbExclamation
      End If
   End If
   SrcZoroFeeChk = bCheck
End Function

'Added by Morgan 2020/5/22
'自動新增【提供書狀意見】
Private Sub SrcAutoAdd()
Dim strCP10 As String
Dim ii As Integer, jj As Integer
    
   m_strCaseCPM = GetAllCaseCPM() 'Add By Sindy 2022/8/29 取得案件性質代碼
   
   If Trim(Text1(6).Text) = "P" Then
      strCP10 = "225"
   Else
      strCP10 = "212"
   End If
'   jj = 0
'   For ii = 1 To 4
'      If Left(Combo1(ii), 3) = strCP10 Then
'         Exit For
'      ElseIf Combo1(ii) = "" And jj = 0 Then
'         jj = ii
'      End If
'   Next
   If InStr(m_strCaseCPM, strCP10) > 0 Then
      Exit Sub
   End If
'   If ii = 5 Then
'      If jj = 0 Then
'         MsgBox "案件性質超過，無法新增【提供書狀意見】！", vbExclamation
'      Else
'         For ii = 0 To Combo1(jj).ListCount - 1
'            If Left(Combo1(jj).List(ii), 3) = strCP10 Then
'               Combo1(jj).ListIndex = ii
'               Combo1_LostFocus jj
'               Text1(101 + 3 * (jj - 1)).Text = "0"
'               Text1(102 + 3 * (jj - 1)).Text = "0"
'               Text1(103 + 3 * (jj - 1)).Text = "0.000"
'               Exit For
'            End If
'         Next
         For ii = 0 To Combo1(1).ListCount - 1
            If Left(Combo1(1).List(ii), 3) = strCP10 Then
               Combo1(1).ListIndex = ii
               Combo1_LostFocus 1
'               Text1(101).Text = "0"
'               Text1(102).Text = "0"
'               Text1(103).Text = "0.000"
               Exit For
            End If
         Next
'      End If
'   End If
End Sub
'Added by Morgan 2020/7/13
'自動新增【準備程序】【言詞辯論】
Private Sub SrcAutoAdd2(pCP10s As String)
Dim arrCP10() As String, strCP10 As String, strCP10n As String
Dim ii As Integer, jj As Integer, kk As Integer
   
   m_strCaseCPM = GetAllCaseCPM() 'Add By Sindy 2022/8/29 取得案件性質代碼
   
   arrCP10() = Split(pCP10s, ",")
   
   For kk = LBound(arrCP10) To UBound(arrCP10)
      strCP10 = arrCP10(kk)
      strCP10n = ""
      For ii = 0 To Combo1(1).ListCount - 1
         If Left(Combo1(1).List(ii), 3) = strCP10 Then
            strCP10n = Mid(Combo1(1).List(ii), 5)
            Exit For
         End If
      Next
      If strCP10 <> "" Then
'         jj = 0
'         For ii = 1 To 4
'            If Left(Combo1(ii), 3) = strCP10 Then
'               Exit For
'            ElseIf Combo1(ii) = "" And jj = 0 Then
'               jj = ii
'            End If
'         Next
         If InStr(m_strCaseCPM, strCP10) > 0 Then
            Exit Sub
         End If
'         If ii = 5 Then
'            If jj = 0 Then
'               MsgBox "案件性質超過，無法新增【" & strCP10n & "】！", vbExclamation
'               Exit For
'            Else
'               For ii = 0 To Combo1(jj).ListCount - 1
'                  If Left(Combo1(jj).List(ii), 3) = strCP10 Then
'                     Combo1(jj).ListIndex = ii
'                     Combo1_LostFocus jj
'                     Text1(101 + 3 * (jj - 1)).Text = "0"
'                     Text1(102 + 3 * (jj - 1)).Text = "0"
'                     Text1(103 + 3 * (jj - 1)).Text = "0.000"
'                     Exit For
'                  End If
'               Next
               For ii = 0 To Combo1(1).ListCount - 1
                  If Left(Combo1(1).List(ii), 3) = strCP10 Then
                     Combo1(1).ListIndex = ii
                     Combo1_LostFocus 1
'                     Text1(101).Text = "0"
'                     Text1(102).Text = "0"
'                     Text1(103).Text = "0.000"
                     Exit For
                  End If
               Next
'            End If
'         End If
      End If
   Next
End Sub

'Added by Morgan 2020/4/28
'B1類智財民刑事:P226/T213配合開庭，自動補上提供書狀意見(P225/T212)，兩項費用規費均設0 <--案件性質抓特殊設定 B1P, B1T
'B2類商標行政訴訟/專利上訴...，費用0: <--案件性質抓特殊設定 B2P, B2T
'C類專利行政訴訟(含參加)...: <--案件性質抓特殊設定 CP, CT --取消 2021/7/13
Private Function SrcSetLCase(Optional pSaveCheck As Boolean = False, Optional pIndex As Integer = -1) As Boolean
   Dim ii As Integer
   Dim bCheck As Boolean
   Dim strCode As String, strCode2 As String, strTemp As String, arrCP10() As String
   Dim strChuTing As String, bolChuTing As Boolean
   Dim strSuYuan As String
   Dim strAutoAdd2Pty As String
   
   If strLOS15 <> "" Then SrcSetLCase = True: Exit Function
   
   m_strCaseCPM = GetAllCaseCPM 'Add By Sindy 2022/9/12 取得案件性質代碼
   strLSourceType = ""
   bolSuYuan = False
   bolIsB2CourtFee = False
   Erase strLCaseNo
   If Left(Combo1(0), 3) = "000" And (Trim(Text1(6).Text) = "P" Or Trim(Text1(6).Text) = "T" Or Trim(Text1(6).Text) = "TC") Then
   
      '設定案件性質(準備程序,言詞辯論,訴願)
      If Trim(Text1(6).Text) = "P" Then
         strChuTing = "211,212"
         'Modified by Morgan 2021/7/13 專利訴願不再是案源(改專利師處理不需律師)
         'strSuYuan = "501,505"
         strSuYuan = ""
         'end 2021/7/13
         strAutoAdd2Pty = "503,506" 'Modified by Morgan 2022/11/11 +506
      Else
         strChuTing = "204,205"
         strSuYuan = "401,406"
         strAutoAdd2Pty = "403" '行政訴訟
      End If
      
      m_strCaseCPM = m_strCaseCPM & IIf(Trim(Left(Trim(Combo1(1)), 4)) <> "", "," & Trim(Left(Trim(Combo1(1)), 4)), "")
      'If Left(m_strCaseCPM, 1) = "," Then m_strCaseCPM = Mid(m_strCaseCPM, 2)
      arrCP10 = Split(m_strCaseCPM, ",")
      'For ii = 1 To 4
      For ii = 0 To UBound(arrCP10) '- 1
         strTemp = arrCP10(ii) 'Left(Combo1(ii), 3)
         If strTemp <> "" Then
            '是否有收準備程序,言詞辯論
            If InStr(strChuTing, strTemp) > 0 Then bolChuTing = True
            '是否有收訴願
            If InStr(strSuYuan, strTemp) > 0 Then bolSuYuan = True
            
            If Trim(Text1(6).Text) = "P" Then
               strCode2 = PUB_GetLOSkind("P", strTemp)
            Else
               strCode2 = PUB_GetLOSkind("T", strTemp)
            End If
            If strCode2 <> "" Then
               If strCode = "" Then
                  strCode = strCode2
               ElseIf strCode <> strCode2 Then
                  MsgBox "不可同時收文" & Left(strCode, Len(strCode) - 1) & "類及" & Left(strCode2, Len(strCode2) - 1) & "類案源的案件性質！", vbExclamation
                  Exit Function
               End If
            End If
         End If
      Next ii
      'B1
      If Left(strCode, 2) = "B1" Then
         strLSourceType = "B1"
         If pSaveCheck = True Then
            If SrcZoroFeeChk() = False Then Exit Function
            
         ElseIf pIndex > -1 Then
         
            Text1(101).Text = "0" ' + 3 * (pIndex - 1)
            Text1(102).Text = "0" ' + 3 * (pIndex - 1)
            Text1(103).Text = "0.000" ' + 3 * (pIndex - 1)
            Call cmdUpd_Click '新增
            SrcAutoAdd '自動新增【提供書狀意見】
         End If
         
      'B2
      ElseIf Left(strCode, 2) = "B2" Then
         strLSourceType = "B2"
         If pSaveCheck = True Then
            If SrcZoroFeeChk() = False Then Exit Function
            
         ElseIf pIndex > -1 Then
         
            Text1(101).Text = "0" ' + 3 * (pIndex - 1)
            Text1(102).Text = "0" ' + 3 * (pIndex - 1)
            Text1(103).Text = "0.000" ' + 3 * (pIndex - 1)
            Call cmdUpd_Click '新增
            
            '新案收文行政訴訟自動新增【準備程序】【言詞辯論】
            'Modified by Morgan 2022/11/11
            'If Option1(0).Value = True And Left(Combo1(pIndex), 3) = strAutoAdd2Pty Then
'            If Option1(0).Value = True And InStr(strAutoAdd2Pty, Left(Combo1(pIndex), 3)) > 0 Then
            'end 2022/11/11
            'If Option1(0).Value = True And InStr(m_strCaseCPM, strAutoAdd2Pty) > 0 Then
            'Modify By Sindy 2022/11/11
            arrCP10 = Split(strAutoAdd2Pty, ",")
            For ii = 0 To UBound(arrCP10)
               strTemp = arrCP10(ii)
               If strTemp <> "" Then
                  'If Option1(0).Value = True And InStr(m_strCaseCPM, strAutoAdd2Pty) > 0 Then
                  If Option1(0).Value = True And InStr(m_strCaseCPM, strTemp) > 0 Then
                     SrcAutoAdd2 strChuTing
                     Exit For
                  End If
               End If
            Next ii
            '2022/11/11 END
         End If
         
      'C
      'Removed by Morgan 2021/7/19 取消 C類案源
      'ElseIf Left(strCode, 1) = "C" Then
      '   strLSourceType = "C"
      '   If pIndex > -1 Then
      '      '新案收文行政訴訟自動新增【準備程序】【言詞辯論】
      '      If Option1(0).Value = True And Left(Combo1(pIndex), 3) = strAutoAdd2Pty Then
      '         SrcAutoAdd2 strChuTing
      '      End If
      '   End If
      'end 2021/7/19
      
      '非案源但有收準備程序或言詞辯論
      ElseIf strLSourceType = "" And bolChuTing = True Then
         '本單有收訴願
         If bolSuYuan = True Then
            'Modified by Morgan 2021/7/19 商標訴願程序參加訴願言詞辯論由律師協助處理-法律所收文開收據
            'strLSourceType = "C"
            strLSourceType = "B2"
            'end 2021/7/19
         End If
         
      End If
      
      'Added by Morgan 2020/6/3
      '準備程序/言詞辯論是否補收出庭費
      If strLSourceType = "" And Option1(1).Value = True And bolChuTing = True And pSaveCheck = True Then
         'B2類案源補收文
         If SrcCheckExistLOS(Text1(6), Text1(7), Text1(8), Text1(9), strLCaseNo, "B2", strTemp) = True Then
            If strLCaseNo(1) = "L" Then '法務案要已收文否則會變成新案號
               If SrcZoroFeeChk() = False Then Exit Function
               If MsgBox("是否有【補收款】？" & vbCrLf & vbCrLf & "若【是】，請繼續操作法務案補收款接洽單", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
                  strLSourceType = "B2"
                  bolIsB2CourtFee = True
               End If
            End If
         
         '非C類案源補收文
         'Removed by Morgan 2021/7/19 取消 C類案源
         'ElseIf SrcCheckExistLOS(Text1(6), Text1(7), Text1(8), Text1(9), strLCaseNo, "C") = False Then
         '   '是否訴願案
         '   bolSuYuan = SrcChkIsSuYuan(Text1(6), Text1(7), Text1(8), Text1(9), strSuYuan)
         '   If bolSuYuan = True Then
         '      strLSourceType = "C"
         '   End If
         'end 2021/7/19
         End If
      End If
      'end 2020/6/3
   
   'Add By Sindy 2022/9/28
   ElseIf Text1(6) = "L" And Text1(7) = "888888" And Combo1(1).Text <> "" Then
         Text1(101).Text = "0" ' + 3 * (pIndex - 1)
         Text1(102).Text = "0" ' + 3 * (pIndex - 1)
         Text1(103).Text = "0.000" ' + 3 * (pIndex - 1)
         Call cmdUpd_Click '新增
         '2022/9/28 END
   End If
   
   SrcSetLCase = True
End Function

'Added by Morgan 2020/5/4
Private Function SrcGetLOS15(pCRL01 As String) As String
   Dim stSQL As String, intQ As Integer
   Dim RsQ As ADODB.Recordset
   
On Error GoTo ErrHnd
   
   stSQL = "select LOS15 From LawOfficeSource where LOS17='" & pCRL01 & "'"
   intQ = 1
   Set RsQ = ClsLawReadRstMsg(intQ, stSQL)
   If intQ = 1 Then
      SrcGetLOS15 = "" & RsQ("LOS15")
   End If
   
ErrHnd:
   Set RsQ = Nothing
End Function

'Added by Morgan 2020/5/6
'預設案件說明事項
'blnOverseas:是否外對台, bolChange:判斷有改過說明時不重新預設
Private Sub SrcSetMemo(Optional blnOverseas As Boolean = False, Optional bolChange As Boolean = False)
   If bolChange = True And Text1(119).Tag <> Text1(119) Then Exit Sub
   
   'Add By Sindy 2022/10/14
   If Text1(6) = "ACS" Then
      Text1(119).Text = ""
      '2022/10/14 END
      
   ElseIf InStr(UCase(Text1(6).Text), "L") > 0 Then
      If PUB_ChkLCompStaff(Text1(10).Text) = True And strLOS15 = "" And strSrvDate(1) < 法律所案源收文啟用日 Then
         Text1(119).Text = "智權人員："
      Else
         Text1(119).Text = ""
      End If
      
   Else
      'Modify By Sindy 2022/10/14 取消 "商品類別：" & vbCrLf "彼所案號：" & vbCrLf: 改存欄位
      Text1(119).Text = "如案件內容摘要、引用條文與客戶洽談要旨等等......" & vbCrLf & _
                        "商品名稱：" & vbCrLf & _
                        "註冊號數：" & vbCrLf & _
                        "優先權日：" & vbCrLf & _
                        "優先權號：" & vbCrLf
      If blnOverseas = True Then
         Text1(119).Text = Text1(119).Text & _
                        "聯絡人：" & vbCrLf '& _
                        '"彼所案號：" & vbCrLf
      End If
   End If
   
   Text1(119).Tag = Text1(119)
End Sub

Private Sub SrcSetField(pSys As String, Optional pNewCase As Boolean)
   
   'Modified by Morgan 2020/8/3 +C類也可能有法務舊案
   'Removed by Morgan 2021/7/19 取消 C類案源
   'If (bolIsB2CourtFee = True Or strLSourceType = "B1" Or strLSourceType = "B2" Or strLSourceType = "C") And strLCaseNo(1) <> "" Then
   If (bolIsB2CourtFee = True Or strLSourceType = "B1" Or strLSourceType = "B2") And strLCaseNo(1) <> "" Then
      '本所案號
      Option1(1).Value = True
      Text1(6).Text = strLCaseNo(1)
      Text1(7).Text = strLCaseNo(2)
      Text1(8).Text = strLCaseNo(3)
      Text1(9).Text = strLCaseNo(4)
      Text1_LostFocus 9
   Else
      'Option1(1).Value = True
      Option1(0).Value = True '預設新案
      '要先設國家再設系統別才會載入案件性質清單
      '國家
      'Modify By Sindy 2024/2/22
      'SetComboCase 0, "000"
      'SetComboCase 0, Combo1(0).Text
      Call frm090801_New_SetComboCase(0, "000", Combo1(0), Me.Text1(6).Text, _
                  Combo1(0).Text, ChkPCT, Text1(101), Text1(102), Text1(103))
      Call frm090801_New_SetComboCase(0, Combo1(0).Text, Combo1(0), Me.Text1(6).Text, _
                  Combo1(0).Text, ChkPCT, Text1(101), Text1(102), Text1(103))
      '2024/2/22 END
      '本所案號
      Text1(6).Text = pSys
      Text1_LostFocus 6
      Text1(7).Text = ""
      Text1(8).Text = ""
      Text1(9).Text = ""
      '主題
      'If pSys = "L" Then Text1(11) = ""
'      If Text1(11).Enabled Then Text1(11).SetFocus
   End If
   
   '鎖住欄位
   SrcLockField True, pNewCase
   
   If pSys = "L" Then
      If strPTSysCode = "P" Then Check3(0).Value = vbChecked
      If strPTSysCode = "T" Then Check3(1).Value = vbChecked
      If strPTSysCode = "TC" Then Check3(2).Value = vbChecked
      SrcSetButton
   End If
End Sub

'Added by Morgan 2020/5/6
'案源輸入欄位控制(BC輸法務案或B1補PT配合開庭)
Private Sub SrcLockField(pLocked As Boolean, Optional pNewPTCase As Boolean)
   
   If pLocked = True Then
      '舊案
      If Option1(1).Value = True Then
         If Text1(6) <> "L" Then
            Option1(0).Enabled = False
            Text1(7).Enabled = False
            Text1(8).Enabled = False
            Text1(9).Enabled = False
         End If
         
      '新案
      'P/T案為新案時,法務案也要是新案
      ElseIf Text1(6) = "L" And pNewPTCase = True Then
         Option1(1).Enabled = False
      End If
   Else
      Option1(0).Enabled = True
      Option1(1).Enabled = True
      If Option1(1).Value = True Then
         Text1(7).Enabled = pLocked
         Text1(8).Enabled = pLocked
         Text1(9).Enabled = pLocked
      End If
   End If
   
   Text1(6).Locked = pLocked '系統別
   Combo1(0).Locked = pLocked '申請國家
   Text1(10).Locked = pLocked '介紹人(智權人員)
End Sub
'Added by Morgan 2020/5/29
'以案源單號載入法務案接洽單預設資料
Private Sub SrcLoadLCasebyLOS15(pLOS15 As String)
   Dim stSQL As String, intQ As Integer
   Dim RsQ As ADODB.Recordset
   
   stSQL = "select LOS02,CP01,CP31,nvl(PA26,TM23) Cus1,nvl(PA27,TM78) Cus2" & _
      ",nvl(PA28,TM79) Cus3,nvl(PA29,TM80) Cus4,nvl(PA30,TM81) Cus5" & _
      " from LawOfficeSource,caseprogress,patent,trademark" & _
      " where los15='" & pLOS15 & "' and cp09(+)=los01" & _
      " and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04" & _
      " and tm01(+)=cp01 and tm02(+)=cp02 and tm03(+)=cp03 and tm04(+)=cp04"
   intQ = 1
   Set RsQ = ClsLawReadRstMsg(intQ, stSQL)
   If intQ = 1 Then
      With RsQ
      strLSourceType = "" & .Fields("LOS02")
      strPTSysCode = .Fields("CP01")
      'P/T是否新案
      If .Fields("CP31") = "Y" Then
         bolPTIsNew = True
      Else
         bolPTIsNew = False
      End If
      
      SrcSetField "L", bolPTIsNew
      strCUNo(1) = "" & .Fields("Cus1")
      strCUNo(2) = "" & .Fields("Cus2")
      strCUNo(3) = "" & .Fields("Cus3")
      strCUNo(4) = "" & .Fields("Cus4")
      strCUNo(5) = "" & .Fields("Cus5")
      SrcSetCustByVar
      
      End With
   End If
   Set RsQ = Nothing
End Sub

'Adde by Morgan 2020/5/29
Private Sub SrcSetCustByVar()
   '申請人1
   If strCUNo(1) <> "" Then
      Option31(1).Value = True
      Text1(12) = strCUNo(1)
   End If
   '申請人2
   If strCUNo(2) <> "" Then
      Option32(1).Value = True
      Text1(28) = strCUNo(1)
   End If
   '申請人3
   If strCUNo(3) <> "" Then
      Option33(1).Value = True
      Text1(44) = strCUNo(3)
   End If
   '申請人4
   If strCUNo(4) <> "" Then
      Option34(1).Value = True
      Text1(60) = strCUNo(4)
   End If
   '申請人5
   If strCUNo(5) <> "" Then
      Option35(1).Value = True
      Text1(76) = strCUNo(5)
   End If
'   SrcCustValidate
End Sub
'Added by Morgan 2020/5/6
'載入申請人資料 (參照QueryData)
Private Sub SrcLoadCustbyCRL01(pCRL01 As String)
   Dim stSQL As String, intQ As Integer
   Dim RsQ As ADODB.Recordset
   Dim intIndex As Integer, intIndex2 As Integer
   
   If pCRL01 = "" Then Exit Sub
   
   stSQL = "select * from ConsultRecApp C where CRA01='" & "" & pCRL01 & "' order by cra02 asc"
   intQ = 1
   Set RsQ = ClsLawReadRstMsg(intQ, stSQL)
   If intQ = 1 Then
      With RsQ
      .MoveFirst
      Do While Not .EOF
         '申請人1
         If .Fields("CRA02") = "1" Then
            Me.SSTab2.Tab = 0
            If "" & .Fields("CRA03") = "Y" Then Option31(0).Value = True Else Option31(1).Value = True
            If "" & .Fields("CRA25") = "Y" Then
               optCP811(0).Value = True
            ElseIf "" & .Fields("CRA25") = "N" Then
               optCP811(1).Value = True
            End If
         '申請人2
         ElseIf .Fields("CRA02") = "2" Then
            Me.SSTab2.Tab = 1
            If "" & .Fields("CRA03") = "Y" Then Option32(0).Value = True Else Option32(1).Value = True
            If "" & .Fields("CRA25") = "Y" Then
               optCP812(0).Value = True
            ElseIf "" & .Fields("CRA25") = "N" Then
               optCP812(1).Value = True
            End If
         '申請人3
         ElseIf .Fields("CRA02") = "3" Then
            Me.SSTab2.Tab = 2
            If "" & .Fields("CRA03") = "Y" Then Option33(0).Value = True Else Option33(1).Value = True
            If "" & .Fields("CRA25") = "Y" Then
               optCP813(0).Value = True
            ElseIf "" & .Fields("CRA25") = "N" Then
               optCP813(1).Value = True
            End If
         '申請人4
         ElseIf .Fields("CRA02") = "4" Then
            Me.SSTab2.Tab = 3
            If "" & .Fields("CRA03") = "Y" Then Option34(0).Value = True Else Option34(1).Value = True
            If "" & .Fields("CRA25") = "Y" Then
               optCP814(0).Value = True
            ElseIf "" & .Fields("CRA25") = "N" Then
               optCP814(1).Value = True
            End If
         '申請人5
         ElseIf .Fields("CRA02") = "5" Then
            Me.SSTab2.Tab = 4
            If "" & .Fields("CRA03") = "Y" Then Option35(0).Value = True Else Option35(1).Value = True
            If "" & .Fields("CRA25") = "Y" Then
               optCP815(0).Value = True
            ElseIf "" & .Fields("CRA25") = "N" Then
               optCP815(1).Value = True
            End If
         End If
         intIndex = .Fields("CRA02")
         If Not IsNull(.Fields("CRA04")) Then Text1(IIf(intIndex = 1, 198, IIf(intIndex = 2, 298, IIf(intIndex = 3, 398, IIf(intIndex = 4, 498, 598))))) = .Fields("CRA04")
         If Not IsNull(.Fields("CRA05")) Then Text1(IIf(intIndex = 1, 12, IIf(intIndex = 2, 28, IIf(intIndex = 3, 44, IIf(intIndex = 4, 60, 76))))) = .Fields("CRA05") & .Fields("CRA06")
         
         If Not IsNull(.Fields("CRA07")) Then
            intIndex2 = IIf(intIndex = 1, 21, IIf(intIndex = 2, 37, IIf(intIndex = 3, 53, IIf(intIndex = 4, 69, 85))))
            Text1(intIndex2) = .Fields("CRA07")
            Text1(intIndex2).Tag = Text1(intIndex2)
         End If
         If Not IsNull(.Fields("CRA08")) Then
            intIndex2 = IIf(intIndex = 1, 22, IIf(intIndex = 2, 38, IIf(intIndex = 3, 54, IIf(intIndex = 4, 70, 86))))
            Text1(intIndex2) = .Fields("CRA08")
            Text1(intIndex2).Tag = Text1(intIndex2)
         End If
         
         If Not IsNull(.Fields("CRA09")) Then Text1(IIf(intIndex = 1, 23, IIf(intIndex = 2, 39, IIf(intIndex = 3, 55, IIf(intIndex = 4, 71, 87))))) = .Fields("CRA09")
         If Not IsNull(.Fields("CRA10")) Then cboContact(intIndex) = .Fields("CRA10")
         If Not IsNull(.Fields("CRA11")) Then Text1(IIf(intIndex = 1, 34, IIf(intIndex = 2, 50, IIf(intIndex = 3, 66, IIf(intIndex = 4, 82, 98))))) = .Fields("CRA11")
         If Not IsNull(.Fields("CRA12")) Then Text1(IIf(intIndex = 1, 92, IIf(intIndex = 2, 93, IIf(intIndex = 3, 94, IIf(intIndex = 4, 95, 96))))) = .Fields("CRA12")
         If Not IsNull(.Fields("CRA13")) Then Text1(IIf(intIndex = 1, 14, IIf(intIndex = 2, 30, IIf(intIndex = 3, 46, IIf(intIndex = 4, 62, 78))))) = .Fields("CRA13")
         If Not IsNull(.Fields("CRA14")) Then Text1(IIf(intIndex = 1, 15, IIf(intIndex = 2, 31, IIf(intIndex = 3, 47, IIf(intIndex = 4, 63, 79))))) = .Fields("CRA14")
         If Not IsNull(.Fields("CRA15")) Then Text1(IIf(intIndex = 1, 16, IIf(intIndex = 2, 32, IIf(intIndex = 3, 48, IIf(intIndex = 4, 64, 80))))) = .Fields("CRA15")
         If Not IsNull(.Fields("CRA16")) Then Text1(IIf(intIndex = 1, 17, IIf(intIndex = 2, 33, IIf(intIndex = 3, 49, IIf(intIndex = 4, 65, 81))))) = .Fields("CRA16")
         If Not IsNull(.Fields("CRA17")) Then Text1(IIf(intIndex = 1, 19, IIf(intIndex = 2, 35, IIf(intIndex = 3, 51, IIf(intIndex = 4, 67, 83))))) = .Fields("CRA17")
         If Not IsNull(.Fields("CRA18")) Then Text1(IIf(intIndex = 1, 20, IIf(intIndex = 2, 36, IIf(intIndex = 3, 52, IIf(intIndex = 4, 68, 84))))) = .Fields("CRA18")
         If Not IsNull(.Fields("CRA19")) Then Text1(IIf(intIndex = 1, 120, IIf(intIndex = 2, 121, IIf(intIndex = 3, 122, IIf(intIndex = 4, 123, 124))))) = .Fields("CRA19")
         If Not IsNull(.Fields("CRA20")) Then Text1(IIf(intIndex = 1, 27, IIf(intIndex = 2, 43, IIf(intIndex = 3, 59, IIf(intIndex = 4, 75, 91))))) = .Fields("CRA20")
         If Not IsNull(.Fields("CRA21")) Then Text1(IIf(intIndex = 1, 125, IIf(intIndex = 2, 141, IIf(intIndex = 3, 157, IIf(intIndex = 4, 173, 189))))) = .Fields("CRA21")
         If Not IsNull(.Fields("CRA22")) Then Text1(IIf(intIndex = 1, 25, IIf(intIndex = 2, 41, IIf(intIndex = 3, 57, IIf(intIndex = 4, 73, 89))))) = .Fields("CRA22")
         If Not IsNull(.Fields("CRA23")) Then Text1(IIf(intIndex = 1, 26, IIf(intIndex = 2, 42, IIf(intIndex = 3, 58, IIf(intIndex = 4, 74, 90))))) = .Fields("CRA23")
         If Not IsNull(.Fields("CRA24")) Then Text1(IIf(intIndex = 1, 24, IIf(intIndex = 2, 40, IIf(intIndex = 3, 56, IIf(intIndex = 4, 72, 88))))) = .Fields("CRA24")
         .MoveNext
      Loop
      End With
'      SrcCustValidate
   End If
   Set RsQ = Nothing
End Sub

'Added by Morgan 2020/5/20
'檢查客戶是否與接洽單相同
Private Function SrcCheckSameCust(pCRA01 As String) As Boolean
   Dim ii As Integer, strID As String, strNameC As String, strNameE As String, strMsg As String
   
   If pCRA01 <> "" Then
      strExc(0) = "select * from ConsultRecApp where CRA01='" & "" & pCRA01 & "' order by cra02 asc"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         With RsTemp
         Do While Not .EOF
            ii = .Fields("cra02")
            strID = Text1(12 + 16 * (ii - 1))
            strNameC = Text1(21 + 16 * (ii - 1))
            strNameE = Text1(22 + 16 * (ii - 1))
            '有客戶編號
            If .Fields("cra05") & .Fields("cra06") <> "" Then
               If .Fields("cra05") & .Fields("cra06") <> strID Then
                  SSTab1.Tab = 1
                  SSTab2.Tab = ii - 1
                  MsgBox "申請人" & ii & "【" & strID & "】與 P/T案【" & .Fields("cra05") & .Fields("cra06") & "】不同，請確認！"
                  Exit Function
                  'Mark by Lydia 2021/09/07 (保留)法務舊案客戶和P/T案舊案客戶編號不一致，改成詢問；ex.P-074838用X29213000(個人), L-006229用X29213120(磐石)
                  '-------目前不啟用的原因：因為該案是去年案源未上線前收文，所以理論不應收；修改程式先保留
                  'If MsgBox("申請人" & ii & "【" & strID & "】與 P/T案【" & .Fields("cra05") & .Fields("cra06") & "】不同，請確認是否繼續？", vbExclamation + vbYesNo + vbDefaultButton2) = vbNo Then
                  '    Exit Function
                  'End If
                  'end 2021/09/07
               End If
            '中文名稱
            ElseIf "" & .Fields("CRA07") <> "" Then
               If .Fields("CRA07") <> strNameC Then
                  SSTab1.Tab = 1
                  SSTab2.Tab = ii - 1
                  MsgBox "申請人" & ii & "【" & strNameC & "】與 P/T案【" & .Fields("cra07") & "】不同，請確認！"
                  Exit Function
               End If
            '英文名稱
            ElseIf "" & .Fields("CRA08") <> "" Then
               If .Fields("CRA08") <> strNameE Then
                  SSTab1.Tab = 1
                  SSTab2.Tab = ii - 1
                  MsgBox "申請人" & ii & "【" & strNameE & "】與 P/T案【" & .Fields("cra08") & "】不同，請確認！"
                  Exit Function
               End If
            End If
            .MoveNext
         Loop
         End With
         
         If ii < 5 Then
            ii = ii + 1
            strNameC = Text1(21 + 16 * (ii - 1))
            strNameE = Text1(22 + 16 * (ii - 1))
            If strNameC <> "" Then
               SSTab1.Tab = 1
               SSTab2.Tab = ii - 1
               MsgBox "P/T案無申請人" & ii & "【" & strNameC & "】，請確認！"
               Exit Function
            ElseIf strNameE <> "" Then
               SSTab1.Tab = 1
               SSTab2.Tab = ii - 1
               MsgBox "P/T案無申請人" & ii & "【" & strNameE & "】，請確認！"
               Exit Function
            End If
         End If
      End If
      
   ElseIf strLOS15 <> "" Then
         
      For ii = 1 To 5
         strID = Text1(12 + 16 * (ii - 1))
         strNameC = Text1(21 + 16 * (ii - 1))
         '舊客戶
         If strID <> "" Then
            If strID <> strCUNo(ii) Then
               SSTab1.Tab = 1
               SSTab2.Tab = ii - 1
               If strCUNo(ii) <> "" Then
                  If Text1(6) = "L" Then
                     strExc(1) = "P/T"
                  Else
                     strExc(1) = "法務"
                  End If
                  strMsg = "申請人" & ii & "(" & strID & ")與" & strExc(1) & "案(" & strCUNo(ii) & ")不同，請確認！"
                  MsgBox strMsg, vbExclamation
               Else
                  MsgBox strExc(1) & "案並無申請人" & ii & "，請確認！", vbExclamation
               End If
               Exit Function
            End If
            
         '新客戶
         ElseIf strNameC <> "" Then
            SSTab1.Tab = 1
            SSTab2.Tab = ii - 1
            If strCUNo(ii) = "" Then
               MsgBox "申請人" & ii & "與" & strExc(1) & "案(" & strCUNo(ii) & ")不同，請確認！", vbExclamation
            Else
               MsgBox strExc(1) & "案並無申請人" & ii & "，請確認！", vbExclamation
            End If
            Exit Function
         End If
      Next
      
   End If
   SrcCheckSameCust = True
End Function
'案源變數清除
Private Sub SrcLOSReset()
   Erase strLCaseNo
   Erase strCUNo
      
   bolSuYuan = False
   bolIsB2CourtFee = False
   bolPTIsNew = False
   
   strLCaseCP10 = ""
   strLSourceType = ""
   strCheckStatus = ""
   strPTSysCode = ""
   strPTCP10 = ""
   strLOS15 = ""
   strLOS17 = ""
   strLOS18 = ""
   
   SrcLockField False
End Sub

'檢查是否有案源
Private Function SrcCheckExistLOS(PTCaseNo1 As String, PTCaseNo2 As String, PTCaseNo3 As String, PTCaseNo4 As String, LCaseNo() As String, Optional strType As String, Optional Pty As String) As Boolean
   Dim strCon As String
   
   If strType <> "" Then strCon = strCon & " and los02='" & strType & "'"
   
   strExc(0) = "select cpm03,c2.cp01,c2.cp02,c2.cp03,c2.cp04 from caseprogress c1,lawofficesource,casepropertymap,caseprogress c2" & _
      " where c1.cp01='" & PTCaseNo1 & "' and c1.cp02='" & PTCaseNo2 & "' and c1.cp03='" & PTCaseNo3 & "' and c1.cp04='" & PTCaseNo4 & "'" & _
      " and c1.cp159=0 and los01(+)=c1.cp09 and c2.cp09(+)=LOS06 and c2.cp09 is not null" & strCon & " and cpm01(+)=c1.cp01 and cpm02(+)=c1.cp10 order by c1.cp05 desc"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   '有案源
   If intI > 0 Then
      SrcCheckExistLOS = True
      LCaseNo(1) = "" & RsTemp("cp01")
      LCaseNo(2) = "" & RsTemp("cp02")
      LCaseNo(3) = "" & RsTemp("cp03")
      LCaseNo(4) = "" & RsTemp("cp04")
      Pty = "" & RsTemp("cpm03")
   End If
End Function

'檢查是否訴願階段:有收文訴願且無准駁(有可能發文後才收文言詞辯論 Ex:T-225478)
Private Function SrcChkIsSuYuan(pCP01 As String, pCP02 As String, pCP03 As String, pCP04 As String, pCP10s As String) As Boolean
   Dim stSQL As String, intQ As Integer
   Dim RsQ As ADODB.Recordset
   Dim stB2 As String
      
   stSQL = "select * from caseprogress a" & _
      " where " & ChgCaseprogress(pCP01 & pCP02 & pCP03 & pCP04) & _
      " and cp159=0 and instr('" & pCP10s & "',cp10)>0 and cp24 is null"
   intQ = 1
   Set RsQ = ClsLawReadRstMsg(intQ, stSQL)
   If intQ = 1 Then
      SrcChkIsSuYuan = True
   End If
   Set RsQ = Nothing
End Function
'抓法務相關P/T案
Private Function SrcGetPTCase(pLCase() As String, pPTCaseNo() As String) As Boolean
   Dim stSQL As String, intQ As Integer
   Dim RsQ As ADODB.Recordset
   
   stSQL = "select b.cp01,b.cp02,b.cp03,b.cp04 from caseprogress a,caseprogress b" & _
      " where a.cp01='" & pLCase(1) & "' and a.cp02='" & pLCase(2) & "' and a.cp158=0" & _
      " and a.cp162 is not null and b.cp162(+)=a.cp162 and b.cp01<>a.cp01 and b.cp158=0" & _
      " union select b.cp01,b.cp02,b.cp03,b.cp04 from caseprogress a,acc0n0,caseprogress b" & _
      " where a.cp01='" & pLCase(1) & "' and a.cp02='" & pLCase(2) & "' and a.cp158=0" & _
      " and a0n01(+)=a1.cp09 and b.cp09(+)=a0n02 and b.cp09<>a0n01"
   intQ = 1
   Set RsQ = ClsLawReadRstMsg(intQ, stSQL)
   If intQ = 1 Then
      pPTCaseNo(1) = RsQ("cp01")
      pPTCaseNo(2) = RsQ("cp02")
      pPTCaseNo(3) = RsQ("cp03")
      pPTCaseNo(4) = RsQ("cp04")
      SrcGetPTCase = True
   End If
   Set RsQ = Nothing
End Function

Private Function SrcGetProperty(pSys As String, pCP10s As String) As String
   Dim stSQL As String, intQ As Integer
   Dim RsQ As ADODB.Recordset
   Dim stTemp As String
   
   stSQL = "select cpm03 from casepropertymap" & _
      " where cpm01='" & pSys & "' and cpm02 in (" & pCP10s & ")" & _
      " order by instr('" & pCP10s & "',cpm02) asc"
   intQ = 1
   Set RsQ = ClsLawReadRstMsg(intQ, stSQL)
   If intQ = 1 Then
      With RsQ
      stTemp = "" & .Fields(0)
      .MoveNext
      Do While Not .EOF
         stTemp = stTemp & "、" & .Fields(0)
         .MoveNext
      Loop
      End With
      SrcGetProperty = stTemp
   End If
   
   Set RsQ = Nothing
End Function
'Added by Morgan 2020/6/17
Private Function SrcGetLC47() As String
   Dim stSQL As String, intQ As Integer
   Dim RsQ As ADODB.Recordset
   
   If Option1(0).Value = True Then
      SrcGetLC47 = Text1(127)
   Else
      If bolPrintNewCase = True And Text1(8).Tag <> "" Then
         stSQL = " select lc47 from lawcase where lc01='" & Text1(6) & "' and lc02='" & Text1(7) & "' and lc03='" & Text1(8).Tag & "' and lc04='" & Text1(9) & "'"
      Else
         stSQL = "select lc47 from lawcase where lc01='" & Text1(6) & "' and lc02='" & Text1(7) & "' and lc03='" & Text1(8) & "' and lc04='" & Text1(9) & "'"
      End If
      intQ = 1
      Set RsQ = ClsLawReadRstMsg(intQ, stSQL)
      If intQ = 1 Then
         SrcGetLC47 = "" & RsQ(0)
      End If
   End If
   Set RsQ = Nothing
End Function

'Added by Morgan 2020/6/17
'法務案是否為智財權類
'pAddElse
Private Function SrcChkLIsIPCase(pLC47 As String, Optional pPTOnly As Boolean = True) As Boolean
   If InStr(pLC47, "專利") > 0 Or InStr(pLC47, "商標") > 0 Or InStr(pLC47, "著作權") > 0 Then
      SrcChkLIsIPCase = True
   ElseIf pPTOnly = False Then
      If InStr(pLC47, "智財權") > 0 Then
         SrcChkLIsIPCase = True
      End If
   End If
End Function
'Added by Morgan 2020/6/17
'法務案是否為訴訟案
Private Function SrcChkLIsSuitCase() As Boolean
   Dim ii As Integer
   Dim arrTmp() As String
   
   arrTmp = Split(m_strCaseCPM, ",")
   'For ii = 1 To 4
   For ii = 0 To UBound(arrTmp)
      'If Combo1(ii).Text <> "" Then
      If arrTmp(ii) <> "" Then
         'arrTmp = Split(Combo1(ii), " ")
         'If SrcChkIsB1CP10(Text1(6), arrTmp(0)) = True Then
         If SrcChkIsB1CP10(Text1(6), arrTmp(ii)) = True Then
            SrcChkLIsSuitCase = True
            Exit For
         End If
      End If
   Next ii
End Function

'Added by Morgan 2020/6/18
'以法務案類別設定PT案系統別及預設收文案件性質
Private Sub SetPTByLC47(pLC47 As String, pPTSys As String, pPTCP10 As String)
   If InStr(pLC47, "專利") > 0 Then
      pPTSys = "P"
      pPTCP10 = "226,225"
   ElseIf InStr(pLC47, "商標") > 0 Then
      pPTSys = "T"
      pPTCP10 = "213,212"
   Else
      pPTSys = "TC"
      pPTCP10 = "213,212"
   End If
End Sub

'Added by Morgan 2020/6/17
'法務案收文性質是否屬B1類(訴訟案件性質)
Private Function SrcChkIsB1CP10(pSys As String, pCP10 As String) As Boolean
   Dim stSQL As String, intQ As Integer
   Dim RsQ As ADODB.Recordset
   
   stSQL = "select cpm11 from casepropertymap where cpm01='" & pSys & "'and cpm02='" & pCP10 & "' and cpm11 in ('414101','416111','416121')"
   intQ = 1
   Set RsQ = ClsLawReadRstMsg(intQ, stSQL)
   If intQ = 1 Then
      SrcChkIsB1CP10 = True
   End If
   Set RsQ = Nothing
End Function

'Add by Amy 2020/09/01 判斷是否為大陸地址
Private Function ChkCN(ByVal stCuAddr As String) As Boolean
    Dim RsQ As New ADODB.Recordset
    Dim intQ As Integer
    Dim strQ As String, stAddr As String
    
    ChkCN = False
    stAddr = stCuAddr
    '去除中國大陸/大陸/中國 文字,取其前3個字或2個字查國家檔域別為大陸地區
    If Left(stAddr, 4) = "中國大陸" Then stAddr = Mid(stAddr, 5)
    If Left(stAddr, 2) = "中國" Then stAddr = Mid(stAddr, 3)
    If Left(stAddr, 2) = "大陸" Then stAddr = Mid(stAddr, 3)
    
    strQ = "Select * From Nation Where Na03='" & Left(stAddr, 3) & "' And SubStr(Na02,1,1)='B' "
    intQ = 1
    Set RsQ = ClsLawReadRstMsg(intQ, strQ)
    If intQ = 1 Then
        ChkCN = True
    Else
        strQ = "Select * From Nation Where Na03='" & Left(stAddr, 2) & "' And SubStr(Na02,1,1)='B' "
        intQ = 1
        Set RsQ = ClsLawReadRstMsg(intQ, strQ)
        If intQ = 1 Then
            ChkCN = True
        End If
    End If
    
    Set RsQ = Nothing
End Function

'Added by Morgan 2020/11/9
'TW-SUPA僅適用於沒有主張優先權的台灣案
Private Function ChkTwSupa() As Boolean
   Dim ii As Integer, jj As Integer
   Dim stCP10 As String, stCP10_1 As String, stCP10_2 As String
   Dim bCancel As Boolean
   
   ChkTwSupa = True
   'For ii = 1 To 4
      'stCP10 = Trim(Left(Me.Combo1(ii).Text, 4))
      'If stCP10 = "434" Or stCP10 = "106" Then
      If InStr(m_strCaseCPM, "434") > 0 Or InStr(m_strCaseCPM, "106") > 0 Then
         'If stCP10 = "434" Then
         If InStr(m_strCaseCPM, "434") > 0 Then
            stCP10_2 = "106"
         Else
            stCP10_2 = "434"
         End If
         
         'For jj = 1 To 4
            'stCP10 = Trim(Left(Me.Combo1(jj).Text, 4))
            'If stCP10 = stCP10_2 Then
            If InStr(m_strCaseCPM, stCP10_2) > 0 Then
               ChkTwSupa = False
               'Exit For
            End If
         'Next
         
         If ChkTwSupa = True Then
            '舊案再檢查收文
            If Option1(1).Value = True Then
               strExc(0) = "select cp09 from caseprogress where " & ChgCaseprogress(Text1(6) & Text1(7) & Text1(8) & Text1(9)) & " and cp10='" & stCP10_2 & "' and cp57 is null"
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
               If intI = 1 Then
                  ChkTwSupa = False
               End If
            End If
         End If
         
         If ChkTwSupa = False Then
            If stCP10_2 = "434" Then
               MsgBox "此案有申請TW-SUPA，不應主張國際優先權，請確認!!", vbCritical
            Else
               MsgBox "此案有主張國際優先權，不得申請TW-SUPA!!", vbCritical
            End If
            Exit Function
         End If
         
      End If
   'Next
End Function

'Added by Lydia 2021/02/24
Private Sub Text7_KeyPress(KeyAscii As MSForms.ReturnInteger)
   If Text1(6).Text = "CFT" Then
      KeyAscii = UpperCase(KeyAscii)
   End If
   '原本主題有下列控制，先保留
   'If KeyAscii = 91 Or KeyAscii = 93 Or KeyAscii = -24219 Or KeyAscii = -24218 Then
   '     MsgBox Replace(fra47Title.Caption & "禁止打中括號!!!", "：", ""), vbExclamation + vbOKOnly
   '     KeyAscii = 0
   'End If
End Sub

Private Sub Text7_GotFocus()
    TextInverse Text7
End Sub

'Added by Lydia 2021/05/07 ACS智財顧問專業分配比例管制
Private Sub SetACS112data()
Dim pCP09 As String, pCP15 As String, pCP53 As String, pCP54 As String
Dim pACS01 As String, pACS02 As String, pACS03 As String, pACS04 As String
Dim pWHours As String, pWTimes As String, pNTimes As String
    
    m_ACS112msg = "": m_ACS112chk = ""
    If strSrvDate(1) < ACS_PFrateStart Then Exit Sub

    'Modified by Lydia 2021/07/05 +L,CFL,
    If Option1(1).Value = True And InStr("CFP,CFT,P,T,ACS,L,CFL,", Text1(6) & ",") > 0 And Len(Trim(Text1(7))) = 6 Then
        If Pub_GetACS112Range(Text1(6), Text1(7), IIf(Trim(Text1(8)) <> "", Trim(Text1(8)), "0"), IIf(Trim(Text1(9)) <> "", Trim(Text1(9)), "00"), pCP09, pCP15, pCP53, pCP54, pACS01, pACS02, pACS03, pACS04) = True Then
             If strSrvDate(1) >= pCP54 Then
                 If Text1(6) = "ACS" Then
                     If InStr(m_strCaseCPM & ",", "112,") = 0 Then
                         m_ACS112chk = "此案智財顧問期間已過期，除智財顧問112外，不可再收文！"
                     End If
                 Else
                     m_ACS112chk = "此案之相關智財顧問" & pACS01 & "-" & pACS02 & IIf(pACS03 <> "0", "-" & pACS03, "") & IIf(pACS04 <> "00", "-" & pACS04, "") & _
                                   "，顧問期間：" & ChangeWStringToTDateString(pCP53) & "~" & ChangeWStringToTDateString(pCP54) & "已過期，不可再收文！"
                 End If
                 If m_ACS112chk <> "" Then
                    MsgBox m_ACS112chk, vbExclamation, "ACS智財顧問管制"
                    Exit Sub
                 End If
             End If
             
             If Text1(6) = "ACS" Then
                m_ACS112msg = ACS112STATISTICS("1", Text1(6), Text1(7), IIf(Trim(Text1(8)) <> "", Trim(Text1(8)), "0"), IIf(Trim(Text1(9)) <> "", Trim(Text1(9)), "00"), pCP09, pCP15, pCP53, pCP54, pWHours, pWTimes, pNTimes)
             Else
                m_ACS112msg = ACS112STATISTICS("1", pACS01, pACS02, pACS03, pACS04, pCP09, pCP15, pCP53, pCP54, pWHours, pWTimes, pNTimes)
             End If
             If Val(pWHours) >= Val(pCP15) Then
                 If Text1(6) = "ACS" Then
                     If InStr(m_strCaseCPM & ",", "112,") = 0 Then
                          m_ACS112chk = "此案智財顧問案總工作時數已達簽約時數，除智財顧問112外，不可再收文！"
                     End If
                 Else
                    m_ACS112chk = "此案之相關智財顧問" & pACS01 & "-" & pACS02 & IIf(pACS03 <> "0", "-" & pACS03, "") & IIf(pACS04 <> "00", "-" & pACS04, "") & _
                                   "，總工作時數已達簽約時數，不可再收文！"
                 End If
                 If m_ACS112chk <> "" Then
                    MsgBox m_ACS112chk, vbExclamation, "ACS智財顧問管制"
                    Exit Sub
                 End If
             End If
             
             If strSrvDate(1) >= pCP53 And strSrvDate(1) < pCP54 Then
                 If Text1(6) <> "ACS" Then  '追加相關ACS案號
                     'Modified by Lydia 2021/07/05 + vbCrlf
                     m_ACS112msg = "智財顧問案：" & pACS01 & "-" & pACS02 & IIf(pACS03 <> "0", "-" & pACS03, "") & IIf(pACS04 <> "00", "-" & pACS04, "") & "，" & vbCrLf & m_ACS112msg
                 End If
                 MsgBox m_ACS112msg, vbInformation, "ACS智財顧問管制"
             End If
        End If
    End If
End Sub

'Added by Lydia 2021/11/23 判斷是否有商標查名 (因為T案增加737智財協作,調整判斷)
Private Function GetTMQArea() As Boolean
Dim intP As Integer, strP1 As String
Dim arrText As Variant
   
   m_strCaseCPM = GetAllCaseCPM() 'Add By Sindy 2022/9/12 取得案件性質代碼
   GetTMQArea = False
   strP1 = ""
   'Modified by Lydia 2022/02/24 debug : Text1(6).Text = "T" Or Text1(6).Text = "TS" => (Text1(6).Text = "T" Or Text1(6).Text = "TS")
   If strSrvDate(1) >= TMQ電子化啟用日 And (Text1(6).Text = "T" Or Text1(6).Text = "TS") And _
      Left(Trim(Combo1(0)), 3) = "000" And Option1(0).Value = True Then
      'Modify By Sindy 2022/12/24
      If Me.Combo1(1).Text <> "" Then
         m_strCaseCPM = m_strCaseCPM & IIf(m_strCaseCPM <> "", ",", "") & Trim(Left(Combo1(1).Text, 4))
      End If
      arrText = Split(m_strCaseCPM, ",")
      '2022/12/24 END
      'For intP = 1 To 4
      For intP = 0 To UBound(arrText)
         'If Trim(Left(Trim(Combo1(intP)), 4)) <> "" Then
         If arrText(intP) <> "" Then
            If Text1(6).Text = "T" Then
               'If InStr(TMQ_T案, Trim(Left(Trim(Combo1(intP)), 4))) > 0 Then
               If InStr(TMQ_T案, arrText(intP)) > 0 Then
                  strP1 = "Y"
                  Exit For 'Add By Sindy 2022/12/24
               End If
            ElseIf Text1(6).Text = "TS" Then
               'If InStr(TMQ_TS案, Trim(Left(Trim(Combo1(intP)), 4))) > 0 Then
               If InStr(TMQ_TS案, arrText(intP)) > 0 Then
                  strP1 = "Y"
                  Exit For 'Add By Sindy 2022/12/24
               End If
            End If
         End If
      Next intP
   End If
   
   If strP1 = "Y" Then GetTMQArea = True
End Function

Private Sub SetGrd2()
   Dim arrGridHeadText, arrGridHeadWidth
   Dim iRow As Integer

   arrGridHeadText = Array("簽核人員", "身份", "日期", "時間", "簽核結果", "B1104")
   arrGridHeadWidth = Array(1050, 600, 800, 800, 800, 0)
   GRD1.Visible = False
   GRD1.Cols = UBound(arrGridHeadText) + 1
   For iRow = 0 To GRD1.Cols - 1
      GRD1.row = 0
      GRD1.col = iRow
      GRD1.Text = arrGridHeadText(iRow)
      GRD1.ColWidth(iRow) = arrGridHeadWidth(iRow)
      GRD1.CellAlignment = flexAlignCenterCenter
   Next
   GRD1.Visible = True
End Sub

'Add By Sindy 2022/9/21
' 更新各控制項的狀態
Private Sub SetCtrlReadOnly_Flow(ByVal bEnable As Boolean)
   Me.txtF0306.Locked = bEnable
'   Me.txtF0306.Enabled = Not bEnable
'   If bEnable = True Then
'      txtF0306.BackColor = &H8000000F
'   Else
'      txtF0306.BackColor = &H80000005 '白底
'   End If
End Sub

Private Sub ClearField_Flow()
   txtF0306 = Empty
   txtF0310 = Empty
   txtF0310_2 = Empty
   txtF0309 = Empty
   txtNote = Empty
   txtF0407 = Empty
   GRD1.Clear
   SetGrd2
      
   '退回原因
   Label1(130).Visible = False
   LblReason.Visible = False
   LblReason.Caption = ""
   '您的意見
   Label34.Visible = False
   txtNote.Visible = False
End Sub

Public Sub QueryData_Flow()
Dim rsTmp As New ADODB.Recordset
Dim RsQ As New ADODB.Recordset
Dim strSql As String
Dim strQ As String, intQ As Integer
Dim strF0305 As String
   
   ClearField_Flow '清空欄位值
   If Text5.Text = "" Then Exit Sub
   
   Screen.MousePointer = vbHourglass
   Me.Enabled = False
   
   '案件表單主檔
   strSql = "select flow003.*,decode(F0309," & ShowFlow表單狀態中文 & ") as F0309NM,decode(F0305,null,'',AC02||'--'||AC03) as BReason" & _
            " from flow003,allcode" & _
            " where f0301='" & Text5.Text & "' and f0302='" & Flow_接洽單 & "' And AC01(+)='12' And F0305=AC02(+)"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      txtF0306 = "" & rsTmp.Fields("F0306")
      strF0305 = "" & rsTmp.Fields("F0305")
      m_F0316 = "" & rsTmp.Fields("F0316") '智權人員
      txtF0310 = "" & rsTmp.Fields("F0310"): txtF0310_2 = GetPrjSalesNM("" & rsTmp.Fields("F0310"))
      If m_F0316 <> "" And txtF0310 <> "" Then
         If m_F0316 <> txtF0310 Then
            txtF0310.ForeColor = &HFF0000
            txtF0310_2.ForeColor = &HFF0000
         Else
            txtF0310.ForeColor = &H80000008
            txtF0310_2.ForeColor = &H80000008
         End If
      End If
      LblReason = "" & rsTmp.Fields("BReason") '退回原因
      m_F0307 = "" & rsTmp.Fields("F0307") '上一處理人員
      m_F0308 = "" & rsTmp.Fields("F0308") '下一處理人員
      m_F0309 = "" & rsTmp.Fields("F0309") '目前狀態
      txtF0309 = "" & rsTmp.Fields("F0309") & " " & rsTmp.Fields("F0309NM")
      Call UpdateCUID(rsTmp)
      If m_F0309 = "" Then Label29.Visible = False
   End If
   rsTmp.Close
   
   'Add By Sindy 2023/2/20
   If strLSourceType <> "" And m_F0309 = Flow_已完成 Then '已收文
      Me.cmdAddAtt.Visible = False
   End If
   '2023/2/20 END
   
   'Add By Sindy 2025/7/1
   If Check10.Value = 1 Then
      Me.Check10.BackColor = QBColor(14) '淡黃色
   End If
   '2025/7/1 END
   
   'If InStr(UCase(TypeName(m_PrevForm)), UCase("frm090801")) = 0 And UCase(TypeName(m_PrevForm)) <> "NOTHING" Then
   '退回原因
   Label1(130).Visible = True
   LblReason.Visible = True
   SSTab1.Tab = 0
   
   If (Text1(6) = "P" Or Text1(6) = "PS" Or Text1(6) = "CFP" Or Text1(6) = "CPS") And (m_F0309 >= Flow_待分案 Or InStr("'A5','A6','A7'", m_F0307) > 0) Then
      Frame57.Visible = True
   Else
      Frame57.Visible = False
   End If
   
   Me.Caption = "檢視接洽單 (本所案號：" & Text1(6) & "-" & Text1(7) & "-" & Text1(8) & "-" & Text1(9) & ")"
   'Added by Lydia 2024/05/16 主管分案作業->預設畫面改為說明處理事項；---李柏翰
   If UCase(TypeName(m_PrevForm)) = "FRM210156" Then
      SSTab1.Tab = 2
   Else
   'end 2024/05/16
      If Not (m_F0309 = Flow_已收文 Or m_F0309 = Flow_已分案 Or InStr(UCase(TypeName(m_PrevForm)), UCase("frmacc")) > 0) Then
         SSTab1.Tab = 5
      End If
   End If
   If InStr(UCase(TypeName(m_PrevForm)), UCase("frmacc")) > 0 Then
      SSTab1.TabVisible(3) = False
      SSTab1.TabVisible(4) = False
      SSTab3.TabVisible(1) = False
      If Text1(119) <> "" Then
         SSTab1.Tab = 2
      Else
         SSTab1.Tab = 0
      End If
   End If
      
   '案件表單流程備註檔
   SetFlow004TextBox txtF0407, Text5
   '案件表單簽核檔
   strSql = "SELECT decode(F0204," & ShowFlow特殊簽核人員 & ",ST02)||nvl(F0208,'') 簽核人員" & _
            ",decode(F0202," & ShowFlow簽核人員身份 & ") 身份,sqldateT(F0205) 日期,sqltime6(F0206) 時間" & _
            ",decode(F0207," & ShowFlow簽核結果 & ") 簽核結果,F0204" & _
            " FROM FLOW002,Staff WHERE F0201='" & Text5 & "' and F0204=ST01(+)" & _
            " order by decode(F0205,null,2,1) asc,F0205||sqltime6(F0206) asc,F0202,F0203 asc" 'order by F0205,F0202,F0203 asc
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      Set GRD1.Recordset = rsTmp
   End If
   
   '若有資料游標停在第一筆
   GRD1.Visible = False
   GRD1.col = 0
   GRD1.row = 1
   GRD1.Visible = True
   
   Screen.MousePointer = vbDefault
   Me.Enabled = True
   
EXITSUB:
   Set rsTmp = Nothing
   Set RsQ = Nothing
End Sub

Private Sub grd1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim nCol As Long, nRow As Long
getGrdColRow GRD1, x, y, nCol, nRow
GRD1.col = nCol
GRD1.row = nRow
End Sub

Private Sub grd1_SelChange()
Dim i As Integer, dblRow As Double
Dim k As Integer

GRD1.Visible = False
If GRD1.MouseRow <> 0 Then
   dblRow = 0
   For k = 1 To GRD1.Rows - 1
      GRD1.col = 1
      GRD1.row = k
      If GRD1.CellBackColor = &HFFC0C0 Then
         dblRow = k
         Exit For
      End If
   Next k
   '上一筆資料列清除反白
   If dblRow > 0 Then
      GRD1.col = 2
      GRD1.row = dblRow
      For i = 0 To GRD1.Cols - 1
         GRD1.col = i
         GRD1.CellBackColor = QBColor(15)
      Next i
   End If
   '目前資料列反白
   GRD1.col = 0
   GRD1.row = GRD1.MouseRow
   For i = 0 To GRD1.Cols - 1
      GRD1.col = i
      GRD1.CellBackColor = &HFFC0C0
   Next i
End If
GRD1.Visible = True
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
   
   If IsNull(rsSrcTmp.Fields("F0310")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("F0310")) = False Then
         strCName = GetStaffName(rsSrcTmp.Fields("F0310"), True)
      End If
   End If
   If IsNull(rsSrcTmp.Fields("F0311")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("F0311")) = False Then
         strTemp = TAIWANDATE(rsSrcTmp.Fields("F0311"))
         strCDate = Format(strTemp, "###/##/##")
      End If
   End If
   If IsNull(rsSrcTmp.Fields("F0312")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("F0312")) = False Then
         strTemp = rsSrcTmp.Fields("F0312")
         strCTime = Format(strTemp, "##:##:##")
      End If
   End If
   If IsNull(rsSrcTmp.Fields("F0313")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("F0313")) = False Then
         strUName = GetStaffName(rsSrcTmp.Fields("F0313"), True)
      End If
   End If
   If IsNull(rsSrcTmp.Fields("F0314")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("F0314")) = False Then
         strTemp = TAIWANDATE(rsSrcTmp.Fields("F0314"))
         strUDate = Format(strTemp, "###/##/##")
      End If
   End If
   If IsNull(rsSrcTmp.Fields("F0315")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("F0315")) = False Then
         strTemp = rsSrcTmp.Fields("F0315")
         strUTime = Format(strTemp, "##:##:##")
      End If
   End If

   ' 設定CUID中的文字
   Label29.Caption = "CREATE : " & strCName & " " & _
              " " & strCDate & " " & _
              " " & strCTime & String(10, " ") & _
              "UPDATE : " & strUName & " " & _
              " " & strUDate & " " & _
              " " & strUTime
End Sub

Private Sub txtF0306_Change()
   PUB_RefreshText txtF0306
End Sub
Private Sub txtF0306_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   If Button = 2 Then Forms(0).PopupMenu2 txtNote
End Sub
Private Sub txtNote_Change()
   PUB_RefreshText txtNote
End Sub
Private Sub txtNote_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   If Button = 2 Then Forms(0).PopupMenu2 txtNote
End Sub
