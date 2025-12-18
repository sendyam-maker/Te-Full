VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm12040163 
   BorderStyle     =   1  '單線固定
   Caption         =   "風險檢查資料維護"
   ClientHeight    =   5532
   ClientLeft      =   108
   ClientTop       =   936
   ClientWidth     =   9144
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5540.681
   ScaleMode       =   0  '使用者自訂
   ScaleWidth      =   9047.999
   Begin VB.CommandButton CmdDBN 
      Appearance      =   0  '平面
      BackColor       =   &H00C0C0FF&
      Caption         =   "正式DB"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   399
      Left            =   7883
      Style           =   1  '圖片外觀
      TabIndex        =   61
      Top             =   120
      Visible         =   0   'False
      Width           =   1162
   End
   Begin VB.CommandButton CmdTranNotAg 
      BackColor       =   &H00FFFFC0&
      Caption         =   "風險檢查對象轉為不得代理名單"
      Height          =   300
      Left            =   6240
      Style           =   1  '圖片外觀
      TabIndex        =   59
      Top             =   849
      Visible         =   0   'False
      Width           =   2779
   End
   Begin VB.TextBox textRCL01 
      Height          =   264
      Left            =   756
      MaxLength       =   5
      TabIndex        =   0
      Top             =   660
      Width           =   735
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4543
      Left            =   60
      TabIndex        =   28
      Top             =   936
      Width           =   9048
      _ExtentX        =   15960
      _ExtentY        =   8022
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   5
      TabHeight       =   420
      TabCaption(0)   =   "基本資料"
      TabPicture(0)   =   "frm12040163.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label41(15)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label41(14)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label41(13)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label41(10)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label30(0)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label29"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label27"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label2(0)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label5"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label6"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "LabRCL21_2"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label1(9)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label1(6)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label3(0)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Label4"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "textRCL22_2"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "textRCL02"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "textRCL07"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "textRCL23"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "LabRCL18_2"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Label1(1)"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "Label1(10)"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "Label3(1)"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "Label7"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "textRCL26_2"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "LabMsg"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "textRCL06"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "textRCL05"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "textRCL04"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "textRCL03"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "textRCL24"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "textRCL22"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "textRCL21"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "textRCL18"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "textRCL19"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "textRCL17"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "textRCL20"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "CmdNotCancel"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "CmdExtension"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).Control(39)=   "textRCL26"
      Tab(0).Control(39).Enabled=   0   'False
      Tab(0).Control(40)=   "cboRCL25"
      Tab(0).Control(40).Enabled=   0   'False
      Tab(0).Control(41)=   "CboRCL08"
      Tab(0).Control(41).Enabled=   0   'False
      Tab(0).ControlCount=   42
      TabCaption(1)   =   "地址"
      TabPicture(1)   =   "frm12040163.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "textRCL11"
      Tab(1).Control(1)=   "textRCL12"
      Tab(1).Control(2)=   "textRCL13"
      Tab(1).Control(3)=   "textRCL14"
      Tab(1).Control(4)=   "textRCL10"
      Tab(1).Control(5)=   "textRCL15"
      Tab(1).Control(6)=   "textRCL16"
      Tab(1).Control(7)=   "textRCL09"
      Tab(1).Control(8)=   "Label18"
      Tab(1).Control(9)=   "Label16"
      Tab(1).Control(10)=   "Label13"
      Tab(1).Control(11)=   "Label41(32)"
      Tab(1).Control(12)=   "Label41(2)"
      Tab(1).Control(13)=   "Label41(3)"
      Tab(1).Control(14)=   "Label41(4)"
      Tab(1).Control(15)=   "Label41(5)"
      Tab(1).Control(16)=   "Label41(6)"
      Tab(1).ControlCount=   17
      Begin VB.ComboBox CboRCL08 
         Height          =   276
         Left            =   1368
         TabIndex        =   7
         Top             =   1620
         Width           =   2625
      End
      Begin VB.ComboBox cboRCL25 
         Height          =   276
         ItemData        =   "frm12040163.frx":0038
         Left            =   1368
         List            =   "frm12040163.frx":0042
         TabIndex        =   16
         Text            =   "cboRCL25"
         Top             =   3900
         Width           =   2250
      End
      Begin VB.TextBox textRCL26 
         Height          =   270
         Left            =   1368
         MaxLength       =   6
         TabIndex        =   17
         Top             =   4200
         Width           =   650
      End
      Begin VB.CommandButton CmdExtension 
         BackColor       =   &H00FFFFC0&
         Caption         =   "延展"
         Enabled         =   0   'False
         Height          =   300
         Left            =   4680
         Style           =   1  '圖片外觀
         TabIndex        =   62
         Top             =   2520
         Width           =   600
      End
      Begin VB.CommandButton CmdNotCancel 
         BackColor       =   &H00FFFFC0&
         Caption         =   "不撤銷通知"
         Enabled         =   0   'False
         Height          =   300
         Left            =   2400
         Style           =   1  '圖片外觀
         TabIndex        =   60
         Top             =   3560
         Width           =   1200
      End
      Begin VB.TextBox textRCL20 
         Height          =   300
         Left            =   4260
         MaxLength       =   2
         TabIndex        =   14
         Text            =   "RC"
         Top             =   2520
         Width           =   400
      End
      Begin VB.TextBox textRCL17 
         Height          =   270
         Left            =   6060
         MaxLength       =   18
         TabIndex        =   8
         Top             =   1620
         Width           =   1650
      End
      Begin VB.TextBox textRCL19 
         Height          =   300
         Left            =   1368
         MaxLength       =   7
         TabIndex        =   13
         Top             =   2520
         Width           =   975
      End
      Begin VB.TextBox textRCL18 
         Height          =   270
         Left            =   1368
         MaxLength       =   8
         TabIndex        =   9
         Top             =   1920
         Width           =   1300
      End
      Begin VB.TextBox textRCL21 
         Height          =   270
         Left            =   4260
         Locked          =   -1  'True
         MaxLength       =   3
         TabIndex        =   12
         Top             =   2220
         Width           =   500
      End
      Begin VB.TextBox textRCL22 
         Height          =   270
         Left            =   1368
         MaxLength       =   6
         TabIndex        =   11
         Top             =   2220
         Width           =   650
      End
      Begin VB.TextBox textRCL11 
         Height          =   270
         Left            =   -70005
         MaxLength       =   30
         TabIndex        =   21
         Top             =   690
         Width           =   3360
      End
      Begin VB.TextBox textRCL12 
         Height          =   270
         Left            =   -73740
         MaxLength       =   30
         TabIndex        =   22
         Top             =   990
         Width           =   3360
      End
      Begin VB.TextBox textRCL13 
         Height          =   270
         Left            =   -70005
         MaxLength       =   30
         TabIndex        =   23
         Top             =   990
         Width           =   3360
      End
      Begin VB.TextBox textRCL14 
         Height          =   270
         Left            =   -73740
         MaxLength       =   30
         TabIndex        =   24
         Top             =   1290
         Width           =   3360
      End
      Begin VB.TextBox textRCL10 
         Height          =   270
         Left            =   -73740
         MaxLength       =   30
         TabIndex        =   20
         Top             =   690
         Width           =   3360
      End
      Begin VB.TextBox textRCL15 
         Height          =   270
         Left            =   -70005
         TabIndex        =   25
         Top             =   1290
         Width           =   3360
      End
      Begin VB.TextBox textRCL24 
         Height          =   270
         Left            =   1368
         MaxLength       =   7
         TabIndex        =   15
         Top             =   3560
         Width           =   975
      End
      Begin MSForms.TextBox textRCL03 
         Height          =   300
         Left            =   1368
         TabIndex        =   2
         Top             =   660
         Width           =   3360
         VariousPropertyBits=   671105051
         MaxLength       =   30
         Size            =   "5927;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textRCL04 
         Height          =   300
         Left            =   5112
         TabIndex        =   3
         Top             =   660
         Width           =   3360
         VariousPropertyBits=   671105051
         MaxLength       =   30
         Size            =   "5927;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textRCL05 
         Height          =   300
         Left            =   1368
         TabIndex        =   4
         Top             =   984
         Width           =   3360
         VariousPropertyBits=   671105051
         MaxLength       =   30
         Size            =   "5927;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textRCL06 
         Height          =   300
         Left            =   5112
         TabIndex        =   5
         Top             =   984
         Width           =   3360
         VariousPropertyBits=   671105051
         MaxLength       =   30
         Size            =   "5927;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label LabMsg 
         Caption         =   "(撤銷日期有值無法再修改此欄,需修改請通知電腦中心)"
         ForeColor       =   &H000000FF&
         Height          =   252
         Left            =   3600
         TabIndex        =   65
         Top             =   3600
         Visible         =   0   'False
         Width           =   4500
      End
      Begin MSForms.TextBox textRCL26_2 
         Height          =   300
         Left            =   2040
         TabIndex        =   64
         TabStop         =   0   'False
         Top             =   4200
         Width           =   924
         VariousPropertyBits=   671105055
         BackColor       =   -2147483633
         Size            =   "1630;529"
         Value           =   "textRCL26_2"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label7 
         Caption         =   "撤銷人員："
         Height          =   252
         Left            =   96
         TabIndex        =   63
         Top             =   4200
         Width           =   912
      End
      Begin VB.Label Label3 
         Caption         =   "延展次數："
         Height          =   252
         Index           =   1
         Left            =   3324
         TabIndex        =   57
         Top             =   2520
         Width           =   912
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "身分證字號/統一編號："
         Height          =   180
         Index           =   10
         Left            =   4200
         TabIndex        =   56
         Top             =   1620
         Width           =   1848
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "要求檢查編號："
         Height          =   180
         Index           =   1
         Left            =   100
         TabIndex        =   55
         Top             =   1920
         Width           =   1250
      End
      Begin MSForms.Label LabRCL18_2 
         Height          =   300
         Left            =   2760
         TabIndex        =   54
         Top             =   1920
         Width           =   6150
         Caption         =   "LabRCL18_2"
         Size            =   "10848;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textRCL16 
         Height          =   300
         Left            =   -73740
         TabIndex        =   26
         Top             =   1620
         Width           =   7116
         VariousPropertyBits=   671105051
         MaxLength       =   70
         Size            =   "12559;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textRCL09 
         Height          =   300
         Left            =   -73740
         TabIndex        =   19
         Top             =   360
         Width           =   7116
         VariousPropertyBits=   671105051
         MaxLength       =   70
         Size            =   "12559;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textRCL23 
         Height          =   708
         Left            =   1368
         TabIndex        =   10
         Top             =   2856
         Width           =   7116
         VariousPropertyBits=   -1467989989
         MaxLength       =   500
         ScrollBars      =   2
         Size            =   "12552;1244"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textRCL07 
         Height          =   300
         Left            =   1368
         TabIndex        =   6
         Top             =   1296
         Width           =   7116
         VariousPropertyBits=   671105051
         MaxLength       =   80
         Size            =   "12559;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textRCL02 
         Height          =   300
         Left            =   1368
         TabIndex        =   1
         Top             =   348
         Width           =   7116
         VariousPropertyBits=   671105051
         MaxLength       =   80
         Size            =   "12559;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textRCL22_2 
         Height          =   300
         Left            =   2040
         TabIndex        =   53
         TabStop         =   0   'False
         Top             =   2220
         Width           =   996
         VariousPropertyBits=   671105055
         BackColor       =   -2147483633
         Size            =   "1764;529"
         Value           =   "textRCL22_2"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label4 
         Caption         =   "撤銷原因："
         Height          =   252
         Left            =   96
         TabIndex        =   52
         Top             =   3900
         Width           =   912
      End
      Begin VB.Label Label3 
         Caption         =   "下次提醒日："
         Height          =   252
         Index           =   0
         Left            =   96
         TabIndex        =   51
         Top             =   2520
         Width           =   1104
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "負責同仁："
         Height          =   180
         Index           =   6
         Left            =   96
         TabIndex        =   50
         Top             =   2220
         Width           =   912
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "部門別："
         Height          =   180
         Index           =   9
         Left            =   3516
         TabIndex        =   49
         Top             =   2220
         Width           =   720
      End
      Begin VB.Label LabRCL21_2 
         AutoSize        =   -1  'True
         Caption         =   "LabRCL21_2"
         Height          =   180
         Left            =   4800
         TabIndex        =   48
         Top             =   2220
         Width           =   2004
      End
      Begin VB.Label Label18 
         Caption         =   "地址(中)："
         Height          =   255
         Left            =   -74820
         TabIndex        =   47
         Top             =   420
         Width           =   1005
      End
      Begin VB.Label Label16 
         Caption         =   "地址(英)："
         Height          =   255
         Left            =   -74820
         TabIndex        =   46
         Top             =   690
         Width           =   900
      End
      Begin VB.Label Label13 
         Caption         =   "地址(日)："
         Height          =   255
         Left            =   -74820
         TabIndex        =   45
         Top             =   1620
         Width           =   1005
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "1"
         Height          =   180
         Index           =   32
         Left            =   -73845
         TabIndex        =   44
         Top             =   690
         Width           =   90
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "3"
         Height          =   180
         Index           =   2
         Left            =   -73845
         TabIndex        =   43
         Top             =   990
         Width           =   90
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "5"
         Height          =   180
         Index           =   3
         Left            =   -73845
         TabIndex        =   42
         Top             =   1290
         Width           =   90
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "2"
         Height          =   180
         Index           =   4
         Left            =   -70125
         TabIndex        =   41
         Top             =   690
         Width           =   90
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "4"
         Height          =   180
         Index           =   5
         Left            =   -70125
         TabIndex        =   40
         Top             =   990
         Width           =   90
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "6"
         Height          =   180
         Index           =   6
         Left            =   -70125
         TabIndex        =   39
         Top             =   1290
         Width           =   90
      End
      Begin VB.Label Label6 
         Caption         =   "備註："
         Height          =   252
         Left            =   96
         TabIndex        =   38
         Top             =   2850
         Width           =   912
      End
      Begin VB.Label Label5 
         Caption         =   "撤銷日期："
         Height          =   252
         Left            =   96
         TabIndex        =   37
         Top             =   3560
         Width           =   912
      End
      Begin VB.Label Label2 
         Caption         =   "國籍："
         Height          =   252
         Index           =   0
         Left            =   100
         TabIndex        =   36
         Top             =   1620
         Width           =   912
      End
      Begin VB.Label Label27 
         Caption         =   "名稱(中)："
         Height          =   252
         Left            =   100
         TabIndex        =   35
         Top             =   350
         Width           =   912
      End
      Begin VB.Label Label29 
         Caption         =   "名稱(英)："
         Height          =   252
         Left            =   100
         TabIndex        =   34
         Top             =   660
         Width           =   912
      End
      Begin VB.Label Label30 
         Caption         =   "名稱(日)："
         Height          =   252
         Index           =   0
         Left            =   100
         TabIndex        =   33
         Top             =   1300
         Width           =   912
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "1"
         Height          =   180
         Index           =   10
         Left            =   1248
         TabIndex        =   32
         Top             =   660
         Width           =   96
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "2"
         Height          =   180
         Index           =   13
         Left            =   4992
         TabIndex        =   31
         Top             =   660
         Width           =   96
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "3"
         Height          =   180
         Index           =   14
         Left            =   1248
         TabIndex        =   30
         Top             =   984
         Width           =   96
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "4"
         Height          =   180
         Index           =   15
         Left            =   4992
         TabIndex        =   29
         Top             =   984
         Width           =   96
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7200
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
            Picture         =   "frm12040163.frx":0062
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040163.frx":037E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040163.frx":069A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040163.frx":0876
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040163.frx":0B92
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040163.frx":0EAE
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040163.frx":11CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040163.frx":14E6
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040163.frx":1802
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040163.frx":1B1E
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040163.frx":1E3A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar TBar1 
      Align           =   1  '對齊表單上方
      Height          =   576
      Left            =   0
      TabIndex        =   58
      Top             =   0
      Width           =   9144
      _ExtentX        =   16129
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
      Height          =   300
      Left            =   2436
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   624
      Width           =   6064
      VariousPropertyBits=   671107103
      BackColor       =   -2147483633
      Size            =   "10696;529"
      Value           =   "textCUID"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      Caption         =   "編號："
      Height          =   252
      Index           =   0
      Left            =   156
      TabIndex        =   27
      Top             =   660
      Width           =   552
   End
End
Attribute VB_Name = "frm12040163"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Create by Amy 2023/12/12
Option Explicit

Dim m_FieldList() As FIELDITEM
' 變數宣告區
Dim m_EditMode As Integer, TB_All As Integer
' 第一筆資料/最後一筆/目前正在顯示
Dim m_FirstKEY(1) As String, m_LastKEY(1) As String, m_CurrKEY(1) As String
'執行各項功能的權限
Dim m_bInsert As Boolean, m_bUpdate As Boolean, m_bDelete As Boolean, m_bQuery As Boolean
Dim intInputState As Integer '0-只能讀/1-能改「撒銷」欄/2-可改「全部」欄
Dim m_RCL24 As String, m_RCL27 As String, strSt52List As String, strA0924 As String, strSupMan As String '撒銷日/新增人員/登入者(帶人主管)帶的人/部門主管/最高主管
Dim i As Integer, strQ As String, m_MeTrackMode As String 'Form2.0 記錄鍵盤傳入順序
'Dim bolNameExist As Boolean, stSameNameData As String '與客戶／代理人／潛在客戶同名同姓且存檔者通知特定人員 'Mark by Amy 2024/04/30不使用

' 初始化欄位陣列
Private Sub InitialField()
   CheckOC2
   strQ = ""
   '新增 or 修改 人/日/時 及 中/英/日 名稱(去符號)-Trigger
   For i = 1 To 26
      strQ = strQ & ",RCL" & Format(i, "00")
   Next i
   strQ = "Select " & Mid(strQ, 2) & " From RiskCheckList Where RowNum<2 "
   adoRecordset1.CursorLocation = adUseClient
   adoRecordset1.Open strQ, cnnConnection, adOpenStatic, adLockReadOnly
   TB_All = adoRecordset1.Fields.Count
   CheckOC2
   ReDim m_FieldList(TB_All) As FIELDITEM
   Call Pub_InitialField(1, Me.Name, m_FieldList, adoRecordset1)
End Sub

'Add by Amy 2024/02/06 國籍改下拉
Private Sub CboRCL08_Validate(Cancel As Boolean)
   Dim strTit As String, strMsg As String, stCountry As String, nResponse
   
   If Trim(CboRCL08) = MsgText(601) Then Exit Sub

   stCountry = ChgType(1, CboRCL08)
   If m_EditMode <> 1 And m_EditMode <> 2 Then Exit Sub
   If m_FieldList(7).fiOldData = Trim(CboRCL08) Then Exit Sub

   strTit = "檢核資料"
   strMsg = "國籍"
   If IsEmptyText(CboRCL08) = False And stCountry = MsgText(601) Then
      Cancel = True
      nResponse = MsgBox(strMsg & "不正確", vbOKOnly, strTit)
      SSTab1.Tab = 0
      CboRCL08.SetFocus
      Exit Sub
   Else
      CboRCL08 = stCountry
   End If
End Sub

'延展 鈕
Private Sub CmdExtension_Click()
   Dim stDate As String, intFeq As Integer, stTmp As String
   Dim stTO As String, stSubject As String, stContext As String, ArrMail
   
   If MsgBox("確定要[延展]？", vbYesNo + vbCritical) = vbNo Then
      Exit Sub
   End If
   
   Screen.MousePointer = vbHourglass
   stDate = Val(textRCL19) + 19110000
   '延期=提醒日+3個月後的工作天
   stDate = CompWorkDay(1, DBDATE(DateAdd("m", 3, Format(stDate, "####/##/##"))), 1)
   '新增時,次數為Null,第1次延期才上1
   intFeq = Val(textRCL20) + 1
   '下次提醒日
   Pub_SetFieldNewData Me.Name, m_FieldList, "RCL19", DBDATE(stDate)
   Pub_SetFieldNewData Me.Name, m_FieldList, "RCL20", intFeq
   If ActRecord(2) = False Then Screen.MousePointer = vbDefault: Exit Sub
   UpdateCtrlData
   Call SetMailTo(1)
   
   Screen.MousePointer = vbDefault
End Sub

'不撤銷通知 鈕
Private Sub CmdNotCancel_Click()
   Dim stTO As String, stSubject As String, stContext As String
   
   If MsgBox("確定要發[不撤銷通知信]？", vbYesNo + vbCritical) = vbNo Then
      Exit Sub
   End If
   
   Screen.MousePointer = vbHourglass
   Call SetMailTo(3)
   Screen.MousePointer = vbDefault
End Sub

'Add by Amy 2024/07/19 風險檢查對象轉為不得代理名單 鈕
Private Sub CmdTranNotAg_Click()
   Dim oForm As Form
   Dim stMemo As String '不得代理名單TB沒有的欄位資料
   
   If MsgBox("資料確定轉到不得代理名單中？", vbYesNo + vbCritical) = vbNo Then
      Exit Sub
   End If

   If PUB_CheckFormExist("frm12040154") = True Then
      Forms(0).GetForm ("")
      MsgBox "請先關閉「不得代理案件之客戶或代理人資料維護」！", vbExclamation
      Exit Sub
   End If

   Me.Hide
   Set oForm = Forms(0).GetForm("frm12040154")

   With oForm
      Set .m_PrevForm = Me
      .m_RCL01 = textRCL01 '編號
      .textNT02 = textRCL02 '名稱(中)
      .textNT03 = textRCL03 '名稱(英)
      .textNT04 = textRCL04
      .textNT05 = textRCL05
      .textNT06 = textRCL06
      .textNT07 = textRCL07 '名稱(日)
      If CboRCL08 <> MsgText(601) Then
         .textNT08 = Mid(CboRCL08, 1, Val(InStr(CboRCL08, " ")) - 1) '國籍
      End If
      .textNT17 = textRCL21 '部門別
      .LabNT17_2 = LabRCL21_2
      .textNT18 = textRCL22 '負責同仁
      .LabNT18_2 = textRCL22_2
      .textNT19 = "風險檢查資料轉入" '原因
       '地址
      .textNT09 = textRCL09 '中
      .textNT10 = textRCL10 '英
      .textNT11 = textRCL11
      .textNT12 = textRCL12
      .textNT13 = textRCL13
      .textNT14 = textRCL14
      .textNT15 = textRCL15
      .textNT16 = textRCL16 '日
      '*** 不得代理資料檔TB沒有的欄位資料,寫入備註 ***
      If textRCL17 <> MsgText(601) Then stMemo = stMemo & ",身份證/統編:" & textRCL17
      If textRCL18 <> MsgText(601) Then stMemo = stMemo & ",要求檢查對象:" & textRCL18
      If textRCL19 <> MsgText(601) Then stMemo = stMemo & ",下次提醒日:" & textRCL19
      If textRCL20 <> MsgText(601) Then stMemo = stMemo & ",延展次數:" & textRCL20
      If textRCL24 <> MsgText(601) Then stMemo = stMemo & ",撤銷日期:" & textRCL24
      If cboRCL25 <> MsgText(601) Then stMemo = stMemo & ",撤銷原因:" & cboRCL25
      If textRCL26 <> MsgText(601) Then stMemo = stMemo & ",撤銷人員:" & textRCL26
      '*** End不得代理資料檔TB沒有的欄位資料,寫入備註 ***
      '備註
      If stMemo <> MsgText(601) Then
         '不加風險檢查編號,因新增取號是Max(),可能刪最後一筆再新增對應資料會不正確
         .textNT20 = "原風險檢查資料-" & Mid(stMemo, 2) & ";" & textRCL23
      End If
      'Added by Lydia 2025/07/29 管制對象
      If Trim(textRCL18) <> "" Then
         .Text1 = textRCL18
         Call .cmdAddNT35_Click
      End If
      'end 2025/07/29
      .Show
   End With
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   Call PUB_SaveMeTrackMode(m_MeTrackMode, 1, KeyCode)
   
    Screen.MousePointer = vbHourglass
    Select Case KeyCode
        Case vbKeyF2: OnAction KeyCode '新增
        Case vbKeyF3: OnAction KeyCode '修改
        Case vbKeyF5: OnAction KeyCode '刪除
        Case vbKeyF4: OnAction KeyCode '查詢
        Case vbKeyHome: OnAction KeyCode '第一筆
        Case vbKeyPageUp: OnAction KeyCode '前一筆
        Case vbKeyPageDown: OnAction KeyCode '後一筆
        Case vbKeyEnd: OnAction KeyCode '最後筆
        'Case vbKeyF9: OnAction KeyCode '確定 '以ENTER控制為換行的功能
        Case vbKeyF10: OnAction KeyCode '取消
        Case vbKeyEscape: OnAction KeyCode '結束
    End Select
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
   If Pub_StrUserSt03 = "M51" And UCase(strServerName) = "LIVE" Then
      CmdDBN.Visible = True
   End If
   MoveFormToCenter Me
   
   SSTab1.Tab = 0
   m_EditMode = 0
   CmdNotCancel.Visible = False
   If InStr(Replace(Pub_GetSpecMan("風險檢查對象可撤銷人員"), ";", ","), strUserNum) > 0 Then
      intInputState = 1
      '[風險檢查對象可撤銷人員]才有 [不撤銷通知] 鈕-秀玲
      CmdNotCancel.Visible = True '不撤銷通知 鈕
      CmdNotCancel.Enabled = True
      LabMsg.Visible = True '撤銷日期訊息通知
   ElseIf Pub_StrUserSt03 = "M51" Then
      LabMsg.Visible = True '撤銷日期訊息通知
   End If
   InitialField
   Call SetCboNation(1, "") 'Add by Amy 2024/02/06 國籍改下拉
   SetCboRCL25
   RefreshRange
   ShowFirstRecord
   SetCtrlReadOnly True
End Sub

'設定權限
Private Sub SetLimit()
   Dim stDate As String '日期
   
   '最高主管
   If Left(PUB_GetST93(m_RCL27), 1) = "S" Then
      strSupMan = Pub_GetSpecMan("全所智權部主管")
   Else
      strSupMan = Pub_GetSpecMan("總經理員工編號")
   End If
   strA0924 = GetDeptMan(PUB_GetST93(m_RCL27), 2) '目前資料建立者的部門主管
   strSt52List = GetST52SelfList(m_RCL27, "st52,st53,st54,st55") '目前資料建立者帶的人
   
   '新增 / 查詢 鈕(每個人都可操作)
   m_bQuery = True: m_bInsert = True '新增/查詢
   m_bUpdate = False: m_bDelete = False '依權限設定
   '只有電腦中心有[刪除]及[風險檢查對象轉為不得代理名單]鈕 權限
   If Pub_StrUserSt03 = "M51" Then
      m_bDelete = True: m_bUpdate = True
      intInputState = 2
      CmdTranNotAg.Visible = True
   End If
   '修改 新增人員本人及其各級主管 且 撒銷日為空
   If (m_RCL27 = strUserNum Or strA0924 = strUserNum Or InStr(strSt52List, strUserNum) > 0 Or strSupMan = strUserNum) _
     And m_RCL24 = MsgText(601) Then
      m_bUpdate = True
   End If
   '瀏覽
   If m_EditMode = 0 Then
      SetCtrlReadOnly True
      '*** 延展 鈕 ***
      CmdExtension.Enabled = False
      If Trim(textRCL19) <> MsgText(601) Then
         stDate = Val(textRCL19) + 19110000
         '下次提醒日 前10個工作天~下次提醒日 當天
         If Val(strSrvDate(1)) >= Val(CompWorkDay(10, stDate, 1)) And Val(strSrvDate(1)) <= Val(stDate) Then
            '登入者 為 此筆資料 建立者的[部門最高主管] Or 有修改權限 且 有下次提醒日,才可操作延展
            '智權部最高主管=Pub_GetSpecMan("全所智權部主管")) / [非]智權部=總經理或風險檢查對象可撤銷人員
            If ((PUB_GetST93(m_RCL27) = "S" And InStr(Pub_GetSpecMan("全所智權部主管"), strSupMan) > 0 And strSupMan = strUserNum) _
              Or (PUB_GetST93(m_RCL27) <> "S" And ((InStr(Pub_GetSpecMan("總經理員工編號"), strSupMan) > 0 And strSupMan = strUserNum) Or intInputState = 1)) _
              Or m_bUpdate = True) _
              And m_EditMode = 0 And textRCL19 <> MsgText(601) Then
               CmdExtension.Enabled = True
            End If
         End If
         '*** End 延展 鈕 ***
      End If
   ElseIf m_bInsert = True Or m_bUpdate = True Then
      SetCtrlReadOnly False
   End If
  
End Sub

'ToolBar權限及顯示
Private Sub ShowToolBar(ByVal intState As Integer)
   Dim oButton As Button
   
   '1:新增/2:修改/3:刪除/4:查詢
   '6:第一筆/7:前一筆/8:後一筆/9:最後一筆
   For i = 1 To 4
      TBar1.Buttons(i).Enabled = False
      TBar1.Buttons(i + 5).Enabled = False
   Next i
   TBar1.Buttons(11).Enabled = False '確定
   TBar1.Buttons(12).Enabled = False '取消
   TBar1.Buttons(14).Enabled = False '結束
   
   '瀏覽
   If intState = 0 Then
      '新增
      If m_bInsert = True Then
         TBar1.Buttons(1).Enabled = True
      End If
      '修改
      If m_bUpdate = True Or (intInputState = 1 And m_RCL24 = MsgText(601)) Then
         TBar1.Buttons(2).Enabled = True
      End If
      '刪除
      If m_bDelete = True Then
         TBar1.Buttons(3).Enabled = True
      End If
      '查詢
      If m_bQuery = True Then
         TBar1.Buttons(4).Enabled = True
         For i = 6 To 9
            TBar1.Buttons(i).Enabled = True
         Next i
      End If
      TBar1.Buttons(14).Enabled = True '結束
   '按下 新增/修改/查詢 後
   Else
      TBar1.Buttons(11).Enabled = True '確定
      TBar1.Buttons(12).Enabled = True '取消
      TBar1.Buttons(14).Enabled = False '結束
   End If
End Sub

Private Sub SetCboRCL25()
   cboRCL25.Clear
   cboRCL25.AddItem ""
   cboRCL25.AddItem "逾管制期限未收文"
   cboRCL25.AddItem "已收文"
End Sub

'重新抓第一筆及最後一筆
Private Sub RefreshRange()
   Dim rsTmp As New ADODB.Recordset, strSql As String
   
   strSql = "Select Count(*) From RiskCheckList "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 And Val(rsTmp.Fields(0)) > 0 Then
      rsTmp.Close
      strSql = "Select Min(RCL01) From RiskCheckList "
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         If IsNull(rsTmp.Fields(0)) = False Then: m_FirstKEY(0) = rsTmp.Fields(0)
      End If
      rsTmp.Close
      
      strSql = "Select Max(RCL01) From RiskCheckList "
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         If IsNull(rsTmp.Fields(0)) = False Then: m_LastKEY(0) = rsTmp.Fields(0)
      End If
   End If
   rsTmp.Close
   
   Set rsTmp = Nothing
End Sub

' 顯示第一筆資料
Private Sub ShowFirstRecord()
   m_CurrKEY(0) = m_FirstKEY(0)
   UpdateCtrlData
End Sub

' 顯示上一筆資料
Private Sub ShowPrevRecord()
   Dim rsTmp As New ADODB.Recordset, strSql As String
   
   If m_CurrKEY(0) = m_FirstKEY(0) Then
      ShowMsg MsgText(9008)
      GoTo EXITSUB
   End If

   strSql = "Select RCL01 From RiskCheckList " & _
            "Where RCL01 = (Select Max(RCL01) From RiskCheckList " & _
                           "Where RCL01 < '" & m_CurrKEY(0) & "') "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("RCL01")) = False Then: m_CurrKEY(0) = rsTmp.Fields("RCL01")
   End If
   rsTmp.Close
   UpdateCtrlData
   
EXITSUB:
   Set rsTmp = Nothing
End Sub

' 顯示下一筆資料
Private Sub ShowNextRecord()
   Dim rsTmp As New ADODB.Recordset, strSql As String
   
   If m_CurrKEY(0) = m_LastKEY(0) Then
      ShowMsg MsgText(9009)
      GoTo EXITSUB
   End If
   
   strSql = "Select RCL01 From RiskCheckList " & _
            "Where RCL01 = (Select Min(RCL01) From RiskCheckList " & _
                           "Where RCL01 > '" & m_CurrKEY(0) & "') "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("RCL01")) = False Then: m_CurrKEY(0) = rsTmp.Fields("RCL01")
   End If
   rsTmp.Close
   
   UpdateCtrlData
   
EXITSUB:
   Set rsTmp = Nothing
End Sub

' 顯示最後一筆資料
Private Sub ShowLastRecord()
   m_CurrKEY(0) = m_LastKEY(0)
   UpdateCtrlData
End Sub

' 顯示資料
Private Sub ShowCurrRecord(ByVal strKEY01 As String)
   Dim rsTmp As New ADODB.Recordset
   
   If ExistCheck("RiskCheckList", "RCL01", strKEY01, strExc(1), False) = True Then
      m_CurrKEY(0) = strKEY01
   Else
      strSql = "Select RCL01 From RiskCheckList " & _
                      "Where RCL01 = '" & m_CurrKEY(0) & "' "
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         If IsNull(rsTmp.Fields("RCL01")) = False Then: m_CurrKEY(0) = rsTmp.Fields("RCL01")
         rsTmp.Close
         UpdateCtrlData
         GoTo EXITSUB
      End If
      rsTmp.Close
      
      strSql = "Select RCL01 From RiskCheckList " & _
                     "Where RCL01 = (Select Min(RCL01) From RiskCheckList " & _
                                                   "Where RCL01 > '" & m_CurrKEY(0) & "') "
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         If IsNull(rsTmp.Fields("RCL01")) = False Then: m_CurrKEY(0) = rsTmp.Fields("RCL01")
      Else
         ShowLastRecord
         GoTo EXITSUB
      End If
      rsTmp.Close
   End If
   UpdateCtrlData
   
EXITSUB:
End Sub

' 將資料庫中的資料更新到所有欄位中
Private Sub UpdateCtrlData(Optional ByVal stKey As String = "")
   Dim rsTmp As New ADODB.Recordset, strSql As String
   
   Call ClearField: m_RCL24 = "": m_RCL27 = ""
   If stKey = MsgText(601) Then
      stKey = m_CurrKEY(0)
   End If
   strSql = "Select * From RiskCheckList Where RCL01= '" & stKey & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If m_CurrKEY(0) <> stKey Then m_CurrKEY(0) = stKey
      If IsNull(rsTmp.Fields("RCL01")) = False Then: textRCL01 = rsTmp.Fields("RCL01")
      If IsNull(rsTmp.Fields("RCL02")) = False Then: textRCL02 = rsTmp.Fields("RCL02")
      If IsNull(rsTmp.Fields("RCL03")) = False Then: textRCL03 = rsTmp.Fields("RCL03")
      If IsNull(rsTmp.Fields("RCL04")) = False Then: textRCL04 = rsTmp.Fields("RCL04")
      If IsNull(rsTmp.Fields("RCL05")) = False Then: textRCL05 = rsTmp.Fields("RCL05")
      If IsNull(rsTmp.Fields("RCL06")) = False Then: textRCL06 = rsTmp.Fields("RCL06")
      If IsNull(rsTmp.Fields("RCL07")) = False Then: textRCL07 = rsTmp.Fields("RCL07")
      'Modify by Amy 2024/02/06 國籍改下拉 原:textRCL08 = rsTmp.Fields("RCL08"): Call textRCL08_Validate(False)
      If IsNull(rsTmp.Fields("RCL08")) = False Then: CboRCL08 = SetCboNation(3, rsTmp.Fields("RCL08"))
      If IsNull(rsTmp.Fields("RCL09")) = False Then: textRCL09 = rsTmp.Fields("RCL09")
      If IsNull(rsTmp.Fields("RCL10")) = False Then: textRCL10 = rsTmp.Fields("RCL10")
      If IsNull(rsTmp.Fields("RCL11")) = False Then: textRCL11 = rsTmp.Fields("RCL11")
      If IsNull(rsTmp.Fields("RCL12")) = False Then: textRCL12 = rsTmp.Fields("RCL12")
      If IsNull(rsTmp.Fields("RCL13")) = False Then: textRCL13 = rsTmp.Fields("RCL13")
      If IsNull(rsTmp.Fields("RCL14")) = False Then: textRCL14 = rsTmp.Fields("RCL14")
      If IsNull(rsTmp.Fields("RCL15")) = False Then: textRCL15 = rsTmp.Fields("RCL15")
      If IsNull(rsTmp.Fields("RCL16")) = False Then: textRCL16 = rsTmp.Fields("RCL16")
      If IsNull(rsTmp.Fields("RCL17")) = False Then: textRCL17 = rsTmp.Fields("RCL17") '身分證/統編
      If IsNull(rsTmp.Fields("RCL18")) = False Then: textRCL18 = rsTmp.Fields("RCL18"): Call textRCL18_Validate(False) '要求檢查對象
      If IsNull(rsTmp.Fields("RCL19")) = False Then: textRCL19 = TAIWANDATE(rsTmp.Fields("RCL19")) '下次提醒日
      If IsNull(rsTmp.Fields("RCL20")) = False Then: textRCL20 = rsTmp.Fields("RCL20") '延展次數
      If IsNull(rsTmp.Fields("RCL21")) = False Then: textRCL21 = rsTmp.Fields("RCL21"): Call textRCL21_Validate(False) '部門
      If IsNull(rsTmp.Fields("RCL22")) = False Then: textRCL22 = rsTmp.Fields("RCL22"): Call textRCL22_Validate(False) '負責同仁
      If IsNull(rsTmp.Fields("RCL23")) = False Then: textRCL23 = rsTmp.Fields("RCL23") '備註
      If IsNull(rsTmp.Fields("RCL24")) = False Then: textRCL24 = TAIWANDATE(rsTmp.Fields("RCL24")): m_RCL24 = textRCL24 '撤銷日期
      If IsNull(rsTmp.Fields("RCL25")) = False Then: cboRCL25 = rsTmp.Fields("RCL25") '撤銷原因
      'Memo 撤銷人員 系統自動上操作撤銷日期/撤銷原因 的人員,有值後,只允許電腦中心修改
      If IsNull(rsTmp.Fields("RCL26")) = False Then: textRCL26 = rsTmp.Fields("RCL26"): Call textRCL26_Validate(False)
      If IsNull(rsTmp.Fields("RCL27")) = False Then: m_RCL27 = rsTmp.Fields("RCL27") '新增人員
      ' 更新CUID
      UpdateCUID rsTmp
      ' 更新暫存區的資料
      Call Pub_InitialField(2, Me.Name, m_FieldList, rsTmp)
   End If
   rsTmp.Close
   SetLimit
   ShowToolBar 0
   
EXITSUB:
   Set rsTmp = Nothing
End Sub

' 更新欄位的內容
Private Sub UpdateFieldNewData()
   Dim stRCL08 As String 'Add by Amy  2024/02/06
   
   '編號(新增時,textRCL01為空,於取號後才有值)
   If IsEmptyText(textRCL01) = False Then
      Pub_SetFieldNewData Me.Name, m_FieldList, "RCL01", textRCL01
   End If
   Pub_SetFieldNewData Me.Name, m_FieldList, "RCL02", textRCL02 '流水號
   Pub_SetFieldNewData Me.Name, m_FieldList, "RCL03", textRCL03 '中文名稱
   Pub_SetFieldNewData Me.Name, m_FieldList, "RCL04", textRCL04
   Pub_SetFieldNewData Me.Name, m_FieldList, "RCL05", textRCL05
   Pub_SetFieldNewData Me.Name, m_FieldList, "RCL06", textRCL06
   Pub_SetFieldNewData Me.Name, m_FieldList, "RCL07", textRCL07
   'Modify by Amy 2024/02/06 國籍改下拉
   'Pub_SetFieldNewData Me.Name, m_FieldList, "RCL08", textRCL08 '國籍
   stRCL08 = CboRCL08
   If stRCL08 <> MsgText(601) Then stRCL08 = Mid(stRCL08, 1, Val(InStr(stRCL08, " ")) - 1)
   Pub_SetFieldNewData Me.Name, m_FieldList, "RCL08", stRCL08
   'end 2024/02/06
   Pub_SetFieldNewData Me.Name, m_FieldList, "RCL09", textRCL09 '中文地址
   Pub_SetFieldNewData Me.Name, m_FieldList, "RCL10", textRCL10
   Pub_SetFieldNewData Me.Name, m_FieldList, "RCL11", textRCL11
   Pub_SetFieldNewData Me.Name, m_FieldList, "RCL12", textRCL12
   Pub_SetFieldNewData Me.Name, m_FieldList, "RCL13", textRCL13
   Pub_SetFieldNewData Me.Name, m_FieldList, "RCL14", textRCL14
   Pub_SetFieldNewData Me.Name, m_FieldList, "RCL15", textRCL15
   Pub_SetFieldNewData Me.Name, m_FieldList, "RCL16", textRCL16
   Pub_SetFieldNewData Me.Name, m_FieldList, "RCL17", textRCL17 '身份證/統編
   Pub_SetFieldNewData Me.Name, m_FieldList, "RCL18", textRCL18 '要求檢查對象
   '下次提醒日
   If IsEmptyText(textRCL19) = False Then
      Pub_SetFieldNewData Me.Name, m_FieldList, "RCL19", DBDATE(textRCL19)
   Else
      Pub_SetFieldNewData Me.Name, m_FieldList, "RCL19", textRCL19
   End If
   Pub_SetFieldNewData Me.Name, m_FieldList, "RCL20", textRCL20 '延展次數
   Pub_SetFieldNewData Me.Name, m_FieldList, "RCL21", textRCL21 '部門
   Pub_SetFieldNewData Me.Name, m_FieldList, "RCL22", textRCL22 '負責同仁
   Pub_SetFieldNewData Me.Name, m_FieldList, "RCL23", textRCL23 '備註
   '撤銷日期
   If IsEmptyText(textRCL24) = False Then
      Pub_SetFieldNewData Me.Name, m_FieldList, "RCL24", DBDATE(textRCL24)
   Else
      Pub_SetFieldNewData Me.Name, m_FieldList, "RCL24", textRCL24
   End If
   Pub_SetFieldNewData Me.Name, m_FieldList, "RCL25", cboRCL25 '原因
   '撤銷日期 由 空->有值且 [撤銷人員] 為空,寫入 操作人員 員編
   If m_FieldList(23).fiOldData = Empty And m_FieldList(23).fiOldData <> Trim(textRCL24) _
     And IsEmptyText(textRCL26) = True Then
      textRCL26 = strUserNum
   End If
   Pub_SetFieldNewData Me.Name, m_FieldList, "RCL26", textRCL26 '撤銷人員
End Sub

' 清除欄位內的資料內容
Private Sub ClearField()
   Dim nIdx As Integer
   
   For nIdx = 1 To TB_All
      m_FieldList(nIdx - 1).fiOldData = Empty
      m_FieldList(nIdx - 1).fiNewData = Empty
   Next nIdx
   textRCL01 = Empty
   textRCL02 = Empty
   textRCL03 = Empty
   textRCL04 = Empty
   textRCL05 = Empty
   textRCL06 = Empty
   textRCL07 = Empty
   'Modify by Amy 2024/02/06 國籍改下拉
   'textRCL08 = Empty:  textRCL08_2 = Empty
   CboRCL08 = Empty
   textRCL09 = Empty
   textRCL10 = Empty
   textRCL11 = Empty
   textRCL12 = Empty
   textRCL13 = Empty
   textRCL14 = Empty
   textRCL15 = Empty
   textRCL16 = Empty
   textRCL17 = Empty
   textRCL18 = Empty: LabRCL18_2 = Empty '要求檢查對象
   textRCL19 = Empty
   textRCL20 = Empty
   textRCL21 = Empty: LabRCL21_2 = Empty '部門
   textRCL22 = Empty: textRCL22_2 = Empty '負責同人
   textRCL23 = Empty
   textRCL24 = Empty
   cboRCL25 = Empty
   textRCL26 = Empty: textRCL26_2 = Empty '撤銷人員
   textCUID = ""
End Sub
   
' 更新 Create 及 Update 的人
Private Sub UpdateCUID(ByRef rsSrcTmp As ADODB.Recordset)
   Dim strTemp As String
   Dim strCName As String, strCDate As String, strCTime As String
   Dim strUName As String, strUDate As String, strUTime As String
   
   For i = 27 To 32
      strTemp = "" & rsSrcTmp.Fields("RCL" & Format(i, "00"))
      If strTemp <> MsgText(601) Then
         Select Case i
            Case 27, 30
               strTemp = GetStaffName(strTemp, True)
               If i = 27 Then
                  strCName = strTemp
               Else
                  strUName = strTemp
               End If
            Case 28, 31
               strTemp = Format(TAIWANDATE(strTemp), "###/##/##")
               If i = 28 Then
                  strCDate = strTemp
               Else
                  strUDate = strTemp
               End If
            Case 29, 32
               strTemp = Format(strTemp, "0#:##") 'Modify by Amy 2024/12/31 原:##:## ex:00005-RCL32,只顯示:8
               If i = 29 Then
                  strCTime = strTemp
               Else
                  strUTime = strTemp
               End If
         End Select
      End If
   Next i
   
   ' 設定CUID中的文字
   textCUID = "CREATE : " & strCName & " " & _
              " : " & strCDate & " " & _
              " : " & strCTime & String(6, " ")
   If strUName <> MsgText(601) Then
      textCUID = textCUID & "UPDATE : " & strUName & " " & _
              " : " & strUDate & " " & _
              " : " & strUTime
   End If
End Sub

' 更新各控制項的狀態
Private Sub SetCtrlReadOnly(ByVal bEnable As Boolean)
   Dim bSetEnable As Boolean
   
   bSetEnable = bEnable
   '「撤銷」欄人員且不是改自己建的,只能輸「撤銷」欄
   If bEnable = False And m_bUpdate = False And (intInputState = 1 And m_EditMode <> 1) Then
      bSetEnable = True
   End If
   
   textRCL01.Locked = bSetEnable
   textRCL02.Locked = bSetEnable
   textRCL03.Locked = bSetEnable
   textRCL04.Locked = bSetEnable
   textRCL05.Locked = bSetEnable
   textRCL06.Locked = bSetEnable
   textRCL07.Locked = bSetEnable
   'Modify by Amy 2024/02/06 國籍改下拉
   'textRCL08.Locked = bSetEnable
   CboRCL08.Locked = bSetEnable
   textRCL09.Locked = bSetEnable
   textRCL10.Locked = bSetEnable
   textRCL11.Locked = bSetEnable
   textRCL12.Locked = bSetEnable
   textRCL13.Locked = bSetEnable
   textRCL14.Locked = bSetEnable
   textRCL15.Locked = bSetEnable
   textRCL16.Locked = bSetEnable
   textRCL17.Locked = bSetEnable
   textRCL18.Locked = bSetEnable
   textRCL19.Locked = True '下次提醒日
   textRCL20.Locked = True '延展次數
   textRCL21.Locked = True '部門
   textRCL22.Locked = True '負責同仁
   textRCL23.Locked = bSetEnable
   textRCL24.Locked = True '撤銷日期
   cboRCL25.Locked = True '撤銷原因
   textRCL26.Locked = True '撤銷人員
   
   '只有電腦中心可調整 (intInputState=2)
   If bEnable = False And intInputState = 2 Then
      '下次提醒日 / 延展次數 由系統計算
      textRCL19.Locked = False '下次提醒日
      textRCL20.Locked = False '延展次數
      '部門 / 負責同仁 於新增時預帶 不可修改
      textRCL21.Locked = False '部門
      textRCL22.Locked = False '負責同仁
      textRCL26.Locked = False '撤銷人員
   End If
   '風險檢查對象可撤銷人員 未輸過[撤銷]欄 資料 or 電腦中心 才可輸 [撤銷日期] 及 [撤銷原因]
   If bEnable = False Then
      If (intInputState = 1 And m_FieldList(23).fiOldData = Empty) Or intInputState = 2 Then
         textRCL24.Locked = False
         cboRCL25.Locked = False
      End If
   End If
End Sub

' 更新Key的狀態
Private Sub SetKeyReadOnly(ByVal bEnable As Boolean)
   textRCL01.Locked = bEnable
End Sub

' 按下按鍵
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   Call PUB_SaveMeTrackMode(m_MeTrackMode, 1, KeyCode)
End Sub

'Enter 事件，等於存檔，做完取消，不然 form 內其他物件有寫 keycode 或是 keyascii 事件的話，也會做到Private Sub Form_KeyPress(KeyAscii As Integer)
Private Sub Form_KeyPress(KeyAscii As Integer)
   '備註欄按enter鍵維持換行功能而不是存檔功能
   If KeyAscii = 13 And UCase(Me.ActiveControl.Name) = UCase("textRCL23") Then
      Exit Sub
   End If
   
   Select Case KeyAscii
      Case vbKeyReturn:
         If m_EditMode <> 0 Then
            KeyAscii = 0
            OnAction vbKeyF9
         End If
   End Select
End Sub

' 執行指令
Private Sub OnAction(ByVal KeyCode As Integer)
   Dim strTit As String, strMsg As String, nResponse
   
   'Form2.0記錄鍵盤傳入順序，判斷是否可執行
   If PUB_ChkMeTrackMode(m_MeTrackMode) = False Then
        Exit Sub
   End If
   
   Select Case KeyCode
      ' 新增
      Case vbKeyF2:
         m_EditMode = 1
         ClearField
         SetCtrlReadOnly False
         ShowToolBar KeyCode
         SetInputData
      ' 修改
      Case vbKeyF3:
         m_EditMode = 2
         SetCtrlReadOnly False
         SetKeyReadOnly True
         ShowToolBar KeyCode
         SetInputData
      ' 刪除
      Case vbKeyF5:
         strTit = "詢問"
         strMsg = "是否要刪除此筆資料?"
         nResponse = MsgBox(strMsg, vbYesNo + vbCritical + vbDefaultButton2, strTit)
         If nResponse = vbYes Then
            m_EditMode = 3
            OnWork
         End If
      ' 查詢
      Case vbKeyF4:
         m_EditMode = 4
         SetCtrlReadOnly True
         SetKeyReadOnly False
         ClearField
         ShowToolBar KeyCode
         SetInputData
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
         PUB_FilterFormText Me '修正畫面所有含跳行符號的文字框
         If CheckDataValid = False Then
            Exit Sub
         End If
         UpdateFieldNewData
         OnWork
         UpdateCtrlData
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
                  ShowToolBar 0
               End If
            Case Else
               m_EditMode = 0
               UpdateCtrlData
               SetCtrlReadOnly True
               ShowToolBar 0
         End Select
      ' 離開
      Case vbKeyEscape:
         Unload Me
   End Select
End Sub

' 開始輸入資料
Private Sub SetInputData()
   Select Case m_EditMode
      Case 1 '新增
         CmdExtension.Enabled = False '新增時鎖住
         '下次提醒日=系統日+3個月後的工作天
         textRCL19 = TAIWANDATE(CompWorkDay(1, DBDATE(DateAdd("m", 3, Format(strSrvDate(1), "####/##/##"))), 1))
         '負責同仁及部門 預帶 操作人員及其 部門(鎖住)
         textRCL22 = strUserNum: textRCL22_2 = GetPrjSalesNM(textRCL22) '負責同仁
         textRCL21 = PUB_GetST93(textRCL22): LabRCL21_2.Caption = ChgType(6, textRCL21) '部門
         textRCL02.SetFocus '中文名稱
      Case 2 '修改
         textRCL02.SetFocus
         If m_bUpdate = False And intInputState = 1 Then
            textRCL24.SetFocus '撒銷日期
         End If
      Case 4 '查詢
         textRCL01.SetFocus
   End Select
End Sub

' 使用者按下確定的按紐
Private Sub OnWork()
   Dim strKey As String, strMsg As String, strTit As String, nResponse
   
   Select Case m_EditMode
      Case 1: '新增
         If ActRecord(1, strKey) = False Then Exit Sub
         If (strKey < m_FirstKEY(0)) Or (strKey > m_LastKEY(0)) Then
            RefreshRange '重新取得第一筆及最後一筆編號
         End If
         ShowCurrRecord strKey
      Case 2: '修改
         If ActRecord(2, strKey) = False Then Exit Sub
         '可撤銷人員 加 撒銷日期 則發信通知 新增人員(電腦中心拿掉撒銷日期/原因/人員 欄不自動發信)
         If intInputState = 1 And m_FieldList(23).fiOldData <> Trim(textRCL24) Then
            Call SetMailTo(2)
         End If
         ShowCurrRecord strKey
      Case 3: '刪除
         strKey = textRCL01
         If ActRecord(3, strKey) = False Then Exit Sub
         'Modify by Amy 2024/07/19 原程式改至AfterDelShowData,讓其他支程式可呼叫
         Call AfterDelShowData(strKey)
      Case 4: '查詢
         strKey = Format(textRCL01, "00000")
         If ExistCheck("RiskCheckList", "RCL01", strKey, strExc(1), False) = False Then
            strMsg = "無此資料"
            strTit = "查詢資料"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         Else
            UpdateCtrlData strKey
         End If
   End Select
   If m_EditMode = 1 Or m_EditMode = 2 Then
      'Modify by Amy 2024/04/30 bolNameExist 不使用
      'If bolNameExist = True Then
         Call SetMailTo(4)
      'End If
   End If
   m_EditMode = 0
   SetCtrlReadOnly True
End Sub

'intChoose:1-新增/2-修改/3-刪除 記錄
Private Function ActRecord(intChoose As Integer, Optional ByRef stKey As String = "") As Boolean
   
On Error GoTo ErrHand
   
   ActRecord = False
   cnnConnection.BeginTrans
  
   '刪除
   If intChoose = 3 Then
      strSql = "Delete From RiskCheckList Where RCL01= '" & stKey & "' "
   Else
      If intChoose = 1 Then
         '**** 新增 *****
         If textRCL01 = MsgText(601) Then
            '自動取號
            stKey = GetNextAutoNo("RiskCheckList", "RCL01")
            If stKey <> MsgText(601) Then
               stKey = Format(stKey, "00000")
               textRCL01 = stKey
               m_CurrKEY(0) = textRCL01
               Pub_SetFieldNewData Me.Name, m_FieldList, "RCL01", textRCL01
            Else
               cnnConnection.RollbackTrans
               ShowMsg "自動取號錯誤，請洽電腦中心!"
               Exit Function
            End If
         End If
         '**** End 新增 *****
      Else
         stKey = textRCL01
      End If
      strSql = Pub_GetFieldItemSql(intChoose, Me.Name, m_FieldList, "RiskCheckList", stKey)
   End If
   Pub_SeekTbLog strSql
   cnnConnection.Execute strSql
   cnnConnection.CommitTrans
   ActRecord = True
   Exit Function
   
ErrHand:
   cnnConnection.RollbackTrans
   If Err.Number <> 0 Then
      MsgBox "存檔有誤(" & Err.Description & ")" & vbCrLf & "請洽電腦中心！", vbCritical
   End If
End Function

Private Function ChgType(ByVal Sty As Integer, ByVal stNo As String) As String
   Dim strTmp As String, strTmp1 As String, bolMsgOnly As Boolean
   
   ChgType = ""
   Select Case Sty
      Case 1 '國籍名稱
         'Modify by Amy 2024/02/06 原:GetNationName(stNo, 0)
         strTmp = CboRCL08
         ChgType = SetCboNation(2, strTmp)
         If CboRCL08 = MsgText(601) Then
            ChgType = SetCboNation(3, strTmp)
         End If
      Case 2 '客戶名稱
         ChgType = GetCustomerName(stNo, 1)
      Case 3 '代理人名稱
         ChgType = GetPrjName1(stNo)
      Case 4 '潛在客戶名稱
         ChgType = GetPotCustomerName(1, stNo)
      Case 5 '員工姓名
         ChgType = GetPrjSalesNM(stNo)
      Case 6 '部門名稱
         ChgType = GetPrjSalesBlack(stNo, True)
   End Select
End Function

Private Function CheckDataValid() As Boolean
   Dim strTit As String, strMsg As String, bCancel As Boolean, strTmp(2) As String, nResponse
   Dim strState As String, stSameNTp As String, stCon As String, bolChkN As Boolean 'Add by Amy 2024/04/30
   
   CheckDataValid = False
   'bolNameExist = False:stSameNameData = ""'Mark by Amy 2024/04/30 不使用
   
'*** 新增或修改 ***
   If m_EditMode = 1 Or m_EditMode = 2 Then
      If textRCL02.Locked = False Then
         '中文名稱, 英文名稱, 日文名稱不可全為空白
         strTit = "檢核資料"
         If IsEmptyText(textRCL02) = True And IsEmptyText(textRCL03) = True And IsEmptyText(textRCL07) = True Then
            strMsg = "中文名稱, 英文名稱, 日文名稱不可全為空白！"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            SSTab1.Tab = 0
            textRCL02.SetFocus
            Exit Function
         End If
         '中文名稱
         If IsEmptyText(textRCL02) Then
            textRCL02_Validate bCancel
            If bCancel = True Then
               Exit Function
            End If
         End If
         '日文名稱
         If IsEmptyText(textRCL07) Then
            textRCL07_Validate bCancel
            If bCancel = True Then
               Exit Function
            End If
         End If
      End If 'textRCL02.Locked = False
      '國籍 (非必填)
      'Memo 不得代理於2020/12 外專提出資料有[無國籍]的狀況,故不控制必輸-秀玲:與不得代理同
      'Modify by Amy 國籍改下拉
      If CboRCL08.Locked = False Then
         If IsEmptyText(CboRCL08) = False Then
            CboRCL08_Validate bCancel
            If bCancel = True Then
               Exit Function
            End If
         End If
      End If
      '身份證字號/統編 (非必填)
      If textRCL17.Locked = False Then
         strTit = "身份證字號/統一編號檢查"
         strMsg = "身份證字號/統一編號"
         If IsEmptyText(textRCL17) = False Then
            Call textRCL17_Validate(bCancel)
            If bCancel = True Then Exit Function
         End If
         '客戶檔已有一樣的ID不可存檔-文雄
         strTmp(0) = ""
         'Modify by Amy 2024/06/13 +Me.Name
         If ChkCU11Same("", "", textRCL17, strTmp(0), 1, Me.Name) = True And m_FieldList(16).fiOldData <> Trim(textRCL17) Then
            strMsg = strMsg & " [" & textRCL17 & "] 與" & IIf(InStr(strTmp(0), ",") > 0, vbCrLf, "") & " [" & strTmp(0) & "] 相同" & vbCrLf & vbCrLf & _
                              "不可存檔"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            Exit Function
         End If
      End If
      '要求檢查對象(必填)
      If textRCL18.Locked = False Then
         strTit = "檢核資料"
         If IsEmptyText(textRCL18) = True Then
            strMsg = "要求檢查對象不可為空白！"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            SSTab1.Tab = 0
            textRCL18.SetFocus
            Exit Function
         Else
            Call textRCL18_Validate(bCancel)
            If bCancel = True Then Exit Function
         End If
      End If
      '備註(必填)
      If textRCL23.Locked = False Then
         strTit = "檢核資料"
         If IsEmptyText(textRCL23) = True Then
            strMsg = "備註不可空白！"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            SSTab1.Tab = 0
            textRCL23.SetFocus
            Exit Function
         End If
      End If
      
'****** 檢查名稱資料重覆 (最後檢查,彈訊息存檔後發mail) ******
      If textRCL02.Locked = False Then
         Call DelR12040163
         strTit = "名稱相同檢查": strMsg = ""
         '英文名稱
         strTmp(1) = Trim(textRCL03) & Trim(textRCL04) & Trim(textRCL05) & Trim(textRCL06)
         strTmp(2) = Trim(m_FieldList(2).fiOldData) & Trim(m_FieldList(3).fiOldData) & Trim(m_FieldList(4).fiOldData) & Trim(m_FieldList(5).fiOldData)
         'Modify by Amy 2024/04/30 stSameNameData 不使用
         '新增時
         If m_EditMode = 1 Then
            bolChkN = True
            If Trim(textRCL02) <> MsgText(601) Then strTmp(0) = strTmp(0) & "★中-" & Trim(textRCL02)
            If Trim(strTmp(1)) <> MsgText(601) Then strTmp(0) = strTmp(0) & "★英-" & strTmp(1)
            If Trim(textRCL07) <> MsgText(601) Then strTmp(0) = strTmp(0) & "★日-" & Trim(textRCL07)
         '修改時
         Else
            stCon = " And RCL01<>'" & textRCL01 & "' "
            If Trim(textRCL02) <> MsgText(601) And m_FieldList(1).fiOldData <> Trim(textRCL02) Then strTmp(0) = strTmp(0) & "★中-" & Trim(textRCL02)
            If Trim(strTmp(1)) <> MsgText(601) And strTmp(2) <> Trim(strTmp(1)) Then strTmp(0) = strTmp(0) & "★英-" & strTmp(1)
            If Trim(textRCL07) <> MsgText(601) And m_FieldList(6).fiOldData <> Trim(textRCL07) Then strTmp(0) = strTmp(0) & "★日-" & Trim(textRCL07)
            '名稱有修改才需檢查
            If strTmp(0) <> MsgText(601) Then bolChkN = True
         End If
         If bolChkN = True Then
            strTmp(0) = Mid(strTmp(0), 2)
            '名稱相同檢查
            If ChkNameExist(0, m_EditMode, strUserNum, strTmp(0), strMsg, stCon, stSameNTp, strState) = True Then
               If InStr(strMsg, "不可重覆建立") > 0 Then
                  nResponse = MsgBox(strMsg, vbOKOnly, strTit)
                  Exit Function
               Else
                  nResponse = MsgBox(strMsg, vbYesNo + vbCritical + vbDefaultButton2, strTit)
                  If nResponse = vbNo Then Exit Function
               End If
            End If
            '原 與客戶／代理人／潛在客戶同名同姓且存檔者通知特定人員,改依stMailState不同發信
            If stSameNTp <> MsgText(601) And strState = "8.1" Then
               '已是風險檢查對象,需再檢查一次 for最後發信
               Call ChkNameExist(2, m_EditMode, strUserNum, strTmp(0), , , stSameNTp, strState)
            End If 'stSameNTp <> MsgText(601)
         End If 'bolChkN = True
         'end 2024/04/30
      End If
'****** End 檢查名稱資料重覆 ******
   
'****** 風險檢查對象可撤銷人員 才可輸 [撤銷日期] 及 [撤銷原因] ******
      '撤銷日期或撤銷原因必須同時輸入或同時不輸
      If textRCL24.Locked = False Then
         'Memo 撒銷日期/原因 有輸 系統會自動帶[撒銷人員]
         If (IsEmptyText(textRCL24) = False And m_FieldList(23).fiOldData <> Trim(textRCL24)) _
           Or (IsEmptyText(cboRCL25) = False And m_FieldList(24).fiOldData <> Trim(cboRCL25)) Then
            strTit = "檢核資料"
            '撤銷日期
            If IsEmptyText(textRCL24) = True Then
               strMsg = "有撤銷原因，撤銷日期不可空白！"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               SSTab1.Tab = 0
               textRCL24.SetFocus
               Exit Function
            Else
               Call textRCL24_Validate(bCancel)
               If bCancel = True Then Exit Function
            End If
            '撤銷原因
            If IsEmptyText(cboRCL25) = True Then
               strMsg = "有撤銷日期，撤銷原因不可空白！"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               SSTab1.Tab = 0
               cboRCL25.SetFocus
               Exit Function
            End If
         End If
      End If
'****** End 風險檢查對象可撤銷人員 才可輸 [撤銷日期] 及 [撤銷原因] ******

'****** 只有電腦中心可操作 ******
      '撤銷日期 改為空 撒銷人員也要改為空
      If textRCL26.Locked = False Then
         strTit = "檢核資料"
         If IsEmptyText(textRCL24) = False And IsEmptyText(textRCL26) = True Then
            strMsg = "[撤銷日期]為空白" & vbCrLf & "[撤銷人員]不可有值！"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            SSTab1.Tab = 0
            textRCL26.SetFocus
            Exit Function
         End If
      End If
      '下次提醒日
      If textRCL19.Locked = False Then
         strTit = "檢核資料"
         If IsEmptyText(textRCL19) = True Then
            strMsg = "下次提醒日不可空白！"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            SSTab1.Tab = 0
            textRCL19.SetFocus
            Exit Function
         Else
            Call textRCL19_Validate(bCancel)
            If bCancel = True Then Exit Function
         End If
      End If
      '延展次數 (新增時,下次提醒日由系統帶,延展次數為null,按過[延展] 鈕 延展次數才有值)
      If textRCL20.Locked = False Then
         strTit = "檢核資料"
         If IsEmptyText(textRCL20) = False Then
            Call textRCL20_Validate(bCancel)
            If bCancel = True Then Exit Function
         End If
      End If
      '部門
      If textRCL21.Locked = False Then
         If IsEmptyText(textRCL21) = True Then
            strTit = "檢核資料"
            strMsg = "部門不可空白！"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            SSTab1.Tab = 0
            textRCL21.SetFocus
            Exit Function
         Else
            Call textRCL21_Validate(bCancel)
            If bCancel = True Then Exit Function
         End If
      End If
      '負責同仁
      If textRCL22.Locked = False Then
         If IsEmptyText(textRCL22) = True Then
            strTit = "檢核資料"
            strMsg = "負責同仁不可空白！"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            SSTab1.Tab = 0
            textRCL22.SetFocus
            Exit Function
         End If
      Else
         Call textRCL22_Validate(bCancel)
         If bCancel = True Then Exit Function
      End If
'****** End 只有電腦中心可操作 ******
'*** 查詢 ***
   ElseIf m_EditMode = 4 Then
      '編號不可空白
      strTit = "檢核資料"
      If IsEmptyText(textRCL01) = True Then
         strMsg = "請輸入編號"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         SSTab1.Tab = 0
         textRCL01.SetFocus
         Exit Function
      End If
   End If
         
   '檢查畫面的 TextBox,是否含有Unicode文字
   If PUB_ChkUniText(Me, , True, "TextBox") = False Then
       Exit Function
   End If
      
   CheckDataValid = True
End Function

Private Sub Form_Unload(Cancel As Integer)
   Set frm12040163 = Nothing
End Sub

Private Sub textRCL01_GotFocus()
   InverseTextBox textRCL01
End Sub

Private Sub textRCL02_GotFocus()
   InverseTextBox textRCL02
   OpenIme
End Sub

Private Sub textRCL03_GotFocus()
   InverseTextBox textRCL03
End Sub

Private Sub textRCL04_GotFocus()
   InverseTextBox textRCL04
End Sub

Private Sub textRCL05_GotFocus()
   InverseTextBox textRCL05
End Sub

Private Sub textRCL06_GotFocus()
   InverseTextBox textRCL06
End Sub

Private Sub textRCL07_GotFocus()
   InverseTextBox textRCL07
   OpenIme
End Sub

Private Sub textRCL09_GotFocus()
   InverseTextBox textRCL09
   OpenIme
End Sub

Private Sub textRCL10_GotFocus()
   InverseTextBox textRCL10
End Sub

Private Sub textRCL11_GotFocus()
   InverseTextBox textRCL11
End Sub

Private Sub textRCL12_GotFocus()
   InverseTextBox textRCL12
End Sub

Private Sub textRCL13_GotFocus()
   InverseTextBox textRCL13
End Sub

Private Sub textRCL14_GotFocus()
   InverseTextBox textRCL14
End Sub

Private Sub textRCL15_GotFocus()
   InverseTextBox textRCL15
End Sub

Private Sub textRCL16_GotFocus()
   InverseTextBox textRCL16
   OpenIme
End Sub

Private Sub textRCL17_GotFocus()
   InverseTextBox textRCL17
End Sub

Private Sub textRCL18_GotFocus()
   InverseTextBox textRCL18
End Sub

Private Sub textRCL19_GotFocus()
   InverseTextBox textRCL19
End Sub

Private Sub textRCL20_GotFocus()
   InverseTextBox textRCL20
   OpenIme
End Sub

Private Sub textRCL21_GotFocus()
   InverseTextBox textRCL21
End Sub

Private Sub textRCL22_GotFocus()
   InverseTextBox textRCL22
End Sub

Private Sub textRCL23_GotFocus()
   InverseTextBox textRCL23
End Sub

Private Sub textRCL24_GotFocus()
   InverseTextBox textRCL24
End Sub

Private Sub textRCL26_GotFocus()
   InverseTextBox textRCL26
End Sub

Private Sub textRCL02_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = ChangeZIP(KeyAscii, textRCL02)
End Sub

Private Sub textRCL07_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = ChangeZIP(KeyAscii, textRCL07)
End Sub

Private Sub CboRCL08_KeyPress(KeyAscii As Integer)
   KeyAscii = Pub_NumAscii(KeyAscii)
End Sub

Private Sub textRCL09_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = ChangeZIP(KeyAscii, textRCL09)
End Sub

'日文地址要轉全形
Private Sub textRCL16_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = ChangeZIP(KeyAscii, textRCL16)
End Sub

Private Sub textRCL17_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textRCL18_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textRCL20_KeyPress(KeyAscii As Integer)
   KeyAscii = Pub_NumAscii(KeyAscii)
End Sub

Private Sub textRCL21_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textRCL22_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textRCL26_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

'名稱(中)
Private Sub textRCL02_Validate(Cancel As Boolean)
   Dim strTit As String, strMsg As String, nResponse
   
   If Trim(textRCL02) = MsgText(601) Then Exit Sub
   If m_EditMode <> 1 And m_EditMode <> 2 Then Exit Sub
   
   strTit = "檢核資料"
   strMsg = "名稱(中)"
   If StrLength(textRCL02) > textRCL02.MaxLength Then
      Cancel = True
      nResponse = MsgBox(strMsg & "內容太長", vbOKOnly, strTit)
      SSTab1.Tab = 0
      textRCL02_GotFocus
      Exit Sub
   End If
End Sub

'名稱(日)
Private Sub textRCL07_Validate(Cancel As Boolean)
   Dim strTit As String, strMsg As String, nResponse
   
   If Trim(textRCL07) = MsgText(601) Then Exit Sub
   If m_EditMode <> 1 And m_EditMode <> 2 Then Exit Sub
   
   strTit = "檢核資料"
   strMsg = "名稱(日)"
   If StrLength(textRCL07) > textRCL07.MaxLength Then
      Cancel = True
      nResponse = MsgBox(strMsg & "內容太長", vbOKOnly, strTit)
      SSTab1.Tab = 0
      textRCL07_GotFocus
      Exit Sub
   End If
End Sub

'地址(中)
Private Sub textRCL09_Validate(Cancel As Boolean)
   Dim strTit As String, strMsg As String, nResponse
   
   If Trim(textRCL09) = MsgText(601) Then Exit Sub
   If m_EditMode <> 1 And m_EditMode <> 2 Then Exit Sub
   
   strTit = "檢核資料"
   strMsg = "地址(中)"
   If StrLength(textRCL09) > textRCL09.MaxLength Then
      Cancel = True
      nResponse = MsgBox(strMsg & "內容太長", vbOKOnly, strTit)
      SSTab1.Tab = 1
      textRCL09_GotFocus
      Exit Sub
   End If
End Sub

'地址(日)
Private Sub textRCL16_Validate(Cancel As Boolean)
   Dim strTit As String, strMsg As String, nResponse
   
   If Trim(textRCL09) = MsgText(601) Then Exit Sub
   If m_EditMode <> 1 And m_EditMode <> 2 Then Exit Sub
   
   strTit = "檢核資料"
   strMsg = "地址(日)"
   If StrLength(textRCL16) > textRCL16.MaxLength Then
      Cancel = True
      nResponse = MsgBox(strMsg & "內容太長", vbOKOnly, strTit)
      SSTab1.Tab = 1
      textRCL16_GotFocus
      Exit Sub
   End If
End Sub

'身份證/統編(檢查同客戶檔)
Private Sub textRCL17_Validate(Cancel As Boolean)
   Dim strMsg As String, ii As Integer
   
   If Trim(textRCL17) = MsgText(601) Then Exit Sub
   If m_EditMode <> 1 And m_EditMode <> 2 Then Exit Sub
   If m_FieldList(16).fiOldData = Trim(textRCL17) Then Exit Sub
 
   'Memo 按了「是」Focus還在textRCL17,會一直觸發上面「檢查有誤,…」訊息無法跳離,故傳textRCL18 跳至此
   'Modify by Amy 2024/02/06 國籍改下拉
   If Pub_CheckIDAll(0, Me.Name, Trim(textRCL17), CboRCL08.Text, , textRCL18) = False Then
      Cancel = True
      SSTab1.Tab = 0
      'textRCL17_GotFocus
      Exit Sub
   End If
End Sub

'要求檢查對象
Private Sub textRCL18_Validate(Cancel As Boolean)
   Dim strTit As String, strMsg As String, nResponse
   Dim strNo As String, strName As String, strTp As String
   
   LabRCL18_2.Caption = Empty
   If textRCL18 = MsgText(601) Then Exit Sub
   
   strMsg = Mid(Label1(1).Caption, 1, Len(Label1(1).Caption) - 1)
   strTit = "檢核資料"
   If Left(textRCL18, 1) <> 客戶編號 And Left(textRCL18, 1) <> 代理人編號 And Left(textRCL18, 1) <> "R" Then
      Cancel = True
      strMsg = strMsg & "欄位" & vbCrLf & _
                        "只可輸入客戶 or 代理人 or 潛在客戶 編號"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      SSTab1.Tab = 0
      textRCL18_GotFocus
      Exit Sub
   End If
   
   strNo = ChangeCustomerL(textRCL18)
   Select Case Left(strNo, 1)
      Case 客戶編號
         strTp = "客戶編號"
         strName = ChgType(2, strNo)
      Case 代理人編號
         strTp = "代理人編號"
         strName = ChgType(3, strNo)
      Case "R"
         strTp = "潛在客戶編號"
         strName = ChgType(4, strNo)
   End Select
   
   If strName <> MsgText(601) Then
      strNo = Left(strNo, 8)
      LabRCL18_2 = strName
   Else
      strNo = ""
   End If
   If m_EditMode <> 1 And m_EditMode <> 2 Then Exit Sub
   If m_FieldList(17).fiOldData = Trim(textRCL18) Then Exit Sub
   
   If strNo = MsgText(601) And Trim(textRCL18) <> MsgText(601) Then
      Cancel = True
      strMsg = "[" & strMsg & "]欄位" & vbCrLf & _
                        " 輸入的" & strTp & "不存在，請確認！"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      SSTab1.Tab = 0
      textRCL18_GotFocus
      Exit Sub
   ElseIf strNo <> MsgText(601) Then
      textRCL18 = strNo
   End If
   
End Sub

'下次提醒日(電腦中心用)
Private Sub textRCL19_Validate(Cancel As Boolean)
   Dim strTit As String, strMsg As String, nResponse
   
   If Trim(textRCL19) = MsgText(601) Then Exit Sub
   If m_EditMode <> 1 And m_EditMode <> 2 Then Exit Sub
   If m_FieldList(18).fiOldData = Trim(textRCL19) Then Exit Sub
   
   strTit = "檢核資料"
   strMsg = "下次提醒日"
   If CheckIsTaiwanDate(textRCL19, False) = False Then
      Cancel = True
      nResponse = MsgBox(strMsg & "格式有誤", vbOKOnly, strTit)
      SSTab1.Tab = 0
      textRCL19_GotFocus
      Exit Sub
   End If
End Sub

'延展次數(電腦中心用)
Private Sub textRCL20_Validate(Cancel As Boolean)
   Dim strTit As String, strMsg As String, nResponse
   
   If Trim(textRCL20) = MsgText(601) Then Exit Sub
   If m_EditMode <> 1 And m_EditMode <> 2 Then Exit Sub
   If m_FieldList(19).fiOldData = Trim(textRCL20) Then Exit Sub
   
   strTit = "檢核資料"
   strMsg = "延展次數"
   If IsNumeric(textRCL20) = False Then
      Cancel = True
      nResponse = MsgBox(strMsg & "只能輸入數字", vbOKOnly, strTit)
      SSTab1.Tab = 0
      textRCL20_GotFocus
      Exit Sub
   End If
End Sub

'部門-記錄新增時負責同仁的 [部門] (電腦中心用)
Private Sub textRCL21_Validate(Cancel As Boolean)
   Dim strTit As String, strMsg As String, nResponse
   
   LabRCL21_2.Caption = Empty
   If Trim(textRCL21) = MsgText(601) Then Exit Sub
   
   LabRCL21_2.Caption = ChgType(6, textRCL21)
   If m_EditMode <> 1 And m_EditMode <> 2 Then Exit Sub
   If m_FieldList(20).fiOldData = Trim(textRCL21) Then Exit Sub
   
   strTit = "檢核資料"
   strMsg = "部門"
   If IsEmptyText(LabRCL21_2) = True Then
      Cancel = True
      nResponse = MsgBox(strMsg & "輸入不正確", vbOKOnly, strTit)
      SSTab1.Tab = 0
      textRCL21_GotFocus
      Exit Sub
   End If
End Sub

'負責同仁-記錄新增時 [負責同仁] (電腦中心用)
Private Sub textRCL22_Validate(Cancel As Boolean)
   Dim strTit As String, strMsg As String, nResponse
   
   textRCL22_2 = Empty
   If Trim(textRCL22) = MsgText(601) Then Exit Sub
   
   textRCL22_2 = ChgType(5, textRCL22)
   If m_EditMode <> 1 And m_EditMode <> 2 Then Exit Sub
   If m_FieldList(21).fiOldData = Trim(textRCL22) Then Exit Sub
   
   strTit = "檢核資料"
   strMsg = "負責同仁"
   If IsEmptyText(textRCL22_2) = True Then
      Cancel = True
      nResponse = MsgBox(strMsg & "輸入不正確", vbOKOnly, strTit)
      SSTab1.Tab = 0
      textRCL22_GotFocus
      Exit Sub
   End If
End Sub

'撤銷日期
Private Sub textRCL24_Validate(Cancel As Boolean)
   Dim strTit As String, strMsg As String, nResponse
   
   If Trim(textRCL24) = MsgText(601) Then Exit Sub
   If m_EditMode <> 1 And m_EditMode <> 2 Then Exit Sub
   If m_FieldList(23).fiOldData = Trim(textRCL24) Then Exit Sub
   
   strTit = "檢核資料"
   strMsg = "撤銷日期"
   If CheckIsTaiwanDate(textRCL24, False) = False Then
      Cancel = True
      nResponse = MsgBox(strMsg & "格式有誤", vbOKOnly, strTit)
      SSTab1.Tab = 0
      textRCL24_GotFocus
      Exit Sub
   End If
   
End Sub

'撤銷人員-電腦中心用
Private Sub textRCL26_Validate(Cancel As Boolean)
   Dim strTit As String, strMsg As String, nResponse
   
   textRCL26_2 = Empty
   If Trim(textRCL26) = MsgText(601) Then Exit Sub
   
   textRCL26_2 = ChgType(5, textRCL26)
   If m_EditMode <> 1 And m_EditMode <> 2 Then Exit Sub
   If m_FieldList(25).fiOldData = Trim(textRCL26) Then Exit Sub
   
   strTit = "檢核資料"
   strMsg = "撤銷人員"
   '撒銷日期有修改且不為空,撤銷人員不可為空
   If m_FieldList(23).fiOldData <> Trim(textRCL24) And IsEmptyText(textRCL24) = False And IsEmptyText(textRCL26_2) = True Then
      Cancel = True
      nResponse = MsgBox(strMsg & "不可為空", vbOKOnly, strTit)
      SSTab1.Tab = 0
      textRCL26_GotFocus
      Exit Sub
   End If
End Sub

Private Sub Tbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
   
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

'Mark by Amy 2023/03/11
'傳入查詢名稱,確認是否於客戶/代理人/潛在客戶/對造資料是否存在
'm_EditMode:目前狀態 / stPNo:人員編號 / stName:名稱(多筆:中/英/日+[-]★區隔) / stMsg:回傳訊息 / stCon:其他條件
'stBackTxt as string:回傳其他資料
Private Function ChkNameExist_Old(m_EditMode As Integer, stPNo As String, stName As String, ByRef stMsg As String, Optional ByVal stCon As String = "", _
  Optional ByRef stBackTxt As String = "") As Boolean
'   Dim rsA As New ADODB.Recordset, strA As String, intA As Integer, ii As Integer
'   Dim strTmp(5) As String, strBack As String, strFindTxt As String, arrTmp
'
'   ChkNameExist = False
'   stBackTxt = ""
'
'   '風險檢查對象資料中是否已有資料
'   Call ChkRiskData(2, Me.Name, , , stName, strTmp(1), ",Decode(RCL27,'" & stPNo & "',1,2) as Sort", m_EditMode, stCon)
'   If strTmp(1) <> MsgText(601) Then
'      stMsg = Replace(strTmp(1), "<br>", vbCrLf)
'      If InStr(stMsg, "不可重覆建立") = 0 Then
'         stMsg = stMsg & vbCrLf & "欲重覆建,請按「是」繼續存檔"
'      End If
'      ChkNameExist = True
'      Exit Function
'   End If
'
'   '共同查詢名稱(不檢查[風險檢查對象]-上面已檢查)
'   arrTmp = Split(stName, "★")
'   For ii = LBound(arrTmp) To UBound(arrTmp)
'      strTmp(0) = arrTmp(ii)
'      strFindTxt = Mid(strTmp(0), Val(InStr(strTmp(0), "-")) + 1)
'      strTmp(0) = Replace(strTmp(0), "-" & strFindTxt, "") '顯示 中/英/日 哪個欄位
'      'Memo by Amy 語法抓共用名稱查詢,若中/英/日寫成一句怕超過字串字數限制
'      strA = "Select Distinct FNo,FName,MailAddr as CaseNo,SField From(" & GetSearchNameSql(Me.Name, strFindTxt, "=", True, True) & ")"
'      intA = 1
'      Set rsA = ClsLawReadRstMsg(intA, strA)
'      If intA = 1 Then
'         rsA.MoveFirst
'         Do While rsA.EOF = False
'            strTmp(2) = "" & rsA.Fields("SFIELD")
'            Select Case strTmp(2)
'               Case "1"
'                  strTmp(2) = "中"
'               Case "2"
'                  strTmp(2) = "英"
'               Case "3"
'                  strTmp(2) = "日"
'            End Select
'            If InStr(strTmp(4), "<br>名稱：" & strFindTxt) = 0 Then
'               strTmp(4) = "<br>名稱：" & strFindTxt & vbCrLf
'               strTmp(1) = strTmp(1) & strTmp(4)
'            End If
'            If rsA.Fields("FNO") = "對造" Then
'               strTmp(3) = rsA.Fields("CaseNo")
'               strTmp(1) = strTmp(1) & "已存在[對造]案號 [" & strTmp(3) & "] " & strTmp(2) & "文欄位中" & vbCrLf
'            Else
'               strTmp(1) = strTmp(1) & "已存在本所編號 [" & rsA.Fields("FNO") & "] " & strTmp(2) & "文欄位中" & vbCrLf
'               '發給 風險檢查對象可撤銷人員 的信件,只需顯示 客戶／代理人／潛在客戶 資訊
'               'Modify by Amy 2023/02/02 不得代理也不寄
'               If rsA.Fields("FNO") <> "R" And Len(rsA.Fields("FNO")) <> 3 Then
'                  If InStr(strTmp(5), "<br>名稱：" & strFindTxt) = 0 Then
'                     strTmp(5) = "<br>名稱：" & strFindTxt & vbCrLf
'                     strBack = strBack & strTmp(5)
'                  End If
'                  strBack = strBack & "已存在本所編號 [" & rsA.Fields("FNO") & "] " & strTmp(2) & "文欄位中" & vbCrLf
'               End If
'            End If
'            rsA.MoveNext
'         Loop
'         If strTmp(1) <> MsgText(601) Then
'            stMsg = Replace(Replace(Mid(strTmp(1), 5), ",名稱：", vbCrLf & "名稱："), "<br>", vbCrLf) & _
'                              vbCrLf & "待確認或尚未進行利益衝突協商" & vbCrLf & _
'                           "請按[否],回前畫面"
'         End If
'         If strBack <> MsgText(601) Then
'            stBackTxt = Replace(Replace(Mid(strBack, 5), ",名稱：", vbCrLf & "名稱："), "<br>", vbCrLf)
'         End If
'         ChkNameExist = True
'      End If
'   Next ii
'   Set rsA = Nothing
End Function

'通知寄信人員
'intChoose:1-延展 鈕 / 2-操作撤銷日期  / 3-不撒銷通知 通知信 / 4-客戶、代理人、潛在客戶同名同姓且存檔者通知特定人員
Private Sub SetMailTo(ByVal intChoose As Integer)
   Dim stTO As String, stCC As String, stSubject As String, stContext As String
   Dim stTmp As String, stUpdN As String, stCrtN As String, ArrMail
   
   '名稱 順序:中->英->日
   If Trim(textRCL02) <> MsgText(601) Then
      stTmp = Trim(textRCL02)
   ElseIf Trim(textRCL03) <> MsgText(601) Then
      stTmp = Trim(textRCL03) & Trim(textRCL04) & Trim(textRCL05) & Trim(textRCL06)
   Else
      stTmp = Trim(textRCL07)
   End If
   stCrtN = GetPrjSalesNM(m_RCL27)
   Select Case intChoose
      Case 1
      '****** 延展 ******
         stSubject = "貴屬已延展原設定風險檢查對象之管制期限，請知悉。"
         stContext = "貴屬 [" & stCrtN & "] 已延展原設定風險檢查對象(" & textRCL01 & " " & stTmp & ")之管制期限，" & vbCrLf & _
                                "若認為已無管制必要，欲取消延展，請通知【可撤銷人員】處理。"
                                
         '建立者自行操作,通知抓順序 st52->a0924(區主管)
         If strUserNum = m_RCL27 Then
            '區主管 操作自己的延展,發給 [最高]主管
            If strUserNum = strA0924 Then
               '已是部門[最高]主管
               If strSupMan = strA0924 Then
                  stTO = Pub_GetSpecMan("總經理員工編號")
               Else
                  stTO = strSupMan
               End If
               stSubject = Replace(stSubject, "貴屬", "[" & stCrtN & GetStaffST20(strUserNum) & "]")
               stContext = Replace(stContext, "貴屬 [" & stCrtN & "] ", "[" & stCrtN & GetStaffST20(strUserNum) & "]")
               '若最高主管為總經理,cc給[險檢查對象可撤銷人員]-給總經理的信原就會cc 給美珍和文雄
               If strSupMan = Pub_GetSpecMan("總經理員工編號") Then
                  stCC = Pub_GetSpecMan("風險檢查對象可撤銷人員")
               End If
            '沒有第2~4級主管,發給 [區]主管
            ElseIf strSt52List = MsgText(601) And m_RCL27 <> strA0924 Then
               stTO = strA0924
            '發給 [2級]主管
            Else
               ArrMail = Split(strSt52List, ",")
               stTO = ArrMail(0)
            End If
         '智權部最高主管代為操作智權部[區主管]之延展 or [非]智權部最高主管(總經理)或[風險檢查對象可撤銷人員],發信通知建立者
         Else
            stUpdN = "[" & GetPrjSalesNM(strUserNum) & GetStaffST20(strUserNum) & "] "
            If (PUB_GetST93(m_RCL27) = "S" And InStr(strSupMan, strUserNum) > 0) _
              Or (PUB_GetST93(m_RCL27) <> "S" And (intInputState = 1 Or InStr(strSupMan, strUserNum) > 0)) Then
               stTO = m_RCL27
               stSubject = Replace(stSubject, "貴屬", "您")
               stContext = Replace(stContext, "貴屬 [" & stCrtN & "] 已", stUpdN & "已替您")
            '[非]建立者自行操作,發信通知建立者
            Else
               stTO = m_RCL27
               stSubject = Replace(stSubject, "貴屬", "您")
               stContext = Replace(stContext, "貴屬 [" & stCrtN & "] 已", stUpdN & "已替您")
            End If
         End If
      '****** End 延展 ******
      Case 2
      '****** 操作撤銷日期 ******
         If cboRCL25 = "已收文" Then
            stSubject = "設定風險檢查對象已取消，請知悉。"
            stContext = "原設定風險檢查對象 (" & textRCL01 & " " & stTmp & ") 因有關聯案件收文，" & vbCrLf & _
                                    "故取消風險檢查對象設定。"
         '逾管制期限未收文
         Else
            stSubject = "原設定風險檢查對象已取消管制，請知悉。"
            stContext = "您先前設定風險檢查對象 (" & textRCL01 & " " & stTmp & ") " & vbCrLf & _
                                    "已取消設定，請知悉。"
         End If
         'Memo 文雄自已建的資料,自已撤銷仍發信給自己-秀玲:仍發可留下記錄
         stTO = m_RCL27
      '****** End 操作撤銷日期 ******
      Case 3
      '****** 不撒銷通知 ******
         stSubject = "原設定風險檢查對象若有轉為不得代理名單，請為後續之處理。"
         stContext = "原設定風險檢查對象 (" & textRCL01 & " " & stTmp & ") 雖有關聯案件收文，仍設定為風險檢查對象，" & vbCrLf & _
                                 "若該風險檢查對象有列入不得代理名單之必要，請依「不得代理名單」簽核流程呈報核可。 "
         stTO = m_RCL27
      '****** End 不撒銷通知 ******
      Case 4
      '****** 同名同姓且存檔者通知特定人員 ******
         'Modify by Amy 2024/04/30 原為「客戶、代理人、潛在客戶同名同姓且存檔者通知特定人員」,改為依條件通知不同人員,且原通知人員設為副本
'         strExc(1) = Mid(stSameNameData, 1, InStr(stSameNameData, ";"))
'         stSameNameData = Replace(stSameNameData, strExc(1), "")
'         stSubject = Left(strExc(1), Len(strExc(1)) - 1) & "風險檢查對象(編號:" & textRCL01 & ")與本所既有客戶/代理人/潛在客戶同名同姓，請確認並為必要處理。"
'         stContext = stSameNameData
'         stTO = Pub_GetSpecMan("風險檢查對象可撤銷人員")
         Call SetSameNameMail
         Exit Sub
         'end 2024/04/30
      '****** End 同名同姓且存檔者通知特定人員 ******
   End Select
   If stTO = MsgText(601) Then stTO = Pub_GetSpecMan("程式管理人員")
   PUB_SendMail strUserNum, stTO, "", stSubject, stContext, , , , , , stCC
End Sub

'Add by Amy 2024/02/06 國籍改下拉 同接洽單
'intChoose:1-設定選單/2-以名稱查資料/3-以編號查資料
Private Function SetCboNation(intChoose As Integer, stText As String) As String
   Dim rsA As New ADODB.Recordset, intA As Integer, sta As String, stWhere As String, stText1 As String
   
   SetCboNation = ""
   Select Case intChoose
      Case 1
         stWhere = " And NA01 not in('224')"
      Case 2
         stWhere = " And (na01||' '||na03 like '%" & stText & "%')"
         stText1 = PUB_ConvFullChar(stText)
         If stText1 <> stText Then
            stWhere = stWhere & " Or na01||' '||na03 like '%" & stText1 & "%'"
         End If
      Case 3
         stWhere = " And Na01='" & stText & "'"
   End Select
   sta = "Select NA01, NA03 From Nation Where Length(NA01)=3 And NA01<='9999' And substr(NA02,3,1)='0'" & _
                stWhere
   If intChoose = 1 Then
      If Me.CboRCL08.ListIndex >= 0 Then
         Me.CboRCL08.Tag = Me.CboRCL08.List(Me.CboRCL08.ListIndex)
      End If
      sta = sta & " Order By NA01"
      Me.CboRCL08.Clear
   Else
      CboRCL08.Text = ""
   End If
   intA = 1
   Set rsA = ClsLawReadRstMsg(intA, sta)
   If intA = 1 Then
      rsA.MoveFirst
      Do While rsA.EOF = False
         If intChoose = 1 Then
            Me.CboRCL08.AddItem "" & rsA.Fields(0).Value & " " & rsA.Fields(1).Value
         Else
            SetCboNation = "" & rsA.Fields(0).Value & " " & rsA.Fields(1).Value
         End If
         rsA.MoveNext
      Loop
   End If
   If stText <> "" Then
      If Me.CboRCL08.ListCount = 1 Then
         Me.CboRCL08.ListIndex = 0
         Me.CboRCL08.Tag = Me.CboRCL08.Text
      Else
         Me.CboRCL08 = stText
      End If
   End If
   Set rsA = Nothing
End Function

'Modify by Amy 2024/04/30
'傳入查詢名稱,確認資料是否存在,並依規則回傳訊息及狀態,以利後續發信
'intChoose:0-全查 / 1-只查風險檢查對象資料 / 2-不查風險檢查對象資料
'm_EditMode:目前狀態 / stPNo:人員編號 / stName:名稱(多筆:中/英/日+[-]★區隔) / stMsg:回傳訊息 / stCon:其他條件 / stBackTxt :回傳其他資料
'stBackState:回傳狀態
'[不]發mail：9.1-不得代理 / 9.2-對造 (X編號同且狀態為[設為對造]身份證或統編與風險檢查對象相[同]) / X編號同且身份證或統編與風險檢查對象相[同] (無法存檔)
'[要]發mail：名稱與 本所客戶 X編號同 1.1-[無]身份證或統編 / 1.2-身份證或統編與風險檢查對象相[不同]
'                      名稱與 本所客戶 R編號同 2.1-名稱字數為[個人] / 2.2-非[個人]
'                      名稱與 本所客戶 Y編號同 3.2
'                      名稱與 本所客戶 4.1-X編號聯絡人同 / 4.2-R編號聯絡人同 / 4.3-Y編號聯絡人同
'                      名稱 8.1-已是風險檢查對象 / 8.2-與 待活化客戶 同
'                      8.9-非本所客戶
Private Function ChkNameExist(intChoose As Integer, m_EditMode As Integer, stPNo As String, stName As String, Optional ByRef stMsg As String, Optional ByVal stCon As String = "", _
  Optional ByRef stBackTxt As String, Optional ByRef stBackState As String) As Boolean
   Dim rsA As New ADODB.Recordset, strA As String, intA As Integer, ii As Integer, arrTmp
   Dim strFindTxt As String, strNowState As String, strCmd As String, stOrder As String, strTmp(5) As String
   
   ChkNameExist = False
   stBackTxt = "": stBackState = ""
   
   If intChoose = 0 Or intChoose = 1 Then
      '風險檢查對象資料中是否已有資料
      Call ChkRiskData(2, Me.Name, , , stName, strTmp(1), ",Decode(RCL27,'" & stPNo & "',1,2) as Sort", m_EditMode, stCon)
      If strTmp(1) <> MsgText(601) Then
         stMsg = Replace(strTmp(1), "<BR>", vbCrLf)
         If InStr(stMsg, "不可重覆建立") = 0 Then
            stMsg = stMsg & vbCrLf & "欲重覆建,請按「是」繼續存檔"
            stBackState = "8.1" '已是風險檢查對象
         End If
         ChkNameExist = True
         Exit Function
      End If
   End If
   
   arrTmp = Split(stName, "★")
   If intChoose = 0 Or intChoose = 2 Then
      '共同查詢名稱(不檢查[風險檢查對象]-上面已檢查)
      For ii = LBound(arrTmp) To UBound(arrTmp)
         strNowState = "8.9" '非本所客戶-預設
         strTmp(0) = arrTmp(ii)
         strFindTxt = Mid(strTmp(0), Val(InStr(strTmp(0), "-")) + 1)
         strTmp(0) = Replace(strTmp(0), "-" & strFindTxt, "") '顯示 中/英/日 哪個欄位
         'Memo by Amy 語法抓共用名稱查詢,若中/英/日寫成一句怕超過字串字數限制,排序為 國立臺灣大學/國立台灣大學
         strA = Replace(UCase(GetSearchNameSql(Me.Name, strFindTxt, "=", False, True)), " AS STATE", " AS STATE,'" & strTmp(0) & "' as InpField")
         'Modify by Amy 2024/06/13 避免抓不到對造資料,因對造暫存檔ID=strUserNum+小寫表單名
         strA = Replace(strA, UCase(Me.Name), Me.Name)
         strA = strA & " Order by FName,FNo"
         intA = 1
         Set rsA = ClsLawReadRstMsg(intA, strA)
         If intA = 1 Then
            rsA.MoveFirst
            Do While rsA.EOF = False
               strTmp(2) = "" & rsA.Fields("SFIELD") '找到的欄位
               Select Case strTmp(2)
                  Case "1"
                     strTmp(2) = "中"
                  Case "2"
                     strTmp(2) = "英"
                  Case "3"
                     strTmp(2) = "日"
               End Select
               If InStr(strTmp(4), Trim("<BR>名稱：" & rsA.Fields("FName"))) = 0 Then
                  strTmp(4) = Trim("<BR>名稱：" & rsA.Fields("FName"))
                  strTmp(1) = strTmp(1) & strTmp(4) & vbCrLf
               End If
               strTmp(3) = "已存在本所編號 [" & rsA.Fields("FNO") & "] " & strTmp(2) & "文欄位中" & vbCrLf
               Select Case "" & rsA.Fields("State")
                  Case "客戶"
                  '*** 本所客戶 X編號 ***
                     '都無身份證或統編 Or 風險檢查 無 但客戶檔 有 身份證或統編 Or 風險檢查 有 但客戶檔 無 身份證或統編 ,需由文雄確認
                     If (Trim(textRCL17) = MsgText(601) And Trim("" & rsA.Fields("mID")) = MsgText(601)) _
                       Or (Trim(textRCL17) = MsgText(601) And Trim("" & rsA.Fields("mID")) <> MsgText(601)) _
                       Or (Trim(textRCL17) <> MsgText(601) And Trim("" & rsA.Fields("mID")) = MsgText(601)) Then
                        strNowState = "1.1"
                     'Modify by Amy 2024/06/13 身份證或統編 相同 且客戶狀態[是]設為對造,屬於[對造],要可存檔-文雄
                     ElseIf Trim(textRCL17) = Trim(rsA.Fields("mID")) _
                       And Pub_GetField("Customer", "cu01||cu02='" & rsA.Fields("FNO") & "'", "cu80") = "設為對造" Then
                        'Memo 1.9-身份證或統編與風險檢查對象相[同]且客戶狀態[不是]設為對造 (無法存檔)
                        strTmp(3) = Replace(strTmp(3), "欄位中", "欄位中,狀態為[設為對造]")
                        strNowState = "9.2"
                     '身份證或統編 不同
                     Else
                        strNowState = "1.2"
                     End If
                  Case "國外潛客", "國內潛客"
                  '*** '本所客戶 R編號 ***
                     '個人,需由文雄確認-字數 同 客戶檔判斷且為中文欄位
                     If GetTextLength(rsA.Fields("FName")) <= 6 And strTmp(2) = "中" Then
                        strNowState = "2.1"
                     Else
                        strNowState = "2.2"
                     End If
                  Case "代理人"
                  '*** 代理人***
                     strNowState = "3.2"
                  Case "聯絡人"
                  '*** 聯絡人***
                     Select Case Left(rsA.Fields("FNO"), 1)
                        Case "X"
                           strNowState = "4.1"
                        Case "R"
                           strNowState = "4.2"
                        Case "Y"
                           strNowState = "4.3"
                     End Select
                  Case "不得代理"
                     strNowState = "9.1"
                  Case "對造"
                  '*** 對造***
                     strTmp(3) = "已存在[對造]案號 [" & rsA.Fields("FNO") & "] " & strTmp(2) & "文欄位中" & vbCrLf
                     strNowState = "9.2"
                  Case Else
               End Select
               'ex:中、英文欄輸相同會重覆顯示
               If InStr(strTmp(1), strTmp(3)) = 0 Then
                  strTmp(1) = strTmp(1) & strTmp(3)
               End If
               '*** 發信寫入暫存檔 ***
               strExc(3) = "" & rsA.Fields("SalesNo")
               '聯絡人
               If Left(strNowState, 1) = "4" Then
                  strExc(3) = GetSalesNo(rsA.Fields("FNO")) '客戶智權人員
               End If
               '                                                                           狀態                                  風險檢查名稱                                          輸入的欄位
               strTmp(5) = "'" & strUserNum & "','" & strNowState & "','" & ChgSQL(rsA.Fields("FName")) & "','" & rsA.Fields("InpField") & "'"
               '                                                        客戶編號 Or 案號                         查到的欄位                               身份證或統編                    智權人員
               strTmp(5) = strTmp(5) & ",'" & rsA.Fields("FNO") & "','" & rsA.Fields("SFIELD") & "','" & rsA.Fields("mID") & "','" & strExc(3) & "' "
               strCmd = "Insert Into R12040163 (ID,State,R001,R002,R003,R004,R005,R006) Values(" & strTmp(5) & ")"
               cnnConnection.Execute strCmd
               'X編號要再確認是否為[待活化客戶]
               If strNowState = "1.1" Then
                  '待活化客戶 需再新增一筆,發全所
                  If ChkOldCustomer(rsA.Fields("FNO")) = True Then
                     strNowState = "8.2"
                     cnnConnection.Execute Replace(strCmd, ",'1.1',", ",'" & strNowState & "',")
                  End If
               End If
            
               '*** End 發信寫入暫存檔 ***
               rsA.MoveNext
            Loop
            If strTmp(1) <> MsgText(601) Then
               stMsg = Replace(Replace(Mid(strTmp(1), 5), ",名稱：", vbCrLf & "名稱："), "<BR>", vbCrLf) & _
                              vbCrLf & "待確認或尚未進行利益衝突協商" & vbCrLf & _
                              "請按[否],回前畫面"
            End If
            ChkNameExist = True
         '未找到資料
         Else
            Select Case strTmp(0)
               Case "中"
                  strTmp(4) = textRCL02
               Case "英"
                  strTmp(4) = textRCL03
                  If textRCL04 <> MsgText(601) Then
                     strTmp(4) = strTmp(4) & " " & textRCL04
                     If textRCL05 <> MsgText(601) Then
                        strTmp(4) = strTmp(4) & " " & textRCL05
                        If textRCL06 <> MsgText(601) Then
                           strTmp(4) = strTmp(4) & " " & textRCL06
                        End If
                     End If
                  End If
               Case "日"
                  strTmp(4) = textRCL07
            End Select
            '建立者
            If m_EditMode = "1" Then
               strExc(3) = strUserNum
            Else
               strExc(3) = m_RCL27
            End If
             '                                                                           狀態                                  風險檢查名稱                     輸入的欄位
            strTmp(5) = "'" & strUserNum & "','" & strNowState & "','" & ChgSQL(strTmp(4)) & "','" & strTmp(0) & "'"
            '                                                        客戶編號 Or 案號  查到的欄位    身份證或統編                    智權人員
            strTmp(5) = strTmp(5) & "," & CNULL(textRCL01) & ",Null," & CNULL(textRCL17) & ",'" & strExc(3) & "' "
            strCmd = "Insert Into R12040163 (ID,State,R001,R002,R003,R004,R005,R006) Values(" & strTmp(5) & ")"
            cnnConnection.Execute strCmd
         End If
      Next ii
   End If
   
   stBackState = strNowState
   Set rsA = Nothing
End Function

'依風險檢查對象名稱對應之資料,判斷如何發mail
Private Sub SetSameNameMail()
   Dim rsA As New ADODB.Recordset, intA As Integer, strA As String, strCmd As String, strWhere As String, intCnt As Integer, intCusCnt As Integer
   Dim strDW(1 To 3) As String, strSubN As String, strOldName As String, strEndTxt(1 To 3) As String, strReason As String
   Dim strTo As String, strCC As String, strSubject As String, strContext As String, strTemplate(1 To 3) As String, strTemplateTB(1 To 2) As String
   Dim strSales As String, strCreateN As String, strOldMailState As String, strOldCuNo As String, strOldSales As String
   Dim strTp(4) As String, bolTest As Boolean, strToTXT As String
   
   bolTest = False '測式用設True (為制作給文雄文件用)
   strWhere = "And ID='" & strUserNum & "' "
   '*** 設定信件內容 ***
   '發[風險檢查對象可撤銷人員]確認信
   strTemplate(1) = "風險檢查對象「<名稱>」與本所編號 <客戶編號> [<欄位>]文名稱相同，<原因>"
   strEndTxt(1) = "確認後請通知 <業務> (<業務編號>)。 "
   strTemplateTB(1) = "<tr><td><業務>(<業務編號>)</td><td><客戶編號></td><td>[<欄位>]　<名稱></td></tr>"
   '發[全所]洽案人
   strTemplate(2) = " 「<名稱>」為風險檢查對象，非經協商，請勿接洽「<名稱>」。"
   '發[案件]洽案人
   strTemplate(3) = "風險檢查對象「<名稱>」與您客戶 <客戶編號> [<欄位>]文名稱相同，非經協商，請勿接洽「<名稱>」"
   strEndTxt(3) = "對上述設定若有疑慮，請洽<建立者> (<建立者員編>)。"
    
'*** 只發一次 ***
'發[風險檢查對象可撤銷人員]確認
   strContext = "": strTp(0) = "": intCnt = 0: strReason = ""
   'State='1.1' ->風險檢查及客戶檔只要任一者[無]身份證或統編 / SubStr(State,1,1)='4'->客戶X or Y or R 編號聯絡人 / R006->智權 or 客戶Y編號[無]開發者及建立者
   'Memo 目前只發現代理人檔[無]開發者及建立者 ex:A. Y51333000
   strDW(1) = "And (State='1.1' Or SubStr(State,1,1)='4' Or R006 is null) "
   strA = "Select Distinct '1' as MailState,R001 as RiskName,R003 as CuNo,Decode(R004,'1','中','2','英','3','日',R004) as FindField,R006 as SalesNo,Cnt,CusCnt " & _
               "From R12040163 ,(Select Count(*) as Cnt From R12040163 Where 1=1 " & strDW(1) & strWhere & " ) " & _
               ",(Select Sum(Count(Distinct R003)) as CusCnt From R12040163 Where 1=1 " & strDW(1) & strWhere & " Group by R003) " & _
               "Where 1=1 " & strDW(1) & strWhere & " Group by R001,R003,Decode(r004,'1','中','2','英','3','日',R004),R006,Cnt,CusCnt "
   'Modify by Amy 2024/03/27 先判斷需[風險檢查對象可撤銷人員]確認 者,其他的就不發
   '               ex:郭怡瑩 建中文「台光電子材料股份有限公司」->發文雄 / 英文「Elite Material Co., Ltd.」->發全所 (已由文雄確認不該再發)
   strSql = strA & "Order by SalesNo,CuNo,FindField "
   intA = 1
   Set rsA = ClsLawReadRstMsg(intA, strSql)
   If intA = 1 Then
      rsA.MoveFirst
      strCreateN = GetPrjSalesNM(m_RCL27) & GetStaffST20(m_RCL27)
      intCusCnt = rsA.Fields("CusCnt") '對應到的客戶數
      Do While rsA.EOF = False
         strTp(1) = "" & rsA.Fields("FindField")
         strTp(2) = Replace(Trim("" & rsA.Fields("RiskName")), vbCrLf, "")
         strTp(3) = "" & rsA.Fields("CuNo")
         strSales = "" & rsA.Fields("SalesNo")
         strTp(4) = strSales
         
         'X/Y/R編號聯絡人or X編號[無]身份證或統編資料 or Y編號[無]開發及建立人員 會發[風險檢查對象可撤銷人員]確認信
         If InStr("" & rsA.Fields("CuNo"), "-") Then
            'X/Y/R編號聯絡人
            If strReason <> MsgText(601) Then strReason = strReason & "或"
            strReason = strReason & "為客戶聯絡人"
         'X編號[無]身份證或統編資料
         ElseIf Left("" & rsA.Fields("CuNo"), 1) = "X" And InStr(strReason, "無身份證或統編資料") = 0 Then
            If strReason <> MsgText(601) Then strReason = strReason & "或"
            strReason = strReason & "無身份證或統編資料"
         'Y編號[無]開發及建立人員
         ElseIf Left("" & rsA.Fields("CuNo"), 1) = "Y" And strSales = MsgText(601) Then
            If strReason <> MsgText(601) Then strReason = strReason & "或"
            strReason = strReason & "無開發及建立人員"
         End If
         '先判斷多筆且非最後一筆,才會組多筆的tag ex:吳彩菱 建「保土谷化學工業股份有限公司」中/英/日 名稱
         If Val(rsA.Fields("CNT")) > 1 And Val(rsA.Fields("CNT")) > intCnt Then
            strTp(0) = strTemplateTB(1)
         ElseIf Val(rsA.Fields("CusCNT")) = 1 Then
            strTp(0) = strTemplate(1)
         End If
         If strOldCuNo = "" & rsA.Fields("CuNo") Then strTp(3) = ""
         If strOldSales = "" & rsA.Fields("SalesNo") Then strTp(4) = ""
      
         strTp(0) = Replace(strTp(0), "<名稱>", strTp(2))
         strTp(0) = Replace(strTp(0), "<欄位>", strTp(1))
         strTp(0) = Replace(strTp(0), "<客戶編號>", strTp(3))
         If strTp(4) = MsgText(601) Then
            strTp(0) = Replace(strTp(0), "(<業務編號>)", strTp(4))
         Else
            strTp(0) = Replace(strTp(0), "<業務編號>", strTp(4))
         End If
         If strTp(4) <> MsgText(601) Then
            strTp(4) = GetPrjSalesNM(strTp(4))
            '發[風險檢查對象可撤銷人員] 且只對應到1筆名稱
            If "" & rsA.Fields("MailState") = "1" And Val(rsA.Fields("CNT")) = 1 Then
               strTp(4) = strTp(4) & GetStaffST20(strTp(4))
            End If
         End If
         strTp(0) = Replace(strTp(0), "<業務>", strTp(4))
         
        strContext = strContext & strTp(0) & vbCrLf
        'ex:吳彩菱 建「保土谷化學工業股份有限公司」中/英/日 名稱與  X38411000 都同,<原因>沒被取代
        strContext = Replace(strContext, "<原因>", "因" & strReason)
         
         intCnt = intCnt + 1
         If strSubN = MsgText(601) Then
            '顯示於主旨的名稱 ex:保土谷化學工業股份有限公司/HODOGAYA CHEMICAL CO., LTD. 顯示中文
            strSubN = "" & rsA.Fields("RiskName")
         End If
         strOldName = "" & rsA.Fields("RiskName")
         strOldCuNo = "" & rsA.Fields("CuNo")
         strOldSales = "" & rsA.Fields("SalesNo")
         rsA.MoveNext
      Loop
      If InStr(strContext, "<業務>") > 0 Then strContext = strContext & Replace(strEndTxt(1), "<業務> ", GetPrjSalesNM(strOldSales))
      If InStr(strContext, "<業務編號>") > 0 Then strContext = Replace(strContext, "<業務編號>", strOldSales)
      If InStr(strContext, "<原因>") > 0 Then strContext = Replace(strContext, "<原因>", "因" & strReason)
      
      strTo = Pub_GetSpecMan("風險檢查對象可撤銷人員")
      '需確認者,副本通知 Pub_GetSpecMan("程式管理人員"),協助林特助是否已完成資料確認 (1140327與秀玲、林特助討論結果)
      strCC = Pub_GetSpecMan("程式管理人員")
      Call SetSubjectAndContext(1, strSubject, strContext, strTo, strSubN, m_RCL27, IIf(intCnt > 1, True, False))
      If bolTest = False Then
         'bolAbsenceSys設True-不發職代
         PUB_SendMail strUserNum, strTo, "", strSubject, strContext, , , , , , strCC, , , , True
      Else
         strToTXT = "發信給：" & strTo & " CC:" & strCC & vbCrLf '測式發信用
         strTo = "A2004": strCC = ""
         PUB_SendMail strUserNum, strTo, "", strSubject, strToTXT & strContext, , , , , , strCC
      End If
      Exit Sub
   End If
   
   'Modify by Amy 2025/03/28 [不需] 風險檢查對象可撤銷人員 確認,才往下做
   '[全所]洽案人 FindField抓風險檢查對象欄位
   strDW(2) = "And (SubStr(State,1,1)='8' Or State='1.2') "
   'strA = strA & " Union "
   strA = "Select Distinct '2' as MailState,R001 as RiskName,'' as CuNo,Decode(R002,'1','中','2','英','3','日',R002) as FindField,'' as SalesNo,Cnt,0 as CusCnt " & _
               "From R12040163,(Select Count(*) as Cnt From R12040163 Where 1=1 " & strDW(2) & strWhere & " ) " & _
               "Where 1=1 " & strDW(2) & strWhere & " Group by R001,Decode(R002,'1','中','2','英','3','日',R002),Cnt " & _
               "Order by MailState,SalesNo,CuNo,FindField "
   'end 2023/03/27
   intA = 1
   Set rsA = ClsLawReadRstMsg(intA, strA)
   If intA = 1 Then
      rsA.MoveFirst
      strCreateN = GetPrjSalesNM(m_RCL27) & GetStaffST20(m_RCL27)
      intCusCnt = rsA.Fields("CusCnt") '對應到的客戶數
      Do While rsA.EOF = False
      '*** 發信 ***
         If strOldMailState <> MsgText(601) And strOldMailState <> "" & rsA.Fields("MailState") Then
            strCC = Pub_GetSpecMan("風險檢查對象可撤銷人員")
            Call SetSubjectAndContext(Val(strOldMailState), strSubject, strContext, strTo, strSubN, m_RCL27, IIf(intCnt > 1, True, False))
            '[風險檢查對象可撤銷人員]確認信 結尾內容
            'Modify by Amy 2025/02/13 拿掉 intCusCnt = 1 ex:吳彩菱 建「保土谷化學工業股份有限公司」中/英/日 名稱與  X38411000 都同,<原因>沒被取代
            'Mark by Amy 2025/03/28 目前使用不到,因風險檢查對象可撤銷人員]確認 者,其他先不發
'            If strOldMailState = "1" Then
'               If InStr(strContext, "<業務>") > 0 Then strContext = strContext & Replace(strEndTxt(1), "<業務> ", GetPrjSalesNM(strOldSales))
'               If InStr(strContext, "<業務編號>") > 0 Then strContext = Replace(strContext, "<業務編號>", strOldSales)
'               If InStr(strContext, "<原因>") > 0 Then strContext = Replace(strContext, "<原因>", "因" & strReason)
'            End If
            'End 2025/03/28
            If strTo = strCC Then strCC = ""
            If bolTest = False Then
               'Modify by Amy 2024/07/19 bolAbsenceSys設True-不發職代
               PUB_SendMail strUserNum, strTo, "", strSubject, strContext, , , , , , strCC, , , , True
            Else
               strToTXT = "發信給：" & strTo & " CC:" & strCC & vbCrLf '測式發信用
               strTo = "A2004": strCC = ""
               PUB_SendMail strUserNum, strTo, "", strSubject, strToTXT & strContext, , , , , , strCC
            End If
            strReason = "": strOldName = "": strTo = "": strSubject = "": strContext = "": intCnt = 0
         End If
      '*** End 發信 ***
         strTp(1) = "" & rsA.Fields("FindField")
         strTp(2) = Replace(Trim("" & rsA.Fields("RiskName")), vbCrLf, "")
         strTp(3) = "" & rsA.Fields("CuNo")
         strSales = "" & rsA.Fields("SalesNo")
         strTp(4) = strSales
         
         '發[風險檢查對象可撤銷人員]確認信
         If "" & rsA.Fields("MailState") = "1" Then
            'Makr by Amy 2024/03/27 目前使用不到,因風險檢查對象可撤銷人員]確認 者,其他先不發
'            '目前只有X及X/Y/R編號聯絡人 會發[風險檢查對象可撤銷人員]確認信
'            If InStr("" & rsA.Fields("CuNo"), "-") Then
'               'Memo by Amy 2024/05/02 X/Y/R編號聯絡人原發[客戶洽案人]->改發[風險檢查對象可撤銷人員]-文雄
'               If strReason <> MsgText(601) Then strReason = strReason & "或"
'               strReason = strReason & "為客戶聯絡人"
'            ElseIf Left("" & rsA.Fields("CuNo"), 1) = "X" And InStr(strReason, "無身份證或統編資料") = 0 Then
'               If strReason <> MsgText(601) Then strReason = strReason & "或"
'               strReason = strReason & "無身份證或統編資料"
'            'Mark by Amy 2024/05/02 R編號為潛在客戶[個人]可直接發給客戶洽案人-文雄
''            ElseIf Left("" & rsA.Fields("CuNo"), 1) = "R" And InStr(strReason, "依字數判斷為個人") = 0 Then
''               If strReason <> MsgText(601) Then strReason = strReason & "或"
''               strReason = strReason & "依字數判斷為個人"
'            End If
'            'Modify by Amy 2025/02/13 改先判斷多筆且非最後一筆,才會組多筆的tag ex:吳彩菱 建「保土谷化學工業股份有限公司」中/英/日 名稱
'            '多筆且非最後一筆
'            If Val(rsA.Fields("CNT")) > 1 And Val(rsA.Fields("CNT")) > intCnt Then
'               strTp(0) = strTemplateTB(1)
'            ElseIf Val(rsA.Fields("CusCNT")) = 1 Then
'               strTp(0) = strTemplate(1)
'            End If
'            'end 2025/02/13
'            If strOldCuNo = "" & rsA.Fields("CuNo") Then strTp(3) = ""
'            If strOldSales = "" & rsA.Fields("SalesNo") Then strTp(4) = ""
            'end 2025/03/28
         '發[全所]洽案人+全所智權人員
         Else
            strTp(0) = strTemplate(2)
         End If
         strTp(0) = Replace(strTp(0), "<名稱>", strTp(2))
         strTp(0) = Replace(strTp(0), "<欄位>", strTp(1))
         strTp(0) = Replace(strTp(0), "<客戶編號>", strTp(3))
         If strTp(4) = MsgText(601) Then
            strTp(0) = Replace(strTp(0), "(<業務編號>)", strTp(4))
         Else
            strTp(0) = Replace(strTp(0), "<業務編號>", strTp(4))
         End If
         If strTp(4) <> MsgText(601) Then
            strTp(4) = GetPrjSalesNM(strTp(4))
            'Mark by Amy 2025/03/28 目前使用不到,因風險檢查對象可撤銷人員]確認 者,其他先不發
'            '發[風險檢查對象可撤銷人員] 且只對應到1筆名稱
'            If "" & rsA.Fields("MailState") = "1" And Val(rsA.Fields("CNT")) = 1 Then
'               strTp(4) = strTp(4) & GetStaffST20(strTp(4))
'            End If
         End If
         strTp(0) = Replace(strTp(0), "<業務>", strTp(4))
         
        strContext = strContext & strTp(0) & vbCrLf
        'Add by Amy 2025/02/13 ex:吳彩菱 建「保土谷化學工業股份有限公司」中/英/日 名稱與  X38411000 都同,<原因>沒被取代
        strContext = Replace(strContext, "<原因>", "因" & strReason)
         
         intCnt = intCnt + 1
         'Modify by Amy 2025/03/28 使用對應到的名稱 ex:郭怡瑩 建中文「台光電子材料股份有限公司」 / 英文「Elite Material Co., Ltd.」
         '              發信主旨:郭怡瑩 已設定「台光電子材料股份有限公司」為… /內文:郭怡瑩已設定 「Elite Material Co., Ltd.」為風險檢查對象，非經協商…->與主旨不合
         'If strSubN = MsgText(601) Then
            '顯示於主旨的名稱
            strSubN = "" & rsA.Fields("RiskName")
         'End If
         strOldMailState = "" & rsA.Fields("MailState")
         strOldName = "" & rsA.Fields("RiskName")
         strOldCuNo = "" & rsA.Fields("CuNo")
         strOldSales = "" & rsA.Fields("SalesNo")
         rsA.MoveNext
      Loop
   End If
   '最後一筆發信
   If strContext <> MsgText(601) Then
      strCC = Pub_GetSpecMan("風險檢查對象可撤銷人員")
      Call SetSubjectAndContext(Val(strOldMailState), strSubject, strContext, strTo, strSubN, m_RCL27, IIf(intCnt > 1, True, False))
      'Mark by Amy 2025/03/28 目前使用不到,因風險檢查對象可撤銷人員]確認 者,其他先不發
      'Modify by Amy 2025/02/13 ex:吳彩菱 建「保土谷化學工業股份有限公司」中/英/日 名稱與  X38411000 都同,<原因>沒被取代
'      If strOldMailState = "1" Then
'         If InStr(strContext, "<業務>") > 0 Then strContext = strContext & Replace(strEndTxt(1), "<業務> ", GetPrjSalesNM(strOldSales))
'         If InStr(strContext, "<業務編號>") > 0 Then strContext = Replace(strContext, "<業務編號>", strOldSales)
'         If InStr(strContext, "<原因>>") > 0 Then strContext = Replace(strContext, "<原因>", "因" & strReason)
'      End If
      'end 2025/03/28
      
      If strTo = strCC Then strCC = ""
      If bolTest = False Then
         'Add by Amy 2024/07/17 sales_all@taie.com.tw失效
         If InStr(strTo, "sales_all@taie.com.tw") > 0 Then
            strTo = Replace(strTo, "sales_all@taie.com.tw;", "")
         End If
         If strTo = MsgText(601) Then strTo = "A2004"
         'end 2024/07/17
         'Modify by Amy 2024/07/19 bolAbsenceSys設True-不發職代
         PUB_SendMail strUserNum, strTo, "", strSubject, strContext, , , , , , strCC, , , , True
      Else
         If InStr(strTo, "sales_all@taie.com.tw") > 0 Then
            strToTXT = "發信給：[全所]洽案人+全所智權人員" & " CC:" & strCC & vbCrLf
         Else
            strToTXT = "發信給：" & strTo & " CC:" & strCC & vbCrLf '測式發信用
         End If
         strTo = "A2004": strCC = ""
         PUB_SendMail strUserNum, strTo, "", strSubject, strToTXT & strContext, , , , , , strCC
      End If
   End If
'*** End 只發一次 ***
   
'*** 發[案件]洽案人 ***
'R編號/ Y編號
   strReason = "": strOldSales = "": strTo = "": strSubject = "": strContext = "": intCnt = 0
   strDW(3) = "And (SubStr(State,1,1)='2' Or State='3.2') "
   strA = "Select R001 as RiskName,R003 as CuNo,Decode(R004,'1','中','2','英','3','日',R004) as FindField,R006 as SalesNo " & _
               "From R12040163 ,(Select SubStr(R003,1,8) as CR003,Count(*) as Cnt From R12040163 " & _
                  "Where 1=1 " & strDW(3) & strWhere & " Group by SubStr(R003,1,8) ) " & _
               "Where 1=1 " & strDW(3) & strWhere & " And SubStr(R003,1,8)=CR003(+) " & strWhere & " Group by R006,R001,R003,Decode(R004,'1','中','2','英','3','日',R004),Cnt " & _
               "Order by SalesNo,CuNo,FindField "
   intA = 1
   Set rsA = ClsLawReadRstMsg(intA, strA)
   If intA = 1 Then
      rsA.MoveFirst
      strCreateN = GetPrjSalesNM(m_RCL27) & GetStaffST20(m_RCL27)
      Do While rsA.EOF = False
      '*** 發信 ***
         If strOldSales <> MsgText(601) And strOldSales <> "" & rsA.Fields("SalesNo") Then
            strCC = Pub_GetSpecMan("風險檢查對象可撤銷人員")
            strTo = strOldSales
            Call SetSubjectAndContext(3, strSubject, strContext, strTo, strSubN, m_RCL27, IIf(intCnt > 1, True, False))
            strContext = strContext & Replace(strEndTxt(3), "<建立者> ", GetPrjSalesNM(m_RCL27) & GetStaffST20(m_RCL27))
            strContext = Replace(strContext, "<建立者員編>", m_RCL27)
            'Modify by Amy 2025/03/28 原:strCC = "" ,發確認者信件,副本加發Pub_GetSpecMan("程式管理人員"),協助林特助是否已完成資料確認
            '              ex:Y51333000 [無]開發者及建立者
            If strTo = strCC Then strCC = Pub_GetSpecMan("程式管理人員")
            If bolTest = False Then
               'Modify by Amy 2024/07/19 bolAbsenceSys設True-不發職代
               PUB_SendMail strUserNum, strTo, "", strSubject, strContext, , , , , , strCC, , , , True
            Else
               strToTXT = "發信給：" & strTo & " CC:" & strCC & vbCrLf '測式發信用
               strTo = "A2004": strCC = ""
               PUB_SendMail strUserNum, strTo, "", strSubject, strToTXT & strContext, , , , , , strCC
            End If
            strReason = "": strOldName = "": strTo = "": strSubject = "": strContext = "": intCnt = 0
         End If
      '*** End 發信 ***
         strTp(0) = strTemplate(3)
         strTp(1) = "" & rsA.Fields("FindField")
         strTp(2) = Replace(Trim("" & rsA.Fields("RiskName")), vbCrLf, "")
         strTp(3) = "" & rsA.Fields("CuNo")
         If InStr("" & rsA.Fields("CuNo"), "-") > 0 Then
            strTp(3) = strTp(3) & " 之聯絡人"
         End If
         
         strTp(0) = Replace(strTp(0), "<名稱>", strTp(2))
         strTp(0) = Replace(strTp(0), "<欄位>", strTp(1))
         strTp(0) = Replace(strTp(0), "<客戶編號>", strTp(3))
         
         strContext = strContext & strTp(0) & vbCrLf
         
         intCnt = intCnt + 1
         'Modify by Amy 2025/03/28 使用對應到的名稱 (參閱同一天說明)
         'If strSubN = MsgText(601) Then
            '顯示於主旨的名稱
            strSubN = "" & rsA.Fields("RiskName")
         'End If
         strOldSales = "" & rsA.Fields("SalesNo")
         rsA.MoveNext
      Loop
   End If
   '最後一筆發信
   If strContext <> MsgText(601) Then
      strCC = Pub_GetSpecMan("風險檢查對象可撤銷人員")
      strTo = strOldSales
      Call SetSubjectAndContext(3, strSubject, strContext, strTo, strSubN, m_RCL27, IIf(intCnt > 1, True, False))
      strContext = strContext & Replace(strEndTxt(3), "<建立者> ", GetPrjSalesNM(m_RCL27) & GetStaffST20(m_RCL27))
      strContext = Replace(strContext, "<建立者員編>", m_RCL27)
      
      'Modify by Amy 2025/03/28 原:strCC = "" ,發確認者信件,副本加發Pub_GetSpecMan("程式管理人員"),協助林特助是否已完成資料確認
      '              ex:Y51333000 [無]開發者及建立者
      If strTo = strCC Then strCC = Pub_GetSpecMan("程式管理人員")
      If bolTest = False Then
         'Modify by Amy 2024/07/19 bolAbsenceSys設True-不發職代
         PUB_SendMail strUserNum, strTo, "", strSubject, strContext, , , , , , strCC, , , , True
      Else
         strToTXT = "發信給：" & strTo & " CC:" & strCC & vbCrLf '測式發信用
         strTo = "A2004": strCC = ""
         PUB_SendMail strUserNum, strTo, "", strSubject, strToTXT & strContext, , , , , , strCC
      End If
   End If
'*** End 發[案件]洽案人 ***

   Set rsA = Nothing
End Sub

'確認是否為待活化客戶
Private Function ChkOldCustomer(ByVal stNo As String) As Boolean
   Dim rsA As New ADODB.Recordset, intA As Integer, sta As String
   
   ChkOldCustomer = False
   If stNo = MsgText(601) Then Exit Function
   
   sta = "Select Ocu01 From OldCustomer Where Ocu01='" & Left(stNo, 8) & "' and Ocu03 is null "
   intA = 1
   Set rsA = ClsLawReadRstMsg(intA, sta)
   If intA = 1 Then
      ChkOldCustomer = True
   End If
   Set rsA = Nothing
End Function

'全所洽案人(參考frm090635.cmdOK_Click-有權限查詢價目表之人員) +全所業務
Private Function GetAllCaseContactP() As String
   Dim rsA As New ADODB.Recordset, intA As Integer, sta As String, stField As String
   
   stField = ",Decode(SubStr(st15,1,1),'S','1'||st15,'2'||st15) as Sort"
   'Modify by Amy 2024/07/17 +全所業務(在職)-plq02:P-個人)/D-部門
   sta = "Select st01" & stField & " From Staff Where SubStr(st15,1,1)='S' And st04='1' And st01>'6' And st01<'F' And SubStr(st01,4,1)<>'9' " & _
"Union Select  st01" & stField & " From PriceListquery, Staff Where plq02='P' And st04='1' And Instr(','||plq03||',',','||st01||',')>0 " & _
"Union Select  st01" & stField & " From PriceListquery, Staff Where plq02='D' and st04='1' And Instr(','||plq03||',',','||st03||',')>0 " & _
               "And st01>'6' And st01<'F' And SubStr(st01,4,1)<>'9' "
   'end 2024/07/17
   'Add by Amy 2024/10/18 +國外部會收文人員 (ST93為F21(外專承辦)/J21(日專承辦)/T21(商二組MC-FC)/T22(商二組FC英文)/T23(商二組FC日文))
   sta = sta & "Union Select st01,'2'||st93 as Sort From Staff Where st93 in ('F21','J21','T21','T22','T23') And st04='1' "
   sta = "Select Distinct st01,Sort From(" & sta & ") Order by sort,st01"
   
   intA = 1
   Set rsA = ClsLawReadRstMsg(intA, sta)
   If intA = 1 Then
      rsA.MoveFirst
      Do While rsA.EOF = False
         GetAllCaseContactP = GetAllCaseContactP & ";" & rsA.Fields("st01")
         rsA.MoveNext
      Loop
   End If
      
   '不拿掉 sales_all@taie.com.tw 測式發信用
   If GetAllCaseContactP <> MsgText(601) Then
      GetAllCaseContactP = "sales_all@taie.com.tw" & GetAllCaseContactP
   Else
      GetAllCaseContactP = "sales_all@taie.com.tw"
   End If
   
   Set rsA = Nothing
End Function

'傳入編號取得目前業務
Private Function GetSalesNo(ByVal stNo As String) As String
   Dim rsA As New ADODB.Recordset, intA As Integer, sta As String
   
   sta = "Select cu13 as SalesNo From Customer Where cu01='" & Left(stNo, 8) & "' and cu02='0' " & _
   "Union Select fa94 as SalesNo From Fagent Where fa01='" & Left(stNo, 8) & "' and fa02='0' " & _
   "Union Select pcu38 as SalesNo From PotCustomer Where pcu01='" & Left(stNo, 8) & "' and pcu02='0' " & _
   "Union Select poc13 as SalesNo From PotCustomer1 Where poc01='" & Left(stNo, 8) & "' and poc02='0' "
   intA = 1
   Set rsA = ClsLawReadRstMsg(intA, sta)
   If intA = 1 Then
      rsA.MoveFirst
      GetSalesNo = "" & rsA.Fields("SalesNo")
   End If
   
   Set rsA = Nothing
End Function

'刪除發信暫存檔
Private Sub DelR12040163()
   strExc(3) = "Delete From R12040163 Where ID='" & strUserNum & "' "
   cnnConnection.Execute strExc(3)
End Sub

Private Sub SetSubjectAndContext(intState As Integer, ByRef stSubject As String, ByRef stContext As String, ByRef stTO As String, stRiskName As String, stCreateP As String, IsMany As Boolean)
   Dim stCreateN As String, stTmp As String, ii As Integer, arrTmp
   
   stCreateN = GetPrjSalesNM(stCreateP) & GetStaffST20(stCreateP)
   Select Case intState
      Case 1 '發[風險檢查對象可撤銷人員]確認
         stTO = Pub_GetSpecMan("風險檢查對象可撤銷人員")
         If IsMany = True Then
            stContext = "<BODY><P>" & stCreateN & " 設定之風險檢查對象，同名同姓資料如下：</P>" & _
                                  "<TABLE　BORDER=1><tr><th>業務(編號)</th><th>客戶編號</th><th>　名　　稱　</th></tr>" & _
                                    stContext & "</TABLE><P>確認後請通知上述業務人員</P></BODY>" & vbCrLf & vbCrLf
         Else
            stContext = stCreateN & " 設定之" & stContext
         End If
         'Modify by Amy 2025/03/28 +if Y編號[無]開發及建立人員,主旨顯示不同
         stSubject = stCreateN & " 已設定「" & stRiskName & "」為風險檢查對象"
         If InStr(stContext, "同名同姓") > 0 Then
            stSubject = stSubject & "，屬同名同姓資料"
         ElseIf InStr(stContext, "無開發及建立人員") > 0 Then
            stSubject = stSubject & "，無開發及建立人員"
         End If
         stSubject = stSubject & "，需確認"
      Case 2 '發[全所]洽案人+全所智權人員
         stTO = GetAllCaseContactP
         stContext = stCreateN & "已設定" & IIf(IsMany = True, vbCrLf, "") & stContext
         stSubject = stCreateN & " 已設定「" & stRiskName & "」為風險檢查對象，請知悉"
      Case 3 '發[案件]洽案人
         stTO = Replace(stTO, ",", ";")
         stSubject = stCreateN & " 已設定「" & stRiskName & "」為風險檢查對象，請知悉"
   End Select
   'Memo X/R編號 業務帶智權人/開發者；Y編號 業務開發者(fa94)->建立者(fa46)=>離職人員抓其主管
   If intState = 3 Then
      If InStr(stTO, ";") = 0 Then
         If ChkStaffST04(stTO, False) = True Then
            stTmp = GetDeptMan(GetST15(stTO))
         End If
      Else
         arrTmp = Split(stTO, ";")
         stTO = ""
         For ii = LBound(arrTmp) To UBound(arrTmp)
            '已離職
            If ChkStaffST04("" & arrTmp(ii), False) = True Then
               stTmp = stTmp & ";" & GetDeptMan(GetST15(arrTmp(ii)))
            Else
               stTmp = stTmp & ";" & arrTmp(ii)
            End If
         Next ii
         stTmp = Mid(stTmp, 2)
      End If
      If stTmp <> MsgText(601) Then stTO = stTmp
   End If
   
   If stTO = MsgText(601) Then stTO = Pub_GetSpecMan("風險檢查對象可撤銷人員") 'ex:Y51333000 (無 開發者及建立者)
End Sub

'Add by Amy 2024/07/19 刪除後重新抓資料
Public Sub AfterDelShowData(ByVal strKey As String)
   If (strKey = m_FirstKEY(0)) Or strKey = m_LastKEY(0) Then
      RefreshRange '重新取得第一筆及最後一筆編號
      If strKey = m_FirstKEY(0) Then
         strKey = m_FirstKEY(0)
      Else
         strKey = m_LastKEY(0)
      End If
   End If
   ShowCurrRecord strKey
End Sub

