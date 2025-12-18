VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm071002 
   BorderStyle     =   1  '單線固定
   Caption         =   "法務－分案"
   ClientHeight    =   6420
   ClientLeft      =   336
   ClientTop       =   576
   ClientWidth     =   9240
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6420
   ScaleWidth      =   9240
   Begin TabDlg.SSTab SSTab1 
      Height          =   5355
      Left            =   30
      TabIndex        =   57
      Top             =   1050
      Width           =   9195
      _ExtentX        =   16214
      _ExtentY        =   9440
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabMaxWidth     =   3528
      TabCaption(0)   =   "第一頁"
      TabPicture(0)   =   "frm071002.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Text(7)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Text(9)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Text(8)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Text(12)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Text(13)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Text(10)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label10"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label3(0)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "lbe(7)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label14"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "lbe(8)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label20"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label11"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label24"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "lbe(10)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "lbe(9)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Text(18)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Label31"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Text(19)"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Label13"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Frame1"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "Command4"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "Command2"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).ControlCount=   23
      TabCaption(1)   =   "第二頁"
      TabPicture(1)   =   "frm071002.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Text(14)"
      Tab(1).Control(1)=   "Text(16)"
      Tab(1).Control(2)=   "Text(15)"
      Tab(1).Control(3)=   "Label30"
      Tab(1).Control(4)=   "Label26"
      Tab(1).Control(5)=   "Label23"
      Tab(1).Control(6)=   "Label15"
      Tab(1).Control(7)=   "Label5(0)"
      Tab(1).Control(8)=   "Label21(1)"
      Tab(1).Control(9)=   "Label27"
      Tab(1).Control(10)=   "Label29"
      Tab(1).Control(11)=   "Label6"
      Tab(1).Control(12)=   "Text(17)"
      Tab(1).Control(13)=   "Label22"
      Tab(1).Control(14)=   "Label21(0)"
      Tab(1).Control(15)=   "Label19"
      Tab(1).Control(16)=   "Text(11)"
      Tab(1).Control(17)=   "Label8"
      Tab(1).Control(18)=   "lbeCost"
      Tab(1).Control(19)=   "Label4(1)"
      Tab(1).Control(20)=   "Label9"
      Tab(1).Control(21)=   "lbePointNum"
      Tab(1).Control(22)=   "Label16"
      Tab(1).Control(23)=   "lbeMoney"
      Tab(1).Control(24)=   "MSHFlexGrid1"
      Tab(1).Control(25)=   "txtcp01"
      Tab(1).Control(26)=   "txtcp02"
      Tab(1).Control(27)=   "txtcp03"
      Tab(1).Control(28)=   "txtcp04"
      Tab(1).ControlCount=   29
      Begin VB.TextBox txtcp04 
         Height          =   285
         Left            =   -71895
         MaxLength       =   2
         TabIndex        =   47
         Top             =   1320
         Width           =   495
      End
      Begin VB.TextBox txtcp03 
         Height          =   285
         Left            =   -72255
         MaxLength       =   1
         TabIndex        =   46
         Top             =   1320
         Width           =   255
      End
      Begin VB.TextBox txtcp02 
         Height          =   285
         Left            =   -73350
         MaxLength       =   6
         TabIndex        =   45
         Top             =   1320
         Width           =   975
      End
      Begin VB.TextBox txtcp01 
         Height          =   285
         Left            =   -73830
         MaxLength       =   3
         TabIndex        =   44
         Top             =   1320
         Width           =   375
      End
      Begin VB.CommandButton Command2 
         Caption         =   "相對人資料(&B)"
         Height          =   270
         Left            =   6045
         TabIndex        =   30
         Top             =   3900
         Width           =   1335
      End
      Begin VB.CommandButton Command4 
         Caption         =   "出庭律師(&L)"
         Height          =   270
         Left            =   7410
         TabIndex        =   31
         Top             =   3900
         Width           =   1485
      End
      Begin VB.Frame Frame1 
         Height          =   2860
         Left            =   120
         TabIndex        =   58
         Top             =   330
         Width           =   8895
         Begin VB.CheckBox Check1 
            BackColor       =   &H00C0FFFF&
            Caption         =   "一般"
            Height          =   215
            Index           =   4
            Left            =   4560
            TabIndex        =   15
            Top             =   2190
            Width           =   735
         End
         Begin VB.Frame Frame2 
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   0  '沒有框線
            Height          =   700
            Left            =   4560
            TabIndex        =   95
            Top             =   2130
            Width           =   4200
            Begin VB.CheckBox Check2 
               BackColor       =   &H00C0FFFF&
               Caption         =   "勞資糾紛"
               Height          =   180
               Index           =   7
               Left            =   2760
               TabIndex        =   23
               Top             =   480
               Width           =   1425
            End
            Begin VB.CheckBox Check2 
               BackColor       =   &H00C0FFFF&
               Caption         =   "買賣糾紛"
               Height          =   180
               Index           =   6
               Left            =   1410
               TabIndex        =   22
               Top             =   480
               Width           =   1245
            End
            Begin VB.CheckBox Check2 
               BackColor       =   &H00C0FFFF&
               Caption         =   "醫療糾紛"
               Height          =   180
               Index           =   5
               Left            =   0
               TabIndex        =   21
               Top             =   480
               Width           =   1245
            End
            Begin VB.CheckBox Check2 
               BackColor       =   &H00C0FFFF&
               Caption         =   "消費糾紛"
               Height          =   180
               Index           =   4
               Left            =   2760
               TabIndex        =   20
               Top             =   270
               Width           =   1425
            End
            Begin VB.CheckBox Check2 
               BackColor       =   &H00C0FFFF&
               Caption         =   "車禍糾紛"
               Height          =   180
               Index           =   3
               Left            =   1410
               TabIndex        =   19
               Top             =   270
               Width           =   1245
            End
            Begin VB.CheckBox Check2 
               BackColor       =   &H00C0FFFF&
               Caption         =   "土地爭議"
               Height          =   180
               Index           =   2
               Left            =   0
               TabIndex        =   18
               Top             =   270
               Width           =   1245
            End
            Begin VB.CheckBox Check2 
               BackColor       =   &H00C0FFFF&
               Caption         =   "公寓大廈糾紛"
               Height          =   180
               Index           =   1
               Left            =   2760
               TabIndex        =   17
               Top             =   60
               Width           =   1425
            End
            Begin VB.CheckBox Check2 
               BackColor       =   &H00C0FFFF&
               Caption         =   "承攬糾紛"
               Height          =   180
               Index           =   0
               Left            =   1410
               TabIndex        =   16
               Top             =   60
               Width           =   1245
            End
         End
         Begin VB.CheckBox Check1 
            Caption         =   "其他智財權"
            Height          =   215
            Index           =   3
            Left            =   6810
            TabIndex        =   12
            Top             =   1695
            Width           =   1215
         End
         Begin VB.CheckBox Check1 
            Caption         =   "著作權"
            Height          =   215
            Index           =   2
            Left            =   5970
            TabIndex        =   11
            Top             =   1695
            Width           =   855
         End
         Begin VB.CheckBox Check1 
            Caption         =   "商標"
            Height          =   215
            Index           =   1
            Left            =   5250
            TabIndex        =   10
            Top             =   1695
            Width           =   735
         End
         Begin VB.CheckBox Check1 
            Caption         =   "專利"
            Height          =   215
            Index           =   0
            Left            =   4530
            TabIndex        =   9
            Top             =   1695
            Width           =   735
         End
         Begin VB.CheckBox Check1 
            Caption         =   "營業秘密法"
            Height          =   215
            Index           =   5
            Left            =   4530
            TabIndex        =   13
            Top             =   1935
            Width           =   1365
         End
         Begin VB.CheckBox Check1 
            Caption         =   "公平交易法"
            Height          =   215
            Index           =   6
            Left            =   5970
            TabIndex        =   14
            Top             =   1935
            Width           =   1365
         End
         Begin VB.Label lblMemo 
            Caption         =   "可以直接輸入，屬性之間用逗號,區隔。"
            Height          =   765
            Left            =   90
            TabIndex        =   105
            Top             =   1980
            Width           =   945
         End
         Begin MSForms.TextBox Text 
            Height          =   285
            Index           =   52
            Left            =   3555
            TabIndex        =   6
            Top             =   1395
            Width           =   375
            VariousPropertyBits=   671105051
            MaxLength       =   1
            Size            =   "661;503"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin VB.Label lblName 
            Caption         =   "案件名稱(日)："
            Height          =   255
            Index           =   1
            Left            =   90
            TabIndex        =   68
            Top             =   1110
            Width           =   1275
         End
         Begin VB.Label lblName 
            Caption         =   "案件名稱(英)："
            Height          =   255
            Index           =   0
            Left            =   90
            TabIndex        =   67
            Top             =   825
            Width           =   1275
         End
         Begin VB.Label lblName 
            Caption         =   "案件名稱(中)："
            Height          =   255
            Index           =   3
            Left            =   90
            TabIndex        =   66
            Top             =   525
            Width           =   1275
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "(Y:是)"
            Height          =   180
            Index           =   1
            Left            =   1650
            TabIndex        =   65
            Top             =   1440
            Width           =   465
         End
         Begin VB.Label Label25 
            Caption         =   "智財權案："
            Height          =   255
            Left            =   90
            TabIndex        =   64
            Top             =   1410
            Width           =   1005
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "分所案號："
            Height          =   180
            Index           =   0
            Left            =   4800
            TabIndex        =   63
            Top             =   1440
            Width           =   900
         End
         Begin MSForms.Label lbe 
            Height          =   285
            Index           =   1
            Left            =   2190
            TabIndex        =   62
            Top             =   210
            Width           =   6255
            VariousPropertyBits=   27
            Caption         =   "lbe(1)"
            Size            =   "11033;494"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin VB.Label Label17 
            Caption         =   "當  事  人："
            Height          =   285
            Left            =   90
            TabIndex        =   61
            Top             =   210
            Width           =   975
         End
         Begin VB.Label Label12 
            BackColor       =   &H00C0FFC0&
            Caption         =   "案件屬性："
            Height          =   210
            Left            =   90
            TabIndex        =   60
            Top             =   1740
            Width           =   945
         End
         Begin VB.Label Label52 
            Caption         =   "專案服務案：         (Y:是)"
            Height          =   180
            Left            =   2490
            TabIndex        =   59
            Top             =   1455
            Width           =   1995
         End
         Begin MSForms.TextBox Text 
            Height          =   285
            Index           =   4
            Left            =   1410
            TabIndex        =   4
            Top             =   1095
            Width           =   7410
            VariousPropertyBits=   671105051
            BackColor       =   16777215
            MaxLength       =   160
            Size            =   "13070;503"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox Text 
            Height          =   285
            Index           =   3
            Left            =   1410
            TabIndex        =   3
            Top             =   810
            Width           =   7410
            VariousPropertyBits=   671105051
            BackColor       =   16777215
            MaxLength       =   160
            Size            =   "13070;503"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox Text 
            Height          =   285
            Index           =   2
            Left            =   1410
            TabIndex        =   2
            Top             =   510
            Width           =   7410
            VariousPropertyBits=   671105051
            BackColor       =   16777215
            MaxLength       =   160
            Size            =   "13070;503"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox Text 
            Height          =   285
            Index           =   6
            Left            =   5760
            TabIndex        =   7
            Top             =   1410
            Width           =   3000
            VariousPropertyBits=   671105051
            MaxLength       =   50
            Size            =   "5292;503"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox Text 
            Height          =   285
            Index           =   5
            Left            =   1110
            TabIndex        =   5
            Top             =   1380
            Width           =   375
            VariousPropertyBits=   671105051
            MaxLength       =   1
            Size            =   "661;503"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox Text 
            Height          =   300
            Index           =   1
            Left            =   1050
            TabIndex        =   1
            Top             =   180
            Width           =   1095
            VariousPropertyBits=   671105051
            MaxLength       =   9
            Size            =   "1931;529"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox Text 
            Height          =   1065
            Index           =   28
            Left            =   1050
            TabIndex        =   8
            Top             =   1680
            Width           =   3465
            VariousPropertyBits=   -1467989989
            MaxLength       =   200
            Size            =   "6112;1879"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   1110
         Left            =   -74400
         TabIndex        =   50
         Top             =   2070
         Width           =   8355
         _ExtentX        =   14732
         _ExtentY        =   1969
         _Version        =   393216
         Cols            =   12
         FixedCols       =   0
         BackColorBkg    =   16772048
         ScrollTrack     =   -1  'True
         AllowUserResizing=   1
         RowSizingMode   =   1
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
         _Band(0).Cols   =   12
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin VB.Label lbeMoney 
         AutoSize        =   -1  'True
         Caption         =   "lbeMoney"
         Height          =   180
         Left            =   -70500
         TabIndex        =   102
         Top             =   450
         Width           =   705
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "後金："
         Height          =   180
         Left            =   -71085
         TabIndex        =   101
         Top             =   450
         Width           =   540
      End
      Begin VB.Label lbePointNum 
         AutoSize        =   -1  'True
         Caption         =   "lbePointNum"
         Height          =   180
         Left            =   -72390
         TabIndex        =   100
         Top             =   450
         Width           =   915
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "點數："
         Height          =   180
         Left            =   -72945
         TabIndex        =   99
         Top             =   450
         Width           =   540
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "費用："
         Height          =   180
         Index           =   1
         Left            =   -74910
         TabIndex        =   98
         Top             =   450
         Width           =   540
      End
      Begin VB.Label lbeCost 
         AutoSize        =   -1  'True
         Caption         =   "lbeCost"
         Height          =   180
         Left            =   -74280
         TabIndex        =   97
         Top             =   450
         Width           =   525
      End
      Begin VB.Label Label8 
         Caption         =   "當事人稱謂："
         Height          =   255
         Left            =   -74910
         TabIndex        =   96
         Top             =   720
         Width           =   1095
      End
      Begin MSForms.TextBox Text 
         Height          =   285
         Index           =   11
         Left            =   -73560
         TabIndex        =   41
         Top             =   690
         Width           =   1695
         VariousPropertyBits=   671105051
         BackColor       =   16777215
         MaxLength       =   10
         Size            =   "2990;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label19 
         Caption         =   "本案期限"
         Height          =   405
         Left            =   -74880
         TabIndex        =   94
         Top             =   2100
         Width           =   405
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "是否取締案："
         Height          =   180
         Index           =   0
         Left            =   -74910
         TabIndex        =   93
         Top             =   1702
         Width           =   1065
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "(Y:取締)"
         Height          =   180
         Left            =   -73380
         TabIndex        =   92
         Top             =   1702
         Width           =   645
      End
      Begin MSForms.TextBox Text 
         Height          =   285
         Index           =   17
         Left            =   -73830
         TabIndex        =   49
         Top             =   1650
         Width           =   375
         VariousPropertyBits=   671105051
         MaxLength       =   1
         Size            =   "661;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label6 
         Caption         =   "轉本所案號："
         Height          =   255
         Left            =   -74910
         TabIndex        =   91
         Top             =   1335
         Width           =   1095
      End
      Begin VB.Label Label29 
         Caption         =   "(Y:取消閉卷)"
         Height          =   255
         Left            =   -69150
         TabIndex        =   90
         Top             =   1335
         Width           =   1305
      End
      Begin VB.Label Label27 
         Caption         =   "相關總收文號："
         Height          =   255
         Left            =   -70905
         TabIndex        =   89
         Top             =   1005
         Width           =   1335
      End
      Begin VB.Label Label21 
         Caption         =   "是否取消閉卷："
         Height          =   255
         Index           =   1
         Left            =   -70905
         TabIndex        =   88
         Top             =   1335
         Width           =   1335
      End
      Begin VB.Label Label5 
         Caption         =   "(N：不算)"
         Height          =   180
         Index           =   0
         Left            =   -72810
         TabIndex        =   87
         Top             =   1042
         Width           =   855
      End
      Begin VB.Label Label15 
         Caption         =   "是否算案件數："
         Height          =   180
         Left            =   -74910
         TabIndex        =   86
         Top             =   1042
         Width           =   1335
      End
      Begin VB.Label Label23 
         Caption         =   "-"
         Height          =   255
         Left            =   -73425
         TabIndex        =   85
         Top             =   1335
         Width           =   45
      End
      Begin VB.Label Label26 
         Caption         =   "-"
         Height          =   255
         Left            =   -72345
         TabIndex        =   84
         Top             =   1335
         Width           =   45
      End
      Begin VB.Label Label30 
         Caption         =   "-"
         Height          =   255
         Left            =   -71970
         TabIndex        =   83
         Top             =   1335
         Width           =   45
      End
      Begin MSForms.TextBox Text 
         Height          =   285
         Index           =   15
         Left            =   -69570
         TabIndex        =   43
         Top             =   990
         Width           =   2175
         VariousPropertyBits=   671105051
         MaxLength       =   9
         Size            =   "3836;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text 
         Height          =   285
         Index           =   16
         Left            =   -69570
         TabIndex        =   48
         Top             =   1320
         Width           =   375
         VariousPropertyBits=   671105051
         MaxLength       =   1
         Size            =   "661;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text 
         Height          =   285
         Index           =   14
         Left            =   -73560
         TabIndex        =   42
         Top             =   990
         Width           =   375
         VariousPropertyBits=   671105051
         MaxLength       =   1
         Size            =   "661;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label13 
         Caption         =   "案件備註："
         Height          =   255
         Left            =   120
         TabIndex        =   82
         Top             =   4800
         Width           =   1065
      End
      Begin MSForms.TextBox Text 
         Height          =   525
         Index           =   19
         Left            =   1230
         TabIndex        =   33
         Top             =   4770
         Width           =   7845
         VariousPropertyBits=   -1466941413
         MaxLength       =   2000
         ScrollBars      =   2
         Size            =   "13838;926"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         Caption         =   "進度備註："
         Height          =   180
         Left            =   120
         TabIndex        =   81
         Top             =   4245
         Width           =   900
      End
      Begin MSForms.TextBox Text 
         Height          =   525
         Index           =   18
         Left            =   1230
         TabIndex        =   32
         Top             =   4230
         Width           =   7845
         VariousPropertyBits=   -1466941413
         MaxLength       =   2000
         ScrollBars      =   2
         Size            =   "13838;926"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbe 
         Height          =   285
         Index           =   9
         Left            =   1890
         TabIndex        =   80
         Top             =   3555
         Width           =   2415
         VariousPropertyBits=   27
         Caption         =   "lbe(9)"
         Size            =   "4260;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbe 
         Height          =   285
         Index           =   10
         Left            =   6450
         TabIndex        =   79
         Top             =   3555
         Width           =   1935
         VariousPropertyBits=   27
         Caption         =   "lbe(10)"
         Size            =   "3413;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label24 
         Caption         =   "智權人員："
         Height          =   255
         Left            =   4470
         TabIndex        =   78
         Top             =   3570
         Width           =   900
      End
      Begin VB.Label Label11 
         Caption         =   "法定期限："
         Height          =   255
         Left            =   2910
         TabIndex        =   77
         Top             =   3915
         Width           =   900
      End
      Begin VB.Label Label20 
         Caption         =   "本所期限："
         Height          =   255
         Left            =   120
         TabIndex        =   76
         Top             =   3915
         Width           =   975
      End
      Begin MSForms.Label lbe 
         Height          =   285
         Index           =   8
         Left            =   6450
         TabIndex        =   75
         Top             =   3270
         Width           =   1935
         VariousPropertyBits=   27
         Caption         =   "lbe(8)"
         Size            =   "3413;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label14 
         Caption         =   "協辦人員："
         Height          =   255
         Left            =   4470
         TabIndex        =   74
         Top             =   3285
         Width           =   900
      End
      Begin MSForms.Label lbe 
         Height          =   285
         Index           =   7
         Left            =   2400
         TabIndex        =   73
         Top             =   3270
         Width           =   1965
         VariousPropertyBits=   27
         Caption         =   "lbe(7)"
         Size            =   "3466;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label3 
         Caption         =   "案件性質："
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   72
         Top             =   3570
         Width           =   975
      End
      Begin VB.Label Label10 
         Caption         =   "承辦人："
         Height          =   255
         Left            =   120
         TabIndex        =   71
         Top             =   3285
         Width           =   975
      End
      Begin MSForms.TextBox Text 
         Height          =   285
         Index           =   10
         Left            =   5385
         TabIndex        =   27
         Top             =   3555
         Width           =   975
         VariousPropertyBits=   671105051
         MaxLength       =   6
         Size            =   "1720;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text 
         Height          =   285
         Index           =   13
         Left            =   3825
         TabIndex        =   29
         Top             =   3900
         Width           =   1095
         VariousPropertyBits=   671105051
         MaxLength       =   7
         Size            =   "1931;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text 
         Height          =   285
         Index           =   12
         Left            =   1230
         TabIndex        =   28
         Top             =   3900
         Width           =   1095
         VariousPropertyBits=   671105051
         MaxLength       =   7
         Size            =   "1931;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text 
         Height          =   285
         Index           =   8
         Left            =   5385
         TabIndex        =   25
         Top             =   3270
         Width           =   975
         VariousPropertyBits=   671105051
         MaxLength       =   6
         Size            =   "1720;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text 
         Height          =   285
         Index           =   9
         Left            =   1230
         TabIndex        =   26
         Top             =   3555
         Width           =   615
         VariousPropertyBits=   671105051
         MaxLength       =   4
         Size            =   "1085;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text 
         Height          =   285
         Index           =   7
         Left            =   1230
         TabIndex        =   24
         Top             =   3270
         Width           =   1095
         VariousPropertyBits=   671105051
         MaxLength       =   6
         Size            =   "1931;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
   End
   Begin VB.CommandButton CmdOK1 
      Caption         =   "案源卷宗區(&C)"
      Height          =   400
      Left            =   1830
      TabIndex        =   37
      Top             =   0
      Width           =   1350
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "下一筆(&N)"
      Height          =   400
      Left            =   3240
      TabIndex        =   38
      Top             =   0
      Width           =   900
   End
   Begin VB.CommandButton ComBack 
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   8235
      TabIndex        =   36
      Top             =   0
      Width           =   756
   End
   Begin VB.CommandButton cmdPrePic 
      Caption         =   "回前畫面(&U)"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   7124
      TabIndex        =   35
      Top             =   0
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Left            =   6361
      TabIndex        =   34
      Top             =   0
      Width           =   756
   End
   Begin VB.CommandButton Command3 
      Caption         =   "案件進度(&C)"
      Height          =   400
      Left            =   5254
      TabIndex        =   40
      Top             =   0
      Width           =   1100
   End
   Begin VB.CommandButton Command1 
      Caption         =   "相關卷號(&F)"
      Height          =   400
      Left            =   4147
      TabIndex        =   39
      Top             =   0
      Width           =   1100
   End
   Begin VB.Label lbeCloseDate 
      Caption         =   "lbeCloseDate"
      Height          =   285
      Left            =   5700
      TabIndex        =   104
      Top             =   750
      Width           =   1575
   End
   Begin VB.Label Label28 
      Caption         =   "取消收文日期："
      Height          =   255
      Left            =   4380
      TabIndex        =   103
      Top             =   750
      Width           =   1260
   End
   Begin VB.Label lbeNumber 
      Height          =   280
      Left            =   1185
      TabIndex        =   70
      Top             =   780
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "本所案號："
      Height          =   255
      Index           =   0
      Left            =   180
      TabIndex        =   69
      Top             =   795
      Width           =   975
   End
   Begin MSForms.TextBox Text 
      Height          =   285
      Index           =   0
      Left            =   5655
      TabIndex        =   0
      Top             =   450
      Width           =   1095
      VariousPropertyBits=   671105051
      MaxLength       =   7
      Size            =   "1931;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label lblLOS01 
      Caption         =   "lblLOS01"
      Height          =   285
      Left            =   3540
      TabIndex        =   56
      Top             =   465
      Width           =   915
   End
   Begin VB.Label LBL01 
      Caption         =   "案源總收文號： "
      Height          =   255
      Left            =   2250
      TabIndex        =   55
      Top             =   465
      Width           =   1275
   End
   Begin VB.Label lblClose 
      Caption         =   "lblClose"
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
      Height          =   285
      Left            =   6840
      TabIndex        =   54
      Top             =   450
      Width           =   2055
   End
   Begin VB.Label lbePaperNum 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """#-##-######"""
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1028
         SubFormatType   =   0
      EndProperty
      Height          =   280
      Left            =   1185
      TabIndex        =   53
      Top             =   450
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "收  文  號： "
      Height          =   255
      Index           =   1
      Left            =   180
      TabIndex        =   52
      Top             =   465
      Width           =   975
   End
   Begin VB.Label Label7 
      Caption         =   "收  文  日："
      Height          =   255
      Index           =   0
      Left            =   4740
      TabIndex        =   51
      Top             =   450
      Width           =   975
   End
End
Attribute VB_Name = "frm071002"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/09/14 改成Form2.0 ; MSHFlexGrid1、lbe(index)、Text(index)
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/26 員工編號欄已修改
'Memo By Sindy 2010/8/3 日期欄已修改
Option Explicit

Dim Rs As New ADODB.Recordset, strCP09() As String, t As Integer, blnIsSave As Boolean
Dim strDate As String, LcTmp As String, strPubcp10() As String, lC() As String
Dim intLastRow As Integer, blnOKtoShow As Boolean, intCols As Integer, strOldLc As String
Dim blnIsNew As Boolean, lc01 As String, lc02 As String, lc03 As String, lc04 As String
Dim m_ODate As String '本所期限
Dim m_LDate As String '法定期限
Dim m_CurrSel As Integer
Dim m_CPCount As Integer
Dim m_Cpindex As Integer
'910703 Sieg 701
Dim m_CP60 As String, m_LC11 As String
'Add By Cheng 2002/08/22
'Dim m_strCust1 As String 'Mark by Lydia 2024/06/13
Dim m_LC22 As String 'Added by Lydia 2023/03/02 FC代理人
'add by nickc 2005/03/17 加乘註記
Dim m_CP98 As String
Dim m_CP101 As String
Dim m_CP104 As String
Dim m_CP65 As String 'Add By Sindy 2010/8/6
Dim strTemp As Variant 'Add By Sindy 2011/6/8
Dim m_CL02 As String, m_Text7 As String, m_CP75 As Double 'Add By Sindy 2011/6/8
Dim m_Text7_2 As String 'Add By Sindy 2012/2/21
Dim m_CP27 As String  'Add By Sindy 2012/6/1
Dim m_CP31 As String 'Added by Lydia 2020/08/18
Dim strChkCuAreaMail As String, strChkCuAreaMailTo As String 'Added by Lydia 2017/06/19 檢查收文智權人員和客戶智權人員不同業務區要發mail的內文和通知人員
'Added by Lydia 2020/05/20 法律所案源收文
Dim m_LOS15 As String '案源單號
Dim m_LOS01 As String '案源總收文號
Dim m_LOS01cp01 As String, m_LOS01cp02 As String, m_LOS01cp03 As String, m_LOS01cp04 As String '案源總收文號之本所案號
Dim m_LOS02 As String '案源案件類型
Dim t_LOS02 As String '需要變更：案源案件類型
Dim m_LOS04 As String, m_LOS04_1 As String  '介紹人、介紹人(第一位)
Dim m_LOS05 As String 'Added by Lydia 2021/06/18 介紹客戶編號
Dim m_LOS06 As String '法律所總收文號1
Dim m_LOS10 As String '收據總收文號
Dim m_CRL84 As String 'Added by Lydia 2020/10/07 接洽記錄單-法務案件屬性
Dim oObj As Control 'Added by Lydia 2022/08/10
Dim m_CP162 As String 'Added by Lydia 2023/08/14 (案件進度)案源單號
'Added by Lydia 2024/09/30 (113/11/01上線)
Dim bChkPaid As Boolean, m_CCP60 As String  '是否已付款, 收款之收據/請款單號
Dim bolActCaseLawer As Boolean '是否進入出庭律師維護
Dim m_LOS01fa As String '案源之FC代理人
'end 2024/09/30

'Add By Sindy 2011/6/8
Private Sub Check1_Click(Index As Integer)
   If Check1(Index).Value = 1 Then
      If InStr(Text(28).Text, Trim(Check1(Index).Caption)) = 0 Then
         If Text(28).Text = "" Then
            Text(28).Text = Trim(Check1(Index).Caption)
         Else
            Text(28).Text = Text(28).Text & "," & Trim(Check1(Index).Caption)
         End If
      End If
      'Added by Lydia 2023/03/14 一般案件屬性
      If Index = 4 Then
          Frame2.Enabled = True
      End If
      'end 2023/03/14
   Else
      '案件屬性=xx,xx,xx
      If Left(Text(28), Len(Trim(Check1(Index).Caption))) = Trim(Check1(Index).Caption) Then
         Text(28).Text = Replace(Text(28).Text, Trim(Check1(Index).Caption) & ",", "")
         Text(28).Text = Replace(Text(28).Text, Trim(Check1(Index).Caption), "")
      Else
         Text(28).Text = Replace(Text(28).Text, "," & Trim(Check1(Index).Caption), "")
      End If
      'Added by Lydia 2023/03/14 一般案件屬性
      If Index = 4 Then
         Frame2.Enabled = False
         For Each oObj In Check2
            If oObj.Value = 1 Then
                oObj.Value = 0
                Call Check2_Click(oObj.Index)
            End If
         Next
      End If
      'end 2023/03/14
   End If
   If InStr(Text(28).Text, "專利") > 0 Or InStr(Text(28).Text, "商標") > 0 Or _
      InStr(Text(28).Text, "著作權") > 0 Or InStr(Text(28).Text, "智財權") > 0 Then
      Text(5) = "Y"
   Else
      Text(5) = ""
   End If
End Sub

Private Sub cmdNext_Click()
Dim i As Integer
  
  'Added by Lydia 2021/09/14
  If m_Cpindex = 0 And m_CPCount = 1 Then
      MsgBox "已經是最後一筆!", vbInformation
      CmdNext.Enabled = False
      Exit Sub
  End If
  'end 2021/09/14
  
  ClearForm
  m_Cpindex = m_Cpindex + 1
  If m_Cpindex = m_CPCount - 1 Then
     CmdNext.Enabled = False
  ElseIf m_Cpindex = m_CPCount Then
     Exit Sub
  End If
  If UCase(Left(lC(m_Cpindex), 2)) = "LA" And strPubcp10(m_Cpindex) = "顧問聘任" Then
      intForm = 2
      For i = 0 To UBound(strCP09)
         ReDim Preserve strArryCP09(i)
            strArryCP09(i) = strCP09(i)
         ReDim Preserve strCP10(i)
            strCP10(i) = strPubcp10(i)
         ReDim Preserve strCaseKind(i)
            strCaseKind(i) = lC(i)
      Next
      intNowRec = m_Cpindex
      t = 0
      frm071003.Show
      Unload Me
  Else
     GetData (m_Cpindex)
  End If
  
End Sub

Private Sub cmdok_Click()
Dim yn As Integer, i As Integer, intWarnMsg As Integer
Dim oSubject As String, oContext As String, strText As String 'Add By Sindy 2012/2/21
  
   If AllTextBeforeSaveCheck Then Exit Sub
   'Add By Cheng 2002/05/24
   '重新檢查欄位有效性
   If TxtValidate = False Then Exit Sub
'   'Add By Cheng 2002/08/23
'   If Me.txtcp01.Text <> "" And Me.txtcp02.Text <> "" Then
'      MsgBox "本次只更新轉本所案號的資料，不更新其他欄位的修改!!!", vbCritical
'   End If
   
   '910703 Sieg 701F
   If txtCP01 <> "" And txtCP02 <> "" Then
      strExc(1) = txtCP01
      strExc(2) = txtCP02
      strExc(3) = txtCP03
      If strExc(3) = "" Then strExc(3) = "0"
      strExc(4) = txtCP04
      If strExc(4) = "" Then strExc(4) = "00"
      
      strExc(5) = Text(9).Text '案件性質
      strExc(6) = lbe(9) '案件性質名稱
      strExc(7) = Text(0) '收文日
      strExc(8) = lbePaperNum '總收文號
      '911118 nick 新增申請人
      strExc(9) = m_LC11
      'edit by nickc 2007/02/07 不用 dll 了
      'If Not objLawDll.ChkSameCase(strExc) Then Exit Sub
      If Not ClsLawChkSameCase(strExc) Then Exit Sub
      'Added by Lydia 2020/08/18 更新相關卷號前,先檢查是否有重複
      If m_CP31 = "Y" Then
          If PUB_ChkUpdCR(lc01, lc02, lc03, lc04, strExc(1), strExc(2), strExc(3), strExc(4)) = False Then
              Exit Sub
          End If
      End If
      'end 2020/08/18
   End If
   
   'Added by Lydia 2023/08/14 針對L-006685(AB2030779)先設出庭律師，後修改為不可輸入出庭律師案件性質的管制
   strExc(2) = ""
   If bolActCaseLawer = True Then   'Added by Lydia 2024/09/30
      If PUB_SaveCaseLawer(lbePaperNum, Mid(Me.Tag, InStr(Me.Tag, "|") + 1), , , True) = True Then
         strExc(2) = "Y"
      End If
   End If 'Added by Lydia 2024/09/30
   If Pub_ChkPtyCL(lc01, Trim(Text(9))) = False Then
      If m_CL02 <> "" Or strExc(2) = "Y" Then   '原本有「出庭律師」or有點選
         If MsgBox("【" & Trim(Text(9)) & " " & lbe(9) & "】不可輸入出庭律師，" & vbCrLf & "存檔將會刪除出庭律師記錄，是否繼續存檔？", vbExclamation + vbYesNo + vbDefaultButton2, "出庭律師檢查") = vbNo Then
            Exit Sub
         End If
      End If
   Else
      If Trim(Text(7)) = "" And (m_CL02 <> "" Or strExc(2) = "Y") Then
          MsgBox "已有出庭律師記錄，承辦人不可空白！", vbExclamation, "出庭律師檢查"
          Exit Sub
      End If
      'Added by Lydia 2024/09/30 (113/11/01上線) 出庭費領取：1.新增會計科目為2201131案件性質發文檢查沒有CaseLawer的設定(Pub_ChkPtyCL內含「出庭費特殊性質」檢查)
      strText = "Y"
      '2.檢查案源是否可以輸入出庭律師
      If Pub_ChkLosToCL(lbePaperNum.Caption, False, strExc(1)) = False Then
         strText = ""
         '另外檢查是否存在出庭費
         If strExc(1) <> "" Then
            strExc(0) = "select cl02,cl03 from caselawer where cl01='" & Trim(lbePaperNum.Caption) & "' and nvl(cl03,0)> 0 "
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               MsgBox strExc(1), vbExclamation + vbOKOnly, "出庭律師檢查"
               Exit Sub
            End If
         End If
      End If
      If strText = "Y" Then
         If bolActCaseLawer = True Then
            strExc(0) = "select rtrim(substr(R002,1,6)) as R002,to_number(R003) as R003 from rdatafactory where id = '" & strUserNum & "' and formname='frm071018' and seqno='" & Mid(Me.Tag, InStr(Me.Tag, "|") + 1) & "' and nvl(r003,'0')>'0' "
         Else
            strExc(0) = "select cl02,cl03 from caselawer where cl01='" & Trim(lbePaperNum.Caption) & "' and nvl(cl03,0)> 0 "
         End If
         '3.增加特定案件性質可輸出庭費，但也可以輸0表示有輸過。(出庭費可以輸入0的狀況1)
         strExc(2) = Pub_GetSpecMan("出庭費特殊性質")
         If InStr(lc01, "L") > 0 And InStr(";" & strExc(2) & ";", ";" & Text(9) & ";") > 0 Then
            '出庭費特殊性質的控制，因為此類案件性質多數不必輸出庭費，請改為開放可輸入出庭費，但不必檢查設定出庭費=0
         Else
            '4.案源為商標且有FC代理人之法務案34行政訴訟程序若已輸入0則不必再提醒。(出庭費可以輸入0的狀況2)
            If m_LOS01cp01 <> "" And m_LOS01cp01 <> "TT" And InStr(m_LOS01cp01, "T") > 0 And m_LOS01fa <> "" And Text(9) = "34" Then
                strExc(0) = Replace(strExc(0), "and nvl(cl03,0)>0", "")
                strExc(0) = Replace(strExc(0), "and nvl(r003,'0')>'0'", "")
            End If
   
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 0 Then
               If MsgBox("尚未設定出庭律師和出庭費，是否繼續存檔？", vbInformation + vbYesNo + vbDefaultButton2, "出庭律師檢查") = vbNo Then
                  Exit Sub
               End If
            Else
               strExc(1) = RsTemp.GetString(adClipString, , , ",")
               If Text(7) <> "" And InStr(strExc(1), Text(7)) = 0 Then
                  If MsgBox("承辦人不在出庭律師內，是否繼續存檔？", vbExclamation + vbYesNo + vbDefaultButton2, "出庭律師檢查") = vbNo Then
                     Exit Sub
                  End If
               End If
            End If
         End If
      End If
      'end 2024/09/30
   End If
   'end 2023/08/14
   
   
   'Add By Cheng 2002/11/18
   If Me.txtCP01.Text <> "" And Me.txtCP02.Text <> "" Then
      MsgBox "本次只更新轉本所案號的資料，不更新其他欄位的修改!!!", vbCritical
      
   'Added by Lydia 2023/12/14
   Else
     '檢查智財協作在分案時若未建立相關案號(caserelation1)時則跳提醒程序人員，但可選擇輸或不輸 !
     'Modified by Lydia 2023/12/15 PS及CPS之智財協作967，TT及S之智財協作737，L之智財協作7601，(也可用案件性質中文判斷)在分案時若未建立相關案號且為ACS且為TIPS的案件時，提醒文字：「案件性質為智財協作，請先依接洽單輸入相關卷號資料」。
     ' If lc01 = "L" And Text(9) = "7601" Then
      '   If PUB_IfCaseRelation1Exists(lc01, lc02, lc03, lc04) = False Then
     '       If MsgBox("案件性質為" & lbe(9).Caption & "，請確認接洽單是否有相關案號，是否補輸入？", vbExclamation + vbYesNo + vbDefaultButton1) = vbYes Then
     '          Exit Sub
     '       End If
      '   End If
     ' End If
     If lc01 = "L" And InStr(lbe(9).Caption, "智財協作") > 0 Then
        If PUB_ChkACSforTIPS(lc01 & lc02 & lc03 & lc04, , True) = False Then
           MsgBox "案件性質為" & lbe(9).Caption & "，請先依接洽單輸入相關卷號資料", vbExclamation
           Exit Sub
        End If
     End If
     'end 2023/12/15
   'end 2023/12/14
   End If
  
   'Added by Lydia 2017/06/19 (存檔前)檢查收文智權人員和客戶智權人員不同業務區要發mail的內文和通知人員
   'Modified by Lydia 2021/06/18 +案源之 介紹人(第一位) m_LOS04_1
   strChkCuAreaMail = PUB_ChkSameCustSales(lc01, lc02, lc03, lc04, lbePaperNum, Trim(Text(1)), "", "", "", "", strChkCuAreaMailTo, m_LOS04_1)
   
    'Added by Lydia 2020/01/21 所有系統別分案，若該收文號存在於下一程序(NP24)，若修改案件性質則彈訊息
    If Me.Text(9).Tag <> Me.Text(9).Text Then
        If Pub_CheckNP24Exists(lbePaperNum.Caption) = True Then
        End If
    End If
    'end 2020/01/21
      
   'Added by Lydia 2020/05/20 法律所案源收文：案件屬性有勾專利或商標或著作權時，若案件性質為1101~1104(民事委任律師)時，若案源非B1時詢問 是否需智慧所配合開庭？」，若選擇要則將案源更新為B1
   'Modified by Lydia 2020/05/29 分案和配合開庭通知整合為一封email
   'Mark by Lydia 2020/06/16 改成frm077005「智財訴訟案需專業部配合通知補收文作業」
   'If strSrvDate(1) >= 法律所案源收文啟用日 And m_LOS01 <> "" And (Check1(0).Value = 1 Or Check1(1).Value = 1 Or Check1(2).Value = 1) And InStr("1101,1102,1103,1104,", Format(Text(9), "0000") & ",") > 0 And m_LOS02 <> "B1" And m_Text7 = "" Then
   '    If MsgBox("是否需智慧所配合開庭？", vbInformation + vbYesNo + vbDefaultButton1) = vbYes Then
   '        If Text(7).Text = "" Then
   '            MsgBox "請輸入承辦人！", vbExclamation
   '            Text(7).SetFocus
   '            Text_GotFocus 7
   '            Exit Sub
   '        End If
   '        t_LOS02 = "B1"
   '    End If
   'End If
   ''end 2020/05/20
   'end 2020/06/16
   
   If Not SaveData Then
      DataErrorMessage (3)
   Else
      'Added by Lydia 2017/06/19 檢查收文智權人員和客戶智權人員不同業務區要發mail的內文和通知人員
      If strChkCuAreaMail <> "" Then
         'Modified by Lydia 2021/12/24 改主旨「案件收文通知」=>「分案通知」
         PUB_SendMail strUserNum, strChkCuAreaMailTo, "", "分案通知--此案收文非原智權人員(區)！", strChkCuAreaMail
      End If
      'end 2017/06/19
      
      'Add By Sindy 2011/6/8 修改承辦人或其他出庭律師時，若該收文號已收款(CP75>0)則MAIL通知71005
      'Mark by Lydia 2024/09/30 (113/11/01上線)出庭律師或出庭費有異動，發Email通知財務處總帳人員；同時整合原有Email通知：'Added by Lydia 2020/10/15 (法律所案源)若分案時已經收款則同時E-MAIL給系統特殊設定之財務處出納人員及智權人員＋'Add By Sindy 2011/6/8 修改承辦人或其他出庭律師時，若該收文號已收款(CP75>0)則MAIL通知71005
      'If ((m_CL02 <> "" And m_CL02 <> Trim(strPublicTemp)) Or (m_Text7 <> "" And m_Text7 <> Trim(Text(7)))) And _
      '   Val(m_CP75) > 0 Then
      '   oSubject = lc01 & "-" & lc02 & "-" & lc03 & "-" & lc04 & "(" & m_CP60 & ")修改承辦人且已收款，若有必要請自行調整出庭費傳票摘要！"
      '   oContext = oContext & "案件性質：" & lbe(9) & vbCrLf & vbCrLf
      '   'Modify By Sindy 2012/2/21 +信件內容
      '   If (m_Text7 <> "" And m_Text7 <> Trim(Text(7))) Then
      '      'Modified by Lydia 2015/10/05
      '      'oContext = oContext & "原承辦律師：" & m_Text7 & " " & m_Text7_2 & "　　改為：" & Trim(Text(7)) & " " & Trim(lbe(7)) & vbCrLf
      '      oContext = oContext & "原承辦人：" & m_Text7 & " " & m_Text7_2 & "　　改為：" & Trim(Text(7)) & " " & Trim(lbe(7)) & vbCrLf
      '   End If
      '   If (m_CL02 <> "" And m_CL02 <> Trim(strPublicTemp)) Then
      '      strTemp = Split(m_CL02, ",")
      '      strText = ""
      '      For i = 0 To UBound(strTemp) - 1
      '         strText = strText & strTemp(i) & " " & GetStaffName(strTemp(i), True) & ","
      '      Next i
      '      If strText <> "" Then strText = Left(strText, Len(strText) - 1)
      '      oContext = oContext & "原其他出庭律師：" & strText
      '      strTemp = Split(strPublicTemp, ",")
      '      strText = ""
      '      For i = 0 To UBound(strTemp) - 1
      '         'Added by Lydia 2022/12/08 判斷變更出庭費
      '         If strTemp(i) <> "" And InStr(strTemp(i), "|") > 0 Then
      '             strText = strText & Mid(strTemp(i), 1, InStr(strTemp(i), "|") - 1) & " " & GetStaffName(Mid(strTemp(i), 1, InStr(strTemp(i), "|") - 1), True) & "變更" & Mid(strTemp(i), InStr(strTemp(i), "|") + 1) & ","
      '         Else
       '        'end 2022/12/08
      '             strText = strText & strTemp(i) & " " & GetStaffName(strTemp(i), True) & ","
      '         End If 'Added by Lydia 2022/12/08
       '     Next i
      '      If strText <> "" Then strText = Left(strText, Len(strText) - 1)
      '      oContext = oContext & "　　改為：" & Trim(strText) & vbCrLf
      '   End If
       '  '2012/2/21 End
       '  oContext = oContext & vbCrLf & vbCrLf & "此程序已收款且修改承辦人員，若有必要請自行調整出庭費傳票摘要！" & vbCrLf
       '  'Modified by Lydia 2023/01/13
       '  'PUB_SendMail strUserNum, Pub_GetSpecMan("財務處出納人員"), "", oSubject, oContext
       '  PUB_SendMail strUserNum, Pub_GetSpecMan("財務處總帳人員"), "", oSubject, oContext
      'End If
      'end 2024/09/30
   End If
   
   'Added by Lydia 2020/10/07 (10/5) 若案件性質或案件屬性有改時Email通知秀玲提醒確認案源及金額是否需調整。案件屬性第1次設定時要與接洽單檔比較是否不同。
   'Modified by Lydia 2020/11/26  A3類案源為非訴訟案，點數都回智慧所，與屬性是否為智財權無關; ex.L-006316
   'If m_LOS01 <> "" Then
   'Modified by Lydia 2021/09/24 改成有案源就通知
   'If m_LOS01 <> "" And m_LOS02 <> "A3" Then
   'Modified by Lydia 2025/08/18 改用案源單號
   'If m_LOS01 <> "" Then
   If m_LOS15 <> "" Then
       strExc(0) = "": strExc(1) = ""
       If Text(28).Visible = True And Text(28).Locked = False Then 'Added by Lydia 2020/11/03 判斷可維護才檢查
            'Modified by Lydia 2021/09/09 不用與接洽單檔比較; ex. L-006229-1-00先是在基本檔維護拿掉案件屬性有發通知, 又在分案設定案件屬性因為與接洽單一致所以沒有發通知
            'If Text(28).Tag = "" Then  '與接洽單檔比較
            '    If PUB_ChkTwoStrLst(m_CRL84, Text(28).Text) = False Then
            '        strExc(1) = strExc(1) & "、案件屬性"
            '        strExc(0) = strExc(0) & vbCrLf & "原案件屬性：" & m_CRL84 & vbCrLf & "現案件屬性：" & Text(28).Text
            '    End If
            'ElseIf Text(28).Tag <> Text(28).Text Then
            If Text(28).Tag <> Text(28).Text Then
            'end 2021/09/09
                If PUB_ChkTwoStrLst(Text(28).Tag, Text(28).Text) = False Then
                    strExc(1) = strExc(1) & "、案件屬性"
                    'Modified by Lydia 2021/09/09
                    'strExc(0) = strExc(0) & vbCrLf & "原案件屬性：" & Text(28).Tag & vbCrLf & "現案件屬性：" & Text(28).Text
                    strExc(0) = strExc(0) & vbCrLf & "原案件屬性：" & IIf(Trim(Text(28).Tag) <> "", Trim(Text(28).Tag), "(空白)") & vbCrLf & "現案件屬性：" & IIf(Trim(Text(28).Text) <> "", Trim(Text(28).Text), "(空白)")
                End If
            End If
       End If 'Added by Lydia 2020/11/03
       If lc01 <> "LA" Then  'Added by Lydia 2020/11/03 排除顧問案
            If Text(9).Tag <> Text(9).Text Then
                Call ClsPDGetCaseProperty(lc01, Text(9).Tag, strExc(3))
                strExc(1) = strExc(1) & "、案件性質"
                strExc(0) = strExc(0) & vbCrLf & "原案件性質：" & Text(9).Tag & strExc(3) & vbCrLf & "現案件性質：" & Text(9).Text & lbe(9).Caption
                'Added by Lydia 2025/08/18
                'strExc(0) = strExc(0) & vbCrLf & "案源類別A類改成BC類並且專業部收文尚未分案，請一併清除LOS01。" 'Mark by Lydia 2025/10/20
            End If
       End If  'Added by Lydia 2020/11/03
       If strExc(0) <> "" Then
           '主旨
           strExc(1) = "法務分案" & lc01 & "-" & lc02 & IIf(lc03 <> "0", "-" & lc03, "") & IIf(lc04 <> "00", "-" & lc04, "") & "，改變" & Mid(strExc(1), 2)
           '內文
           strExc(2) = "法律所案號：" & lc01 & "-" & lc02 & IIf(lc03 <> "0", "-" & lc03, "") & IIf(lc04 <> "00", "-" & lc04, "") & "(" & lbePaperNum & ")" & vbCrLf & _
                            "專業部案號：" & m_LOS01cp01 & "-" & m_LOS01cp02 & IIf(m_LOS01cp03 <> "0", "-" & m_LOS01cp03, "") & IIf(m_LOS01cp04 <> "00", "-" & m_LOS01cp04, "") & "(" & m_LOS01 & ")" & vbCrLf & _
                             strExc(0)
           strExc(2) = strExc(2) & vbCrLf & vbCrLf & "請確認案源及金額是否需調整。" 'Added by Lydia 2021/09/09 加提醒
           'Added by Lydia 2025/10/20
           If lc01 <> "LA" And Text(9).Tag <> Text(9).Text Then
              'Memo by Lydia 2025/10/20 因為案源A類=TT總收文號，BC類 = P / T案總收文號(分案時回寫)，若修改收文性質可能要變更LOS01,LOS02
              'Ex. 114/8/13追查"Anny 的FCL-011035-1沒有發mail，是因為LawOfficeSource.LOS01非空白並且記錄為TT案收文號
              strExc(2) = strExc(2) & vbCrLf & "若案源類別A類改成BC類，請一併修改LOS01為P/T案收文號。"
              strExc(2) = strExc(2) & vbCrLf & "若案源類別BC類改成A類，請一併修改LOS01為TT案收文號。"
           End If
           'end 2025/10/20
           'Modified by Lydia 2023/02/01 改成系統特殊設定
           'PUB_SendMail strUserNum, "83002", "", strExc(1), strExc(2)
           PUB_SendMail strUserNum, Pub_GetSpecMan("程式管理人員"), "", strExc(1), strExc(2)
       End If
   End If
   'end 2020/10/07
   
   '2010/12/7 add by sonia
   ' 設定滑鼠游標為預設
   Screen.MousePointer = vbDefault
   ' 當轉本所案號時檢查原本所案號是否還有案件進度的資料
   If IsEmptyText(txtCP01) = False Then
      strExc(0) = "SELECT COUNT(*) FROM CASEPROGRESS WHERE " & ChgCaseprogress(lc01 & lc02 & lc03 & lc04)
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
      If RsTemp.Fields(0) < 1 Then
         MsgBox "原本所案號 " & lc01 & lc02 & lc03 & lc04 & "已無案件進度資料，請通知收文人員刪號！", vbInformation
      Else
         MsgBox "原本所案號為 " & lc01 & lc02 & lc03 & lc04 & "，請自行更新原本所案號之下一程序資料 !", vbInformation
      End If
   End If
   '2010/12/7 end
   
   'If UBound(strCP09) = t Then
   If m_Cpindex = m_CPCount - 1 Then
      cmdok.Enabled = False
      intForm = 0
      intNowRec = 0
      blnIsFormBack = True
      Unload Me
      frm071001.Show
      Exit Sub
   End If
   cmdNext_Click
'   t = t + 1
'   If Left(lc(m_Cpindex), 2) = "LA" And strPubcp10(m_Cpindex) = "顧問聘任" Then
'      intForm = 2
'      For i = 0 To UBound(strCP09)
'         ReDim Preserve strArryCP09(i)
'            strArryCP09(i) = strCP09(i)
'         ReDim Preserve strCP10(i)
'            strCP10(i) = strPubcp10(i)
'         ReDim Preserve strCaseKind(i)
'            strCaseKind(i) = lc(i)
'      Next
'      intNowRec = t
'      t = 0
'      frm071003.Show
'      Unload Me
'   Else
'       GetData (t)
'   End If
End Sub

Private Sub cmdPrePic_Click()
Dim yn As Integer
   
   If blnIsSave = False Then
      yn = MsgBox("你並未存檔,確定離開嗎?", vbYesNo + vbCritical + vbDefaultButton2)
      If yn = 7 Then
       Exit Sub
      End If
   End If
   intForm = 0
   intNowRec = 0
   blnIsFormBack = True
   Unload Me
   Unload frm071018
   frm071001.Show
End Sub

Private Sub ComBack_Click()
   If blnIsSave = False Then
      If MsgBox("你並未存檔,確定離開嗎?", vbYesNo + vbCritical + vbDefaultButton2) = vbNo Then
         Exit Sub
      End If
   End If
   intForm = 0
   intNowRec = 0
   Unload Me
   Unload frm071001
End Sub

Private Sub Command1_Click()
Dim i As Integer
Dim strNum As String
Dim strTmp As String
   
   strTmp = lbeNumber.Caption
   If strTmp = "" Then
      Exit Sub
   End If
   
   i = InStr(strTmp, "-")
   If i <> 0 Then
      strNum = Left(strTmp, i - 1)
      strTmp = Mid(strTmp, i + 1)
   End If
   frm1103_2.intWhereComeFrom = 1
   Set frm1103_2.m_form = Me
   frm1103_2.lblSystem = strNum
   'Modified by Lydia 2023/03/14 直接帶入本所案號; ex.L-006556-1顯示為L-006556-1-1
'   i = InStr(strTmp, "-")
'   If i <> 0 Then
'      strNum = Left(strTmp, i - 1)
'      If strTmp <> "" Then
'         strTmp = Mid(strTmp, i + 1)
'      End If
'      frm1103_2.lblCode(0) = strNum
'   Else
'      frm1103_2.lblCode(0) = strTmp
'      strTmp = ""
'   End If
'   If i <> 0 Then
'      i = InStr(strTmp, "-")
'      If i <> 0 Then
'         strNum = Left(strTmp, i - 1)
'         If strTmp <> "" Then
'            strTmp = Mid(strTmp, i + 1)
'         End If
'         frm1103_2.lblCode(1) = strNum
'      Else
'         frm1103_2.lblCode(1) = strTmp
'      End If
'   Else
'         frm1103_2.lblCode(1) = "0"
'   End If
'
'   If strTmp <> "" Then
'      frm1103_2.lblCode(2) = strTmp
'   Else
'      frm1103_2.lblCode(2) = "00"
'   End If
   frm1103_2.lblCode(0) = lc02
   frm1103_2.lblCode(1) = lc03
   frm1103_2.lblCode(2) = lc04
   'end 2023/03/14
   
   frm1103_2.Show
   Me.Hide
End Sub

Private Sub Command2_Click()
   frm07100202.Show
End Sub

Private Sub Command3_Click()
   frm07100203.Show
   If IsNoExistData Then Unload frm07100203
End Sub

'Add By Sindy 2011/6/8
Private Sub Command4_Click()
   'Added by Lydia 2023/03/16 須限制案件性質表規費科目(CPM12)為220113的才可點出庭律師。
   'Modified by Lydia 2023/08/14 改成模組
   'strExc(0) = "select cpm12 from casepropertymap where cpm01='" & lc01 & "' and cpm02='" & Trim(Text(9)) & "' "
   'intI = 1
   'Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   'If InStr(RsTemp.Fields("cpm12") & ",", "220113") = 0 Then
   If Pub_ChkPtyCL(lc01, Trim(Text(9))) = False Then
   'end 2023/08/14
       MsgBox "【" & Trim(Text(9)) & " " & lbe(9) & "】不可輸入出庭律師！", vbExclamation + vbOKOnly, "出庭律師檢查"
       Exit Sub
   End If
   'end 2023/03/16
   'Added by Lydia 2024/07/02 檢查案源是否可以輸入出庭律師
   If Pub_ChkLosToCL(lbePaperNum.Caption, True, strExc(1)) = False Then
      Exit Sub
   End If
   'end 2024/07/02
   
   'Modified by Lydia 2022/12/08
   'frm071018.Hide
   'Set frm071018.UpForm = frm071002
   'frm071018.lbePaperNum = Me.lbePaperNum
   'frm071018.lbeNumber = Me.lbeNumber
   'Modified by Lydia 2024/09/30 (113/11/01上線)傳入收文號、判斷是否進入第2次以上
   'Call frm071018.SetParent(Me, Me.lbePaperNum, IIf(Me.Tag = "", True, False), Trim(Text(7)))
   Call frm071018.SetParent(Me, Me.lbePaperNum, IIf(bolActCaseLawer = False, True, False), Trim(Text(7)), Trim(Text(9)), IIf(bolActCaseLawer = True And Me.Tag = "", True, False))
   'end 2022/12/08
   bolActCaseLawer = True 'Added by Lydia 2024/09/30
   Me.Hide
   frm071018.Show vbModal
End Sub

Private Sub Form_Load()
Dim i As Integer, n As Integer
Dim nPos As Integer
 
   MoveFormToCenter Me
   Call ClearForm 'Added by Lydia 2021/09/14
   
   'Added by Lydia 2020/05/20 法律所案源收文
   If strSrvDate(1) >= 法律所案源收文啟用日 Then
       LBL01.Visible = True
       lblLOS01.Visible = True
       CmdOk1.Visible = True
   Else
       LBL01.Visible = False
       lblLOS01.Visible = False
       CmdOk1.Visible = False
   End If
   'end 2020/05/20
   
   m_CPCount = 0
   t = 0
   blnIsSave = False
   If intForm = 3 Then
       t = intNowRec
       m_Cpindex = intNowRec
     For i = 0 To UBound(strArryCP09)
       ReDim Preserve strCP09(n)
         strCP09(n) = strArryCP09(n)
       ReDim Preserve strPubcp10(n)
         strPubcp10(n) = strCP10(n)
       ReDim Preserve lC(n)
         lC(n) = strCaseKind(n)
         n = n + 1
         m_CPCount = m_CPCount + 1
      Next
      
      If m_Cpindex = m_CPCount - 1 Then
         CmdNext.Enabled = False
      End If
      GetData (m_Cpindex)
   Else
       With frm071001.MSHFlexGrid1
           n = 0
           For i = 1 To .Rows - 1
            .row = i
            .col = 0
               If .Text = "v" Then
                  .col = 2
                      ReDim Preserve strCP09(n)
                      strCP09(n) = .Text
                      m_CPCount = m_CPCount + 1
                  .col = 3
                      ReDim Preserve strPubcp10(n)
                      strPubcp10(n) = .Text
                  .col = 4
                      ReDim Preserve lC(n)
                      nPos = InStr(.Text, "-")
                      If nPos <> 0 Then
                         lC(n) = Left(.Text, nPos - 1)
                      Else
                         lC(n) = ""
                      End If
                        n = n + 1
               End If
           Next
       End With
      GetData (0)
   End If
   If UCase(Left(lbeNumber.Caption, 2)) = "LA" Then
      lblName(0).Visible = False
      lblName(1).Visible = False
      Label5(1).Visible = False
      Label25.Visible = False
      Text(3).Visible = False
      Text(4).Visible = False
      Text(5).Visible = False
      'Add by Amy 2018/08/15 +專案服務案 L用
      Label52.Visible = False
      Text(52).Visible = False
      'end 2018/08/15
      'Add By Sindy 2011/6/8
      Label12.Visible = False
      Text(28).Visible = False
      'Modified by Lydia 2022/08/10
      'For i = 0 To 4
      '   Check1(i).Visible = False
      'Next i
      '2011/6/8 End
      For Each oObj In Check1
          oObj.Visible = False
      Next
      'end 2022/08/10
      'Added by Lydia 2023/03/14 一般案件屬性
      Frame2.Visible = False
      lblMemo.Visible = False
   End If
   intRecount = m_CPCount
   
   SSTab1.Tab = 0 'Added by Lydia 2023/03/14
   
   'Added by Lydia 2023/01/07 2023/01/06 設定可看開庭費及輸入開庭費之法律所同仁為系統特殊設定「出庭費維護」，其他法律所同仁請關閉權限
   If (Left(Pub_StrUserSt03, 1) = "L" And InStr(Pub_GetSpecMan("出庭費維護"), strUserNum) > 0) Or InStr("01,08,09,00", Pub_strUserST05) > 0 Then
      Command4.Visible = True
   Else
      Command4.Visible = False
   End If
   'end 2023/01/07
End Sub

Private Sub Form_Unload(Cancel As Integer)
   PUB_SendMailCache 'Add by Sindy 2010/6/18
   'Set frm071002 = Nothing 'Remove by Lydia 2023/03/14  form2.0會有問題，改在呼叫時清除記憶體變數
   'Add By Sindy 2011/6/8
   strPublicTemp = ""
   Unload frm071018
   '2011/6/8 End
End Sub

Private Sub MSHFlexGrid1_Click()
Dim intRow As Integer
  
  'Add By Cheng 2002/03/25
  intRow = Me.MSHFlexGrid1.MouseRow
  If intRow <= 0 Then Exit Sub
   
   If MSHFlexGrid1.Rows > 1 Then
      If MSHFlexGrid1.row > 0 Then
         If MSHFlexGrid1.TextMatrix(MSHFlexGrid1.row, 0) = "v" Then
            MSHFlexGrid1.TextMatrix(MSHFlexGrid1.row, 0) = Empty
         Else
            MSHFlexGrid1.TextMatrix(MSHFlexGrid1.row, 0) = "v"
         End If
      End If
   End If

   If MSHFlexGrid1.Rows < 2 Then Exit Sub
   'GridClick MSHFlexGrid1, intRow, 0
   'Modify By Cheng 2002/03/25
'   If MSHFlexGrid1.TextMatrix(1, 0) = "v" Then
'      Text(12).Text = ChangeTDateStringToTString(MSHFlexGrid1.TextMatrix(1, 2))
'      Text(13).Text = ChangeTDateStringToTString(MSHFlexGrid1.TextMatrix(1, 3))
'      Text(15).Text = MSHFlexGrid1.TextMatrix(1, 8)
'      'Modify By Cheng 2002/03/25
''      Text(18).Text = MSHFlexGrid1.TextMatrix(1, 13)
'      Text(18).Text = IIf(Len(Me.Text(18).Tag) > 0, Me.Text(18).Tag & "，" & MSHFlexGrid1.TextMatrix(1, 13), MSHFlexGrid1.TextMatrix(1, 13))
'   Else
'      Text(12).Text = ""
'      Text(13).Text = ""
'      Text(15).Text = ""
'      'Modify By Cheng 2002/03/25
''      Text(18).Text = ""
'      Text(18).Text = "" & Me.Text(18).Tag
'   End If
   'Modify By Cheng 2002/03/25
   If MSHFlexGrid1.TextMatrix(intRow, 0) = "v" Then
      Text(12).Text = ChangeTDateStringToTString(MSHFlexGrid1.TextMatrix(intRow, 2))
      Text(13).Text = ChangeTDateStringToTString(MSHFlexGrid1.TextMatrix(intRow, 3))
      Text(15).Text = MSHFlexGrid1.TextMatrix(intRow, 8)
      Text(18).Text = IIf(Len(Me.Text(18).Tag) > 0, Me.Text(18).Tag & IIf(Len(MSHFlexGrid1.TextMatrix(intRow, 13)) > 0, "，" & MSHFlexGrid1.TextMatrix(intRow, 13), ""), MSHFlexGrid1.TextMatrix(intRow, 13))
   Else
      Text(12).Text = ""
      Text(13).Text = ""
      Text(15).Text = ""
      Text(18).Text = "" & Me.Text(18).Tag
   End If
   
End Sub

Private Sub MSHFlexGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeySpace Then
      If MSHFlexGrid1.row > 0 Then
         If MSHFlexGrid1.TextMatrix(MSHFlexGrid1.row, 0) = "V" Then
            MSHFlexGrid1.TextMatrix(MSHFlexGrid1.row, 0) = Empty
         Else
            MSHFlexGrid1.TextMatrix(MSHFlexGrid1.row, 0) = "V"
         End If
      End If
   End If
End Sub

' 將MSHFlexGrid1所選取的列反白, 並將未選取的列設成一般顏色
Private Sub MSHFlexGrid1_ShowSelection()
Dim nCurrSel As Integer
Dim nCol As Integer
   
   nCurrSel = MSHFlexGrid1.row
   
   ' 與前一選擇的列位置相同則不處理
   If m_CurrSel = MSHFlexGrid1.row Then
      GoTo EXITSUB
   End If
   
   ' 將原先選取的列回復到正常的顏色
   If m_CurrSel > 0 And m_CurrSel < MSHFlexGrid1.Rows Then
      MSHFlexGrid1.row = m_CurrSel
      MSHFlexGrid1.col = 1
      If MSHFlexGrid1.CellBackColor <> &H80000005 Then
         For nCol = 1 To MSHFlexGrid1.Cols - 1
            MSHFlexGrid1.col = nCol
            If MSHFlexGrid1.CellBackColor <> &H80000005 Then: MSHFlexGrid1.CellBackColor = &H80000005
            If MSHFlexGrid1.CellForeColor <> &H80000008 Then: MSHFlexGrid1.CellForeColor = &H80000008
         Next nCol
      End If
      MSHFlexGrid1.col = 0
   End If
   ' 設定成所選取的列
   m_CurrSel = nCurrSel
   ' 將所選取的列反白
   If m_CurrSel > 0 And m_CurrSel < MSHFlexGrid1.Rows Then
      MSHFlexGrid1.row = m_CurrSel
      MSHFlexGrid1.col = 1
      For nCol = 1 To MSHFlexGrid1.Cols - 1
         MSHFlexGrid1.col = nCol
        ' MSHFlexGrid1.CellBackColor = &H8000000D
        MSHFlexGrid1.CellBackColor = &HFFC0C0
         MSHFlexGrid1.CellForeColor = &H80000008
      Next nCol
      MSHFlexGrid1.col = 0
   End If
EXITSUB:
End Sub

Private Sub MSHFlexGrid1_SelChange()
  MSHFlexGrid1_ShowSelection
End Sub

Private Sub Text_Change(Index As Integer)
Dim i As Integer
   
   Select Case Index
      Case 1, 7, 8, 9, 10
         If Text(Index) = "" Then lbe(Index) = ""
   End Select
End Sub

Private Sub Text_GotFocus(Index As Integer)
   TextInverse Text(Index)
   Select Case Index
      Case 12
         If Text(Index) <> "" Then strDate = Text(Index)
         'edit by nickc 2007/06/11  切換輸入法改用API
         'Text(Index).IMEMode = 2
         CloseIme
      Case 2, 4, 11, 18, 19
         'edit by nickc 2007/06/11  切換輸入法改用API
         'Text(Index).IMEMode = 1
         OpenIme
      Case Else
         'edit by nickc 2007/06/11  切換輸入法改用API
         'Text(Index).IMEMode = 2
         CloseIme
   End Select
End Sub

'Modified by Lydia 2021/09/14 改成Form 2.0
'Private Sub Text_KeyPress(Index As Integer, KeyAscii As Integer)
Private Sub Text_KeyPress(Index As Integer, KeyAscii As MSForms.ReturnInteger)
   Select Case Index
      'Modify by Amy 2018/08/15 +lc52(專案服務案)
      Case 1, 5, 6, 7, 8, 9, 10, 14, 15, 16, 17, 52
         KeyAscii = UpperCase(KeyAscii)
         'Add By Cheng 2002/04/24
         If Index = 16 Then
            If KeyAscii <> 89 And KeyAscii <> 8 Then
               KeyAscii = 0
            End If
         ElseIf Index = 5 Or Index = 52 Then
            If KeyAscii <> 8 And KeyAscii <> Asc("Y") Then
               KeyAscii = 0
               Beep
            End If
         End If
      'end 2018/08/15
   End Select
End Sub

Private Sub Text_LostFocus(Index As Integer)
Dim blnIsEmpty As Boolean
Dim i As Integer
   
   Select Case Index
         Case 2, 3, 4
            If Text(Index) <> "" Then Text(Index) = UCase(Text(Index))
            If Index = 4 Then
                For i = 2 To 4
                   If Text(i) <> "" Then
                       blnIsEmpty = False
                       Exit For
                   Else
                      blnIsEmpty = True
                    End If
                Next
                If blnIsEmpty Then
                      MsgBox "案件名稱不可同時為空", vbCritical
                     Text(2).SetFocus
                     Exit Sub
                 End If
              End If
   End Select
End Sub

'Added by Lydia 2023/03/14
Private Sub Text_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
   If Index = 28 Then
      If Button = 2 Then Forms(0).PopupMenu2 Text(Index)
   End If
End Sub

Private Sub Text_Validate(Index As Integer, Cancel As Boolean)
Dim strTempNo As String, strTempName As String, i As Integer, blnIsEmpty As Boolean
'Added by Lydia 2019/02/14
Dim m_SalesST15 As String '畫面上智權人員的收文部門
Dim m_Tuser As String '創新業務部預設收文人員
   
   Select Case Index
      Case 0
           If Text(Index) <> "" Then
              If CheckIsTaiwanDate(Text(Index)) Then
                 If Val(GetTaiwanTodayDate) - Val(Text(Index)) < 0 Then
                    DataErrorMessage 2, "收文日期"
                    Cancel = True
                 End If
              Else
                   Cancel = True
              End If
          Else
             MsgBox "收文日不可空白", vbCritical
             Cancel = True
          End If
      Case 1 '當事人
          If Text(Index) <> "" Then
             'edit by nickc 2007/02/07 不用 dll 了
             'If objPublicData.GetCustomer(Text(Index), strTempName) Then
             If ClsPDGetCustomer(Text(Index), strTempName) Then
                lbe(Index) = strTempName
               
                  '910703 Sieg 701
                  If m_CP60 <> "" And InStr(ChangeCustomerL(m_LC11), ChangeCustomerL(Text(Index))) = 0 Then
                     strExc(1) = lc01
                     strExc(2) = lc02
                     strExc(3) = lc03
                     strExc(4) = lc04
                     strExc(5) = m_CP60
                     strExc(6) = Text(Index)
                     strExc(7) = strTempName
                     '911118 nick 新增申請人
                     strExc(8) = m_LC11
                     'edit by nickc 2007/02/07 不用 dll 了
                     'If Not objLawDll.UpdAcc0k0(strExc()) Then
                     If Not ClsLawUpdAcc0k0(strExc()) Then
                        lbe(Index) = ""
                        Cancel = True
                     End If
                  End If
                
             Else
                Cancel = True
                lbe(Index) = ""
             End If
          Else
             If m_LC22 = "" Then 'Added by Lydia 2023/03/03 若案件有FC代理人時，當事人若空白僅提醒即可，不必限制不能分案
                 MsgBox "當事人不可空白", vbCritical
                 Cancel = True
             End If 'Added by Lydia 2023/03/03
          End If
          'Add By Cheng 2002/08/22
          If Cancel = False Then
            'Modified by Lydia 2024/06/13
            'If m_strCust1 <> Me.Text(1).Text Then
            If m_LC11 <> ChangeCustomerL(Me.Text(1).Text) Then
              'Modify By Cheng 2002/12/25
      '         If Not PUB_EditCustOk(Me.lbePaperNum.Caption, lc(1), lc(2), lc(3), lc(4)) Then Cancel = True
               If Not PUB_EditCustOk(Me.lbePaperNum.Caption, lc01, lc02, lc03, lc04) Then Cancel = True
            End If
          End If
          
      Case 2
            If CheckLengthIsOK(Text(2), 160) = False Then
                Cancel = True
            End If
      Case 4
            If CheckLengthIsOK(Text(4), 160) = False Then
                Cancel = True
            End If
      Case 5, 17
          If Text(Index) <> "" Then
             Text(Index) = UCase(Text(Index))
             If Text(Index) <> "Y" Then
                If Index = 5 Then
                   DataErrorMessage 1, "是否為智慧財產權案"
                   Cancel = True
                Else
                   DataErrorMessage 1, "是否取締案"
                   Cancel = True
                End If
             End If
           End If
      Case 6
          If Text(Index) <> "" Then
             Text(Index) = UCase(Text(Index))
             'add by nickc 2005/10/06 加大到 50 就要檢查
             If CheckLengthIsOK(Text(6), 50) = False Then
                Cancel = True
             End If
          End If
          
      Case 7, 8, 10
           '2010/1/28 智權人員欄鎖住
           If Text(Index) <> "" Then
              Text(Index) = UCase(Text(Index))
              'edit by nickc 2007/02/07 不用 dll 了
              'If objPublicData.GetStaff(Text(Index), strTempName) Then Lbe(Index) = strTempName Else Cancel = True: Lbe(Index) = ""
              If ClsPDGetStaff(Text(Index), strTempName) Then lbe(Index) = strTempName Else Cancel = True: lbe(Index) = ""
            
              'Added by Lydia 2019/02/14 創新業務部人員收文控管
              If Index = 10 Then
                 m_SalesST15 = GetST15(Text(Index))
                 If PUB_ChkIsT10T20("2", Text(Index).Text, m_Tuser, strTempName) = True Then
                     Text(Index).Text = m_Tuser
                     lbe(Index).Caption = strTempName
                     Text(Index).SetFocus
                     Call Text_GotFocus(Index)
                     Cancel = True
                     Exit Sub
                 End If
              End If
              'end 2019/02/14
           End If
      Case 9
           If Text(Index) = "" Then
              MsgBox "案件性質不可空白", vbCritical
              lbe(Index) = ""
              Cancel = True
           Else
               'edit by nickc 2007/02/07 不用 dll 了
               'If objPublicData.GetCaseProperty(CheckCaseNum, Text(Index), strTempName, False) Then Lbe(Index) = strTempName Else Cancel = True
               If ClsPDGetCaseProperty(CheckCaseNum, Text(Index), strTempName, False) Then lbe(Index) = strTempName Else Cancel = True
           End If
      Case 11
          If Text(Index) <> "" Then
             Text(Index) = UCase(Text(Index))
          End If
      Case 12
           If Text(Index) <> "" Then
              If CheckIsTaiwanDate(Text(Index)) Then
                  If Text(13) <> "" Then
                     If Val(Text(13)) - Val(Text(Index)) < 0 Then DataErrorMessage 13: Cancel = True
                  End If
              Else
                 Cancel = True
              End If
              If m_ODate <> "" Then
                 If Text(12) <> m_ODate Then
                    i = MsgBox("是否要修改此期限?", vbYesNo, "修改")
                    If i = 7 Then
                       Text(12) = m_ODate
                    End If
                 End If
              End If
           End If
       Case 13
            If Text(Index) <> "" Then
              If CheckIsTaiwanDate(Text(Index)) Then
                  If Text(12) <> "" Then
                     If Val(Text(Index)) - Val(Text(12)) < 0 Then DataErrorMessage 12: Cancel = True
                  End If
              Else
                 Cancel = True
              End If
           End If
           If m_LDate <> "" Then
              If Text(13) <> m_LDate Then
                 i = MsgBox("是否要修改此期限?", vbYesNo, "修改")
                 If i = 7 Then
                    Text(13) = m_LDate
                 End If
              End If
           End If
           
      Case 14
         Text(Index) = UCase(Text(Index))
      
          If Text(Index) <> "" Then
             If Text(Index) <> "N" Then
               DataErrorMessage 1, "是否算案件數"
               Cancel = True
             End If
          End If
      Case 15
          If Text(Index) <> "" Then
             Text(Index) = UCase(Text(Index))
             If Text(Index) = lbePaperNum Then
                MsgBox "且不可為本身之收文號", vbCritical
                Cancel = True
             End If
             'edit by nickc 2007/02/07 不用 dll 了
             'If Not objLawDll.GetRelation(LcTmp, lbePaperNum, Text(Index)) Then Cancel = True
             If Not ClsLawGetRelation(LcTmp, lbePaperNum, Text(Index)) Then Cancel = True
          End If
      Case 16
          If Text(Index) <> "" Then
                Text(Index) = UCase(Text(Index))
              If Text(Index) = "Y" Then
                 i = MsgBox("確定閉卷?", vbYesNo, "詢問")
                 If i = 7 Then Text(Index) = ""
              Else
                 DataErrorMessage 1, "是否閉卷": Cancel = True
              End If
          End If
      Case 18
          If Text(18).Text <> "" Then
             If CheckLengthIsOK(Text(18), 2000) = False Then
                Cancel = True
             End If
          End If
      Case 19
          If Text(19).Text <> "" Then
            If CheckLengthIsOK(Text(19), 2000) = False Then
                Cancel = True
            End If
          End If
      'Added by Lydia 2023/03/14 案件屬性：可直接輸入
      Case 28
          If Text(Index).Locked = False Then
             If Trim(Text(Index)) = "" Then
                For Each oObj In Check1
                   oObj.Value = 0
                Next
                For Each oObj In Check2
                   oObj.Value = 0
                Next
             Else
               strExc(1) = PUB_StringFilter(Text(Index))
               strTemp = Split(strExc(1), ",")
               For i = 0 To UBound(strTemp)
                  If Trim(strTemp(i)) <> "" Then
                      '案件屬性
                      For Each oObj In Check1
                         If Trim(strTemp(i)) = oObj.Caption And oObj.Value = 0 Then
                            oObj.Value = 1
                         ElseIf InStr(strExc(1), oObj.Caption) = 0 Then
                            oObj.Value = 0
                         End If
                      Next
                      '一般案件屬性
                      For Each oObj In Check2
                         If Trim(strTemp(i)) = oObj.Caption And oObj.Value = 0 Then
                            oObj.Value = 1
                            If Check1(4).Value = 0 Then
                                Check1(4).Value = 1
                                strExc(1) = "一般," & strExc(1)
                            End If
                         ElseIf InStr(strExc(1), oObj.Caption) = 0 Then
                            oObj.Value = 0
                         End If
                      Next
                  End If
               Next i
               Text(Index) = strExc(1)
             End If
          End If
      'end 2023/03/14
   End Select
   If Cancel Then TextInverse Text(Index)
End Sub

Private Sub txtcp01_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtcp01_Validate(Cancel As Boolean)
   If txtCP01 <> "" Then
     txtCP01 = UCase(txtCP01)
'     If txtcp01 <> GetCaseNumSysKind(lbeNumber) Then
'     DataErrorMessage 1, "本所案號"
'     Cancel = True
'     End If
   End If
   If Cancel Then TextInverse txtCP01
End Sub

Private Sub txtcp02_Validate(Cancel As Boolean)
Dim strTemp As String, i As Integer, yn As Integer, strlcTemp As String
   
   If txtCP02 <> "" Then
      'edit by nickc 2007/02/07 不用 dll 了
      'If objLawDll.ChkCaseNum(txtcp01, txtcp02) Then
      If ClsPDChkCaseNum(txtCP01, txtCP02) Then
         TextInverse txtCP02
         Cancel = True
      End If
   End If
   If Cancel Then TextInverse txtCP02
End Sub

Private Sub GetData(ByVal intI As Integer)
Dim lc05 As String, lc06 As String, lc07 As String, lc08 As String, _
    LC11 As String, lc13 As String, lc16 As String, lc27 As String
Dim yn As Boolean, i As Integer, j As Integer
Dim Rs As New ADODB.Recordset
Dim St(29) As String
 
   m_Cpindex = intI
   If lC(intI) <> "LA" Then
      i = 0
      'Add By Sindy 2010/8/6 增加CP65
      'Modify By Sindy 2011/6/8 +lc47,cp75,cp31
      'Modify By Sindy 2012/6/1 +cp27
      'Modify by Amy 2018/08/15 +LC52
      'Modified by Lydia 2023/03/03 +LC22
      'Modified by Lydia 2023/08/14 +CP162
      strExc(1) = "select cp01,cp02,cp03,cp04,cp05,cp06,cp07,cp09,cp10,cp13,cp14,cp16,cp18,cp19,cp21,cp26," + _
        "cp29,cp43,cp49,cp57,cp64,lc05,lc06,lc07,lc08,lc11,lc13,lc16,lc27,CP60,CP65,lc47,cp75,cp31,cp27,lc52,lc22,CP162 " + _
        "from lawcase, caseprogress where " + _
        "CP01=LC01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) and cp09='" + strCP09(intI) + "' order by LC01,LC02,LC03,LC04"
   Else
      i = 1
      'Add By Sindy 2010/8/6 增加CP65
      'Modify By Sindy 2011/6/8 +cp75,cp31
      'Modify By Sindy 2012/6/1 +cp27
      'Modified by Lydia 2023/08/14 +CP162
      strExc(1) = "select cp01,cp02,cp03,cp04,cp05,cp06,cp07,cp09,cp10,cp13,cp14,cp16,cp18,cp19,cp21,cp26," + _
        "cp29,cp43,cp49,cp57,cp64,hc05,hc06,hc07,hc09,hc12,CP60,CP65,cp75,cp31,cp27,CP162 " + _
        "from hirecase, caseprogress where " + _
        "CP01=HC01(+) AND CP02=HC02(+) AND CP03=HC03(+) AND CP04=HC04(+) and cp09='" + strCP09(intI) + "' order by HC01,HC02,HC03,HC04"
   End If
   intI = 1
   Set Rs = ClsLawReadRstMsg(intI, strExc(1))    'edit by nickc 2007/02/07 不用 dll 了 Set rs = objLawDll.ReadRstMsg(intI, strExc(1))
   If intI = 1 Then
      Select Case i
      Case 0 'ChangeCustomerL
          lc05 = IIf(IsNull(Rs.Fields!lc05), "", Rs.Fields!lc05)
          lc06 = IIf(IsNull(Rs.Fields!lc06), "", Rs.Fields!lc06)
          lc07 = IIf(IsNull(Rs.Fields!lc07), "", Rs.Fields!lc07)
          lc08 = IIf(IsNull(Rs.Fields!lc08), "", Rs.Fields!lc08)
          'lc11 = IIf(IsNull(Rs.Fields!lc11), "", ChangeCustomerS(Rs.Fields!lc11))
          If IsNull(Rs.Fields!LC11) Then
             LC11 = ""
          Else
             LC11 = ChangeCustomerS(Rs.Fields!LC11)
          End If
          lc13 = IIf(IsNull(Rs.Fields!lc13), "", Rs.Fields!lc13)
          lc16 = IIf(IsNull(Rs.Fields!lc16), "", Rs.Fields!lc16)
          lc27 = IIf(IsNull(Rs.Fields!lc27), "", Rs.Fields!lc27)
          m_LC22 = "" & Rs.Fields("LC22") 'Added by Lydia 2023/03/02 FC代理人
      Case 1
          LC11 = IIf(IsNull(Rs.Fields!hc05), "", ChangeCustomerS(Rs.Fields!hc05)) 'lc11=>hc05
          lc05 = IIf(IsNull(Rs.Fields!hc06), "", Rs.Fields!hc06)   'lc05=>hc06
          lc16 = IIf(IsNull(Rs.Fields!hc07), "", Rs.Fields!hc07)  'lc16=>hc07
          lc08 = IIf(IsNull(Rs.Fields!hc09), "", Rs.Fields!hc09) 'lc08=>hc09
          lc27 = IIf(IsNull(Rs.Fields!hc12), "", Rs.Fields!hc12) 'lc27=>hc12
      End Select
      For j = 1 To 21
         If IsNull(Rs.Fields(j - 1).Value) = False Then
            St(j) = Rs.Fields(j - 1).Value
         Else
            St(j) = ""
         End If
      Next
      
      'Add By Sindy 2011/6/8
      If Not IsNull(Rs.Fields("CP75")) Then
         m_CP75 = Rs.Fields("CP75")
      Else
         m_CP75 = 0
      End If
      
      '910703 Sieg 701
      If Not IsNull(Rs.Fields("CP60")) Then
         m_CP60 = Rs.Fields("CP60")
      Else
         m_CP60 = ""
      End If
      
      ' CreateID Add By Sindy 2010/8/6
      m_CP65 = ""
      If IsNull(Rs.Fields("CP65")) = False Then
         m_CP65 = Rs.Fields("CP65")
      End If
      
      'Add By Sindy 2012/6/1
      m_CP27 = Empty
      If IsNull(Rs.Fields("CP27")) = False Then
         m_CP27 = Rs.Fields("CP27")
      End If
      '2012/6/1 End
      m_CP162 = "" & Rs.Fields("CP162") 'Added by Lydia 2023/08/14 (案件進度)案源單號
      
      lc01 = St(1)
      lc02 = St(2)
      lc03 = St(3)
      lc04 = St(4)
      If lc01 <> "LA" Then
         If Not IsNull(Rs.Fields("LC11")) Then
            m_LC11 = Rs.Fields("LC11")
         Else
            m_LC11 = ""
         End If
      Else
         If Not IsNull(Rs.Fields("hc05")) Then
            m_LC11 = Rs.Fields("hc05")
         Else
            m_LC11 = ""
         End If
      End If
      
      Text(0) = ChangeWStringToTString(St(5))
      lbeNumber = GiveSymbol(St(1), St(2), St(3), St(4), LcTmp)
      lbeNumber.Tag = LcTmp
      Text(1) = IIf(IsNull(LC11), "", LC11): ChgType (1)
      Text(2) = IIf(IsNull(lc05), "", lc05)
      Text(3) = IIf(IsNull(lc06), "", lc06)
      Text(4) = IIf(IsNull(lc07), "", lc07)
      Text(5) = UCase(lc13)
      'Add By Sindy 2011/6/8
      If lc01 <> "LA" Then
         Text(28) = "" & Rs.Fields("lc47")
         Text(28).Tag = Text(28) 'Add by Amy 2018/07/30
         Text(52) = "" & Rs.Fields("lc52") 'Add by Amy 2018/08/15 專案服務案
      End If
      '案件屬性
      'Modified by Lydia 2022/08/10
      'For i = 0 To 4
      '   If InStr(Text(28).Text, Trim(Check1(i).Caption)) > 0 Then
      '      Check1(i).Value = 1
      '   End If
      'Next i
      ''2011/6/8 End
      For Each oObj In Check1
         If InStr(Text(28).Text, Trim(oObj.Caption)) > 0 Then
            oObj.Value = 1
         End If
      Next
      'end 2022/08/10
      'Added by Lydia 2023/03/14 一般案件屬性
      If Check1(4).Value = 1 Then
         Frame2.Enabled = True
      Else
         Frame2.Enabled = False
      End If
      For Each oObj In Check2
         If InStr(Text(28).Text, Trim(oObj.Caption)) > 0 Then
            oObj.Value = 1
            If Frame2.Enabled = False Then
               Frame2.Enabled = True
            End If
         End If
      Next
      'end 2023/03/14
      Text(6) = IIf(IsNull(lc16), "", lc16)
      'Modify By Cheng 2002/04/24
      '是否取消閉卷欄不要顯示資料
      '    Text(16) = UCase(lc08)
      'Add By Cheng 2002/04/22
      If UCase(lc08) = "Y" Then
         Me.lblClose.Caption = "已閉卷"
         Me.Text(16).Visible = True
         Me.Label21(1).Visible = True
         Me.Label29.Visible = True
      Else
         Me.lblClose.Caption = ""
         Me.Text(16).Visible = False
         Me.Label21(1).Visible = False
         Me.Label29.Visible = False
      End If
      Text(19) = IIf(IsNull(lc27), "", lc27)
      Text(12) = ChangeWStringToTString(St(6))
      Text(13) = ChangeWStringToTString(St(7))
      m_ODate = ChangeWStringToTString(St(6))
      m_LDate = ChangeWStringToTString(St(7))
      lbePaperNum = St(8)
      
      If St(9) = "" Then
         Text(9) = ""
         lbe(9) = ""
      Else
         Text(9) = St(9)
         ChgType (9)
      End If
      Text(9).Tag = Text(9).Text 'Added by Lydia 2020/01/21
      bChkPaid = PUB_ChkIsPaid(lbePaperNum, m_CCP60) 'Added by Lydia 2024/09/30 (113/11/01上線)已否已請款、已付款
      
      Text(10) = St(10): ChgType (10)
      Text(7) = St(11): ChgType (7)  '承辦人
      m_Text7 = Trim(Text(7)) 'Add By Sindy 2011/6/8
      m_Text7_2 = Trim(lbe(7)) 'Add By Sindy 2012/2/21
      lbeCost = St(12)
      lbePointNum = St(13)
      lbeMoney = St(14)
      Text(17) = UCase(St(15))
      Text(14) = UCase(St(16))
      Text(8) = St(17): ChgType (8) '協辦人員
      Text(15) = UCase(St(18))
      Text(11) = St(19)
      lbeCloseDate = ChangeWStringToTDateString(St(20))
      Text(18) = St(21)
      'Add By Cheng 2002/03/25
      Me.Text(18).Tag = Me.Text(18).Text
      'Add By Cheng 2002/08/22
      'm_strCust1 = "" & Me.Text(1).Text 'Mark by Lydia 2024/06/13
      
      m_CP31 = "" & Rs.Fields("CP31")  'Added by Lydia 2020/08/18
      'Add By Sindy 2011/6/20
      'CP31為Y時,Shape1內的欄位才可修改,否則鎖住 'Memo by Lydia 2023/03/14 不使用Shape1了
      If "" & Rs.Fields("CP31") = "Y" Then
         Text(1).Locked = False
         Text(2).Locked = False
         Text(3).Locked = False
         Text(4).Locked = False
         Text(5).Locked = False
         Text(6).Locked = False
         Text(28).Locked = False
         Text(52).Locked = False 'Add by Amy 2018/08/15
         'Modified by Lydia 2022/08/10
         'For i = 0 To 4
         '   Check1(i).Enabled = True
         'Next i
         For Each oObj In Check1
             oObj.Enabled = True
         Next
         'end 2022/08/10
         'Added by Lydia 2023/03/14 一般案件屬性
         Frame2.BackColor = &HC0FFFF
         For Each oObj In Check2
             oObj.Enabled = True
         Next
         'end 2023/03/14
      Else
         Text(1).Locked = True
         Text(2).Locked = True
         Text(3).Locked = True
         Text(4).Locked = True
         Text(5).Locked = True
         Text(6).Locked = True
         Text(28).Locked = True
         Text(52).Locked = True 'Add by Amy 2018/08/15
         'Modified by Lydia 2022/08/10
         'For i = 0 To 4
         '   Check1(i).Enabled = False
         'Next i
         For Each oObj In Check1
             oObj.Enabled = False
         Next
         'end 2022/08/10
         'Added by Lydia 2023/03/14 一般案件屬性
         Frame2.BackColor = &H8000000F
         For Each oObj In Check2
             oObj.Enabled = False
         Next
         'end 2023/03/14
      End If
      'Added by Lydia 2016/01/27 L,CFL案要勾選案件屬性
      If lc01 = "L" Then
         If Text(28).Text = "" Then
            Text(28).Locked = False
            'Modified by Lydia 2022/08/10
            'For i = 0 To 4
            '   Check1(i).Enabled = True
            'Next i
            For Each oObj In Check1
                oObj.Enabled = True
            Next
            'end 2022/08/10
         'Mark by Lydia 2023/02/01 前面已有設定Locked
         'Else
         '   Text(28).Locked = True
         '   'Modified by Lydia 2022/08/10
         '   'For i = 0 To 4
         '   '   Check1(i).Enabled = False
          '  'Next i
          '  For Each oObj In Check1
         '       oObj.Enabled = False
         '   Next
         '   'end 2022/08/10
         'end 2023/02/01
         End If
      End If
      'end 2016/01/27
      
      Getrs
      
      'Add By Sindy 2011/6/8
      strPublicTemp = ""
      strExc(0) = "select cl02 from caselawer where cl01='" + lbePaperNum + "' order by cl02 asc "
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         RsTemp.MoveFirst
         Do While Not RsTemp.EOF
            If IsNull(RsTemp.Fields(0).Value) = False Then
               strPublicTemp = strPublicTemp & RsTemp.Fields(0).Value & ","
            End If
            RsTemp.MoveNext
         Loop
      End If
      m_CL02 = Trim(strPublicTemp) 'Add By Sindy 2011/6/8
      
      'Modify By Sindy 2012/6/1 C類來函或已發文案件須鎖住轉本所案號欄位, 若為併號請以聯絡單通知電腦中心處理
      If Trim(lbePaperNum.Caption) < "C" And Val(m_CP27) = 0 Then
          Me.txtCP01.Enabled = True
          Me.txtCP02.Enabled = True
          Me.txtCP03.Enabled = True
          Me.txtCP04.Enabled = True
      Else
          Me.txtCP01.Enabled = False
          Me.txtCP02.Enabled = False
          Me.txtCP03.Enabled = False
          Me.txtCP04.Enabled = False
      End If
      '2012/6/1 End
   End If
   blnIsSave = False
   
   'Add by Morgan 2003/12/07
   Call PUB_CheckSales(lc01, lc02, lc03, lc04, Text(0), Text(10), lbe(10))
   'End 2003/12/07
   
   '2007/8/13 ADD BY SONIA銷卷提醒
   CheckCaseDestroy lc01, lc02, lc03, lc04
   '2007/8/13 END
   
   'Added by Lydia 2020/05/20 法律所案源收文
   If strSrvDate(1) >= 法律所案源收文啟用日 Then
        Call ReadLOS
        lblLOS01.Caption = m_LOS01
        If lblLOS01.Caption <> "" Then
            CmdOk1.Visible = True
        Else
            CmdOk1.Visible = False
        End If
   End If

End Sub

Private Function SaveData() As Boolean
Dim i As Integer, blnIsChange As Boolean
Dim cp01 As String, cp02 As String, cp03 As String, cp04 As String
Dim strTmp As String
Dim strTmp1(1 To 3) As String
'Add By Cheng 2002/08/23
Dim iStep As Integer
'Add By Cheng 2002/09/09
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
'edit by nickc 2007/02/08
'Dim sHC(1 To T_HC) As String
'Dim sLC(1 To T_LC) As String
Dim sHC() As String
Dim sLC() As String
ReDim sHC(1 To TF_HC) As String
ReDim sLC(1 To TF_LC) As String

   'Add By Cheng 2002/11/07
   On Error GoTo ErrorHandler
   SaveData = True
   cnnConnection.BeginTrans
   
   iStep = 1
   '若有輸入轉本所案號
   If Me.txtCP01.Text <> "" And Me.txtCP02.Text <> "" Then
      cp01 = txtCP01
      cp02 = txtCP02
      cp03 = Left(txtCP03 & "0", 1)
      cp04 = Left(txtCP04 & "00", 2)
      blnIsChange = True
      
'cancel by sonia 2024/11/26 已不立卷不必再通知分所收文人員
'      'Add by Sindy 2010/8/6 若為分所收文案件則發Mail通知收文人員
'      strExc(0) = PUB_GetST06(m_CP65)
'      If strExc(0) > "1" Then
'         strExc(1) = "原本所案號 " & lc01 & "-" & lc02 & IIf(lc03 & lc04 = "000", "", "-" & lc03 & "-" & lc04)
'         strExc(1) = strExc(1) & " 已更改為 " & cp01 & "-" & cp02 & IIf(cp03 & cp04 = "000", "", "-" & cp03 & "-" & cp04) & " 。"
'         'Modify By Sindy 2010/12/3
''         strSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08)" & _
''            " values ('" & strUserNum & "','" & m_CP65 & "',to_char(sysdate,'yyyymmdd'),to_char(sysdate,'hh24miss')" & _
''            ",'" & ChgSQL(strExc(1)) & "','如旨' )"
'         strSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08)" & _
'            " values ('" & strUserNum & "','" & m_CP65 & "',to_char(sysdate,'yyyymmdd'),to_char(sysdate,'hh24miss')" & _
'            ",'" & ChgSQL(strExc(1)) & "','總收文號：" & lbePaperNum & " 改本所案號如主旨')"
'         cnnConnection.Execute strSql
'      End If
'      'end 2010/8/6
'end 2024/11/26
      
      'Add By Cheng 2002/08/23
      GoTo ProcChgCaseNum
   '若未輸入轉本所案號
   Else
     cp01 = lc01
     cp02 = lc02
     cp03 = lc03
     cp04 = lc04
     blnIsChange = False
   End If
   
   LcTmp = cp01 + cp02 + cp03 + cp04
   Select Case GetCaseNumSysKind(lbeNumber)
      Case "L"
         i = 0
         'Modify By Cheng 2002/04/24
'         ' 91.04.04 modify by louis (修改單引號)
'         strExc(1) = "update lawcase set lc05=" + CNULL(ChgSQL(Text(2))) + ",lc06=" + CNULL(ChgSQL(Text(3))) + ",lc07=" + CNULL(ChgSQL(Text(4))) + _
'            ", lc08 = " + CNULL(Text(16)) + ",lc11=" + CNULL(ChangeCustomerL(Text(1))) + ",lc13=" + CNULL(Text(5)) + ",lc16=" + CNULL(Text(6)) + _
'            ", lc27=" + CNULL(ChgSQL(Text(19))) + " where " & ChgLawcase(LcTmp)
         
         '910703 Sieg 701
         If Text(1) <> "" Then
            'edit by nickc 2007/02/07 不用 dll 了
            'If objPublicData.GetCustomerNameAndAddress(Text(1).Text, strExc(0), strTmp1(1), strTmp1(2), strTmp1(3)) Then
            If ClsPDGetCustomerNameAndAddress(Text(1).Text, strExc(0), strTmp1(1), strTmp1(2), strTmp1(3)) Then
               '修改申請人時
               If InStr(ChangeCustomerL(m_LC11), ChangeCustomerL(Text(1))) = 0 Then
                  If m_CP60 <> "" Then
                     strExc(1) = lc01
                     strExc(2) = lc02
                     strExc(3) = lc03
                     strExc(4) = lc04
                     strExc(5) = m_CP60
                     strExc(6) = Text(1)
                     strExc(7) = strExc(0)
                     '911118 nick 新增申請人
                     strExc(8) = m_LC11
                     'edit by nickc 2007/02/07 不用 dll 了
                     'If Not objLawDll.UpdAcc0k0(strExc(), True) Then
                     If Not ClsLawUpdAcc0k0(strExc(), True) Then
                        Text(1).SetFocus
                        'Modify By Cheng 2002/11/07
'                        Exit Function
                        GoTo ErrorHandler
                     End If
                  End If
               End If
            End If
         End If
         
         If Me.Text(16).Text = "Y" Then
            strExc(1) = ", LC08=NULL, LC09=NULL, LC10=NULL "
         Else
            strExc(1) = " "
         End If
         
         'Modify By Sindy 2011/6/8 +LC47
         'Modify by Amy 2018/08/15 +LC52
         strExc(1) = "update lawcase set lc05=" + CNULL(ChgSQL(Text(2))) + _
            ",lc06=" + CNULL(ChgSQL(Text(3))) + ",lc07=" + CNULL(ChgSQL(Text(4))) + _
            strExc(1) + ",lc11=" + CNULL(ChangeCustomerL(Text(1))) + _
            ",lc13=" + CNULL(Text(5)) + ",lc16=" + CNULL(Text(6)) + _
            ", lc27=" + CNULL(ChgSQL(Text(19))) + ", lc47=" + CNULL(ChgSQL(Text(28))) + _
            ",lc52=" + CNULL(ChgSQL(Text(52))) + _
            " where " & ChgLawcase(LcTmp)
        'Add by Amy 2018/07/30 有修改lc47寫log
        If Text(28).Visible = True Then
            If Text(28).Tag <> Text(28) Then Pub_SeekTbLog strExc(1)
        End If
        'Add By Cheng 2002/11/07
        cnnConnection.Execute strExc(1)
        
         'Add By Cheng 2002/08/23
         iStep = iStep + 1
      
      Case "LA"
         i = 1
         'Modify By Cheng 2002/04/24
'         ' 91.04.04 modify by louis (修改單引號)
'         strExc(1) = "update hirecase set hc05=" + CNULL(ChangeCustomerL(Text(1))) + ",hc06=" + CNULL(Text(2)) + ",hc07=" + CNULL(Text(6)) + _
'            " ,hc09=" + CNULL(Text(16)) + " ,hc12=" + CNULL(ChgSQL(Text(19))) + " where " & ChgHirecase(LcTmp)
         
         '910703 Sieg 701
         If Text(1) <> "" Then
            'edit by nickc 2007/02/07 不用 dll 了
            'If objPublicData.GetCustomerNameAndAddress(Text(1).Text, strExc(0), strTmp1(1), strTmp1(2), strTmp1(3)) Then
            If ClsPDGetCustomerNameAndAddress(Text(1).Text, strExc(0), strTmp1(1), strTmp1(2), strTmp1(3)) Then
               '修改申請人時
               If InStr(ChangeCustomerL(m_LC11), ChangeCustomerL(Text(1))) = 0 Then
                  If m_CP60 <> "" Then
                     strExc(1) = lc01
                     strExc(2) = lc02
                     strExc(3) = lc03
                     strExc(4) = lc04
                     strExc(5) = m_CP60
                     strExc(6) = Text(1)
                     strExc(7) = strExc(0)
                     '911118 nick 新增申請人
                     strExc(8) = m_LC11
                     'edit by nickc 2007/02/07 不用 dll 了
                     'If Not objLawDll.UpdAcc0k0(strExc(), True) Then
                     If Not ClsLawUpdAcc0k0(strExc(), True) Then
                        Text(1).SetFocus
                        'Modify By Cheng 2002/11/07
'                        Exit Function
                        GoTo ErrorHandler
                     End If
                  End If
               End If
            End If
         End If
         
         If Me.Text(16).Text = "Y" Then
            strExc(1) = ", HC09=NULL, HC10=NULL, HC11=NULL "
         Else
            strExc(1) = " "
         End If
         strExc(1) = "update hirecase set hc05=" + CNULL(ChangeCustomerL(Text(1))) + _
            ",hc06=" + CNULL(Text(2)) + ",hc07=" + CNULL(Text(6)) + strExc(1) + _
            " ,hc12=" + CNULL(ChgSQL(Text(19))) + " where " & ChgHirecase(LcTmp)
        'Add By Cheng 2002/11/07
        cnnConnection.Execute strExc(1)
         'Add By Cheng 2002/08/23
         iStep = iStep + 1
   End Select
      
   '2005/12/20 MODIFY BY SONIA
   'strTmp = "cp05=" + CNULL(IIf(Text(0) = "", "", ChangeTStringToWString(Text(0)))) + ",cp06=" + CNULL(IIf(Text(12) = "", "", ChangeTStringToWString(Text(12)))) + ", cp07=" + CNULL(IIf(Text(13) = "", "", ChangeTStringToWString(Text(13)))) + ",cp10=" + CNULL(Text(9)) + _
   '      ",cp13=" + CNULL(Text(10)) + ",cp14=" + CNULL(Text(7)) + ",cp21=" + CNULL(Text(17)) + ",cp26=" + CNULL(Text(14)) + _
   '      ", cp29=" + CNULL(Text(8)) + ",cp43=" + CNULL(Text(15)) + ",cp64=" + CNULL(Text(18)) + ",cp49=" + CNULL(Text(11)) + " where cp09=" + CNULL(lbePaperNum)
   'edit by nickc 2007/11/13 智權人員有更動時，一併更動業務區
   'strTmp = "cp05=" + CNULL(IIf(Text(0) = "", "", ChangeTStringToWString(Text(0)))) + ",cp06=" + CNULL(IIf(Text(12) = "", "", ChangeTStringToWString(Text(12)))) + ", cp07=" + CNULL(IIf(Text(13) = "", "", ChangeTStringToWString(Text(13)))) + ",cp10=" + CNULL(Text(9)) + _
         ",cp13=" + CNULL(Text(10)) + ",cp14=" + CNULL(Text(7)) + ",cp21=" + CNULL(Text(17)) + ",cp26=" + CNULL(Text(14)) + _
         ", cp29=" + CNULL(Text(8)) + ",cp43=" + CNULL(Text(15)) + ",cp64=" + CNULL(Text(18)) + ",cp49=" + CNULL(Text(11)) + " where cp09=" + CNULL(lbePaperNum)
   'Modified by Lydia 2023/03/20 +chgsql
   strTmp = "cp05=" + CNULL(IIf(Text(0) = "", "", ChangeTStringToWString(Text(0)))) + ",cp06=" + CNULL(IIf(Text(12) = "", "", ChangeTStringToWString(Text(12)))) + ", cp07=" + CNULL(IIf(Text(13) = "", "", ChangeTStringToWString(Text(13)))) + ",cp10=" + CNULL(Text(9)) + _
         ",cp12=" + CNULL(GetST15(Text(10))) + ",cp13=" + CNULL(Text(10)) + ",cp14=" + CNULL(Text(7)) + ",cp21=" + CNULL(Text(17)) + ",cp26=" + CNULL(Text(14)) + _
         ", cp29=" + CNULL(Text(8)) + ",cp43=" + CNULL(Text(15)) + ",cp64=" + CNULL(ChgSQL(Text(18))) + ",cp49=" + CNULL(ChgSQL(Text(11))) + " where cp09=" + CNULL(lbePaperNum)
   '2005/12/20 END
'Add By Cheng 2002/08/23
ProcChgCaseNum:
   
    'Modify By Cheng 2002/11/07
'   SaveData = objLawDll.ExecSQL(2, strExc)
'   SaveData = objLawDll.ExecSQL(iStep - 1, strExc)
   'Add By Cheng 2002/09/09
   If blnIsChange Then
      If Me.txtCP01.Text <> "" And Me.txtCP02.Text <> "" And SaveData = True Then
         '判斷是否新增法務或服務業務基本案
         Select Case lc01
            Case "LA"
               StrSQLa = "SELECT * FROM HIRECASE WHERE " & ChgHirecase(cp01 & cp02 & cp03 & cp04)
            Case Else:
               StrSQLa = "SELECT * FROM LAWCASE WHERE " & ChgLawcase(cp01 & cp02 & cp03 & cp04)
         End Select
         rsA.CursorLocation = adUseClient
         rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
         If rsA.RecordCount <= 0 Then
            Select Case lc01
               Case "LA"
                  If PUB_ReadHireCaseData(sHC(), lc01, lc02, lc03, lc04) Then
                     sHC(1) = Me.txtCP01.Text
                     sHC(2) = Me.txtCP02.Text
                     sHC(3) = Left(Me.txtCP03.Text & "0", 1)
                     sHC(4) = Left(Me.txtCP04.Text & "00", 2)
                     If PUB_AddNewHireCase(sHC()) Then
                        'Add By Cheng 2002/11/07
                        Else
                            GoTo ErrorHandler
                     End If
                  End If
               Case Else:
                  If PUB_ReadLawCaseData(sLC(), lc01, lc02, lc03, lc04) Then
                     sLC(1) = Me.txtCP01.Text
                     sLC(2) = Me.txtCP02.Text
                     sLC(3) = Left(Me.txtCP03.Text & "0", 1)
                     sLC(4) = Left(Me.txtCP04.Text & "00", 2)
                     If PUB_AddNewLawCase(sLC()) Then
                        'Add By Cheng 2002/11/07
                        Else
                            GoTo ErrorHandler
                     End If
                  End If
            End Select
        'Add By Cheng 2002/12/06
        '若基本檔有資料, 若是否續辦欄為'Y'更新為Null
        Else
              strSql = " Update CaseProgress Set CP31=DECODE(CP31,'Y',NULL,CP31) WHERE " & ChgCaseprogress(cp01 & cp02 & cp03 & cp04)
              cnnConnection.Execute strSql
         End If
         If rsA.State <> adStateClosed Then rsA.Close
         Set rsA = Nothing
      End If
   End If
   
   If blnIsChange Then
      'Modify By Cheng 2002/08/23
'      strExc(iStep) = "update caseprogress set cp01=" + CNULL(CP01) + ",cp02=" + CNULL(cp02) + ",cp03=" + CNULL(cp03) + ",cp04=" + CNULL(cp04) + _
'         "," & strTmp
      strExc(iStep) = "update caseprogress set cp01=" + CNULL(cp01) + ",cp02=" + CNULL(cp02) + ",cp03=" + CNULL(cp03) + ",cp04=" + CNULL(cp04) + _
         " WHERE CP09='" & Me.lbePaperNum.Caption & "'"
        'Add By Cheng 2002/11/07
        cnnConnection.Execute strExc(iStep)
      'Add By Cheng 2002/08/23
      iStep = iStep + 1
      strExc(iStep) = "UPDATE CASEPROGRESS SET CP43='' WHERE CP09='" & Me.lbePaperNum.Caption & "'"
      Pub_SeekTbLog strExc(iStep) 'Added by Lydia 2025/09/25 L-006989誤輸入轉入ACS案
        'Add By Cheng 2002/11/07
        cnnConnection.Execute strExc(iStep)
      iStep = iStep + 1
      
      'Added by Lydia 2020/08/18 更新CaseRelation1和DivisionCase
      If m_CP31 = "Y" Then
          Call PUB_UpdateCaseRelation1(lc01, lc02, lc03, lc04, cp01, cp02, cp03, cp04)
      End If
      'end 2020/08/18
      
      'Add by Sindy 2010/8/12
      '更正財務相關資料
      PUB_UpdateAccData Trim(lbePaperNum), lc01 & lc02 & lc03 & lc04
   Else
      strExc(iStep) = "update caseprogress set " & strTmp
      'Added by Lydia 2023/08/14 有修改案件性質寫入記錄
      If Text(9).Text <> Text(9).Tag Then
         Pub_SeekTbLog strExc(iStep)
      End If
      'end 2023/08/14
      
        'Add By Cheng 2002/11/07
        cnnConnection.Execute strExc(iStep)
      'Add By Cheng 2002/08/23
      iStep = iStep + 1
   End If
     
   'Modify By Cheng 2002/08/23
   If Not blnIsChange Then
        'Modify By Cheng 2002/11/07
'      SaveNextProgress
      If SaveNextProgress = False Then GoTo ErrorHandler
   End If
   If SaveData Then blnIsSave = True
   frm071001.SetDataComplete lbePaperNum.Caption
   
   'add by nickc 2005/03/17 加入加乘註記及寄件值
   m_CP98 = "": m_CP101 = "": m_CP104 = ""
   If PUB_GetFlagValue(Me.lbePaperNum.Caption, m_CP98, m_CP101, m_CP104) = True Then
      strSql = "update caseprogress set cp98=" & m_CP98 & ",cp101=" & m_CP101 & ",cp104=" & m_CP104 & " WHERE CP09 = '" & Me.lbePaperNum.Caption & "' "
      cnnConnection.Execute strSql
   End If
      
   'Added by Lydia 2023/08/14 針對L-006685(AB2030779)先設出庭律師，後修改為不可輸入出庭律師案件性質的管制
   If Pub_ChkPtyCL(lc01, Trim(Text(9))) = False Then
      If m_LOS15 <> "" And m_CL02 <> "" Then
         strSql = "delete from CaseLawer where cl01='" & lbePaperNum & "' "
         'Modified by Lydia 2025/07/22 傳入收文號
         'Pub_SeekTbLog strSql
         Pub_SeekTbLog strSql, , , , , lbePaperNum
         cnnConnection.Execute strSql
         Call PUB_UpdateTTFee(m_LOS15)
      End If
   Else
   'end 2023/08/14
      'Add By Sindy 2011/6/8 更新其他出庭律師
      'Modified by Lydia 2022/12/08 配合輸入出庭費,改成先存暫存檔再寫入正式Table
      'strExc(0) = "delete from caselawer where cl01='" & lbePaperNum & "'"
      'cnnConnection.Execute strExc(0)
      'If strPublicTemp <> "" Then
      '   strTemp = Empty 'Added by Lydia 2020/10/15
      '   strTemp = Split(strPublicTemp, ",")
      '   For i = 0 To UBound(strTemp) - 1
      '      strExc(0) = "insert into caselawer values('" & lbePaperNum & "','" & strTemp(i) & "')"
      '      cnnConnection.Execute strExc(0)
      '   Next i
      If Me.Tag <> "" And InStr(Me.Tag, "|") > 0 Then '有點選「出庭律師」
         If PUB_SaveCaseLawer(lbePaperNum, Mid(Me.Tag, InStr(Me.Tag, "|") + 1), strPublicTemp) = True Then
            strTemp = Split(strPublicTemp, ",")
            'Memo by Lydia 2024/09/30 (113/11/01上線)PUB_SaveCaseLawer回傳strPublicTemp有包含變更出庭費;ex. 員工編號|出庭費：15000=>7500 ,
         End If
         '判斷補資料的情況
         'Mark by Lydia 2024/09/30 (113/11/01上線)經過測試不用了
         'If strPublicTemp = m_CL02 & Text(7) & "," Then
         '  m_CL02 = strPublicTemp
         'End If
         'end 2024/09/30
      'Added by Lydia 2024/09/30 (113/11/01上線)刪除所有「出庭律師」記錄
      ElseIf Me.Tag = "" And bolActCaseLawer = True And m_CL02 <> "" Then
         StrSQLa = "select cl01,cl02 from caselawer where cl01='" & lbePaperNum & "' "
         intI = 1
         Set rsA = ClsLawReadRstMsg(intI, StrSQLa)
         If intI = 1 Then
            rsA.MoveFirst
            Do While Not rsA.EOF
               strSql = "Delete from CaseLawer Where CL01='" & rsA.Fields("cl01") & "' and CL02='" & rsA.Fields("cl02") & "' "
               'Modified by Lydia 2025/07/22 傳入收文號
               'Pub_SeekTbLog strSql
               Pub_SeekTbLog strSql, , , , , "" & rsA.Fields("cl01")
               cnnConnection.Execute strSql
               rsA.MoveNext
            Loop
         End If
      'end 2024/09/30
      End If
      If strPublicTemp <> m_CL02 Then 'Memo by Lydia 2024/09/30 (113/11/01上線)回傳strPublicTemp有包含變更出庭費
      'end 2022/12/08
         'Added by Lydia 2024/09/30 (113/11/01上線)出庭律師或出庭費有異動，發Email通知財務處總帳人員；同時整合原有Email通知：'Added by Lydia 2020/10/15 (法律所案源)若分案時已經收款則同時E-MAIL給系統特殊設定之財務處出納人員及智權人員＋'Add By Sindy 2011/6/8 修改承辦人或其他出庭律師時，若該收文號已收款(CP75>0)則MAIL通知71005
         strExc(2) = "": strExc(4) = ""
         strTemp = Split(strPublicTemp, ",")
         For intI = 0 To UBound(strTemp)
            If Trim(strTemp(intI)) <> "" Then
                If InStr(strTemp(intI), "|") > 0 Then
                    strExc(3) = GetStaffName(Mid(Trim(strTemp(intI)), 1, InStr(Trim(strTemp(intI)), "|") - 1), True)
                    strExc(4) = strExc(4) & ";" & Mid(Trim(strTemp(intI)), 1, InStr(Trim(strTemp(intI)), "|") - 1) 'Added by Lydia 2024/11/04
                Else
                    strExc(3) = GetStaffName(strTemp(intI), True)
                    strExc(4) = strExc(4) & ";" & strTemp(intI) 'Added by Lydia 2024/11/04
                End If
                'strExc(4) = strExc(4) & ";" & strExc(3) 'Mark by Lydia 2024/11/04 debug
                If strExc(3) <> lbe(7).Caption Then
                   strExc(2) = strExc(2) & "、" & strExc(3)
                End If
            End If
         Next intI
         If strExc(2) <> "" Then strExc(2) = Mid(strExc(2), 2)  '其他出庭律師(名稱)
         If strExc(4) <> "" Then strExc(4) = Mid(strExc(4), 2)  'Added by Lydia 2024/11/04 承辦人+其他出庭律師
         
         If m_CCP60 = "" Then  '通知出庭律師
            strExc(0) = strExc(4)
         Else
            strExc(0) = Pub_GetSpecMan("財務處總帳人員")
         End If
         If strExc(0) <> "" Then
            'Modiied by Lydia 2025/05/26 將案源的介紹人員改至副本收件人
            'If m_CCP60 <> "" Then strExc(0) = strExc(0) & IIf(m_LOS04 <> "", ";" & Replace(m_LOS04, ",", ";"), "")
            strExc(9) = ""
            If m_CCP60 <> "" Then
               strExc(9) = IIf(m_LOS04 <> "", Replace(m_LOS04, ",", ";"), "")
            End If
            'end 2025/05/26
            
            '主旨
            strExc(1) = lc01 & "-" & lc02 & IIf(lc03 <> "0", "-" & lc03, "") & IIf(lc04 <> "00", "-" & lc04, "") & _
                        IIf(m_LOS01cp01 <> "" And m_LOS01cp01 <> "TT", "(案源案號：" & m_LOS01cp01 & "-" & m_LOS01cp02 & "-" & m_LOS01cp03 & "-" & m_LOS01cp04 & ")", "") & _
                        IIf(Val(m_CP27) > 0, "發文後", "") & "修改承辦人或出庭律師資料，" & _
                        IIf(m_CCP60 = "", "因尚未請款，暫時無法做出庭費之確認！", IIf(bChkPaid = True, "但已收款，請調整收款及傳票資料。", "因尚未收款，請注意是否在待發放出庭費清單內。"))
            '內文
            strExc(3) = "本所案號：" & lc01 & "-" & lc02 & IIf(lc03 <> "0", "-" & lc03, "") & IIf(lc04 <> "00", "-" & lc04, "") & vbCrLf
            If m_LOS01cp01 <> "" And m_LOS01cp01 <> "TT" Then
               strExc(3) = strExc(3) & "案源案號：" & m_LOS01cp01 & "-" & m_LOS01cp02 & IIf(m_LOS01cp03 <> "0", "-" & m_LOS01cp03, "") & IIf(m_LOS01cp04 <> "00", "-" & m_LOS01cp04, "") & vbCrLf
            End If
            strExc(3) = strExc(3) & "出庭律師：" & lbe(7).Caption & IIf(strExc(2) <> "", "、" & strExc(2), "") & vbCrLf
            If m_CCP60 <> "" Then strExc(3) = strExc(3) & IIf(Left(m_CCP60, 1) = "E", "收據號碼：", "請款單號：") & m_CCP60 & vbCrLf
            strExc(3) = strExc(3) & "收款狀態：" & IIf(m_CCP60 = "", "未請款", IIf(bChkPaid = True, "已收款", "未收款")) & vbCrLf
            'Modified by Lydia 2025/05/26 +CC=mc09
            StrSQLa = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08,mc09)" & _
                     " values ('" & strUserNum & "','" & strExc(0) & "'," & strSrvDate(1) & ",to_char(sysdate,'hh24miss')" & _
                     ",'" & ChgSQL(strExc(1)) & "','" & ChgSQL(strExc(3)) & "','" & strExc(9) & "')"
            cnnConnection.Execute StrSQLa
         End If
         'end 2024/09/30
         
         'Added by Lydia 2020/10/15 若有多個出庭律師，每一個律師扣一次15000或5000。案源類型A4、B1、B2要更新回TT-999999案源收文號之費用CP16欄並扣點數CP18
         If strSrvDate(1) >= 法律所案源收文啟用日 And m_LOS15 <> "" And InStr("A4,B1,B2", m_LOS02) > 0 Then
             strExc(1) = "" & UBound(strTemp)
             strExc(2) = "0"
             If m_CL02 <> "" Then
                 strTemp = Split(m_CL02, ",")
                 strExc(2) = "" & UBound(strTemp)
             End If
             '出庭人數有變,才更新扣款
             'Modified by Lydia 2022/12/12 判斷有變更就更新扣款
             'If Val(strExc(1)) <> Val(strExc(2)) Then
             If strPublicTemp <> m_CL02 Then
                  Call PUB_UpdateTTFee(m_LOS15)
                  '若分案時已經收款則同時E-MAIL給系統特殊設定之財務處出納人員及智權人員，主旨：L-XXXXXX有多個出庭律師，但已收款，請調整收款及傳票資料。內文：本所案號、收據號碼、出庭律師。
                  'Mark by Lydia 2024/09/30 (113/11/01上線)出庭律師或出庭費有異動，發Email通知財務處總帳人員；同時整合原有Email通知：'Added by Lydia 2020/10/15 (法律所案源)若分案時已經收款則同時E-MAIL給系統特殊設定之財務處出納人員及智權人員＋'Add By Sindy 2011/6/8 修改承辦人或其他出庭律師時，若該收文號已收款(CP75>0)則MAIL通知71005
                  'If m_CP60 <> "" Then
                  '   StrSQLa = "Select A0m01,A0m02,A0m03,A0m04,A0m05 From Acc0m0 Where a0m02=" & CNULL(m_CP60)
                  '   intI = 1
                  '   Set rsA = ClsLawReadRstMsg(intI, StrSQLa)
                  '   If intI = 1 Then
                  '      'Modified by Lydia 2023/01/13
                  '      'strExc(0) = Pub_GetSpecMan("財務處出納人員")
                  '      strExc(0) = Pub_GetSpecMan("財務處總帳人員")
                  '      If strExc(0) <> "" Then
                  '         strExc(0) = strExc(0) & IIf(m_LOS04 <> "", ";" & Replace(m_LOS04, ",", ";"), "")
                  '         '主旨
                  '         strExc(1) = lc01 & "-" & lc02 & IIf(lc03 <> "0", "-" & lc03, "") & IIf(lc04 <> "00", "-" & lc04, "") & "有多個出庭律師，但已收款，請調整收款及傳票資料。"
                  '         '其他出庭律師
                  '         strExc(2) = ""
                  '         'Modified by Lydia 2022/12/08 分析字串
                  '         'If strPublicTemp <> "" Then
                  '         '    StrSQLa = "select getstaffnamelist(" & CNULL(strPublicTemp) & ") from dual "
                  '         '    intI = 1
                  '         '    Set RsTemp = ClsLawReadRstMsg(intI, StrSQLa)
                  '         '    If intI = 1 Then
                  '         '        strExc(2) = "" & RsTemp(0)
                  '         '    End If
                  '         'End If
                  '         strTemp = Split(strPublicTemp, ",")
                  '         For intI = 0 To UBound(strTemp)
                  '            If Trim(strTemp(intI)) <> "" Then
                  '                If InStr(strTemp(intI), "|") > 0 Then
                  '                    strExc(3) = GetStaffName(Mid(Trim(strTemp(intI)), 1, InStr(Trim(strTemp(intI)), "|") - 1), True)
                  '                Else
                  '                    strExc(3) = GetStaffName(strTemp(intI), True)
                  '                End If
                  '                If strExc(3) <> lbe(7).Caption Then
                  '                   strExc(2) = strExc(2) & "、" & strExc(3)
                  '                End If
                  '            End If
                  '         Next intI
                  '         If strExc(2) <> "" Then strExc(2) = Mid(strExc(2), 2)
                  '         'end 2022/12/08
                  '
                  '         '內文
                  '         strExc(3) = "本所案號：" & lc01 & "-" & lc02 & IIf(lc03 <> "0", "-" & lc03, "") & IIf(lc04 <> "00", "-" & lc04, "") & vbCrLf & _
                  '                          "收據號碼：" & m_CP60 & vbCrLf & "出庭律師：" & lbe(7).Caption & IIf(strExc(2) <> "", "、" & strExc(2), "")
                  '         StrSQLa = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08)" & _
                  '                  " values ('" & strUserNum & "','" & strExc(0) & "'," & strSrvDate(1) & ",to_char(sysdate,'hh24miss')" & _
                  '                  ",'" & ChgSQL(strExc(1)) & "','" & ChgSQL(strExc(3)) & "')"
                  '         cnnConnection.Execute StrSQLa
                  '      End If
                  '   End If
                  'End If 'Mark by Lydia 2024/09/30 (113/11/01上線)整合通知
             End If   'If strPublicTemp <> m_CL02 Then
         End If
         'end 2020/10/15
      End If
   End If 'Added by Lydia 2023/08/14
   
   'PUB_UpdateCaseValue Me.lbePaperNum.Caption 'Remove by Morgan 2005/4/13 改由 trigger 更新
   
   'Added by Lydia 2020/05/20 法律所案源收文
   If strSrvDate(1) >= 法律所案源收文啟用日 And m_LOS15 <> "" Then
       'Mark by Lydia 2020/06/16 改成frm077005「智財訴訟案需專業部配合通知補收文作業」
       'strExc(1) = ""
       'If m_LOS02 <> t_LOS02 Then
       '     '先EMAIL通知智權人員補收配合開庭(因為LOS01會被清空)
       '     Call PUB_AddMailCache_LOS("3", m_LOS15)
       '     '案件屬性有勾專利或商標或著作權時，若案件性質為1101~1104(民事委任律師)時，
       '     '若案源非B1時詢問 是否需智慧所配合開庭？」，若選擇要則將案源更新為B1，同時將案源總收文號LOS01清除；
       '     strSql = "Update LawOfficeSource set LOS01=null, LOS02='B1' where LOS15='" & m_LOS15 & "' "
       '     cnnConnection.Execute strSql, i
       '     strExc(1) = "Y"
       '     If m_LOS10 <> "" Then
       '         '並重新計算TT-999999費用點數更新回去並改案件性質為736(服務費)；再EMAIL通知智權人員補收配合開庭226(B1)
       '         strSql = "Update CaseProgress set CP10='736' where CP09='" & m_LOS10 & "' "
       '         cnnConnection.Execute strSql, i
       '     End If
       'End If
       'end 2020/06/16
    
       '法律所總收文號1=> 分案時E-MAIL給介紹人通知法務案已收文之郵件
       If Me.lbePaperNum.Caption = m_LOS06 And m_Text7 = "" And m_Text7 <> Text(7).Text Then
           'Modified by Lydia 2020/05/29 分案和配合開庭通知整合為一封email
           'Modified by Lydia 2020/06/16
           'If strExc(1) <> "Y" Then Call PUB_AddMailCache_LOS("2", m_LOS15)
           Call PUB_AddMailCache_LOS("2", m_LOS15)
       End If
       'Added by Lydia 2021/06/18 補上介紹客戶 ex.L-006401(AB0024215)收文時只輸入代理人,於分案時輸入客戶
       If m_LOS05 = "" And Text(1) <> "" Then
           strSql = "Update LawOfficeSource set los05='" & ChangeCustomerL(Text(1)) & "' where los15='" & m_LOS15 & "' "
           cnnConnection.Execute strSql
       End If
       'end 2021/06/18
   End If
   'end 2020/05/20
   
   'Add By Cheng 2002/11/07
   cnnConnection.CommitTrans
Exit Function
ErrorHandler:
    cnnConnection.RollbackTrans
    SaveData = False
End Function

Private Sub Getrs()
   'Modified by Lydia 2023/03/22 法律案(L,FCL)不同審級收文期限沖銷：案件不同審級會收文-1、-2…案號，但實際是同一案件。ChgNextProgress(LcTmp)=>IIf(lc01 = "L", "np02='" & lc01 & "' and np03='" & lc02 & "' ", ChgNextProgress(LcTmp))
   strExc(1) = "select '',decode(np02 || np07, cpm01 || CPM02, CPM03, CPM04)," + _
      " decode(np08,null,'',SUBSTR(np08, 1, 4)- 1911 || '/' || SUBSTR(np08, 5, 2)|| '/' || SUBSTR(np08, 7, 2))," + _
      " decode(np09,null,'',SUBSTR(np09, 1, 4)- 1911 || '/' || SUBSTR(np09, 5, 2)|| '/' || SUBSTR(np09, 7, 2))," + _
      " np13,np14,decode(np11,null,'',SUBSTR(np11, 1, 4)- 1911 || '/' || SUBSTR(np11, 5, 2)|| '/' || " + _
      " SUBSTR(np11, 7, 2)),np06,np01,np07,np16,np17,np18,np15,np22 from nextprogress,CASEPROPERTYMAP where" + _
      " " & IIf(lc01 = "L", "np02='" & lc01 & "' and np03='" & lc02 & "' ", ChgNextProgress(LcTmp)) & " and np02=cpm01(+) and np07=cpm02(+) and (np06='N' or np06 is null)"
   intI = 1
   'edit by nickc 2007/02/07 不用 dll 了
   'Set MSHFlexGrid1.Recordset = objLawDll.ReadRstMsg(intI, strExc(1))
   Set MSHFlexGrid1.Recordset = ClsLawReadRstMsg(intI, strExc(1))
   GridHead
   
End Sub

Private Sub GridHead()
Dim i As Integer

   With MSHFlexGrid1
      blnOKtoShow = False
      .row = 0
      .col = 0: .ColWidth(0) = 200: .Text = "v"
      .CellAlignment = flexAlignCenterCenter
      .col = 1: .ColWidth(1) = 900: .Text = "下一程序"
      .col = 2: .ColWidth(2) = 1000: .Text = "本所期限"
      .col = 3: .ColWidth(3) = 900: .Text = "法定期限"
      .col = 4: .ColWidth(4) = 900: .Text = "機關文號"
      .col = 5: .ColWidth(5) = 900: .Text = "相關人"
      .col = 6: .ColWidth(6) = 1500: .Text = "解除期限日期"
      .col = 7: .ColWidth(7) = 0
      .col = 8: .ColWidth(8) = 0
      .col = 9: .ColWidth(9) = 0
      .col = 10: .ColWidth(10) = 0
      .col = 11: .ColWidth(11) = 0
      .col = 12: .ColWidth(12) = 0
      'Modify By Cheng 2002/03/25
      '將備註欄顯示出來
      '.Col = 13: .ColWidth(13) = 0
      .col = 13: .ColWidth(13) = 1500: .Text = "備註"
      .col = 14: .ColWidth(14) = 0
      intLastRow = 0
      blnOKtoShow = True
      '判斷是否有資料
   End With

End Sub

Private Function SaveNextProgress() As Boolean
Dim i As Integer, n As Integer, NP07 As String, np08 As String
Dim np16 As String, np17 As String, np18 As String, np06 As String, np01 As String
Dim np22 As String
   
   'Add By Cheng 2002/11/07
   On Error GoTo ErrorHandler
   SaveNextProgress = True

   With MSHFlexGrid1
      n = 0
         For i = 1 To .Rows - 1
          .row = i
          .col = 0
           If .Text = "v" Then
               .col = 2
                np08 = ChangeTStringToWString(Replace(.Text, "/", ""))
               .col = 8
               np01 = .Text
               .col = 9
               NP07 = .Text
               .col = 10
               np16 = .Text
               .col = 11
               np17 = .Text
               .col = 12
               np18 = .Text
               .col = 14
               np22 = .Text
               np06 = "Y"
               'Modified by Lydia 2023/03/22 法律案(L,FCL)不同審級收文期限沖銷：案件不同審級會收文-1、-2…案號，但實際是同一案件。ChgNextProgress(LcTmp)=>IIf(lc01 = "L", "np02='" & lc01 & "' and np03='" & lc02 & "' ", ChgNextProgress(LcTmp))
               strExc(1) = "update nextprogress set np06=" & CNULL(np06) & _
                  " where np01=" & CNULL(np01) & " and " & IIf(lc01 = "L", "np02='" & lc01 & "' and np03='" & lc02 & "' ", ChgNextProgress(LcTmp)) & _
                  " and np07=" & CNULL(NP07) & " and np08=" & CNULL(np08) & _
                  " and np16=" & CNULL(np16) & " and np17=" & CNULL(np17) & _
                  " and np18=" & CNULL(np18) & " and np22=" & CNULL(np22)
                  
                'Modify By Cheng 2002/11/07
'               SaveNextProgress = objLawDll.ExecSQL(1, strExc)
                cnnConnection.Execute strExc(1)
           End If
         Next
   End With
Exit Function
ErrorHandler:
    SaveNextProgress = False
End Function

Private Function ChangText() As Boolean
Dim i As Integer, strTemp As String, yn As Integer, strlcTemp As String

   strlcTemp = GiveSymbol(txtCP01, txtCP02, txtCP03, txtCP04)
   If lbeNumber = strlcTemp Then
      MsgBox "此本所案號與原本所案號相同", vbCritical
      ChangText = True
      txtCP01 = ""
      txtCP02 = ""
      txtCP03 = ""
      txtCP04 = ""
      Exit Function
   End If
   
   If txtCP02 <> "" Then
      'edit by nickc 2007/02/07 不用 dll 了
      'If objLawDll.ChkCaseNum(txtcp01, txtcp02) Then
      If ClsPDChkCaseNum(txtCP01, txtCP02) Then
            TextInverse txtCP02
            ChangText = True
      Else
         If txtCP01 = "LA" Then i = 2 Else i = 1
            'edit by nickc 2007/02/07 不用 dll 了
            'If Not objLawDll.CheckIsExistCaseNum(i, Replace(strlcTemp, "-", ""), strTemp) Then
            If Not ClsPDCheckIsExistCaseNum(i, Replace(strlcTemp, "-", ""), strTemp) Then
               yn = MsgBox("" + strTemp + ",是否要轉入此本所案號", vbYesNo)
              Select Case yn
               Case 6
                 strOldLc = lbeNumber
                   'Modify By Cheng 2002/12/06
                   '維持原本所案號
   '              lbeNumber = GiveSymbol(txtcp01, txtcp02, txtcp03, txtcp04, LcTmp)
                  blnIsNew = True
                 Getrs
               Case 7
                 txtCP01 = ""
                 txtCP02 = ""
                 txtCP03 = ""
                 txtCP04 = ""
               End Select
           Else
              MsgBox "" + strTemp + "", vbCritical
              ChangText = True
           End If
      End If
   End If
End Function

Private Function AllTextBeforeSaveCheck() As Boolean
Dim i As Integer
Dim strTempName As String
      
      AllTextBeforeSaveCheck = True
      If Text(0) = "" Then
         MsgBox "收文日不可空白", vbCritical
         Text(0).SetFocus
         AllTextBeforeSaveCheck = True
         Exit Function
      Else
        If CheckIsTaiwanDate(Text(0)) Then
           If Val(GetTaiwanTodayDate) - Val(Text(0)) < 0 Then
              DataErrorMessage 2, "收文日期"
              Text(0).SetFocus
              AllTextBeforeSaveCheck = True
              Exit Function
           End If
        Else
           Text(0).SetFocus
           AllTextBeforeSaveCheck = True
           Exit Function
        End If
      End If
      For i = 2 To 4
         If Text(i).Text <> "" Then
            Exit For
         Else
            If i = 4 Then
              MsgBox " 案件名稱不可同時空白", vbCritical
              Text(2).SetFocus
              AllTextBeforeSaveCheck = True
              Exit Function
             End If
         End If
      Next
         
      'Added by Lydia 2023/03/03 若案件有FC代理人時，當事人若空白僅提醒即可，不必限制不能分案
      'If Text(1) = "" Or IsNull(Text(1)) Then
      If m_LC22 <> "" And Trim(Text(1)) = "" Then
          If MsgBox("當事人代號為空白，是否繼續存檔？", vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then
             Text(1).SetFocus
             AllTextBeforeSaveCheck = True
             Exit Function
          End If
      ElseIf Text(1) = "" Or IsNull(Text(1)) Then
      'end 2023/03/03
         MsgBox "當事人代號不可空白", vbCritical
         Text(1).SetFocus
         AllTextBeforeSaveCheck = True
         Exit Function
      End If
      If Text(9) = "" Or IsNull(Text(9)) Then
         MsgBox "案件性質不可空白", vbCritical
         Text(9).SetFocus
         AllTextBeforeSaveCheck = True
         Exit Function
      End If
      
      If Text(5) <> "" Then
         Text(5) = UCase(Text(5))
         If Text(5) <> "Y" Then
               DataErrorMessage 1, "是否為智慧財產權案"
               Text(5).SetFocus
               AllTextBeforeSaveCheck = True
               Exit Function
         End If
      End If
      
      If Text(6) <> "" Then
         Text(6) = UCase(Text(6))
      End If
     If Text(7) <> "" Then
        Text(7) = UCase(Text(7))
        'edit by nickc 2007/02/07 不用 dll 了
        'If objPublicData.GetStaff(Text(7), strTempName) Then
        If ClsPDGetStaff(Text(7), strTempName) Then
           lbe(7) = strTempName
        Else
           lbe(7) = ""
           Text(7).SetFocus
           TextInverse Text(7)
           AllTextBeforeSaveCheck = True
           Exit Function
        End If
     End If
     strTempName = ""
     If Text(8) <> "" Then
        Text(8) = UCase(Text(8))
        'edit by nickc 2007/02/07 不用 dll 了
        'If objPublicData.GetStaff(Text(8), strTempName) Then
        If ClsPDGetStaff(Text(8), strTempName) Then
           lbe(8) = strTempName
        Else
           lbe(8) = ""
           Text(8).SetFocus
           TextInverse Text(8)
           AllTextBeforeSaveCheck = True
           Exit Function
        End If
     End If
     strTempName = ""
     If Text(10) <> "" Then
        Text(10) = UCase(Text(10))
        'edit by nickc 2007/02/07 不用 dll 了
        'If objPublicData.GetStaff(Text(10), strTempName) Then
        If ClsPDGetStaff(Text(10), strTempName) Then
           lbe(10) = strTempName
        Else
           lbe(10) = ""
           Text(10).SetFocus
           TextInverse Text(10)
           AllTextBeforeSaveCheck = True
           Exit Function
        End If
     End If
     strTempName = ""

    If Text(11) <> "" Then
       Text(11) = UCase(Text(11))
    End If

     If Text(12) <> "" Then
        If CheckIsTaiwanDate(Text(12)) Then
            If Text(13) <> "" Then
               If Val(Text(13)) - Val(Text(12)) < 0 Then
                  DataErrorMessage 13
                  Text(12).SetFocus
                  AllTextBeforeSaveCheck = True
                  Exit Function
               End If
            End If
        Else
            Text(12).SetFocus
            AllTextBeforeSaveCheck = True
            Exit Function
        End If
     End If
 
     If Text(13) <> "" Then
        If CheckIsTaiwanDate(Text(13)) Then
            If Text(12) <> "" Then
               If Val(Text(13)) - Val(Text(12)) < 0 Then
                  DataErrorMessage 12
                  Text(13).SetFocus
                  AllTextBeforeSaveCheck = True
                  Exit Function
               End If
            End If
        Else
           Text(13).SetFocus
           AllTextBeforeSaveCheck = True
           Exit Function
        End If
      End If
     
    Text(14) = UCase(Text(14))
    If Text(14) <> "" Then
       If Text(14) <> "N" Then
         DataErrorMessage 1, "是否算案件數"
         AllTextBeforeSaveCheck = True
         Text(14).SetFocus
         Exit Function
       End If
    End If

    If Text(15) <> "" Then
       Text(15) = UCase(Text(15))
       If Text(15) = lbePaperNum Then
          MsgBox "且不可為本身之收文號", vbCritical
          AllTextBeforeSaveCheck = True
          Text(15).SetFocus
          Exit Function
       End If
    End If

    If Text(16) <> "" Then
          Text(16) = UCase(Text(16))
        If Text(16) = "Y" Then
           i = MsgBox("確定取消閉卷?", vbYesNo, "詢問")
           If i = 7 Then Text(16) = ""
        Else
           Text(16).SetFocus
           DataErrorMessage 1, "是否閉卷":
           AllTextBeforeSaveCheck = True
           Exit Function
        End If
    End If
     
      If Text(17) <> "" Then
         Text(17) = UCase(Text(17))
         If Text(17) <> "Y" Then
            DataErrorMessage 1, "是否取締案"
            AllTextBeforeSaveCheck = True
            Text(17).SetFocus
            Exit Function
         End If
      End If
   If txtCP02 <> "" And txtCP03 = "" Then txtCP03 = "0"
   If txtCP02 <> "" And txtCP04 = "" Then txtCP04 = "00"
   
   'Add By Sindy 2010/12/3
   If txtCP02 <> "" Then
      If txtCP01 = lc01 And txtCP02 = lc02 And txtCP03 = lc03 And txtCP04 = lc04 Then
         MsgBox "轉本所案號不可與原本所案號相同 !", vbCritical
         AllTextBeforeSaveCheck = True
         txtCP02.SetFocus
         Exit Function
      End If
   End If
   '2010/12/3 End
   
   AllTextBeforeSaveCheck = False
End Function

Private Function CheckCaseNum() As String
Dim strKind As String, i As Integer
   
   For i = 1 To 4
      If Mid(lbeNumber, i, 1) = "-" Then
         CheckCaseNum = Left(lbeNumber, i - 1)
         Exit For
      End If
   Next
End Function

Private Sub ChgType(i As Integer)
Dim strTempName As String, blnIsEmpty As Boolean
   
   Select Case i
      Case 1
         If Text(i) <> "" Then
            'edit by nickc 2007/02/07 不用 dll 了
            'If objPublicData.GetCustomer(Text(i), strTempName) Then Lbe(i) = strTempName Else Lbe(i) = ""
            If ClsPDGetCustomer(Text(i), strTempName) Then lbe(i) = strTempName Else lbe(i) = ""
         Else
             If m_LC22 = "" Then 'Added by Lydia 2023/03/03 若案件有FC代理人時，當事人若空白僅提醒即可，不必限制不能分案
                MsgBox "當事人不可空白", vbCritical
             End If 'Added by Lydia 2023/03/03
         End If
      Case 7, 8, 10
         If Text(i) <> "" Then
             'edit by nickc 2007/02/07 不用 dll 了
             'If objPublicData.GetStaff(Text(i), strTempName) Then Lbe(i) = strTempName Else Lbe(i) = ""
             If ClsPDGetStaff(Text(i), strTempName) Then
               lbe(i) = strTempName
             Else
               lbe(i) = ""
             End If
         End If
      Case 9
         If Text(i) = "" Then
            MsgBox "案件性質不可空白", vbCritical
            lbe(i) = ""
         Else
            'edit by nickc 2007/02/07 不用 dll 了
            'If objPublicData.GetCaseProperty(CheckCaseNum, Text(i), strTempName, False) Then Lbe(i) = strTempName Else Lbe(i) = ""
            If ClsPDGetCaseProperty(CheckCaseNum, Text(i), strTempName, False) Then lbe(i) = strTempName Else lbe(i) = ""
         End If
   End Select
End Sub

Private Sub txtcp03_Validate(Cancel As Boolean)
   If txtCP02 <> "" And txtCP03 = "" Then txtCP03 = "0"
End Sub

Private Sub txtcp04_Validate(Cancel As Boolean)
   If txtCP02 <> "" And txtCP04 = "" Then txtCP04 = "00"
   If ChangText Then TextInverse txtCP02
   'Add By Cheng 2002/08/23
   If Cancel = False Then
      If Me.txtCP01.Text <> "" And Me.txtCP02.Text <> "" Then
         MsgBox "不可同時做" & """轉本所案號""" & "及" & """修改其他欄位""" & "，否則資料可能有不一致的情形產生!!!", vbExclamation + vbOKOnly
      End If
   End If
   
   If Cancel Then TextInverse txtCP04
End Sub

Private Sub ClearForm()
   
  'Modified by Lydia 2023/03/14
  'For i = 0 To 19
  '    Text(i).Text = ""
  'Next i
  For Each oObj In Text
     oObj.Text = ""
     oObj.Tag = ""
  Next
  'end 2023/03/14
  
  MSHFlexGrid1.Clear
  MSHFlexGrid1.Rows = 2
  txtCP01.Text = ""
  txtCP02.Text = ""
  txtCP03.Text = ""
  txtCP04.Text = ""
  lbePaperNum.Caption = ""
  lbeNumber.Caption = ""
  lbeCloseDate.Caption = ""
  lbeCost.Caption = ""
  lbePointNum.Caption = ""
  lbeMoney.Caption = ""
  'Added by Lydia 2021/09/14 改成Form2.0 ;
  For Each oObj In lbe
     oObj.Caption = ""
  Next
  
  Me.Tag = "" 'Added by Lydia 2022/12/08
  
  m_LC22 = "" 'Added by Lydia 2023/03/03
  'Added by Lydia 2023/03/14
  For Each oObj In Check1
     oObj.Value = 0
  Next
  For Each oObj In Check2
     oObj.Value = 0
  Next
  'end 2023/03/14
  
  'Added by Lydia 2024/09/30
  bChkPaid = False
  bolActCaseLawer = False
  'end 2024/09/30
End Sub

'Add By Cheng 2002/05/24
Private Function TxtValidate() As Boolean
Dim objTxt As Object
Dim ii As Integer
Dim Cancel As Boolean

   TxtValidate = False
   For Each objTxt In Me.Text
      If objTxt.Enabled = True Then
         Cancel = False
         Text_Validate objTxt.Index, Cancel
         If Cancel = True Then
            Exit Function
         End If
      End If
   Next
   
   If Me.txtCP01.Enabled = True Then
      Cancel = False
      txtcp01_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.txtCP02.Enabled = True Then
      Cancel = False
      txtcp02_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.txtCP03.Enabled = True Then
      Cancel = False
      txtcp03_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.txtCP04.Enabled = True Then
      Cancel = False
      txtcp04_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   'Add By Sindy 2011/6/20
   If (InStr(Text(28).Text, "專利") > 0 Or InStr(Text(28).Text, "商標") > 0 Or _
      InStr(Text(28).Text, "著作權") > 0 Or InStr(Text(28).Text, "智財權") > 0) And Text(5) <> "Y" Then
      MsgBox "案件屬性欄位出現專利,商標,著作權或智財權字樣時,智財權案須為Y!!!", vbExclamation + vbOKOnly
      Text(5).SetFocus
      Exit Function
   End If
   
    'Added by Lydia 2016/01/27 L,CFL案必須輸入案件屬性(分配目標點數用)
    If lc01 = "L" And Text(28).Text = "" Then
       MsgBox "法務案請勾選案件屬性!", vbExclamation + vbOKOnly
       Exit Function
    End If
    'end 2016/01/27
    
    'Added by Lydia 2021/09/14 檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
    If PUB_ChkUniText(Me, , True, "TextBox") = False Then
         Exit Function
    End If
    
   'Added by Morgan 2022/11/9 補收款必須輸入相關總收文號(案源TT扣出庭費及收款抓出庭律師要用)
   If lc01 = "L" And Text(9) = "78" And Text(15) = "" Then
      MsgBox lbe(9) & "必須輸入相關總收文號！", vbExclamation
      Text(15).SetFocus
      Exit Function
   End If
   'end 2022/11/9
    
   'Added by Lydia 2022/12/08 修改承辦人檢查; 若已有CaseLawer資料,但是修改承辦人後沒有再進入維護
   'Mark by Lydia 2024/09/30 (113/11/01上線)與出庭費領取的檢查合併
   'If m_CL02 <> "" And InStr(strPublicTemp, Text(7)) = 0 Then
   '   'Modified by Lydia 2023/07/06 改成提醒
   '   'MsgBox "請進入出庭律師資料輸入作業，檢查出庭律師是否正確！", vbExclamation
   '   If MsgBox("承辦人不在出庭律師內，是否繼續存檔？", vbExclamation + vbYesNo + vbDefaultButton2, "出庭律師檢查") = vbNo Then
   '      Command4.SetFocus
   '      Exit Function
   '   End If 'Added by Lydia 2023/07/06
   'End If
   ''end 2022/12/08
   'end 2024/09/30
   
   'Added by Lydia 2024/06/13 檢查更新代理人／申請人狀態排除「不得代理」
   strExc(1) = ChangeCustomerL(Text(1))
   strExc(2) = ChangeCustomerL(m_LC11)
   If strExc(1) <> "" And strExc(1) <> strExc(2) Then
      If GetCustomerAndState(strExc(1), strExc(3), , , , lc01, strExc(8), False, Me.Name, lc02, lc03, lc04) = False Then
         Me.SSTab1.Tab = 0
         Text(1).SetFocus
         Text_GotFocus 1
         Exit Function
      End If
   End If
   'end 2024/06/13
   
   TxtValidate = True
End Function

'Added by Lydia 2020/05/20 法律所案源收文：案源卷宗區
Private Sub cmdok1_Click()
    If lblLOS01 <> "" Then
       If PUB_CheckFormExist("frm100101_L") Then
           MsgBox "請先關閉共同查詢〔卷宗區〕畫面！"
           Exit Sub
       End If
       With frm100101_L
            .m_strKey = lblLOS01
            .SetParent Me
            If .QueryData = True Then
               .Show
               Me.Hide
            End If
       End With
    End If
End Sub

'Added by Lydia 2020/05/20 法律所案源收文：讀取法務案源檔
Private Sub ReadLOS()
Dim stSQL As String, intQ As Integer
Dim RsQ As ADODB.Recordset
   
   m_LOS01 = "": m_LOS01cp01 = "": m_LOS01cp02 = "": m_LOS01cp03 = "": m_LOS01cp04 = ""
   m_LOS02 = ""
   m_LOS04 = "": m_LOS04_1 = ""
   m_LOS05 = "" 'Added by Lydia 2021/06/18
   m_LOS06 = ""
   m_LOS10 = ""
   m_LOS15 = ""
   m_CRL84 = "" 'Added by Lydia 2020/10/07
   
   'Modified by Lydia 2020/10/07 抓接洽單編號 + NVL(X.LOS17,X.LOS18) CRLNO
   'Modified by Lydia 2021/06/18 +LOS05
   'Modified by Lydia 2023/08/14 改用(案件進度)案源單號
   'stSQL = "select nvl(X.LOS07,0) ord1,X.LOS01,X.LOS02,X.LOS04,X.LOS05,X.LOS06,X.LOS10,X.LOS15,CP01,CP02,CP03,CP04,NVL(X.LOS17,X.LOS18) CRLNO " & _
                "from LawOfficeSource X,CaseProgress where X.LOS06='" & lbePaperNum & "' and X.LOS01=CP09(+) order by ord1, X.LOS01 "
   'Modified by Lydia 2024/09/30 (113/11/01上線)案源FC代理人
   'stSQL = "select nvl(X.LOS07,0) ord1,X.LOS01,X.LOS02,X.LOS04,X.LOS05,X.LOS06,X.LOS10,X.LOS15,CP01,CP02,CP03,CP04,NVL(X.LOS17,X.LOS18) CRLNO " & _
                "from LawOfficeSource X,CaseProgress where X.LOS15='" & m_CP162 & "' and X.LOS01=CP09(+) order by ord1, X.LOS01 "
   stSQL = "select nvl(X.LOS07,0) ord1,X.LOS01,X.LOS02,X.LOS04,X.LOS05,X.LOS06,X.LOS10,X.LOS15,CP01,CP02,CP03,CP04,NVL(X.LOS17,X.LOS18) CRLNO,NVL(PA75,TM44) FAGENT " & _
           "from LawOfficeSource X,CaseProgress, patent, trademark where X.LOS15='" & m_CP162 & "' and X.LOS01=CP09(+) " & _
           "and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) and cp01=tm01(+) and cp02=tm02(+) and cp03=tm03(+) and cp04=tm04(+) " & _
           "order by ord1, X.LOS01 "
   intQ = 1
   Set RsQ = ClsLawReadRstMsg(intQ, stSQL)
   If intQ = 1 Then
      '案源總收文號
      m_LOS01 = "" & RsQ.Fields("los01")
      '案源總收文號之本所案號
      m_LOS01cp01 = "" & RsQ.Fields("cp01")
      m_LOS01cp02 = "" & RsQ.Fields("cp02")
      m_LOS01cp03 = "" & RsQ.Fields("cp03")
      m_LOS01cp04 = "" & RsQ.Fields("cp04")
      '(原)案源案件類型
      m_LOS02 = "" & RsQ.Fields("LOS02")
      t_LOS02 = m_LOS02
      m_LOS04 = "" & RsQ.Fields("LOS04") 'Added by Lydia 2020/10/15
      m_LOS01fa = "" & RsQ.Fields("FAGENT") 'Added by Lydia 2024/09/30
      
      '介紹人, 介紹人(第一位)
      If m_LOS04 <> "" And InStr(m_LOS04, ",") > 0 Then
          m_LOS04_1 = Mid(m_LOS04, 1, InStr(m_LOS04, ",") - 1)
      Else
          m_LOS04_1 = m_LOS04
      End If
      m_LOS05 = "" & RsQ.Fields("los05") 'Added by Lydia 2021/06/18
      m_LOS06 = "" & RsQ.Fields("los06")
      m_LOS10 = "" & RsQ.Fields("los10")
      m_LOS15 = "" & RsQ.Fields("los15")
      
      'Added by Lydia 2020/10/07 接洽記錄單-法務案件屬性
      If "" & RsQ.Fields("CRLNO") <> "" Then
           stSQL = "select crl84 from Consultrecordlist where crl01=" & CNULL(RsQ.Fields("CRLNO"))
           intQ = 1
           Set RsQ = ClsLawReadRstMsg(intQ, stSQL)
           If intQ = 1 Then
                m_CRL84 = "" & RsQ.Fields("crl84")
           End If
      End If
      'end 2020/10/07
   End If
   Set RsQ = Nothing
End Sub

'Added by Lydia 2023/03/14
Private Sub Check2_Click(Index As Integer)
   If Check2(Index).Value = 1 Then
      If InStr(Text(28).Text, Trim(Check2(Index).Caption)) = 0 Then
         If Text(28).Text = "" Then
            Text(28).Text = Trim(Check2(Index).Caption)
         Else
            Text(28).Text = Text(28).Text & "," & Trim(Check2(Index).Caption)
         End If
      End If
   Else
      '案件屬性=xx,xx,xx
      If Left(Text(28), Len(Trim(Check2(Index).Caption))) = Trim(Check2(Index).Caption) Then
         Text(28).Text = Replace(Text(28).Text, Trim(Check2(Index).Caption) & ",", "")
         Text(28).Text = Replace(Text(28).Text, Trim(Check2(Index).Caption), "")
      Else
         Text(28).Text = Replace(Text(28).Text, "," & Trim(Check2(Index).Caption), "")
      End If
   End If
End Sub

