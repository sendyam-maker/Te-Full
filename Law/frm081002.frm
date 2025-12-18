VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm081002 
   BorderStyle     =   1  '單線固定
   Caption         =   "法務－分案"
   ClientHeight    =   6540
   ClientLeft      =   240
   ClientTop       =   972
   ClientWidth     =   8988
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6540
   ScaleWidth      =   8988
   Begin TabDlg.SSTab SSTab1 
      Height          =   5535
      Left            =   30
      TabIndex        =   63
      Top             =   960
      Width           =   8925
      _ExtentX        =   15748
      _ExtentY        =   9758
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabMaxWidth     =   3528
      TabCaption(0)   =   "第一頁"
      TabPicture(0)   =   "frm081002.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Text(7)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Text(9)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Text(12)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Text(11)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Text(10)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Text(13)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Text(8)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label8"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "lbeCloseDate"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label28"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "lbe(9)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "lbe(10)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label24"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label11"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Label20"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "lbe(8)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Label14"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "lbe(7)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Label3(0)"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Label10"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Text(18)"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "Text(19)"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "Label13"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "Label12"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "Frame1"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "Command4"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "Command2"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).ControlCount=   27
      TabCaption(1)   =   "第二頁"
      TabPicture(1)   =   "frm081002.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txtcp02"
      Tab(1).Control(1)=   "txtcp03"
      Tab(1).Control(2)=   "txtcp04"
      Tab(1).Control(3)=   "txtcp01"
      Tab(1).Control(4)=   "MSHFlexGrid1"
      Tab(1).Control(5)=   "Label22"
      Tab(1).Control(6)=   "Label21(0)"
      Tab(1).Control(7)=   "Label19"
      Tab(1).Control(8)=   "Label4(1)"
      Tab(1).Control(9)=   "Label9"
      Tab(1).Control(10)=   "lbePointNum"
      Tab(1).Control(11)=   "Label15"
      Tab(1).Control(12)=   "Label5(0)"
      Tab(1).Control(13)=   "Label21(1)"
      Tab(1).Control(14)=   "Label27"
      Tab(1).Control(15)=   "Label29"
      Tab(1).Control(16)=   "Label6"
      Tab(1).Control(17)=   "Label16"
      Tab(1).Control(18)=   "lbeMoney"
      Tab(1).Control(19)=   "Text(17)"
      Tab(1).Control(20)=   "Text(16)"
      Tab(1).Control(21)=   "Text(15)"
      Tab(1).Control(22)=   "lbeCost"
      Tab(1).Control(23)=   "Text(14)"
      Tab(1).ControlCount=   24
      Begin VB.TextBox txtcp02 
         Height          =   300
         Left            =   -72945
         MaxLength       =   6
         TabIndex        =   47
         Top             =   720
         Width           =   855
      End
      Begin VB.TextBox txtcp03 
         Height          =   300
         Left            =   -72000
         MaxLength       =   1
         TabIndex        =   48
         Top             =   720
         Width           =   255
      End
      Begin VB.TextBox txtcp04 
         Height          =   300
         Left            =   -71685
         MaxLength       =   2
         TabIndex        =   49
         Top             =   720
         Width           =   495
      End
      Begin VB.TextBox txtcp01 
         Height          =   300
         Left            =   -73545
         MaxLength       =   3
         TabIndex        =   46
         Top             =   720
         Width           =   550
      End
      Begin VB.CommandButton Command2 
         Caption         =   "相對人資料(&B)"
         Height          =   270
         Left            =   5820
         TabIndex        =   33
         Top             =   3900
         Width           =   1335
      End
      Begin VB.CommandButton Command4 
         Caption         =   "出庭律師(&L)"
         Height          =   270
         Left            =   7170
         TabIndex        =   34
         Top             =   3900
         Width           =   1485
      End
      Begin VB.Frame Frame1 
         Height          =   2595
         Left            =   120
         TabIndex        =   64
         Top             =   330
         Width           =   8685
         Begin VB.CheckBox Check1 
            BackColor       =   &H00C0FFFF&
            Caption         =   "一般"
            Height          =   215
            Index           =   4
            Left            =   3870
            TabIndex        =   17
            Top             =   1920
            Width           =   735
         End
         Begin VB.Frame Frame2 
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   0  '沒有框線
            Height          =   700
            Left            =   3870
            TabIndex        =   107
            Top             =   1860
            Width           =   4200
            Begin VB.CheckBox Check2 
               BackColor       =   &H00C0FFFF&
               Caption         =   "承攬糾紛"
               Height          =   180
               Index           =   0
               Left            =   1410
               TabIndex        =   18
               Top             =   60
               Width           =   1245
            End
            Begin VB.CheckBox Check2 
               BackColor       =   &H00C0FFFF&
               Caption         =   "公寓大廈糾紛"
               Height          =   180
               Index           =   1
               Left            =   2760
               TabIndex        =   19
               Top             =   60
               Width           =   1425
            End
            Begin VB.CheckBox Check2 
               BackColor       =   &H00C0FFFF&
               Caption         =   "土地爭議"
               Height          =   180
               Index           =   2
               Left            =   0
               TabIndex        =   20
               Top             =   270
               Width           =   1245
            End
            Begin VB.CheckBox Check2 
               BackColor       =   &H00C0FFFF&
               Caption         =   "車禍糾紛"
               Height          =   180
               Index           =   3
               Left            =   1410
               TabIndex        =   21
               Top             =   270
               Width           =   1245
            End
            Begin VB.CheckBox Check2 
               BackColor       =   &H00C0FFFF&
               Caption         =   "消費糾紛"
               Height          =   180
               Index           =   4
               Left            =   2760
               TabIndex        =   22
               Top             =   270
               Width           =   1425
            End
            Begin VB.CheckBox Check2 
               BackColor       =   &H00C0FFFF&
               Caption         =   "醫療糾紛"
               Height          =   180
               Index           =   5
               Left            =   0
               TabIndex        =   23
               Top             =   480
               Width           =   1245
            End
            Begin VB.CheckBox Check2 
               BackColor       =   &H00C0FFFF&
               Caption         =   "買賣糾紛"
               Height          =   180
               Index           =   6
               Left            =   1410
               TabIndex        =   24
               Top             =   480
               Width           =   1245
            End
            Begin VB.CheckBox Check2 
               BackColor       =   &H00C0FFFF&
               Caption         =   "勞資糾紛"
               Height          =   180
               Index           =   7
               Left            =   2760
               TabIndex        =   25
               Top             =   480
               Width           =   1425
            End
         End
         Begin VB.CheckBox Check1 
            Caption         =   "專利"
            Height          =   215
            Index           =   0
            Left            =   3855
            TabIndex        =   11
            Top             =   1410
            Width           =   735
         End
         Begin VB.CheckBox Check1 
            Caption         =   "商標"
            Height          =   215
            Index           =   1
            Left            =   4575
            TabIndex        =   12
            Top             =   1410
            Width           =   735
         End
         Begin VB.CheckBox Check1 
            Caption         =   "著作權"
            Height          =   215
            Index           =   2
            Left            =   5295
            TabIndex        =   13
            Top             =   1410
            Width           =   855
         End
         Begin VB.CheckBox Check1 
            Caption         =   "其他智財權"
            Height          =   215
            Index           =   3
            Left            =   6135
            TabIndex        =   14
            Top             =   1410
            Width           =   1215
         End
         Begin VB.ComboBox Combo1 
            Height          =   300
            ItemData        =   "frm081002.frx":0038
            Left            =   1020
            List            =   "frm081002.frx":0045
            Style           =   2  '單純下拉式
            TabIndex        =   2
            Top             =   150
            Width           =   1095
         End
         Begin VB.CheckBox Check1 
            Caption         =   "營業秘密法"
            Height          =   215
            Index           =   5
            Left            =   3855
            TabIndex        =   15
            Top             =   1665
            Width           =   1365
         End
         Begin VB.CheckBox Check1 
            Caption         =   "公平交易法"
            Height          =   215
            Index           =   6
            Left            =   5295
            TabIndex        =   16
            Top             =   1665
            Width           =   1365
         End
         Begin VB.Label lblMemo 
            Caption         =   "可以直接輸入，屬性之間用逗號,區隔。"
            Height          =   765
            Left            =   90
            TabIndex        =   106
            Top             =   1650
            Width           =   945
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "當  事  人："
            Height          =   180
            Left            =   90
            TabIndex        =   75
            Top             =   510
            Width           =   900
         End
         Begin MSForms.Label lbe 
            Height          =   285
            Index           =   1
            Left            =   2160
            TabIndex        =   74
            Top             =   458
            Width           =   1845
            VariousPropertyBits=   27
            Caption         =   "lbe(1)"
            Size            =   "3254;503"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "分所案號："
            Height          =   180
            Index           =   0
            Left            =   4500
            TabIndex        =   73
            Top             =   810
            Width           =   900
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "(Y:是)"
            Height          =   180
            Index           =   1
            Left            =   2460
            TabIndex        =   72
            Top             =   1125
            Width           =   465
         End
         Begin VB.Label lblName 
            AutoSize        =   -1  'True
            Caption         =   "案件名稱："
            Height          =   180
            Index           =   3
            Left            =   90
            TabIndex        =   71
            Top             =   210
            Width           =   900
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            Caption         =   "代  理  人："
            Height          =   180
            Left            =   4455
            TabIndex        =   70
            Top             =   510
            Width           =   900
         End
         Begin MSForms.Label lbe 
            Height          =   285
            Index           =   2
            Left            =   6585
            TabIndex        =   69
            Top             =   458
            Width           =   2055
            VariousPropertyBits=   27
            Caption         =   "lbe(2)"
            Size            =   "3625;503"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin VB.Label lblName 
            AutoSize        =   -1  'True
            Caption         =   "彼所案號："
            Height          =   180
            Index           =   0
            Left            =   90
            TabIndex        =   68
            Top             =   810
            Width           =   900
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            Caption         =   "署  名  人："
            Height          =   180
            Index           =   2
            Left            =   4500
            TabIndex        =   67
            Top             =   1125
            Width           =   900
         End
         Begin VB.Label Label26 
            BackColor       =   &H00C0FFC0&
            Caption         =   "案件屬性："
            Height          =   210
            Left            =   90
            TabIndex        =   66
            Top             =   1425
            Width           =   945
         End
         Begin MSForms.TextBox Text 
            Height          =   300
            Index           =   6
            Left            =   5430
            TabIndex        =   7
            Top             =   750
            Width           =   2055
            VariousPropertyBits=   671105051
            MaxLength       =   50
            Size            =   "3625;529"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox Text 
            Height          =   300
            Index           =   2
            Left            =   2220
            TabIndex        =   3
            Top             =   150
            Width           =   6015
            VariousPropertyBits=   671105051
            MaxLength       =   160
            Size            =   "10610;529"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox Text 
            Height          =   300
            Index           =   21
            Left            =   5430
            TabIndex        =   5
            Top             =   450
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
            Height          =   300
            Index           =   4
            Left            =   5430
            TabIndex        =   9
            Top             =   1065
            Width           =   2055
            VariousPropertyBits=   671105051
            MaxLength       =   10
            Size            =   "3625;529"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            Caption         =   "是否為智慧財產權案："
            Height          =   180
            Index           =   0
            Left            =   90
            TabIndex        =   65
            Top             =   1125
            Width           =   1800
         End
         Begin MSForms.TextBox Text 
            Height          =   300
            Index           =   5
            Left            =   1980
            TabIndex        =   8
            Top             =   1065
            Width           =   375
            VariousPropertyBits=   671105051
            MaxLength       =   1
            Size            =   "661;529"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox Text 
            Height          =   1065
            Index           =   28
            Left            =   1050
            TabIndex        =   10
            Top             =   1365
            Width           =   2775
            VariousPropertyBits=   -1467989989
            MaxLength       =   200
            Size            =   "4895;1879"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox Text 
            Height          =   300
            Index           =   3
            Left            =   1020
            TabIndex        =   6
            Top             =   750
            Width           =   2655
            VariousPropertyBits=   671105051
            MaxLength       =   50
            Size            =   "4683;529"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox Text 
            Height          =   300
            Index           =   1
            Left            =   1020
            TabIndex        =   4
            Top             =   450
            Width           =   1095
            VariousPropertyBits=   671105051
            MaxLength       =   9
            Size            =   "1931;529"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   1365
         Left            =   -74325
         TabIndex        =   52
         Top             =   1440
         Width           =   8085
         _ExtentX        =   14266
         _ExtentY        =   2413
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
      Begin VB.Label Label12 
         Caption         =   "進度備註："
         Height          =   375
         Left            =   150
         TabIndex        =   105
         Top             =   4230
         Width           =   1035
      End
      Begin VB.Label Label13 
         Caption         =   "案件備註："
         Height          =   255
         Left            =   150
         TabIndex        =   104
         Top             =   4860
         Width           =   915
      End
      Begin MSForms.TextBox Text 
         Height          =   585
         Index           =   19
         Left            =   1320
         TabIndex        =   36
         Top             =   4860
         Width           =   7410
         VariousPropertyBits=   -1466941413
         MaxLength       =   2000
         ScrollBars      =   2
         Size            =   "13070;1032"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text 
         Height          =   585
         Index           =   18
         Left            =   1320
         TabIndex        =   35
         Top             =   4230
         Width           =   7410
         VariousPropertyBits=   -1466941413
         MaxLength       =   2000
         ScrollBars      =   2
         Size            =   "13070;1032"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "(N:不收)"
         Height          =   180
         Left            =   -68100
         TabIndex        =   103
         Top             =   1117
         Width           =   645
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "是否向客戶收款："
         Height          =   180
         Index           =   0
         Left            =   -69990
         TabIndex        =   102
         Top             =   1117
         Width           =   1440
      End
      Begin VB.Label Label19 
         Caption         =   "本案期限"
         Height          =   495
         Left            =   -74805
         TabIndex        =   101
         Top             =   1485
         Width           =   375
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "費用："
         Height          =   180
         Index           =   1
         Left            =   -74850
         TabIndex        =   100
         Top             =   1117
         Width           =   540
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "點數："
         Height          =   180
         Left            =   -73260
         TabIndex        =   99
         Top             =   1117
         Width           =   540
      End
      Begin VB.Label lbePointNum 
         Caption         =   "lbePointNum"
         Height          =   255
         Left            =   -72720
         TabIndex        =   98
         Top             =   1117
         Width           =   975
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "是否算案件數："
         Height          =   180
         Left            =   -74850
         TabIndex        =   97
         Top             =   480
         Width           =   1260
      End
      Begin VB.Label Label5 
         Caption         =   "(N：不算)"
         Height          =   180
         Index           =   0
         Left            =   -72960
         TabIndex        =   96
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label21 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         Caption         =   "是否取消閉卷："
         Height          =   180
         Index           =   1
         Left            =   -70545
         TabIndex        =   95
         Top             =   780
         Width           =   1260
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         Caption         =   "相關總收文號："
         Height          =   180
         Left            =   -70545
         TabIndex        =   94
         Top             =   480
         Width           =   1260
      End
      Begin VB.Label Label29 
         Caption         =   "(Y:取消閉卷)"
         Height          =   180
         Left            =   -68670
         TabIndex        =   93
         Top             =   780
         Width           =   1095
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "轉本所案號："
         Height          =   180
         Left            =   -74850
         TabIndex        =   92
         Top             =   780
         Width           =   1080
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "後金："
         Height          =   180
         Left            =   -71670
         TabIndex        =   91
         Top             =   1117
         Width           =   540
      End
      Begin VB.Label lbeMoney 
         Caption         =   "lbeMoney"
         Height          =   255
         Left            =   -71100
         TabIndex        =   90
         Top             =   1117
         Width           =   855
      End
      Begin MSForms.TextBox Text 
         Height          =   300
         Index           =   17
         Left            =   -68490
         TabIndex        =   51
         Top             =   1080
         Width           =   375
         VariousPropertyBits=   671105051
         MaxLength       =   1
         Size            =   "661;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text 
         Height          =   300
         Index           =   16
         Left            =   -69255
         TabIndex        =   50
         Top             =   720
         Width           =   495
         VariousPropertyBits=   671105051
         MaxLength       =   1
         Size            =   "873;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text 
         Height          =   300
         Index           =   15
         Left            =   -69255
         TabIndex        =   45
         Top             =   420
         Width           =   2175
         VariousPropertyBits=   671105051
         MaxLength       =   9
         Size            =   "3836;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label lbeCost 
         Caption         =   "lbeCost"
         Height          =   255
         Left            =   -74250
         TabIndex        =   89
         Top             =   1117
         Width           =   975
      End
      Begin MSForms.TextBox Text 
         Height          =   300
         Index           =   14
         Left            =   -73545
         TabIndex        =   44
         Top             =   420
         Width           =   375
         VariousPropertyBits=   671105051
         MaxLength       =   1
         Size            =   "661;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "承辦人："
         Height          =   180
         Left            =   150
         TabIndex        =   88
         Top             =   3030
         Width           =   720
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "案件性質："
         Height          =   180
         Index           =   0
         Left            =   150
         TabIndex        =   87
         Top             =   3345
         Width           =   900
      End
      Begin MSForms.Label lbe 
         Height          =   285
         Index           =   7
         Left            =   2370
         TabIndex        =   86
         Top             =   2985
         Width           =   1455
         VariousPropertyBits=   27
         Caption         =   "lbe(7)"
         Size            =   "2566;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "協辦人員："
         Height          =   180
         Left            =   4620
         TabIndex        =   85
         Top             =   3030
         Width           =   900
      End
      Begin MSForms.Label lbe 
         Height          =   285
         Index           =   8
         Left            =   6570
         TabIndex        =   84
         Top             =   2985
         Width           =   1935
         VariousPropertyBits=   27
         Caption         =   "lbe(8)"
         Size            =   "3413;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "本所期限："
         Height          =   180
         Left            =   150
         TabIndex        =   83
         Top             =   3945
         Width           =   900
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "法定期限："
         Height          =   180
         Left            =   2760
         TabIndex        =   82
         Top             =   3945
         Width           =   900
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         Caption         =   "智權人員："
         Height          =   180
         Left            =   4620
         TabIndex        =   81
         Top             =   3345
         Width           =   900
      End
      Begin MSForms.Label lbe 
         Height          =   285
         Index           =   10
         Left            =   6570
         TabIndex        =   80
         Top             =   3300
         Width           =   1905
         VariousPropertyBits=   27
         Caption         =   "lbe(10)"
         Size            =   "3360;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbe 
         Height          =   285
         Index           =   9
         Left            =   2010
         TabIndex        =   79
         Top             =   3300
         Width           =   1935
         VariousPropertyBits=   27
         Caption         =   "lbe(9)"
         Size            =   "3413;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         Caption         =   "取消收文日期："
         Height          =   180
         Left            =   4260
         TabIndex        =   78
         Top             =   3645
         Width           =   1260
      End
      Begin VB.Label lbeCloseDate 
         Caption         =   "lbeCloseDate"
         Height          =   255
         Left            =   5550
         TabIndex        =   77
         Top             =   3615
         Width           =   1575
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "當事人稱謂："
         Height          =   180
         Left            =   150
         TabIndex        =   76
         Top             =   3645
         Width           =   1080
      End
      Begin MSForms.TextBox Text 
         Height          =   300
         Index           =   8
         Left            =   5550
         TabIndex        =   28
         Top             =   2970
         Width           =   975
         VariousPropertyBits=   671105051
         MaxLength       =   6
         Size            =   "1720;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text 
         Height          =   300
         Index           =   13
         Left            =   3720
         TabIndex        =   32
         Top             =   3885
         Width           =   1095
         VariousPropertyBits=   671105051
         MaxLength       =   7
         Size            =   "1931;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text 
         Height          =   300
         Index           =   10
         Left            =   5550
         TabIndex        =   26
         Top             =   3285
         Width           =   975
         VariousPropertyBits=   671105051
         MaxLength       =   6
         Size            =   "1720;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text 
         Height          =   300
         Index           =   11
         Left            =   1320
         TabIndex        =   30
         Top             =   3585
         Width           =   2415
         VariousPropertyBits=   671105051
         MaxLength       =   10
         Size            =   "4260;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text 
         Height          =   300
         Index           =   12
         Left            =   1320
         TabIndex        =   31
         Top             =   3885
         Width           =   1095
         VariousPropertyBits=   671105051
         MaxLength       =   7
         Size            =   "1931;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text 
         Height          =   300
         Index           =   9
         Left            =   1320
         TabIndex        =   29
         Top             =   3285
         Width           =   615
         VariousPropertyBits=   671105051
         MaxLength       =   4
         Size            =   "1085;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text 
         Height          =   300
         Index           =   7
         Left            =   1320
         TabIndex        =   27
         Top             =   2970
         Width           =   975
         VariousPropertyBits=   671105051
         MaxLength       =   6
         Size            =   "1720;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
   End
   Begin VB.CommandButton CmdOK1 
      Caption         =   "案源卷宗區(&C)"
      Height          =   300
      Left            =   1500
      TabIndex        =   40
      Top             =   0
      Width           =   1350
   End
   Begin VB.CommandButton ComBack 
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   300
      Left            =   8004
      TabIndex        =   39
      Top             =   0
      Width           =   800
   End
   Begin VB.CommandButton cmdPrePic 
      Caption         =   "回前畫面(&U)"
      CausesValidation=   0   'False
      Height          =   300
      Left            =   6876
      TabIndex        =   38
      Top             =   0
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   300
      Left            =   6048
      TabIndex        =   37
      Top             =   0
      Width           =   800
   End
   Begin VB.CommandButton Command3 
      Caption         =   "案件進度(&C)"
      Height          =   300
      Left            =   4920
      TabIndex        =   43
      Top             =   0
      Width           =   1100
   End
   Begin VB.CommandButton Command1 
      Caption         =   "相關卷號(&F)"
      Height          =   300
      Left            =   3792
      TabIndex        =   42
      Top             =   0
      Width           =   1100
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "下一筆(&N)"
      Height          =   300
      Left            =   2860
      TabIndex        =   41
      Top             =   0
      Width           =   900
   End
   Begin VB.Label lbeNumber 
      Height          =   255
      Left            =   1020
      TabIndex        =   62
      Top             =   653
      Width           =   1185
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
      Height          =   255
      Left            =   1020
      TabIndex        =   61
      Top             =   353
      Width           =   1065
   End
   Begin MSForms.TextBox Text 
      Height          =   300
      Index           =   20
      Left            =   5580
      TabIndex        =   1
      Top             =   630
      Width           =   1095
      VariousPropertyBits=   671105051
      MaxLength       =   4
      Size            =   "1931;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text 
      Height          =   300
      Index           =   0
      Left            =   5580
      TabIndex        =   0
      Top             =   330
      Width           =   1095
      VariousPropertyBits=   671105051
      MaxLength       =   7
      Size            =   "1931;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label LBL01 
      Caption         =   "案源總收文號： "
      Height          =   255
      Left            =   2160
      TabIndex        =   60
      Top             =   353
      Width           =   1275
   End
   Begin VB.Label lblLOS01 
      Caption         =   "lblLOS01"
      Height          =   255
      Left            =   3450
      TabIndex        =   59
      Top             =   360
      Width           =   915
   End
   Begin MSForms.Label lbe 
      Height          =   285
      Index           =   20
      Left            =   6740
      TabIndex        =   57
      Top             =   645
      Width           =   1605
      VariousPropertyBits=   27
      Caption         =   "lbe(20)"
      Size            =   "2831;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label25 
      AutoSize        =   -1  'True
      Caption         =   "相關國家："
      Height          =   180
      Index           =   1
      Left            =   4680
      TabIndex        =   56
      Top             =   690
      Width           =   900
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "收  文  日："
      Height          =   180
      Index           =   0
      Left            =   4650
      TabIndex        =   55
      Top             =   390
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "收  文  號： "
      Height          =   180
      Index           =   1
      Left            =   60
      TabIndex        =   54
      Top             =   390
      Width           =   945
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "本所案號："
      Height          =   180
      Index           =   0
      Left            =   60
      TabIndex        =   53
      Top             =   690
      Width           =   900
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
      Height          =   255
      Left            =   6740
      TabIndex        =   58
      Top             =   353
      Width           =   1605
   End
End
Attribute VB_Name = "frm081002"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/09/22 改成Form2.0 ; Text(index)、lbe(index)、MSHFlexGrid1改字型=新細明體-ExtB
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/26 員工編號欄已修改
'Memo By Sindy 2010/8/3 日期欄已修改
Option Explicit

Dim strCP09() As String, t As Integer
Dim strDate As String, LcTmp As String
Dim intLastRow As Integer, intCols As Integer, strOldLc As String
Dim blnIsNew As Boolean, blnIsSave As Boolean, blnOKtoShow As Boolean
Dim lc01 As String, lc02 As String, lc03 As String, lc04 As String
Dim stName(1 To 3) As String
Dim m_ODate As String '本所期限
Dim m_LDate As String '法定期限
Dim m_CurrSel As Integer
Dim m_CPCount As Integer
Dim m_Cpindex As Integer

'910703 Sieg 701
Dim m_CP60 As String, m_LC11 As String
Dim m_LC22 As String 'Added by Lydia 2024/06/13 FC代理人
'Add By Cheng 2002/08/22
'Dim m_strCust1 As String 'Mark by Lydia 2024/06/13
'add by nickc 2005/03/17 加乘註記
Dim m_CP98 As String
Dim m_CP101 As String
Dim m_CP104 As String
Dim m_CP65 As String 'Add By Sindy 2010/8/6
Dim strTemp As Variant 'Add By Sindy 2011/6/8
Dim m_CP27 As String 'Add By Sindy 2012/6/1
Dim m_CP31 As String 'Add by Amy 2018/10/15
'Added by Lydia 2020/05/20 法律所案源收文
Dim m_LOS15 As String '案源單號
Dim m_LOS01 As String '案源總收文號
'Dim m_LOS01cp60 As String 'Added by Lydia 2022/12/08 PT案之請款單號 'Mark by Lydia 2024/09/30 (113/11/01上線)
Dim m_LOS01cp01 As String, m_LOS01cp02 As String, m_LOS01cp03 As String, m_LOS01cp04 As String '案源總收文號之本所案號
Dim m_LOS02 As String '案源案件類型
Dim t_LOS02 As String '需要變更：案源案件類型
Dim m_LOS04 As String, m_LOS04_1 As String  '介紹人、介紹人(第一位)
Dim m_LOS05 As String 'Added by Lydia 2021/06/18 介紹客戶編號
Dim m_LOS06 As String '法律所總收文號1
Dim m_LOS10 As String '收據總收文號
Dim m_CRL84 As String 'Added by Lydia 2020/10/07 接洽記錄單-法務案件屬性
Dim m_CL02 As String 'Added by Lydia 2020/10/15 其他出庭律師
Dim oObj As Control 'Added by Lydia 2022/08/10
Dim m_CP162 As String 'Added by Lydia 2023/08/14 (案件進度)案源單號
'Added by Lydia 2024/09/30 (113/11/01上線)
Dim bChkPaid As Boolean, m_CCP60 As String  '是否已付款, 收款之收據/請款單號
Dim bolActCaseLawer As Boolean '是否進入出庭律師維護
Dim m_LOS01fa As String '案源之FC代理人
'end 2024/09/30 (113/11/01上線)

'Add By Sindy 2011/5/31
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
  ClearForm
  m_Cpindex = m_Cpindex + 1
  If m_Cpindex = m_CPCount - 1 Then
     CmdNext.Enabled = False
  ElseIf m_Cpindex = m_CPCount Then
     Exit Sub
  End If
  GetData (m_Cpindex)

End Sub

Private Sub cmdok_Click()
Dim oSubject As String, oContext As String, strText As String, intR As Integer 'Added by Lydia 2022/12/08

   If AllTextBeforeSaveCheck Then Exit Sub
   'Add By Cheng 2002/05/24
   '重新檢查欄位有效性
   If TxtValidate = False Then Exit Sub
'   'Add By Cheng 2002/08/22
'   If Me.txtcp01.Text <> "" And Me.txtcp02.Text <> "" Then
'      MsgBox "本次只更新轉本所案號的資料，不更新其他欄位的修改!!!", vbCritical
'   End If
   
   '910703 Sieg 701
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
   If bolActCaseLawer = True Then   'Added by Lydia 2024/09/30 (113/11/01上線)
      If PUB_SaveCaseLawer(lbePaperNum, Mid(Me.Tag, InStr(Me.Tag, "|") + 1), , , True) = True Then
         strExc(2) = "Y"
      End If
   End If 'Added by Lydia 2024/09/30 (113/11/01上線)
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
      'end 2024/09/30 (113/11/01上線)
   End If
   'end 2023/08/14
   
   'Add By Cheng 2002/11/18
   If Me.txtCP01.Text <> "" And Me.txtCP02.Text <> "" Then
      MsgBox "本次只更新轉本所案號的資料，不更新其他欄位的修改!!!", vbCritical
   End If
   
   'Added by Lydia 2020/01/21 所有系統別分案，若該收文號存在於下一程序(NP24)，若修改案件性質則彈訊息
   If Me.Text(9).Tag <> Me.Text(9).Text Then
       If Pub_CheckNP24Exists(lbePaperNum.Caption) = True Then
       End If
   End If
   'end 2020/01/21
   
   'Added by Lydia 2020/05/20 法律所案源收文：案件屬性有勾專利或商標或著作權時，若案件性質為1101~1104(民事委任律師)時，若案源非B1時詢問 是否需智慧所配合開庭？」，若選擇要則將案源更新為B1
   'Modified by Lydia 2020/05/29 分案和配合開庭通知整合為一封email
   'Mark by Lydia 2020/06/16 改成frm077005「智財訴訟案需專業部配合通知補收文作業」
   'If strSrvDate(1) >= 法律所案源收文啟用日 And m_LOS01 <> "" And (Check1(0).Value = 1 Or Check1(1).Value = 1 Or Check1(2).Value = 1) And InStr("1101,1102,1103,1104,", Format(Text(9), "0000") & ",") > 0 And m_LOS02 <> "B1" And Text(7).Tag = "" Then
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
   
   If Not SaveData Then DataErrorMessage (3)
   
   'Added by Lydia 2022/12/08 修改承辦人或其他出庭律師時，若該收文號已收款則MAIL通知71005
          '外法B類案源的費用在PT案=>CP60, 已收款判斷也在PT案
   'Mark by Lydia 2024/09/30 (113/11/01上線) 出庭律師或出庭費有異動，發Email通知財務處總帳人員；同時整合原有Email通知：'Added by Lydia 2020/10/15 (法律所案源)若分案時已經收款則同時E-MAIL給系統特殊設定之財務處出納人員及智權人員＋'Add By Sindy 2011/6/8 修改承辦人或其他出庭律師時，若該收文號已收款(CP75>0)則MAIL通知71005
   'If ((m_CL02 <> "" And m_CL02 <> Trim(strPublicTemp)) Or (Text(7).Tag <> "" And Text(7).Tag <> Trim(Text(7)))) And (m_CP60 <> "" Or m_LOS01cp60 <> "") Then
   '   strExc(0) = "": strExc(1) = ""
   '   If Left(m_LOS02, 1) = "B" Then
   '      strExc(1) = m_LOS01cp60
   '   Else
   '      strExc(1) = m_CP60
   '   End If
   '   If Left(strExc(1), 1) = "E" Then
   '       strExc(0) = "select nvl(cp75,0) amt1 from caseprogress where cp09=" & CNULL(lbePaperNum)
   '   ElseIf Left(strExc(1), 1) = "X" Then
   '       strExc(0) = "select nvl(a1k30,0) amt1 from acc1k0 where a1k01=" & CNULL(strExc(1))
   '   End If
   '   If strExc(0) <> "" Then
   '      intI = 1
   '      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   '      If intI = 1 Then
   '         If Val("" & RsTemp.Fields("amt1")) > 0 Then
   '
   '            oSubject = lc01 & "-" & lc02 & "-" & lc03 & "-" & lc04 & "(" & strExc(1) & IIf(Left(m_LOS02, 1) = "B", "專業部案號：" & m_LOS01cp01 & "-" & m_LOS01cp02 & "-" & m_LOS01cp03 & "-" & m_LOS01cp04, "") & ")修改承辦人且已收款，若有必要請自行調整出庭費傳票摘要！"
   '            oContext = oContext & "案件性質：" & lbe(9) & vbCrLf & vbCrLf
   '            '+信件內容
    '           If (Text(7).Tag <> "" And Text(7).Tag <> Trim(Text(7))) Then
   '               oContext = oContext & "原承辦人：" & Text(7).Tag & " " & GetStaffName(Text(7).Tag, True) & "　　改為：" & Trim(Text(7)) & " " & Trim(lbe(7)) & vbCrLf
   '            End If
   '            If (m_CL02 <> "" And m_CL02 <> Trim(strPublicTemp)) Then
   '               strTemp = Split(m_CL02, ",")
   '               strText = ""
   '               For intR = 0 To UBound(strTemp) - 1
   '                  strText = strText & strTemp(intR) & " " & GetStaffName(strTemp(intR), True) & ","
   '               Next intR
    '              If strText <> "" Then strText = Left(strText, Len(strText) - 1)
   '               oContext = oContext & "原其他出庭律師：" & strText
   '               strTemp = Split(strPublicTemp, ",")
   '               strText = ""
   '               For intR = 0 To UBound(strTemp) - 1
   '                  '判斷變更出庭費
   '                  If strTemp(intR) <> "" And InStr(strTemp(intR), "|") > 0 Then
   '                      strText = strText & Mid(strTemp(intR), 1, InStr(strTemp(intR), "|") - 1) & " " & GetStaffName(Mid(strTemp(intR), 1, InStr(strTemp(intR), "|") - 1), True) & "變更" & Mid(strTemp(intR), InStr(strTemp(intR), "|") + 1) & ","
   '                  Else
   '                      strText = strText & strTemp(intR) & " " & GetStaffName(strTemp(intR), True) & ","
   '                  End If
   '               Next intR
   '               If strText <> "" Then strText = Left(strText, Len(strText) - 1)
   '               oContext = oContext & "　　改為：" & Trim(strText) & vbCrLf
   '            End If
   '            oContext = oContext & vbCrLf & vbCrLf & "此程序已收款且修改承辦人員，若有必要請自行調整出庭費傳票摘要！" & vbCrLf
   '            'Modified by Lydia 2023/01/13
   '            'PUB_SendMail strUserNum, Pub_GetSpecMan("財務處出納人員"), "", oSubject, oContext
   '            PUB_SendMail strUserNum, Pub_GetSpecMan("財務處總帳人員"), "", oSubject, oContext
   '         End If
   '      End If
   '   End If 'If strExc(0) <> "" Then
   'End If
   'end 2022/12/08
   'end 2024/09/30 (113/11/01上線)
    
   'Added by Lydia 2020/10/07 (10/5) 若案件性質或案件屬性有改時Email通知秀玲提醒確認案源及金額是否需調整。案件屬性第1次設定時要與接洽單檔比較是否不同。
   'Modified by Lydia 2020/11/26  A3類案源為非訴訟案，點數都回智慧所，與屬性是否為智財權無關
   'If m_LOS01 <> "" Then
   'Modified by Lydia 2025/08/18 改成有案源就通知
   'If m_LOS01 <> "" And m_LOS02 <> "A3" Then
   If m_LOS15 <> "" Then
       strExc(0) = "": strExc(1) = ""
       If Text(28).Visible = True And Text(28).Locked = False Then 'Added by Lydia 2020/11/03 判斷可維護才檢查
            If Text(28).Tag = "" Then  '與接洽單檔比較
                If PUB_ChkTwoStrLst(m_CRL84, Text(28).Text) = False Then
                    strExc(1) = strExc(1) & "、案件屬性"
                    strExc(0) = strExc(0) & vbCrLf & "原案件屬性：" & m_CRL84 & vbCrLf & "現案件屬性：" & Text(28).Text
                End If
            ElseIf Text(28).Tag <> Text(28).Text Then
                If PUB_ChkTwoStrLst(Text(28).Tag, Text(28).Text) = False Then
                    strExc(1) = strExc(1) & "、案件屬性"
                    strExc(0) = strExc(0) & vbCrLf & "原案件屬性：" & Text(28).Tag & vbCrLf & "現案件屬性：" & Text(28).Text
                End If
            End If
       End If 'Added by Lydia 2020/11/03
       If Text(9).Tag <> Text(9).Text Then
          Call ClsPDGetCaseProperty(lc01, Text(9).Tag, strExc(3))
          strExc(1) = strExc(1) & "、案件性質"
          strExc(0) = strExc(0) & vbCrLf & "原案件性質：" & Text(9).Tag & strExc(3) & vbCrLf & "現案件性質：" & Text(9).Text & lbe(9).Caption
          'Added by Lydia 2025/08/18
          'strExc(0) = strExc(0) & vbCrLf & "案源類別A類改成BC類並且專業部收文尚未分案，請一併清除LOS01。" 'Mark by Lydia 2025/10/20
       End If
       If strExc(0) <> "" Then
           '主旨
           strExc(1) = "法務分案" & lc01 & "-" & lc02 & IIf(lc03 <> "0", "-" & lc03, "") & IIf(lc04 <> "00", "-" & lc04, "") & "，改變" & Mid(strExc(1), 2)
           '內文
           strExc(2) = "法律所案號：" & lc01 & "-" & lc02 & IIf(lc03 <> "0", "-" & lc03, "") & IIf(lc04 <> "00", "-" & lc04, "") & "(" & lbePaperNum & ")" & vbCrLf & _
                            "專業部案號：" & m_LOS01cp01 & "-" & m_LOS01cp02 & IIf(m_LOS01cp03 <> "0", "-" & m_LOS01cp03, "") & IIf(m_LOS01cp04 <> "00", "-" & m_LOS01cp04, "") & "(" & m_LOS01 & ")" & vbCrLf & _
                             strExc(0)
           strExc(2) = strExc(2) & vbCrLf & vbCrLf & "請確認案源及金額是否需調整。" 'Added by Lydia 2025/10/20 參考內法的2021/09/09 加提醒
           'Added by Lydia 2025/10/20
           If Text(9).Tag <> Text(9).Text Then
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

'   If UBound(strCP09) = t Then
    If m_Cpindex = m_CPCount - 1 Then
      cmdok.Enabled = False
      intForm = 0
      intNowRec = 0
      blnIsFormBack = True
      Unload Me
      frm081001.Show
      Exit Sub
   End If
'   t = t + 1
'   GetData (t)
   cmdNext_Click
End Sub

Private Sub cmdPrePic_Click()
   If blnIsSave = False Then
      If MsgBox("你並未存檔,確定離開嗎?", vbYesNo + vbCritical + vbDefaultButton2) = vbNo Then
         Exit Sub
      End If
   End If
   intForm = 0
   intNowRec = 0
   blnIsFormBack = True
   Unload Me
   frm081001.Show
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
   Unload frm081001
End Sub

Private Sub Combo1_Click()
   Text(2).Text = stName(Combo1.ListIndex + 1)
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
   frm08100202.Show
End Sub

Private Sub Command3_Click()
   frm08100203.Show
   If IsNoExistData Then Unload frm08100203
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
   
   'Modified by Lydia 2022/12/05
   'frm071018.Hide
   'Set frm071018.UpForm = frm081002
   'frm071018.lbePaperNum = Me.lbePaperNum
   'frm071018.lbeNumber = Me.lbeNumber
   'Modified by Lydia 2024/09/30 (113/11/01上線) 傳入收文號、判斷是否進入第2次以上
   'Call frm071018.SetParent(Me, Me.lbePaperNum, IIf(Me.Tag = "", True, False), Trim(Text(7)))
   Call frm071018.SetParent(Me, Me.lbePaperNum, IIf(bolActCaseLawer = False, True, False), Trim(Text(7)), Trim(Text(9)), IIf(bolActCaseLawer = True And Me.Tag = "", True, False))
   'end  2022/12/05
   bolActCaseLawer = True 'Added by Lydia 2024/09/30 (113/11/01上線)
   
   Me.Hide
   frm071018.Show vbModal
End Sub

Private Sub Form_Load()
 Dim i As Integer, n As Integer
 m_CPCount = 0
   MoveFormToCenter Me
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
   
   t = 0
   blnIsSave = False
   With frm081001.MSHFlexGrid1
      n = 0
      For i = 1 To .Rows - 1
         .row = i
         .col = 0
         If .Text = "v" Then
            .col = 2
            ReDim Preserve strCP09(n)
            strCP09(n) = .Text
            m_CPCount = m_CPCount + 1
            n = n + 1
         End If
      Next
   End With
   GetData (0)
    'Add By Cheng 2002/12/03
    If m_CPCount = 1 Then Me.CmdNext.Enabled = False
    
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
   'Set frm081002 = Nothing  'Remove by Lydia 2023/03/14  form2.0會有問題，改在呼叫時清除記憶體變數
   'Add By Sindy 2011/6/8
   strPublicTemp = ""
   Unload frm071018
   '2011/6/8 End
End Sub

Private Sub MSHFlexGrid1_Click()
  Dim intRow As Integer
  Dim i As Integer
  
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
      With MSHFlexGrid1
             For i = 1 To .Rows - 1
                If .TextMatrix(i, 0) = "v" Then
                   Text(12).Text = ChangeTDateStringToTString(MSHFlexGrid1.TextMatrix(i, 2))
                   Text(13).Text = ChangeTDateStringToTString(MSHFlexGrid1.TextMatrix(i, 3))
                   Text(15).Text = MSHFlexGrid1.TextMatrix(i, 8)
                  'Modify By Cheng 2002/03/25
'                   Text(18).Text = MSHFlexGrid1.TextMatrix(i, 13)
                   Text(18).Text = IIf(Val(Me.Text(18).Tag) > 0, Me.Text(18).Tag & IIf(Len(Me.MSHFlexGrid1.TextMatrix(i, 13)) > 0, "，" & Me.MSHFlexGrid1.TextMatrix(i, 13), ""), Me.MSHFlexGrid1.TextMatrix(i, 13))
                   Exit For
                Else
                   Text(12).Text = ""
                   Text(13).Text = ""
                   Text(15).Text = ""
                  'Modify By Cheng 2002/03/25
'                   Text(18).Text = ""
                   Text(18).Text = "" & Me.Text(18).Tag
                End If
            Next i
      End With
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
   Select Case Index
      Case 12
         If Text(Index) <> "" Then strDate = Text(Index)
         TextInverse Text(Index)
         'edit by nickc 2007/06/11  切換輸入法改用API
         'Text(Index).IMEMode = 2
         CloseIme
      Case 2
         If Combo1.ListIndex = 0 Then
             'edit by nickc 2007/06/11  切換輸入法改用API
             'Text(Index).IMEMode = 1
             OpenIme
         Else
             'edit by nickc 2007/06/11  切換輸入法改用API
             'Text(Index).IMEMode = 2
             CloseIme
         End If
      Case 4, 11, 18, 19
          'edit by nickc 2007/06/11  切換輸入法改用API
          'Text(Index).IMEMode = 1
          OpenIme
      Case Else
         'edit by nickc 2007/06/11  切換輸入法改用API
         'Text(Index).IMEMode = 2
         CloseIme
         TextInverse Text(Index)
   End Select
End Sub

'Modified by Lydia 2021/09/22 改成Form 2.0
'Private Sub Text_KeyPress(Index As Integer, KeyAscii As Integer)
Private Sub Text_KeyPress(Index As Integer, KeyAscii As MSForms.ReturnInteger)
   Select Case Index
      Case Index
         KeyAscii = UpperCase(KeyAscii)
         'Add By Cheng 2002/04/24
         If Index = 16 Then
            If KeyAscii <> 89 And KeyAscii <> 8 Then
               KeyAscii = 0
            End If
         End If
   End Select
End Sub

Private Sub Text_LostFocus(Index As Integer)
   'edit by nickc 2007/06/11  切換輸入法改用API
   'Text(Index).IMEMode = 2
   CloseIme
End Sub

'Added by Lydia 2023/03/14
Private Sub Text_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
   If Index = 28 Then
      If Button = 2 Then Forms(0).PopupMenu2 Text(Index)
   End If
End Sub

Private Sub Text_Validate(Index As Integer, Cancel As Boolean)
Dim strTempName As String, i As Integer, blnIsEmpty As Boolean
'Added by Lydia 2019/02/14
Dim m_SalesST15 As String '畫面上智權人員的收文部門
Dim m_Tuser As String '創新業務部預設收文人員

 strTempName = ""
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
            Text(Index) = UCase(Text(Index))
            If Left(Text(Index), 1) = "Y" Then
               Text(Index) = "X" & Mid(Text(Index), 2)
            ElseIf Left(Text(Index), 1) <> "X" Then
               MsgBox "當事人代碼輸入錯誤!", vbExclamation, "法務－分案"
               TextInverse Text(Index)
               Cancel = True
               Exit Sub
            End If
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
         '2008/8/13 CANCEL BY SONIA
         'Else
         '   MsgBox "當事人不可空白", vbCritical
         '   Cancel = True
         '2008/8/13 END
         End If
         'Add By Cheng 2002/08/22
         If Cancel = False Then
            'Modified by Lydia 2024/06/13
            'If m_strCust1 <> Me.Text(1).Text Then
            If m_LC11 <> ChangeCustomerL(Me.Text(1).Text) Then
               If Not PUB_EditCustOk(Me.lbePaperNum.Caption, lc01, lc02, lc03, lc04) Then Cancel = True
            End If
         End If
      Case 2
         If Text(2) <> "" Then Text(2) = UCase(Text(2))
         stName(Combo1.ListIndex + 1) = Text(2).Text
         For i = 1 To 3
            If stName(i) <> "" Then
               blnIsEmpty = False
               Exit For
            Else
               blnIsEmpty = True
            End If
         Next i
            
         If blnIsEmpty = True Then
            MsgBox "案件名稱不可全部空白!", vbExclamation, "法務－分案"
            Text(2).SetFocus
            Exit Sub
         End If

      Case 5
         If Text(Index) <> "" Then
            Text(Index) = UCase(Text(Index))
            If Text(Index) <> "Y" Then
               Cancel = True
               DataErrorMessage 1, "是否為智慧財產權案"
            End If
         End If
      Case 6
         If Text(Index) <> "" Then
            Text(Index) = UCase(Text(Index))
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
'              If m_ODate <> "" Then
'                 If Text(12) <> m_ODate Then
'                     i = MsgBox("是否要修改此期限?", vbYesNo, "修改")
'                     If i = 7 Then
'                         Text(12) = m_ODate
'                     End If
'                 End If
'               End If
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
'         If m_LDate <> "" Then
'            If Text(13) <> m_LDate Then
'               i = MsgBox("是否要修改此期限?", vbYesNo, "修改")
'               If i = 7 Then
'                  Text(13) = m_LDate
'               End If
'            End If
'         End If
      
      Case 14
         Text(Index) = UCase(Text(Index))
         If Text(Index) <> "" And Text(Index) <> "N" Then
            Cancel = True
            DataErrorMessage 1, "是否算案件數"
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
      Case 17
         Text(Index) = UCase(Text(Index))
         If Text(Index) <> "" And Text(Index) <> "N" Then
            Cancel = True
            DataErrorMessage 1, "是否向客戶收款"
         End If
      Case 20
         If Text(20) <> "" Then
            'edit by nickc 2007/02/07 不用 dll 了
            'If objPublicData.GetNation(Text(20), strTempName) = True Then
            If ClsPDGetNation(Text(20), strTempName) = True Then
               lbe(20).Caption = strTempName
            Else
               Cancel = True
               lbe(20).Caption = ""
            End If
            'Add By Sindy 2012/4/13
            If (lc01 = "FCL" Or lc01 = "LIN") And Text(20).Text <> 台灣國家代號 Then
               ShowMsg MsgText(9219)
               Cancel = True
            End If
            '2012/4/13 End
         End If
       Case 21
            Text(Index) = UCase(Text(Index))
            lbe(2).Caption = ""
            If Text(Index) <> "" Then
               If Left(Text(Index), 1) = "X" Then
                     Text(Index) = "Y" & Mid(Text(Index), 2)
                     
               ElseIf Left(Text(Index), 1) <> "Y" Then
                      MsgBox "代理人代碼輸入錯誤!", vbExclamation, "法務－分案"
                      TextInverse Text(Index)
                      Cancel = True
                      Exit Sub
                      
               End If
               If ReadFagent(Text(Index)) = False Then
                  Cancel = True
                  Exit Sub
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
    'Add By Cheng 2002/12/06
    KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtcp01_LostFocus()
   If Len(txtCP01) > 0 Then txtCP01 = UCase(txtCP01)
End Sub

Private Sub txtcp01_Validate(Cancel As Boolean)
   If txtCP01 <> "" Then
     txtCP01 = UCase(txtCP01)
     If txtCP01 <> GetCaseNumSysKind(lbeNumber) Then
     DataErrorMessage 1, "本所案號"
     Cancel = True
     End If
   End If
   If Cancel Then TextInverse txtCP01
End Sub

Private Sub txtcp02_Validate(Cancel As Boolean)
 Dim strTemp As String, i As Integer, yn As Integer, strlcTemp As String
   If txtCP02 <> "" Then
      If Len(txtCP02) = 6 Then
         'edit by nickc 2007/02/07 不用 dll 了
         'If objLawDll.ChkCaseNum(txtcp01, txtcp02) Then
         If ClsPDChkCaseNum(txtCP01, txtCP02) Then
            TextInverse txtCP02
            Cancel = True
         Else
            If txtCP03 = "" Then
               strlcTemp = txtCP01 + txtCP02 + "000"
            Else
               strlcTemp = txtCP01 + txtCP02 + txtCP03 + txtCP04
            End If
         End If
      Else
         DataErrorMessage 1, "本所案號"
         Cancel = True
      End If
   End If
   If Cancel Then TextInverse txtCP02
End Sub

Private Sub txtcp03_Validate(Cancel As Boolean)
  If txtCP02 <> "" And txtCP03 = "" Then txtCP03 = "0"
  If txtCP03 <> "" Then
      If Len(txtCP03) > 1 Then
         DataErrorMessage 1, "本所案號"
         Cancel = True
         Exit Sub
      End If
   End If
   If Cancel Then TextInverse txtCP03
End Sub

Private Sub GetData(ByVal intI As Integer)
Dim yn As Boolean, i As Integer, j As Integer
Dim RsTemp As New ADODB.Recordset, St(33) As String

   'Add By Sindy 2010/8/6 增加CP65
   'Modify By Sindy 2011/6/8 +LC47,cp31
   'Modify By Sindy 2012/6/1 +cp27
   'Modified by Lydia 2023/08/14 +CP162
   'Modified by Lydia 2024/07/29
   strExc(1) = "select cp01,cp02,cp03,cp04,cp05,cp06,cp07,cp09,cp10,cp13," & _
      "cp14,cp16,cp18,cp19,cp20,cp26,cp29,cp43,LC22,LC23,cp49,cp57,cp64," & _
      "lc05,lc06,lc07,lc08,lc11,lc13,lc14,lc15,lc16,lc27,CP60,CP65,LC47,cp31,cp27,CP162 from lawcase, caseprogress " & _
      "where cp09='" + strCP09(intI) + "' AND CP01=LC01 AND CP02=LC02 AND CP03=LC03 AND CP04=LC04 " & _
      " order by LC01,LC02,LC03,LC04"
   intI = 0
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(1))    'edit by nickc 2007/02/07 不用 dll 了 Set rstemp = objLawDll.ReadRstMsg(intI, strExc(1))
   If intI = 1 Then
      For j = 1 To 33
         If IsNull(RsTemp.Fields(j - 1).Value) = False Then
            St(j) = RsTemp.Fields(j - 1).Value
         Else
            St(j) = ""
         End If
      Next
      
      '910703 Sieg 701
      If Not IsNull(RsTemp.Fields("CP60")) Then
         m_CP60 = RsTemp.Fields("CP60")
      Else
         m_CP60 = ""
      End If
      
      ' CreateID Add By Sindy 2010/8/6
      m_CP65 = ""
      If IsNull(RsTemp.Fields("CP65")) = False Then
         m_CP65 = RsTemp.Fields("CP65")
      End If
      
      'Add By Sindy 2012/6/1
      m_CP27 = Empty
      If IsNull(RsTemp.Fields("CP27")) = False Then
         m_CP27 = RsTemp.Fields("CP27")
      End If
      '2012/6/1 End
      m_CP162 = "" & RsTemp.Fields("CP162") 'Added by Lydia 2023/08/14 (案件進度)案源單號
      
      If Not IsNull(RsTemp.Fields("LC11")) Then
         m_LC11 = RsTemp.Fields("LC11")
      Else
         m_LC11 = ""
      End If
      
      lc01 = St(1)
      lc02 = St(2)
      lc03 = St(3)
      lc04 = St(4)
      lbeNumber = GiveSymbol(St(1), St(2), St(3), St(4), LcTmp)
      lbeNumber.Tag = LcTmp
      Text(0) = ChangeWStringToTString(St(5))
      Text(12) = ChangeWStringToTString(St(6))
      Text(13) = ChangeWStringToTString(St(7))
      m_ODate = ChangeWStringToTString(St(6))
      m_LDate = ChangeWStringToTString(St(7))
      
      lbePaperNum = St(8)
      Text(9) = St(9): ChgType (9)
      Text(9).Tag = Text(9).Text 'Added by Lydia 2020/01/21
      bChkPaid = PUB_ChkIsPaid(lbePaperNum, m_CCP60) 'Added by Lydia 2024/09/30 (113/11/01上線) 已否已請款、已付款
      
      Text(10) = St(10): ChgType (10)
      Text(7) = St(11): ChgType (7)
      Text(7).Tag = Text(7).Text 'Added by Lydia 2020/05/20
      lbeCost = St(12)
      lbePointNum = St(13)
      lbeMoney = St(14)
      Text(17) = UCase(St(15))
      Text(14) = UCase(St(16))
      Text(8) = St(17): ChgType (8)
      Text(15) = UCase(St(18))
      Text(11) = St(21)
      lbeCloseDate = ChangeWStringToTDateString(St(22))
      Text(18) = St(23)
      'Add By Cheng 2002/03/25
      Me.Text(18).Tag = Me.Text(18).Text
      'Modify By Cheng 2002/04/24
      '是否取消閉卷欄不要顯示資料
      '    Text(16) = UCase(St(27))
      'Add By Cheng 2002/04/24
      If UCase(St(27)) = "Y" Then
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
      
      Text(1) = ChangeCustomerS(St(28)): ChgType (1)
      Text(5) = UCase(St(29))
      Text(28) = "" & RsTemp.Fields("lc47") 'Add By Sindy 2011/6/8
      Text(28).Tag = Text(28) 'Add by Amy 2018/07/30
      '案件屬性
      'Modified by Lydia 2022/08/10
      'For i = 0 To 4
      '   If InStr(Text(28).Text, Trim(Check1(i).Caption)) > 0 Then
      '      Check1(i).Value = 1
      '   End If
      'Next i
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
      Text(6) = St(32)
      Text(19) = St(33)
      j = 0 'Added by Lydia 2023/03/14
      For i = 1 To 3
         stName(i) = ""
         stName(i) = St(23 + i)
         'Added by Lydia 2023/03/14
         If stName(i) <> "" And j = 0 Then
             j = i
         End If
         'end 2023/03/14
      Next
      'Modified by Lydia 2023/03/14 Debug:下一筆分案無法帶出案件名稱
      'Text(2) = St(22)
      Text(2) = stName(j)
      Text(20) = St(31): ChgType (20)
      If St(19) <> "" Then
         Text(21) = ChangeCustomerS(St(19)) ': ChgType (1)
         If ReadFagent(Text(21)) Then
         End If
      End If
      m_LC22 = ChangeCustomerL(Text(21)) 'Added by Lydia 2024/06/13
      Text(3) = St(20)
      Text(4) = St(30)
      'Add By Cheng 2002/08/22
      'm_strCust1 = "" & Me.Text(1).Text 'Mark by Lydia 2024/06/13
     
      Getrs
      'Modified by Lydia 2023/03/14 Debug:下一筆分案無法帶出案件名稱
      'If Combo1.ListCount > 0 Then Combo1.ListIndex = 0
      Combo1.ListIndex = j - 1
      
      m_CP31 = "" & RsTemp.Fields("CP31") 'Add by Amy 2018/10/15
      'CP31為Y時,Shape1內的欄位才可修改,否則鎖住 'Memo by Lydia 2023/03/14 不使用Shape1了
      If "" & RsTemp.Fields("CP31") = "Y" Then
         Text(20).Locked = False
         Combo1.Locked = False
         Text(2).Locked = False
         Text(1).Locked = False
         Text(21).Locked = False
         Text(3).Locked = False
         Text(6).Locked = False
         Text(5).Locked = False
         Text(4).Locked = False
         Text(28).Locked = False
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
         Text(20).Locked = True
         Combo1.Locked = True
         Text(2).Locked = True
         Text(1).Locked = True
         Text(21).Locked = True
         Text(3).Locked = True
         Text(6).Locked = True
         Text(5).Locked = True
         Text(4).Locked = True
         Text(28).Locked = True
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
      'Added by Lydia 2016/01/27 L,CFL案要勾選案性質
      'Modified by Lydia 20202/07/03 +FCL
      If (lc01 = "CFL" Or lc01 = "FCL") Then
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
         '   'Next i
         '   For Each oObj In Check1
         '       oObj.Enabled = False
         '   Next
         '   'end 2022/08/10
         'end 2023/02/01
         End If
      End If
      'end 2016/01/27
      '2007/8/13 ADD BY SONIA銷卷提醒
      CheckCaseDestroy lc01, lc02, lc03, lc04
      '2007/8/13 END
      
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
      m_CL02 = Trim(strPublicTemp) 'Added by Lydia 2020/10/15
      
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
      'end 2020/05/20
   End If
   blnIsSave = False
End Sub

Private Function SaveData() As Boolean
 Dim i As Integer, blnIsChange As Boolean
 Dim cp01 As String, cp02 As String, cp03 As String, cp04 As String
 Dim strTmp As String, iStep As Integer
'Add By Cheng 2002/09/09
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
'edit by nickc 2007/02/07
'Dim sLC(1 To T_LC) As String
Dim sLC() As String
ReDim sLC(1 To TF_LC) As String
Dim strApply As Variant 'Add by Amy 2018/10/15
   
 '911107 nick transation
On Error GoTo CheckingErr
SaveData = True
cnnConnection.BeginTrans

   iStep = 1
   '若有輸入轉本所案號
   If txtCP01 <> "" Then
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
   If cp03 = "" Then cp03 = "0"
   If cp04 = "" Then cp04 = "00"
   LcTmp = cp01 & cp02 & cp03 & cp04
   'Modify By Cheng 2002/04/24
'   ' 91.04.04 modify by louis (單引號)
'   strExc(1) = "update lawcase set lc05=" & CNULL(StName(1)) & ",lc06=" & CNULL(ChgSQL(StName(2))) & _
'      ",lc07=" & CNULL(StName(3)) & ", lc08 = " & CNULL(Text(16)) & _
'      ",lc11=" & CNULL(ChangeCustomerL(Text(1))) & ",lc13=" & CNULL(Text(5)) & _
'      ",lc14=" & CNULL(Text(4)) & ",lc15=" & CNULL(Text(20)) & _
'      ",lc16=" & CNULL(Text(6)) & ", lc27=" & CNULL(ChgSQL(Text(19))) & _
'      " where " & ChgLawcase(LcTmp)

   '910703 Sieg 701
   Dim strTmp1(1 To 3) As String
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
                    'Modify By Cheng 2002/12/03
                    '新增申請人失敗, 應取消存檔
'                  '911202 nick
'                  cnnConnection.CommitTrans
'                  Exit Function
                    GoTo CheckingErr
               End If
            End If
         End If
      End If
   End If

   If Me.Text(16).Text = "Y" Then
      strExc(1) = " , LC08=NULL, LC09=NULL, LC10=NULL "
   Else
      strExc(1) = " "
   End If
   
   strExc(1) = "update lawcase set lc05=" & CNULL(ChgSQL(stName(1))) & ",lc06=" & CNULL(ChgSQL(stName(2))) & _
      ",lc07=" & CNULL(stName(3)) & strExc(1) & _
      ",lc11=" & CNULL(ChangeCustomerL(Text(1))) & ",lc13=" & CNULL(Text(5)) & _
      ",lc14=" & CNULL(Text(4)) & ",lc15=" & CNULL(Text(20)) & _
      ",lc16=" & CNULL(Text(6)) & ",lc22=" & CNULL(ChangeCustomerL(Text(21))) & _
      ",lc23=" & CNULL(Text(3)) & ",lc27=" & CNULL(ChgSQL(Text(19))) & _
      ",lc47=" & CNULL(ChgSQL(Text(28))) & _
      " where " & ChgLawcase(LcTmp)
      
'Add by Amy 2018/07/30 有修改lc47寫log
If Text(28).Visible = True Then
    If Text(28).Tag <> Text(28) Then Pub_SeekTbLog strExc(1)
End If
 '911107 nick transation
cnnConnection.Execute strExc(1)

'   iStep = 2
   iStep = iStep + 1
   
   ' 91.04.04 modify by louis (單引號)
   strTmp = "cp05=" & CNULL(TransDate(Text(0), 2)) & _
      ",cp06=" & CNULL(TransDate(Text(12), 2)) & _
      ",cp07=" & CNULL(TransDate(Text(13), 2)) & _
      ",cp10=" & CNULL(Text(9)) & ",cp12=" & CNULL(GetST15(Text(10))) & ",cp13=" & CNULL(Text(10)) & ",cp14=" & CNULL(Text(7)) & _
      ",cp20=" & CNULL(Text(17)) & ",cp26=" & CNULL(Text(14)) & ",cp29=" & CNULL(Text(8)) & _
      ",cp43=" & CNULL(Text(15)) & ",cp64=" & CNULL(ChgSQL(Text(18))) & ",cp49=" & CNULL(ChgSQL(Text(11))) & _
      " where cp09=" & CNULL(lbePaperNum)

'Add By Cheng 2002/08/23
ProcChgCaseNum:
   
   '911107 nick transation
   'SaveData = objLawDll.ExecSQL(iStep - 1, strExc)
   
   'Add By Cheng 2002/09/09
   If blnIsChange Then
      If Me.txtCP01.Text <> "" And Me.txtCP02.Text <> "" And SaveData = True Then
         StrSQLa = "SELECT * FROM LAWCASE WHERE " & ChgLawcase(cp01 & cp02 & cp03 & cp04)
         rsA.CursorLocation = adUseClient
         rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
         If rsA.RecordCount <= 0 Then
            If PUB_ReadLawCaseData(sLC(), lc01, lc02, lc03, lc04) Then
               sLC(1) = Me.txtCP01.Text
               sLC(2) = Me.txtCP02.Text
               sLC(3) = Left(Me.txtCP03.Text & "0", 1)
               sLC(4) = Left(Me.txtCP04.Text & "00", 2)
               If PUB_AddNewLawCase(sLC()) Then
               'Add By Cheng 2002/12/03
               Else
                    GoTo CheckingErr
               End If
            End If
        'Add By Cheng 2002/12/06
        '若基本檔有資料, 若是否新案欄為'Y'更新為Null
        Else
              'modify by sonia 2019/8/8 要用收文號更新
              'strSql = " Update CaseProgress Set CP31=DECODE(CP31,'Y',NULL,CP31) WHERE " & ChgCaseprogress(cp01 & cp02 & cp03 & cp04)
              strSql = " Update CaseProgress Set CP31=NULL WHERE CP09='" & Me.lbePaperNum.Caption & "'"
              'end 2019/8/8
              cnnConnection.Execute strSql
         End If
         If rsA.State <> adStateClosed Then rsA.Close
         Set rsA = Nothing
      End If
   End If
   
   If blnIsChange Then
      'Modify By Cheng 2002/08/23
'      strExc(iStep) = "update caseprogress set cp01=" & CNULL(CP01) & _
'         ",cp02=" & CNULL(cp02) & ",cp03=" & CNULL(cp03) & _
'         ",cp04=" & CNULL(cp04) & "," & strTmp
      strExc(iStep) = "update caseprogress set cp01=" & CNULL(cp01) & _
         ",cp02=" & CNULL(cp02) & ",cp03=" & CNULL(cp03) & _
         ",cp04=" & CNULL(cp04) & " WHERE CP09='" & Me.lbePaperNum.Caption & "'"
    
    'Modify By Cheng 2002/12/03
'     '911107 nick transation
'    cnnConnection.Execute strExc(1)
    Pub_SeekTbLog strExc(iStep) 'Added by Lydia 2025/09/25 L-006989誤輸入轉入ACS案
    cnnConnection.Execute strExc(iStep)
         
      iStep = iStep + 1
      'Add By Cheng 2002/08/23
      strExc(iStep) = "UPDATE CASEPROGRESS SET CP43='' WHERE CP09='" & Me.lbePaperNum.Caption & "'"
      
    'Modify By Cheng 2002/12/03
'     '911107 nick transation
'    cnnConnection.Execute strExc(1)
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
      
        'Modify By Cheng 2002/12/03
'        '911107 nick transation
'        cnnConnection.Execute strExc(1)
        cnnConnection.Execute strExc(iStep)
      
      iStep = iStep + 1
   End If
   
   'Modify By Cheng 2002/08/23
   If Not blnIsChange Then
        'Modify By Cheng 2002/12/03
'      SaveNextProgress
      If SaveNextProgress = False Then GoTo CheckingErr
   End If
   If SaveData Then blnIsSave = True
   frm081001.SetDataComplete lbePaperNum.Caption
   
   'add by nickc 2005/03/17 加入加乘註記及寄件值
   m_CP98 = "": m_CP101 = "": m_CP104 = ""
   If PUB_GetFlagValue(Me.lbePaperNum.Caption, m_CP98, m_CP101, m_CP104) = True Then
      strSql = "update caseprogress set cp98=" & m_CP98 & ",cp101=" & m_CP101 & ",cp104=" & m_CP104 & " WHERE CP09 = '" & Me.lbePaperNum.Caption & "' "
      cnnConnection.Execute strSql
   End If
   'PUB_UpdateCaseValue Me.lbePaperNum.Caption'Remove by Morgan 2005/4/13 改由 trigger 更新
   
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
      'Add By Sindy 2011/6/8
      'Modified by Lydia 2022/12/08 配合輸入出庭費,改成先存暫存檔再寫入正式Table
      'strExc(0) = "delete from caselawer where cl01='" & lbePaperNum & "'"
      'cnnConnection.Execute strExc(0)
      'If strPublicTemp <> "" Then
      '   strTemp = Split(strPublicTemp, ",")
      '   For i = 0 To UBound(strTemp) - 1
      '      strExc(0) = "insert into caselawer values('" & lbePaperNum & "','" & strTemp(i) & "')"
      '      cnnConnection.Execute strExc(0)
      '   Next i
      If Me.Tag <> "" And InStr(Me.Tag, "|") > 0 Then '有點選「出庭律師」
         If PUB_SaveCaseLawer(lbePaperNum, Mid(Me.Tag, InStr(Me.Tag, "|") + 1), strPublicTemp) = True Then
            strTemp = Split(strPublicTemp, ",")
            'Memo by Lydia 2024/09/30 (113/11/01上線) PUB_SaveCaseLawer回傳strPublicTemp有包含變更出庭費;ex. 員工編號|出庭費：15000=>7500 ,
         End If
         '判斷補資料的情況
         'Mark by Lydia 2024/09/30 (113/11/01上線) 經過測試不用了
         'If strPublicTemp = m_CL02 & Text(7) & "," Then
         '  m_CL02 = strPublicTemp
         'End If
         'end 2024/09/30 (113/11/01上線)
      'Added by Lydia 2024/09/30 (113/11/01上線) 刪除所有「出庭律師」記錄
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
      'end 2024/09/30 (113/11/01上線)
      End If
      If strPublicTemp <> m_CL02 Then  'Memo by Lydia 2024/09/30 (113/11/01上線) 回傳strPublicTemp有包含變更出庭費
         'Added by Lydia 2024/09/30 (113/11/01上線) 出庭律師或出庭費有異動，發Email通知財務處總帳人員；同時整合原有Email通知：'Added by Lydia 2020/10/15 (法律所案源)若分案時已經收款則同時E-MAIL給系統特殊設定之財務處出納人員及智權人員＋'Add By Sindy 2011/6/8 修改承辦人或其他出庭律師時，若該收文號已收款(CP75>0)則MAIL通知71005
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
         'end 2024/09/30 (113/11/01上線)
      'end 2022/12/08
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
                  'Modified by Lydia 2022/12/08 外法B類案源的費用在PT案=>CP60, 已收款判斷也在PT案
                  'If m_CP60 <> "" Then
                  '   StrSQLa = "Select a0z01,a0z02,a0z03,a0z04,a0z05 From acc0z0 Where a0z02=" & CNULL(m_CP60)
                  'Mark by Lydia 2024/09/30 (113/11/01上線) 出庭律師或出庭費有異動，發Email通知財務處總帳人員；同時整合原有Email通知：'Added by Lydia 2020/10/15 (法律所案源)若分案時已經收款則同時E-MAIL給系統特殊設定之財務處出納人員及智權人員＋'Add By Sindy 2011/6/8 修改承辦人或其他出庭律師時，若該收文號已收款(CP75>0)則MAIL通知71005
                  'If m_CP60 <> "" Or m_LOS01cp60 <> "" Then
                  '   strExc(1) = ""
                  '   If Left(m_LOS02, 1) = "B" Then
                  '      strExc(1) = m_LOS01cp60
                  '   Else
                  '      strExc(1) = m_CP60
                  '   End If
                  '   If Left(strExc(1), 1) = "E" Then
                  '       StrSQLa = "select nvl(cp75,0) amt1 from caseprogress where cp09=" & CNULL(lbePaperNum)
                  '   ElseIf Left(strExc(1), 1) = "X" Then
                  '       StrSQLa = "select nvl(a1k30,0) amt1 from acc1k0 where a1k01=" & CNULL(strExc(1))
                  '   End If
                   '  If StrSQLa <> "" Then
                  ''end ------ Modified by Lydia 2022/12/08 外法B類案源的費用在PT案=>CP60, 已收款判斷也在PT案
                  '       intI = 1
                  '       Set rsA = ClsLawReadRstMsg(intI, StrSQLa)
                  '       If intI = 1 Then
                  '          'Modified by Lydia 2023/01/13
                  '          'strExc(0) = Pub_GetSpecMan("財務處出納人員")
                  '          strExc(0) = Pub_GetSpecMan("財務處總帳人員")
                  '          'Modified by Lydia 2022/12/08 +Val("" & RsTemp.Fields("amt1")) > 0
                  '          'If strExc(0) <> "" And Val("" & rsA.Fields("amt1")) > 0 Then
                  '             strExc(0) = strExc(0) & IIf(m_LOS04 <> "", ";" & Replace(m_LOS04, ",", ";"), "")
                  '             '主旨
                  '             'Modified by Lydia 2022/12/08
                  '             'strExc(1) = lc01 & "-" & lc02 & IIf(lc03 <> "0", "-" & lc03, "") & IIf(lc04 <> "00", "-" & lc04, "") & "有多個出庭律師，但已收款，請調整收款及傳票資料。"
                  '             strExc(1) = lc01 & "-" & lc02 & IIf(lc03 <> "0", "-" & lc03, "") & IIf(lc04 <> "00", "-" & lc04, "") & _
                  '                     "(" & IIf(Left(m_LOS02, 1) = "B", "專業部案號：" & m_LOS01cp01 & "-" & m_LOS01cp02 & "-" & m_LOS01cp03 & "-" & m_LOS01cp04, "") & _
                  '                     ")有多個出庭律師，但已收款，請調整收款及傳票資料。"
                  '
                  '             '其他出庭律師
                  '             strExc(2) = ""
                  '             'Modified by Lydia 2022/12/08 分析字串
                   '            'If strPublicTemp <> "" Then
                   '            '    StrSQLa = "select getstaffnamelist(" & CNULL(strPublicTemp) & ") from dual "
                   '            '    intI = 1
                  '             '    Set RsTemp = ClsLawReadRstMsg(intI, StrSQLa)
                  '             '    If intI = 1 Then
                  '             '        strExc(2) = "" & RsTemp(0)
                  '             '    End If
                  '             'End If
                  '             strTemp = Split(strPublicTemp, ",")
                  '             For intI = 0 To UBound(strTemp)
                  '                If Trim(strTemp(intI)) <> "" Then
                  '                    If InStr(strTemp(intI), "|") > 0 Then
                  '                        strExc(3) = GetStaffName(Mid(Trim(strTemp(intI)), 1, InStr(Trim(strTemp(intI)), "|") - 1), True)
                  '                    Else
                  '                        strExc(3) = GetStaffName(strTemp(intI), True)
                   '                   End If
                   '                   If strExc(3) <> lbe(7).Caption Then
                   '                      strExc(2) = strExc(2) & "、" & strExc(3)
                   '                   End If
                   '               End If
                   '            Next intI
                   '            If strExc(2) <> "" Then strExc(2) = Mid(strExc(2), 2)
                   '            'end 2022/12/08
                   '            '內文
                   '            'Modified by Lydia 2022/12/08 m_CP60=> IIf(m_CP60 <> "", m_CP60, m_LOS01cp60)
                   '            strExc(3) = "本所案號：" & lc01 & "-" & lc02 & IIf(lc03 <> "0", "-" & lc03, "") & IIf(lc04 <> "00", "-" & lc04, "") & vbCrLf & _
                   '                             "收據號碼：" & IIf(m_CP60 <> "", m_CP60, m_LOS01cp60) & vbCrLf & "出庭律師：" & lbe(7).Caption & IIf(strExc(2) <> "", "、" & strExc(2), "")
                                                
                   '            StrSQLa = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08)" & _
                   '                     " values ('" & strUserNum & "','" & strExc(0) & "'," & strSrvDate(1) & ",to_char(sysdate,'hh24miss')" & _
                   '                     ",'" & ChgSQL(strExc(1)) & "','" & ChgSQL(strExc(3)) & "')"
                   '            cnnConnection.Execute StrSQLa
                   '         End If
                   '      End If
                   '  End If 'If strExc(0) <> "" Then
                  'End If   'If m_CP60 <> "" Then
             End If   'If strPublicTemp <> m_CL02 Then
         End If
         'end 2020/10/15
      End If
   End If 'Added by Lydia 2023/08/14
   
   'Add by Amy 2018/10/15 MCTF控管
   If m_CP31 = "Y" And Text(21) <> MsgText(601) And Trim(Text(1)) <> MsgText(601) Then
      strApply = Split(ChangeCustomerL(Text(1)), ",")
      If UpdMCTF_Cu13(Mid(lbeNumber, 1, InStr(lbeNumber, "-") - 1), ChangeCustomerL(Text(21)), Trim(Text(20)), strApply, "") = False Then
          GoTo CheckingErr
      End If
   End If
   'Added by Lydia 2020/05/20 法律所案源收文
   If strSrvDate(1) >= 法律所案源收文啟用日 And m_LOS15 <> "" Then
       'Mark by Lydia 2020/06/16 改成frm077005「智財訴訟案需專業部配合通知補收文作業」
       'strExc(1) = ""
       'If m_LOS02 <> t_LOS02 Then
       '     '先EMAIL通知智權人員補收配合開庭(因為LOS01會被清空)
       '     Call PUB_AddMailCache_LOS("3", m_LOS15)
       '     strExc(1) = "Y"
       '     '案件屬性有勾專利或商標或著作權時，若案件性質為1101~1104(民事委任律師)時，
      '      '若案源非B1時詢問 是否需智慧所配合開庭？」，若選擇要則將案源更新為B1，同時將案源總收文號LOS01清除；
     '       strSql = "Update LawOfficeSource set LOS01=null, LOS02='B1' where LOS15='" & m_LOS15 & "' "
     '       cnnConnection.Execute strSql, i
      '      If m_LOS10 <> "" Then
      '          '並重新計算TT-999999費用點數更新回去並改案件性質為736(服務費)；再EMAIL通知智權人員補收配合開庭226(B1)
      '          strSql = "Update CaseProgress set CP10='736' where CP09='" & m_LOS10 & "' "
      '          cnnConnection.Execute strSql, i
      '      End If
      ' End If
       'end 2020/06/16
       
       '法律所總收文號1=> 分案時E-MAIL給介紹人通知法務案已收文之郵件
       If Me.lbePaperNum.Caption = m_LOS06 And Text(7).Tag = "" And Text(7).Tag <> Text(7).Text Then
            'Modified by Lydia 2020/05/29 分案和配合開庭通知整合為一封email
            'Modiffied by Lydia 2020/06/16
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
   
 '911107 nick transation
  cnnConnection.CommitTrans
     Exit Function
CheckingErr:
    SaveData = False
     cnnConnection.RollbackTrans
End Function

Private Sub Getrs()
   'Modified by Lydia 2023/03/22 法律案(L,FCL)不同審級收文期限沖銷：案件不同審級會收文-1、-2…案號，但實際是同一案件。ChgNextProgress(LcTmp)=>IIf(lc01 = "FCL", "np02='" & lc01 & "' and np03='" & lc02 & "' ", ChgNextProgress(LcTmp))
   strExc(1) = "select '',decode(np02||np07,cpm01||CPM02,CPM03,CPM04)," + _
      "decode(np08,null,'',SUBSTR(np08,1,4)-1911||'/'||SUBSTR(np08,5,2)||'/'||SUBSTR(np08,7,2))," + _
      "decode(np09,null,'',SUBSTR(np09,1,4)-1911||'/'||SUBSTR(np09,5,2)||'/'||SUBSTR(np09,7,2))," + _
      "np13,np14,decode(np11,null,'',SUBSTR(np11,1,4)-1911||'/'||SUBSTR(np11,5,2)||'/'||" + _
      "SUBSTR(np11,7,2)),np06,np01,np07,np16,np17,np18,np15,np22 from nextprogress,CASEPROPERTYMAP where " + _
      "" + IIf(lc01 = "FCL", "np02='" & lc01 & "' and np03='" & lc02 & "' ", ChgNextProgress(LcTmp)) + " and (np02=cpm01(+) and np07=cpm02(+)) and (np06='N' or np06 is null)"
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
   '.Col = 13: .ColWidth(13) = 0
   .col = 13: .ColWidth(13) = 1500: .Text = "備註"
   .col = 14: .ColWidth(14) = 0
   intLastRow = 0
   blnOKtoShow = True
End With
End Sub

Private Function SaveNextProgress() As Boolean
Dim i As Integer, n As Integer, NP07 As String, np08 As String
Dim np16 As String, np17 As String, np18 As String, np06 As String, np01 As String
Dim np22 As String
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
               'Modified by Lydia 2023/03/22 法律案(L,FCL)不同審級收文期限沖銷：案件不同審級會收文-1、-2…案號，但實際是同一案件。ChgNextProgress(LcTmp)=>IIf(lc01 = "FCL", "np02='" & lc01 & "' and np03='" & lc02 & "' ", ChgNextProgress(LcTmp))
               strExc(1) = "update nextprogress set np06=" & CNULL(np06) & _
                  " where np01=" & CNULL(np01) & " and " & IIf(lc01 = "FCL", "np02='" & lc01 & "' and np03='" & lc02 & "' ", ChgNextProgress(LcTmp)) & _
                  " and np07=" & CNULL(NP07) & " and np08=" & CNULL(np08) & _
                  " and np16=" & CNULL(np16) & " and np17=" & CNULL(np17) & _
                  " and np18=" & CNULL(np18) & " and np22=" & CNULL(np22)
                  
               '911107 nick transation
               'SaveNextProgress = objLawDll.ExecSQL(1, strExc)
               cnnConnection.Execute strExc(1)
           End If
         Next
   End With
   '911107 nick
SaveNextProgress = True
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
End Function

Private Function AllTextBeforeSaveCheck() As Boolean
Dim i As Integer
Dim strTempName As String
Dim blnIsEmpty As Boolean

strTempName = ""
AllTextBeforeSaveCheck = True
If Text(0) = "" Then
   MsgBox "收文日不可空白", vbCritical
   Text(0).SetFocus
   AllTextBeforeSaveCheck = True
   Exit Function
End If
  
  stName(Combo1.ListIndex + 1) = Text(2).Text
         For i = 1 To 3
            If stName(i) <> "" Then
               blnIsEmpty = False
               Exit For
            Else
               blnIsEmpty = True
            End If
         Next i
            
         If blnIsEmpty = True Then
            MsgBox "案件名稱不可全部空白!", vbExclamation, "法務－分案"
            Text(2).SetFocus
            AllTextBeforeSaveCheck = True
            Exit Function
         End If


If Text(1) = "" Then
   '2008/8/13 ADD BY SONIA
   'MsgBox "當事人代號不可空白", vbCritical
   'Text(1).SetFocus
   'AllTextBeforeSaveCheck = True
   'Exit Function
   '2008/8/13 END
Else
   Text(1) = UCase(Text(1))
   If Left(Text(1), 1) = "Y" Then
      Text(1) = "X" & Mid(Text(1), 2)
   ElseIf Left(Text(1), 1) <> "X" Then
          MsgBox "當事人代碼輸入錯誤!", vbExclamation, "法務－分案"
          TextInverse Text(1)
          AllTextBeforeSaveCheck = True
          Exit Function
   End If
   'edit by nickc 2007/02/07 不用 dll 了
   'If objPublicData.GetCustomer(Text(1), strTempName) Then
   If ClsPDGetCustomer(Text(1), strTempName) Then
      lbe(1) = strTempName
   Else
       lbe(1) = ""
       AllTextBeforeSaveCheck = True
       Exit Function
   End If
End If

If Text(21).Text <> "" Then
   Text(21) = UCase(Text(21))
   lbe(2).Caption = ""
   If Left(Text(21), 1) = "X" Then
      Text(21) = "Y" & Mid(Text(21), 2)
   ElseIf Left(Text(21), 1) <> "Y" Then
          MsgBox "代理人代碼輸入錯誤!", vbExclamation, "法務－分案"
          TextInverse Text(21)
          AllTextBeforeSaveCheck = True
          Exit Function
   End If
   If ReadFagent(Text(21)) = False Then
      AllTextBeforeSaveCheck = True
      Exit Function
   End If
End If

 If Text(5) <> "" Then
    Text(5) = UCase(Text(5))
    If Text(5) <> "Y" Then
       DataErrorMessage 1, "是否為智慧財產權案"
       AllTextBeforeSaveCheck = True
       TextInverse Text(5)
       Exit Function
    End If
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

If Text(9) = "" Or IsNull(Text(9)) Then
   MsgBox "案件性質不可空白", vbCritical
   Text(9).SetFocus
   AllTextBeforeSaveCheck = True
   Exit Function
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

   If Text(12) <> "" Then
      If CheckIsTaiwanDate(Text(12)) Then
          If Text(13) <> "" Then
             If Val(Text(13)) - Val(Text(12)) < 0 Then DataErrorMessage 13
          End If
      Else
        Text(12).SetFocus
        TextInverse Text(12)
        AllTextBeforeSaveCheck = True
        Exit Function
      End If
    End If
      If m_ODate <> "" Then
         If Text(12) <> m_ODate Then
             If MsgBox("是否要修改本所期限?", vbYesNo, "修改") = vbNo Then
                 Text(12) = m_ODate
             End If
         End If
       End If
  

      If Text(13) <> "" Then
        If CheckIsTaiwanDate(Text(13)) Then
            If Text(12) <> "" Then
               If Val(Text(13)) - Val(Text(12)) < 0 Then DataErrorMessage 12
            End If
        Else
           Text(13).SetFocus
           TextInverse Text(13)
           AllTextBeforeSaveCheck = True
           Exit Function
        End If
     End If
   If m_LDate <> "" Then
      If Text(13) <> m_LDate Then
         If MsgBox("是否要修改法定期限?", vbYesNo, "修改") = vbNo Then
            Text(13) = m_LDate
         End If
      End If
   End If
   
   Text(14) = UCase(Text(14))
   If Text(14) <> "" And Text(14) <> "N" Then
       DataErrorMessage 1, "是否算案件數"
       AllTextBeforeSaveCheck = True
       Text(14).SetFocus
       TextInverse Text(14)
       Exit Function
   End If
   
   If Text(15) <> "" Then
      Text(15) = UCase(Text(15))
      If Text(15) = lbePaperNum Then
         MsgBox "且不可為本身之收文號", vbCritical
         AllTextBeforeSaveCheck = True
         Text(15).SetFocus
         TextInverse Text(15)
         Exit Function
      End If
      'edit by nickc 2007/02/07 不用 dll 了
      'If Not objLawDll.GetRelation(LcTmp, lbePaperNum, Text(15)) Then
      If Not ClsLawGetRelation(LcTmp, lbePaperNum, Text(15)) Then
         AllTextBeforeSaveCheck = True
         Text(15).SetFocus
         TextInverse Text(15)
         Exit Function
      End If
   End If

   If Text(16) <> "" Then
         Text(16) = UCase(Text(16))
       If Text(16) = "Y" Then
          i = MsgBox("確定取消閉卷?", vbYesNo, "詢問")
          If i = 7 Then Text(16) = ""
       Else
          DataErrorMessage 1, "是否閉卷"
          AllTextBeforeSaveCheck = True
          Text(16).SetFocus
          TextInverse Text(16)
          Exit Function
       End If
   End If
   
    Text(17) = UCase(Text(17))
    If Text(17) <> "" And Text(17) <> "N" Then
       DataErrorMessage 1, "是否向客戶收款"
       AllTextBeforeSaveCheck = True
       Text(17).SetFocus
       TextInverse Text(17)
       Exit Function
    End If

  strTempName = ""
   If Text(20) <> "" Then
      'edit by nickc 2007/02/07 不用 dll 了
      'If objPublicData.GetNation(Text(20), strTempName) = True Then
      If ClsPDGetNation(Text(20), strTempName) = True Then
         lbe(20).Caption = strTempName
      Else
         Text(20).SetFocus
         TextInverse Text(20)
         lbe(20).Caption = ""
         AllTextBeforeSaveCheck = True
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
   lbe(i) = ""
   Select Case i
      Case 1
         If Text(i) <> "" Then
            'edit by nickc 2007/02/07 不用 dll 了
            'If objPublicData.GetCustomer(Text(i), strTempName) Then
            If ClsPDGetCustomer(Text(i), strTempName) Then
               lbe(i) = strTempName
            End If
         '2008/8/13 CANCEL BY SONIA
         'Else
         '    MsgBox "當事人不可空白", vbCritical
         '2008/8/13 END
         End If
      Case 7, 8, 10
         If Text(i) <> "" Then
            'edit by nickc 2007/02/07 不用 dll 了
            'If objPublicData.GetStaff(Text(i), strTempName) Then
            If ClsPDGetStaff(Text(i), strTempName) Then
               lbe(i) = strTempName
            End If
         End If
      Case 9
         If Text(i) = "" Then
            MsgBox "案件性質不可空白", vbCritical
         Else
            'edit by nickc 2007/02/07 不用 dll 了
            'If objPublicData.GetCaseProperty(CheckCaseNum, Text(i), strTempName, False) Then
            If ClsPDGetCaseProperty(CheckCaseNum, Text(i), strTempName, False) Then
               lbe(i) = strTempName
            End If
         End If
      Case 20
         If Text(i) <> "" Then
            'edit by nickc 2007/02/07 不用 dll 了
            'If objPublicData.GetNation(Text(i), strTempName) = True Then
            If ClsPDGetNation(Text(i), strTempName) = True Then
               lbe(i).Caption = strTempName
            End If
         End If
   End Select
End Sub

Private Sub txtcp04_Validate(Cancel As Boolean)
   If txtCP02 <> "" And txtCP04 = "" Then txtCP04 = "00"
   If txtCP04 <> "" Then
      If Len(txtCP04) <> 2 Then
         DataErrorMessage 1, "本所案號"
         Cancel = True
      Else
         If ChangText Then TextInverse txtCP02
      End If
   End If
   'Add By Cheng 2002/08/22
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
  
  Me.Tag = "" 'Added by Lydia 2022/12/08
  
  'Added by Lydia 2023/03/14
  For Each oObj In Check1
     oObj.Value = 0
  Next
  For Each oObj In Check2
     oObj.Value = 0
  Next
  For Each oObj In lbe
     oObj.Caption = ""
  Next
  'end 2023/03/14
  
  'Added by Lydia 2024/09/30 (113/11/01上線)
  bChkPaid = False
  bolActCaseLawer = False
  'end 2024/09/30 (113/11/01上線)
End Sub

Private Function ReadFagent(ByVal strCP44 As String) As Boolean
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   Dim strFA01 As String
   
   strFA01 = GetNewFagent(strCP44)
   
   strSql = "SELECT nvl(fa05,nvl(fa04,fa06)) FROM FAGENT " & _
            "WHERE FA01 = '" & Left(strCP44, 8) & "'"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.EOF = False Then
      If Not IsNull(rsTmp.Fields(0)) Then
         lbe(2).Caption = rsTmp.Fields(0)
         ReadFagent = True
      End If
   Else
      ReadFagent = False
      MsgBox "代理人代碼不存在!", vbExclamation, "發文"
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Function

'Add By Cheng 2002/05/24
Private Function TxtValidate() As Boolean
Dim objTxt As Object
Dim ii As Integer
Dim Cancel As Boolean
Dim bolData As Boolean, strMCTF(0) As String, strTmp(0) As String

TxtValidate = False
For Each objTxt In Text
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

'Add By Sindy 2011/6/20
If (InStr(Text(28).Text, "專利") > 0 Or InStr(Text(28).Text, "商標") > 0 Or _
   InStr(Text(28).Text, "著作權") > 0 Or InStr(Text(28).Text, "智財權") > 0) And Text(5) <> "Y" Then
   MsgBox "案件屬性欄位出現專利,商標,著作權或智財權字樣時,智財權案須為Y!!!", vbExclamation + vbOKOnly
   Text(5).SetFocus
   Exit Function
End If

'Added by Lydia 2016/01/27 L,CFL案必須輸入案件性質(分配目標點數用)
'Modified by Lydia 20202/07/03 +FCL
If (lc01 = "CFL" Or lc01 = "FCL") And Text(28).Text = "" Then
      MsgBox "法務案請勾選案件屬性!", vbExclamation + vbOKOnly
   Exit Function
End If
'end 2016/01/27

'Add by Amy 2018/10/15  MCTF組別控制(有輸代理人且為MCTF,判斷申請人若與代理人的MCTF組別不同不可收文)
If Len(Trim(Text(21))) > 0 Then
    bolData = GetCusORFagentData(ChangeCustomerL(Text(21)), "FA120", strMCTF())
    If Left(strMCTF(0), 4) = "MCTF" Then
        bolData = GetCusORFagentData(ChangeCustomerL(Text(1)), "CU13", strTmp())
        If strMCTF(0) <> strTmp(0) And Left(strTmp(0), 4) = "MCTF" Then
            MsgBox "當事人智權人員(" & strTmp(0) & ")與代理人" & Text(21) & "商標管控智權人員(" & strMCTF(0) & ")不同，不可存檔！", vbExclamation + vbOKOnly
            Exit Function
        End If
    End If
End If
'end 2018/10/15

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
strExc(1) = ChangeCustomerL(Text(21))
strExc(2) = ChangeCustomerL(m_LC22)
If strExc(1) <> "" And strExc(1) <> strExc(2) Then
   If GetAgentAndState(strExc(1), strExc(3), , , , lc01, strExc(8), False) = False Then
      Me.SSTab1.Tab = 0
      Text(21).SetFocus
      Text_GotFocus 21
      Exit Function
   End If
End If
'end 2024/06/13
   

'Added by Lydia 2021/09/22 檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
If PUB_ChkUniText(Me, , True, "TextBox") = False Then
    Exit Function
End If

   'Added by Morgan 2022/11/9 補收款必須輸入相關總收文號(案源TT扣出庭費及收款抓出庭律師要用)
   If lc01 = "FCL" And Text(9) = "997" And Text(15) = "" Then
      MsgBox lbe(9) & "必須輸入相關總收文號！", vbExclamation
      Text(15).SetFocus
      Exit Function
   End If
   'end 2022/11/9
   
   'Added by Lydia 2022/12/08 修改承辦人檢查; 若已有CaseLawer資料,但是修改承辦人後沒有再進入維護
   'Mark by Lydia 2024/09/30 (113/11/01上線) 與出庭費領取的檢查合併
   'If m_CL02 <> "" And InStr(strPublicTemp, Text(7)) = 0 Then
   '   'Modified by Lydia 2024/07/29 改成提醒
   '   'MsgBox "請進入出庭律師資料輸入作業，檢查出庭律師是否正確！", vbExclamation
   '   If MsgBox("承辦人不在出庭律師內，是否繼續存檔？", vbExclamation + vbYesNo + vbDefaultButton2, "出庭律師檢查") = vbNo Then
   '      Command4.SetFocus
   '      Exit Function
   '   End If 'Added by Lydia 2024/07/29
   'End If
   ''end 2022/12/08
   'end 2024/09/30 (113/11/01上線)
   
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
   'm_LOS01cp60 = "" 'Added by Lydia 2022/12/08 'Mark by Lydia 2024/09/30 (113/11/01上線)
   
   'Modified by Lydia 2020/10/07 抓接洽單編號 + NVL(X.LOS17,X.LOS18) CRLNO
   'Modified by Lydia 2021/06/18 +LOS05
   'Modified by Lydia 2022/12/08 +CP60
   'Modified by Lydia 2023/08/14 改用(案件進度)案源單號
   'stSQL = "select nvl(X.LOS07,0) ord1,X.LOS01,X.LOS02,X.LOS04,X.LOS05,X.LOS06,X.LOS10,X.LOS15,CP01,CP02,CP03,CP04,NVL(X.LOS17,X.LOS18) CRLNO,CP60 " & _
                "from LawOfficeSource X,CaseProgress where X.LOS06='" & lbePaperNum & "' and X.LOS01=CP09(+) order by ord1, X.LOS01 "
   'Modified by Lydia 2024/09/30 (113/11/01上線) 案源FC代理人
   'stSQL = "select nvl(X.LOS07,0) ord1,X.LOS01,X.LOS02,X.LOS04,X.LOS05,X.LOS06,X.LOS10,X.LOS15,CP01,CP02,CP03,CP04,NVL(X.LOS17,X.LOS18) CRLNO,CP60 " & _
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
      'm_LOS01cp60 = "" & RsQ.Fields("CP60") 'Added by Lydia 2022/12/08 'Mark by Lydia 2024/09/30 (113/11/01上線)
      '案源總收文號之本所案號
      m_LOS01cp01 = "" & RsQ.Fields("cp01")
      m_LOS01cp02 = "" & RsQ.Fields("cp02")
      m_LOS01cp03 = "" & RsQ.Fields("cp03")
      m_LOS01cp04 = "" & RsQ.Fields("cp04")
      '(原)案源案件類型
      m_LOS02 = "" & RsQ.Fields("LOS02")
      t_LOS02 = m_LOS02
      m_LOS04 = "" & RsQ.Fields("LOS04") 'Added by Lydia 2020/10/15
      m_LOS01fa = "" & RsQ.Fields("FAGENT") 'Added by Lydia 2024/09/30 (113/11/01上線)
      
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

