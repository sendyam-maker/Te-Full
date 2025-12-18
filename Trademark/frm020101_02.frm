VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm020101_02 
   BorderStyle     =   1  '單線固定
   Caption         =   "分案"
   ClientHeight    =   6084
   ClientLeft      =   228
   ClientTop       =   984
   ClientWidth     =   9120
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6084
   ScaleWidth      =   9120
   Begin VB.TextBox txtF0309 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  '沒有框線
      ForeColor       =   &H000000C0&
      Height          =   180
      Left            =   7365
      Locked          =   -1  'True
      TabIndex        =   192
      Top             =   450
      Width           =   1665
   End
   Begin VB.CommandButton cmdFile 
      Caption         =   "檢視接洽單"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   1800
      TabIndex        =   181
      Top             =   10
      Width           =   1065
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5388
      Left            =   48
      TabIndex        =   91
      Top             =   648
      Width           =   10872
      _ExtentX        =   19177
      _ExtentY        =   9504
      _Version        =   393216
      Tabs            =   5
      Tab             =   3
      TabsPerRow      =   6
      TabHeight       =   420
      TabMaxWidth     =   3175
      TabCaption(0)   =   "第一頁"
      TabPicture(0)   =   "frm020101_02.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "grdList"
      Tab(0).Control(1)=   "Frame1"
      Tab(0).Control(2)=   "textCP122"
      Tab(0).Control(3)=   "FraLOS"
      Tab(0).Control(4)=   "textTM14"
      Tab(0).Control(5)=   "textTM80"
      Tab(0).Control(6)=   "textTM81"
      Tab(0).Control(7)=   "textCP10"
      Tab(0).Control(8)=   "txtDivCaseNo(2)"
      Tab(0).Control(9)=   "txtDivCaseNo(4)"
      Tab(0).Control(10)=   "txtDivCaseNo(3)"
      Tab(0).Control(11)=   "txtDivCaseNo(1)"
      Tab(0).Control(12)=   "txtDivCaseNo(0)"
      Tab(0).Control(13)=   "textCP14"
      Tab(0).Control(14)=   "textTM06"
      Tab(0).Control(15)=   "textCP05"
      Tab(0).Control(16)=   "textTM02_2"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "textTM01"
      Tab(0).Control(18)=   "textTM03"
      Tab(0).Control(19)=   "textTM04"
      Tab(0).Control(20)=   "textCP57"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "textSP59"
      Tab(0).Control(22)=   "textSP58"
      Tab(0).Control(23)=   "cmdNation"
      Tab(0).Control(24)=   "textTM45"
      Tab(0).Control(25)=   "textTM44"
      Tab(0).Control(26)=   "textTM23"
      Tab(0).Control(27)=   "textTM29"
      Tab(0).Control(28)=   "textCP43"
      Tab(0).Control(29)=   "textCP26"
      Tab(0).Control(30)=   "textTM28"
      Tab(0).Control(31)=   "textTM10_2"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "textTM08_2"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "textTM08"
      Tab(0).Control(34)=   "textCP13"
      Tab(0).Control(35)=   "textTM10"
      Tab(0).Control(36)=   "textCP06"
      Tab(0).Control(37)=   "textCP07"
      Tab(0).Control(38)=   "textCP10_2"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).Control(39)=   "textTM02"
      Tab(0).Control(40)=   "Frame21"
      Tab(0).Control(41)=   "Frame2"
      Tab(0).Control(42)=   "cboTM08"
      Tab(0).Control(43)=   "Label1(121)"
      Tab(0).Control(44)=   "Label58"
      Tab(0).Control(45)=   "textTM07"
      Tab(0).Control(46)=   "textTM05"
      Tab(0).Control(47)=   "textTM05_1"
      Tab(0).Control(48)=   "textTM80_2"
      Tab(0).Control(49)=   "textTM81_2"
      Tab(0).Control(50)=   "textSP59_2"
      Tab(0).Control(51)=   "textSP58_2"
      Tab(0).Control(52)=   "textTM44_2"
      Tab(0).Control(53)=   "textTM23_2"
      Tab(0).Control(54)=   "textCP13_2"
      Tab(0).Control(55)=   "textCP14_2"
      Tab(0).Control(56)=   "Label56"
      Tab(0).Control(57)=   "Label27"
      Tab(0).Control(58)=   "Label26"
      Tab(0).Control(59)=   "lblDivCase"
      Tab(0).Control(60)=   "Label42"
      Tab(0).Control(61)=   "Label37"
      Tab(0).Control(62)=   "Label36"
      Tab(0).Control(63)=   "Label35"
      Tab(0).Control(64)=   "Label34"
      Tab(0).Control(65)=   "Label33"
      Tab(0).Control(66)=   "Label40"
      Tab(0).Control(67)=   "Label22"
      Tab(0).Control(68)=   "Label55"
      Tab(0).Control(69)=   "Label14"
      Tab(0).Control(70)=   "Label12"
      Tab(0).Control(71)=   "Label11"
      Tab(0).Control(72)=   "Label13"
      Tab(0).Control(73)=   "Label15"
      Tab(0).Control(74)=   "Label10"
      Tab(0).Control(75)=   "Label8"
      Tab(0).Control(76)=   "Label7"
      Tab(0).Control(77)=   "Label6"
      Tab(0).Control(78)=   "Label1(1)"
      Tab(0).Control(79)=   "Label1(8)"
      Tab(0).Control(80)=   "Label5"
      Tab(0).Control(81)=   "Label25"
      Tab(0).Control(82)=   "Label4"
      Tab(0).Control(83)=   "Label3"
      Tab(0).ControlCount=   84
      TabCaption(1)   =   "第二頁"
      TabPicture(1)   =   "frm020101_02.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "textTM89"
      Tab(1).Control(1)=   "textTM93"
      Tab(1).Control(2)=   "textTM88"
      Tab(1).Control(3)=   "textTM92"
      Tab(1).Control(4)=   "textTM87"
      Tab(1).Control(5)=   "textTM91"
      Tab(1).Control(6)=   "textTM86"
      Tab(1).Control(7)=   "textTM90"
      Tab(1).Control(8)=   "textSP32"
      Tab(1).Control(9)=   "textTM35"
      Tab(1).Control(10)=   "textTM34"
      Tab(1).Control(11)=   "textTM26"
      Tab(1).Control(12)=   "textTM25"
      Tab(1).Control(13)=   "textTM85"
      Tab(1).Control(14)=   "textTM84"
      Tab(1).Control(15)=   "textTM83"
      Tab(1).Control(16)=   "textTM82"
      Tab(1).Control(17)=   "textTM24"
      Tab(1).Control(18)=   "Label51"
      Tab(1).Control(19)=   "Label50"
      Tab(1).Control(20)=   "Label49"
      Tab(1).Control(21)=   "Label48"
      Tab(1).Control(22)=   "Label47"
      Tab(1).Control(23)=   "Label46"
      Tab(1).Control(24)=   "Label45"
      Tab(1).Control(25)=   "Label44"
      Tab(1).Control(26)=   "Label43"
      Tab(1).Control(27)=   "Label23"
      Tab(1).Control(28)=   "Label16"
      Tab(1).Control(29)=   "Label9"
      Tab(1).Control(30)=   "Label21"
      Tab(1).Control(31)=   "Label20"
      Tab(1).Control(32)=   "Label38"
      Tab(1).Control(33)=   "Label18"
      Tab(1).Control(34)=   "Label19"
      Tab(1).Control(35)=   "Label17"
      Tab(1).ControlCount=   36
      TabCaption(2)   =   "第三頁"
      TabPicture(2)   =   "frm020101_02.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "textTM72"
      Tab(2).Control(1)=   "textTM72_2"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "textTM130"
      Tab(2).Control(3)=   "textCP44"
      Tab(2).Control(4)=   "textTM32"
      Tab(2).Control(5)=   "textTM09"
      Tab(2).Control(6)=   "Text1"
      Tab(2).Control(7)=   "Text2"
      Tab(2).Control(8)=   "Text3"
      Tab(2).Control(9)=   "Text4"
      Tab(2).Control(10)=   "cmdPriority"
      Tab(2).Control(11)=   "textTM21S"
      Tab(2).Control(11).Enabled=   0   'False
      Tab(2).Control(12)=   "textTM22S"
      Tab(2).Control(12).Enabled=   0   'False
      Tab(2).Control(13)=   "cboTM72"
      Tab(2).Control(14)=   "textTM05S"
      Tab(2).Control(15)=   "textTM58"
      Tab(2).Control(16)=   "textCP64"
      Tab(2).Control(17)=   "textCP44_2"
      Tab(2).Control(18)=   "Label70"
      Tab(2).Control(19)=   "lblTM130"
      Tab(2).Control(20)=   "Label1(115)"
      Tab(2).Control(21)=   "Label54"
      Tab(2).Control(22)=   "Label52"
      Tab(2).Control(23)=   "Label29"
      Tab(2).Control(24)=   "Label41"
      Tab(2).Control(25)=   "Label30"
      Tab(2).Control(26)=   "Label31"
      Tab(2).Control(27)=   "Label32"
      Tab(2).Control(28)=   "Line1"
      Tab(2).Control(29)=   "Label24"
      Tab(2).Control(30)=   "Label28"
      Tab(2).ControlCount=   31
      TabCaption(3)   =   "聯絡人／TF基礎案"
      TabPicture(3)   =   "frm020101_02.frx":0054
      Tab(3).ControlEnabled=   -1  'True
      Tab(3).Control(0)=   "Label53"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "Label59"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "Label60"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "Label61"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).Control(4)=   "Label62"
      Tab(3).Control(4).Enabled=   0   'False
      Tab(3).Control(5)=   "Label63"
      Tab(3).Control(5).Enabled=   0   'False
      Tab(3).Control(6)=   "textTM38"
      Tab(3).Control(6).Enabled=   0   'False
      Tab(3).Control(7)=   "textTM41"
      Tab(3).Control(7).Enabled=   0   'False
      Tab(3).Control(8)=   "textTM40"
      Tab(3).Control(8).Enabled=   0   'False
      Tab(3).Control(9)=   "textTM43"
      Tab(3).Control(9).Enabled=   0   'False
      Tab(3).Control(10)=   "textTM39"
      Tab(3).Control(10).Enabled=   0   'False
      Tab(3).Control(11)=   "textTM42"
      Tab(3).Control(11).Enabled=   0   'False
      Tab(3).Control(12)=   "Frame22"
      Tab(3).Control(12).Enabled=   0   'False
      Tab(3).Control(13)=   "cmdTFBaseNo"
      Tab(3).Control(13).Enabled=   0   'False
      Tab(3).ControlCount=   14
      TabCaption(4)   =   "簽核"
      TabPicture(4)   =   "frm020101_02.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "txtF0301"
      Tab(4).Control(1)=   "CmdAddInfo"
      Tab(4).Control(2)=   "GRD1"
      Tab(4).Control(3)=   "Label67"
      Tab(4).Control(4)=   "Label68"
      Tab(4).Control(5)=   "txtF0407"
      Tab(4).Control(6)=   "Label66"
      Tab(4).Control(7)=   "txtNote"
      Tab(4).ControlCount=   8
      Begin VB.CommandButton cmdTFBaseNo 
         Caption         =   "TF基礎案號數"
         Height          =   285
         Left            =   144
         Style           =   1  '圖片外觀
         TabIndex        =   209
         Top             =   3000
         Width           =   1404
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdList 
         Height          =   1008
         Left            =   -73812
         TabIndex        =   204
         Top             =   4032
         Width           =   7512
         _ExtentX        =   13250
         _ExtentY        =   1778
         _Version        =   393216
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
      Begin VB.Frame Frame22 
         Caption         =   "大陸查名"
         Height          =   555
         Left            =   90
         TabIndex        =   196
         Top             =   2250
         Width           =   8775
         Begin VB.CheckBox ChkEP43 
            Caption         =   "不查名"
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   3150
            TabIndex        =   200
            Top             =   210
            Width           =   1005
         End
         Begin VB.TextBox textEP43 
            Height          =   300
            Left            =   1260
            TabIndex        =   198
            Top             =   180
            Width           =   730
         End
         Begin MSForms.Label lblEP43 
            Height          =   255
            Left            =   2040
            TabIndex        =   199
            Top             =   210
            Width           =   915
            Size            =   "1614;450"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin VB.Label Label72 
            Caption         =   "查名人員："
            Height          =   195
            Left            =   120
            TabIndex        =   197
            Top             =   240
            Width           =   1005
         End
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  '沒有框線
         Height          =   555
         Left            =   -74970
         TabIndex        =   194
         Top             =   4440
         Visible         =   0   'False
         Width           =   1005
         Begin VB.TextBox textTM136 
            Height          =   264
            Left            =   630
            MaxLength       =   1
            TabIndex        =   39
            Top             =   240
            Width           =   345
         End
         Begin VB.Label Label71 
            AutoSize        =   -1  'True
            Caption         =   "註冊證型式:  1:電子      2:紙本"
            Height          =   540
            Left            =   60
            TabIndex        =   195
            Top             =   30
            Width           =   1035
            WordWrap        =   -1  'True
         End
      End
      Begin VB.TextBox textCP122 
         BackColor       =   &H00FFFFFF&
         Height          =   270
         Left            =   -66750
         MaxLength       =   1
         TabIndex        =   190
         Top             =   270
         Width           =   255
      End
      Begin VB.TextBox txtF0301 
         BorderStyle     =   0  '沒有框線
         Height          =   255
         Left            =   -73770
         Locked          =   -1  'True
         TabIndex        =   188
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton CmdAddInfo 
         Caption         =   "呈分案主管"
         CausesValidation=   0   'False
         Height          =   400
         Left            =   -67290
         TabIndex        =   184
         Top             =   420
         Width           =   1065
      End
      Begin VB.Frame FraLOS 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   0  '沒有框線
         Caption         =   "Frame2"
         Height          =   220
         Left            =   -69630
         TabIndex        =   179
         Top             =   3750
         Width           =   3375
         Begin VB.TextBox txtLOSagree 
            Height          =   270
            Left            =   1920
            MaxLength       =   1
            TabIndex        =   38
            Top             =   -8
            Width           =   405
         End
         Begin VB.Label LBL6 
            Caption         =   "是否需要法律所配合：　　　(Y: 配合) "
            Height          =   200
            Left            =   60
            TabIndex        =   180
            Top             =   30
            Width           =   3140
         End
      End
      Begin VB.TextBox textTM72 
         BackColor       =   &H00FFFFC0&
         Height          =   270
         Left            =   -69168
         MaxLength       =   1
         TabIndex        =   73
         Top             =   5028
         Visible         =   0   'False
         Width           =   285
      End
      Begin VB.TextBox textTM72_2 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   0  '沒有框線
         Height          =   240
         Left            =   -68868
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   177
         TabStop         =   0   'False
         Top             =   5028
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.TextBox textTM42 
         Height          =   270
         Left            =   1350
         TabIndex        =   78
         Top             =   1590
         Width           =   5235
      End
      Begin VB.TextBox textTM39 
         Height          =   270
         Left            =   1350
         TabIndex        =   75
         Top             =   690
         Width           =   5235
      End
      Begin VB.TextBox textTM130 
         Height          =   270
         Left            =   -70350
         MaxLength       =   1
         TabIndex        =   69
         Top             =   1260
         Width           =   255
      End
      Begin VB.TextBox textTM14 
         Height          =   270
         Left            =   -71010
         MaxLength       =   7
         TabIndex        =   19
         Top             =   1620
         Width           =   1095
      End
      Begin VB.TextBox textCP44 
         Height          =   270
         Left            =   -69780
         MaxLength       =   9
         TabIndex        =   67
         Top             =   900
         Width           =   972
      End
      Begin VB.TextBox textTM32 
         Height          =   270
         Left            =   -73635
         MaxLength       =   1500
         TabIndex        =   62
         Top             =   630
         Width           =   7400
      End
      Begin VB.TextBox textTM09 
         Height          =   270
         Left            =   -73635
         MaxLength       =   395
         TabIndex        =   61
         Top             =   330
         Width           =   7400
      End
      Begin VB.TextBox textTM89 
         Height          =   270
         Left            =   -73410
         MaxLength       =   154
         TabIndex        =   56
         Top             =   3816
         Width           =   6972
      End
      Begin VB.TextBox textTM93 
         Height          =   270
         Left            =   -73410
         MaxLength       =   100
         TabIndex        =   57
         Top             =   4083
         Width           =   6972
      End
      Begin VB.TextBox textTM88 
         Height          =   270
         Left            =   -73410
         MaxLength       =   154
         TabIndex        =   53
         Top             =   3000
         Width           =   6972
      End
      Begin VB.TextBox textTM92 
         Height          =   270
         Left            =   -73410
         MaxLength       =   100
         TabIndex        =   54
         Top             =   3267
         Width           =   6972
      End
      Begin VB.TextBox textTM87 
         Height          =   270
         Left            =   -73410
         MaxLength       =   154
         TabIndex        =   50
         Top             =   2184
         Width           =   6972
      End
      Begin VB.TextBox textTM91 
         Height          =   270
         Left            =   -73410
         MaxLength       =   100
         TabIndex        =   51
         Top             =   2451
         Width           =   6972
      End
      Begin VB.TextBox textTM86 
         Height          =   270
         Left            =   -73410
         MaxLength       =   154
         TabIndex        =   47
         Top             =   1368
         Width           =   6972
      End
      Begin VB.TextBox textTM90 
         Height          =   270
         Left            =   -73410
         MaxLength       =   100
         TabIndex        =   48
         Top             =   1635
         Width           =   6972
      End
      Begin VB.TextBox Text1 
         Height          =   270
         Left            =   -73635
         MaxLength       =   3
         TabIndex        =   63
         Top             =   930
         Width           =   612
      End
      Begin VB.TextBox Text2 
         Height          =   270
         Left            =   -73020
         MaxLength       =   6
         TabIndex        =   64
         Top             =   930
         Width           =   1095
      End
      Begin VB.TextBox Text3 
         Height          =   270
         Left            =   -71940
         MaxLength       =   1
         TabIndex        =   65
         Top             =   930
         Width           =   255
      End
      Begin VB.TextBox Text4 
         Height          =   270
         Left            =   -71715
         MaxLength       =   2
         TabIndex        =   66
         Top             =   930
         Width           =   492
      End
      Begin VB.CommandButton cmdPriority 
         Caption         =   "輸入(&V)"
         Height          =   252
         Left            =   -73635
         TabIndex        =   68
         Top             =   1260
         Width           =   1000
      End
      Begin VB.TextBox textTM21S 
         BorderStyle     =   0  '沒有框線
         ForeColor       =   &H80000007&
         Height          =   264
         Left            =   -73635
         Locked          =   -1  'True
         TabIndex        =   81
         TabStop         =   0   'False
         Top             =   4710
         Width           =   1212
      End
      Begin VB.TextBox textTM22S 
         BorderStyle     =   0  '沒有框線
         ForeColor       =   &H80000007&
         Height          =   264
         Left            =   -72060
         Locked          =   -1  'True
         TabIndex        =   82
         TabStop         =   0   'False
         Top             =   4710
         Width           =   1212
      End
      Begin VB.TextBox textSP32 
         Height          =   270
         Left            =   -73410
         MaxLength       =   40
         TabIndex        =   60
         Top             =   4608
         Width           =   2292
      End
      Begin VB.TextBox textTM35 
         Height          =   270
         Left            =   -69660
         MaxLength       =   50
         TabIndex        =   59
         Top             =   4350
         Width           =   2292
      End
      Begin VB.TextBox textTM34 
         Height          =   270
         Left            =   -73410
         MaxLength       =   50
         TabIndex        =   58
         Top             =   4350
         Width           =   2292
      End
      Begin VB.TextBox textTM26 
         Height          =   270
         Left            =   -73410
         MaxLength       =   100
         TabIndex        =   45
         Top             =   819
         Width           =   6972
      End
      Begin VB.TextBox textTM25 
         Height          =   270
         Left            =   -73410
         MaxLength       =   154
         TabIndex        =   44
         Top             =   552
         Width           =   6972
      End
      Begin VB.TextBox textTM80 
         Height          =   270
         Left            =   -74070
         MaxLength       =   9
         TabIndex        =   35
         Top             =   3480
         Width           =   1065
      End
      Begin VB.TextBox textTM81 
         Height          =   270
         Left            =   -69630
         MaxLength       =   9
         TabIndex        =   36
         Top             =   3480
         Width           =   1065
      End
      Begin VB.TextBox textCP10 
         Height          =   270
         Left            =   -73860
         MaxLength       =   6
         TabIndex        =   2
         Top             =   510
         Width           =   732
      End
      Begin VB.TextBox txtDivCaseNo 
         Enabled         =   0   'False
         Height          =   270
         Index           =   2
         Left            =   -67170
         MaxLength       =   1
         TabIndex        =   24
         Top             =   1860
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.TextBox txtDivCaseNo 
         Enabled         =   0   'False
         Height          =   270
         Index           =   4
         Left            =   -66570
         MaxLength       =   2
         TabIndex        =   26
         Top             =   1860
         Visible         =   0   'False
         Width           =   390
      End
      Begin VB.TextBox txtDivCaseNo 
         Enabled         =   0   'False
         Height          =   270
         Index           =   3
         Left            =   -66870
         MaxLength       =   1
         TabIndex        =   25
         Top             =   1860
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.TextBox txtDivCaseNo 
         Enabled         =   0   'False
         Height          =   270
         Index           =   1
         Left            =   -67845
         MaxLength       =   6
         TabIndex        =   23
         Top             =   1860
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox txtDivCaseNo 
         Enabled         =   0   'False
         Height          =   270
         Index           =   0
         Left            =   -68235
         MaxLength       =   3
         TabIndex        =   22
         Top             =   1860
         Visible         =   0   'False
         Width           =   390
      End
      Begin VB.TextBox textCP14 
         Height          =   270
         Left            =   -73860
         MaxLength       =   6
         TabIndex        =   0
         Top             =   270
         Width           =   732
      End
      Begin VB.TextBox textTM06 
         Height          =   270
         Left            =   -73590
         MaxLength       =   60
         TabIndex        =   29
         Top             =   2400
         Width           =   6975
      End
      Begin VB.TextBox textCP05 
         Height          =   270
         Left            =   -70110
         MaxLength       =   7
         TabIndex        =   17
         Top             =   1320
         Width           =   1095
      End
      Begin VB.TextBox textTM02_2 
         Height          =   270
         Left            =   -72420
         Locked          =   -1  'True
         MaxLength       =   1
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   1320
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.TextBox textTM01 
         Height          =   270
         Left            =   -73860
         MaxLength       =   3
         TabIndex        =   12
         Top             =   1320
         Width           =   612
      End
      Begin VB.TextBox textTM03 
         Height          =   270
         Left            =   -72180
         MaxLength       =   1
         TabIndex        =   15
         Top             =   1320
         Width           =   255
      End
      Begin VB.TextBox textTM04 
         Height          =   270
         Left            =   -71940
         MaxLength       =   2
         TabIndex        =   16
         Top             =   1320
         Width           =   492
      End
      Begin VB.TextBox textCP57 
         BorderStyle     =   0  '沒有框線
         ForeColor       =   &H80000007&
         Height          =   240
         Left            =   -67170
         Locked          =   -1  'True
         TabIndex        =   130
         TabStop         =   0   'False
         Top             =   1320
         Width           =   1095
      End
      Begin VB.TextBox textSP59 
         Height          =   270
         Left            =   -69630
         MaxLength       =   9
         TabIndex        =   34
         Top             =   3210
         Width           =   1065
      End
      Begin VB.TextBox textSP58 
         Height          =   270
         Left            =   -74070
         MaxLength       =   9
         TabIndex        =   33
         Top             =   3210
         Width           =   1065
      End
      Begin VB.CommandButton cmdNation 
         Caption         =   "指定國家"
         Height          =   252
         Left            =   -69000
         TabIndex        =   119
         Top             =   270
         Width           =   915
      End
      Begin VB.TextBox textTM45 
         Height          =   270
         Left            =   -73830
         MaxLength       =   100
         TabIndex        =   37
         Top             =   3750
         Width           =   4160
      End
      Begin VB.TextBox textTM44 
         Height          =   270
         Left            =   -69630
         MaxLength       =   9
         TabIndex        =   32
         Top             =   2940
         Width           =   1065
      End
      Begin VB.TextBox textTM23 
         Height          =   270
         Left            =   -74070
         MaxLength       =   9
         TabIndex        =   31
         Top             =   2940
         Width           =   1065
      End
      Begin VB.TextBox textTM29 
         Height          =   270
         Left            =   -73590
         MaxLength       =   1
         TabIndex        =   21
         Top             =   1860
         Width           =   372
      End
      Begin VB.TextBox textCP43 
         Height          =   270
         Left            =   -68352
         MaxLength       =   9
         TabIndex        =   20
         Top             =   1590
         Width           =   2172
      End
      Begin VB.TextBox textCP26 
         Height          =   270
         Left            =   -73590
         MaxLength       =   20
         TabIndex        =   18
         Top             =   1590
         Width           =   372
      End
      Begin VB.TextBox textTM28 
         Height          =   270
         Left            =   -71010
         MaxLength       =   20
         TabIndex        =   11
         Top             =   1050
         Width           =   372
      End
      Begin VB.TextBox textTM10_2 
         BorderStyle     =   0  '沒有框線
         Height          =   264
         Left            =   -70380
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   105
         TabStop         =   0   'False
         Top             =   270
         Width           =   1335
      End
      Begin VB.TextBox textTM08_2 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   0  '沒有框線
         Height          =   264
         Left            =   -72912
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   104
         TabStop         =   0   'False
         Top             =   816
         Visible         =   0   'False
         Width           =   768
      End
      Begin VB.TextBox textTM08 
         BackColor       =   &H00FFFFC0&
         Height          =   270
         Left            =   -73392
         MaxLength       =   20
         TabIndex        =   10
         Top             =   816
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.TextBox textCP13 
         Height          =   270
         Left            =   -71010
         MaxLength       =   6
         TabIndex        =   3
         Top             =   510
         Width           =   852
      End
      Begin VB.TextBox textTM10 
         Height          =   270
         Left            =   -71010
         MaxLength       =   20
         TabIndex        =   1
         Top             =   270
         Width           =   612
      End
      Begin VB.TextBox textCP06 
         Height          =   270
         Left            =   -73860
         MaxLength       =   7
         TabIndex        =   4
         Top             =   780
         Width           =   1215
      End
      Begin VB.TextBox textCP07 
         Height          =   270
         Left            =   -71010
         MaxLength       =   7
         TabIndex        =   5
         Top             =   780
         Width           =   1095
      End
      Begin VB.TextBox textCP10_2 
         BorderStyle     =   0  '沒有框線
         Height          =   264
         Left            =   -73110
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   97
         TabStop         =   0   'False
         Top             =   540
         Width           =   1095
      End
      Begin VB.TextBox textTM02 
         Height          =   270
         Left            =   -73260
         MaxLength       =   6
         TabIndex        =   13
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Frame Frame21 
         BorderStyle     =   0  '沒有框線
         Height          =   750
         Left            =   -68040
         TabIndex        =   165
         Top             =   540
         Width           =   1995
         Begin VB.TextBox textCP143 
            Height          =   210
            Left            =   1290
            MaxLength       =   1
            TabIndex        =   8
            Top             =   450
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.TextBox textEP34 
            Height          =   210
            Left            =   1290
            MaxLength       =   1
            TabIndex        =   7
            Top             =   225
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.TextBox textEP06 
            Height          =   210
            Left            =   1290
            MaxLength       =   1
            TabIndex        =   6
            Top             =   -15
            Width           =   255
         End
         Begin VB.Label Label65 
            AutoSize        =   -1  'True
            Caption         =   "查名是否齊備：       (Y/N)"
            Height          =   180
            Left            =   0
            TabIndex        =   176
            Top             =   480
            Visible         =   0   'False
            Width           =   1980
         End
         Begin VB.Label Label57 
            AutoSize        =   -1  'True
            Caption         =   "是否會稿：       (Y/N)"
            Height          =   180
            Left            =   360
            TabIndex        =   166
            Top             =   270
            Visible         =   0   'False
            Width           =   1680
         End
         Begin VB.Label Label64 
            AutoSize        =   -1  'True
            Caption         =   "資料是否齊備：       (Y/N)"
            Height          =   180
            Left            =   0
            TabIndex        =   167
            Top             =   30
            Width           =   1980
         End
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid GRD1 
         Bindings        =   "frm020101_02.frx":008C
         Height          =   1995
         Left            =   -70380
         TabIndex        =   185
         Top             =   2880
         Width           =   4215
         _ExtentX        =   7451
         _ExtentY        =   3535
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
      Begin VB.Frame Frame2 
         Height          =   405
         Left            =   -74050
         TabIndex        =   201
         Top             =   4950
         Width           =   6860
         Begin VB.Frame Frame3 
            Height          =   340
            Left            =   4200
            TabIndex        =   205
            Top             =   60
            Width           =   2170
            Begin VB.OptionButton Option1 
               Caption         =   "之後"
               Height          =   195
               Index           =   2
               Left            =   1380
               TabIndex        =   208
               Top             =   95
               Width           =   705
            End
            Begin VB.OptionButton Option1 
               Caption         =   "當天"
               Height          =   195
               Index           =   0
               Left            =   0
               TabIndex        =   207
               Top             =   95
               Width           =   705
            End
            Begin VB.OptionButton Option1 
               Caption         =   "之前"
               Height          =   195
               Index           =   1
               Left            =   690
               TabIndex        =   206
               Top             =   95
               Width           =   705
            End
         End
         Begin VB.OptionButton OptSendType 
            Caption         =   "立即"
            Height          =   180
            Index           =   1
            Left            =   30
            TabIndex        =   203
            Top             =   170
            Width           =   1260
         End
         Begin VB.OptionButton OptSendType 
            Caption         =   "收款後"
            Height          =   180
            Index           =   2
            Left            =   1290
            TabIndex        =   40
            Top             =   170
            Width           =   870
         End
         Begin VB.OptionButton OptSendType 
            Caption         =   "指定日期"
            Height          =   180
            Index           =   3
            Left            =   2160
            TabIndex        =   41
            Top             =   170
            Width           =   1065
         End
         Begin VB.TextBox textCP142 
            Height          =   264
            Left            =   3230
            MaxLength       =   7
            TabIndex        =   42
            Top             =   120
            Width           =   945
         End
      End
      Begin MSForms.ComboBox cboTM72 
         Height          =   300
         Left            =   -69192
         TabIndex        =   72
         Top             =   4687
         Width           =   2124
         VariousPropertyBits=   679495707
         DisplayStyle    =   3
         Size            =   "3746;529"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "新細明體"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.ComboBox cboTM08 
         Height          =   285
         Left            =   -73860
         TabIndex        =   9
         Top             =   1056
         Width           =   1932
         VariousPropertyBits=   679495707
         DisplayStyle    =   3
         Size            =   "3408;503"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "新細明體"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         Caption         =   "送件方式 :"
         Height          =   195
         Index           =   121
         Left            =   -74910
         TabIndex        =   202
         Top             =   5100
         Width           =   915
      End
      Begin VB.Label Label58 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  '透明
         Caption         =   "是否急件：       (Y/N)"
         Height          =   180
         Left            =   -67680
         TabIndex        =   191
         Top             =   315
         Width           =   1620
      End
      Begin MSForms.TextBox textTM43 
         Height          =   300
         Left            =   1350
         TabIndex        =   79
         Top             =   1890
         Width           =   7545
         VariousPropertyBits=   679493659
         MaxLength       =   30
         Size            =   "13309;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM40 
         Height          =   300
         Left            =   1350
         TabIndex        =   76
         Top             =   990
         Width           =   7545
         VariousPropertyBits=   679493659
         MaxLength       =   30
         Size            =   "13309;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label67 
         AutoSize        =   -1  'True
         Caption         =   "接洽單編號："
         Height          =   180
         Left            =   -74880
         TabIndex        =   189
         Top             =   390
         Width           =   1080
      End
      Begin VB.Label Label68 
         AutoSize        =   -1  'True
         Caption         =   "內容："
         Height          =   300
         Left            =   -74880
         TabIndex        =   187
         Top             =   2880
         Width           =   540
      End
      Begin MSForms.TextBox txtF0407 
         Height          =   1995
         Left            =   -74340
         TabIndex        =   186
         Top             =   2880
         Width           =   3885
         VariousPropertyBits=   -1466939361
         ScrollBars      =   3
         Size            =   "6853;3528"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label66 
         AutoSize        =   -1  'True
         Caption         =   "呈報內容："
         Height          =   180
         Left            =   -74880
         TabIndex        =   183
         Top             =   660
         Width           =   900
      End
      Begin MSForms.TextBox txtNote 
         Height          =   1500
         Left            =   -74880
         TabIndex        =   182
         Top             =   900
         Width           =   8745
         VariousPropertyBits=   -1466939365
         ScrollBars      =   3
         Size            =   "15425;2646"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM41 
         Height          =   300
         Left            =   1350
         TabIndex        =   77
         Top             =   1290
         Width           =   4000
         VariousPropertyBits=   679493659
         Size            =   "7056;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM38 
         Height          =   300
         Left            =   1350
         TabIndex        =   74
         Top             =   390
         Width           =   4005
         VariousPropertyBits=   679493659
         MaxLength       =   30
         Size            =   "7056;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM85 
         Height          =   285
         Left            =   -73410
         TabIndex        =   55
         Top             =   3534
         Width           =   6972
         VariousPropertyBits=   679493659
         MaxLength       =   100
         Size            =   "12298;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM84 
         Height          =   285
         Left            =   -73410
         TabIndex        =   52
         Top             =   2718
         Width           =   6972
         VariousPropertyBits=   679493659
         MaxLength       =   100
         Size            =   "12298;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM83 
         Height          =   285
         Left            =   -73410
         TabIndex        =   49
         Top             =   1902
         Width           =   6972
         VariousPropertyBits=   679493659
         MaxLength       =   100
         Size            =   "12298;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM82 
         Height          =   285
         Left            =   -73410
         TabIndex        =   46
         Top             =   1086
         Width           =   6972
         VariousPropertyBits=   679493659
         MaxLength       =   100
         Size            =   "12298;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM05S 
         Height          =   792
         Left            =   -73635
         TabIndex        =   80
         TabStop         =   0   'False
         Top             =   3840
         Width           =   7400
         VariousPropertyBits=   -1467989993
         ForeColor       =   -2147483641
         ScrollBars      =   2
         Size            =   "13053;1397"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM24 
         Height          =   285
         Left            =   -73410
         TabIndex        =   43
         Top             =   270
         Width           =   6972
         VariousPropertyBits=   679493659
         MaxLength       =   100
         Size            =   "12298;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM58 
         Height          =   800
         Left            =   -73635
         TabIndex        =   70
         Top             =   1590
         Width           =   7400
         VariousPropertyBits=   -1466941413
         MaxLength       =   2000
         Size            =   "13053;1411"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCP64 
         Height          =   1035
         Left            =   -73635
         TabIndex        =   71
         Top             =   2700
         Width           =   7400
         VariousPropertyBits=   -1466941413
         MaxLength       =   2000
         Size            =   "13053;1826"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM07 
         Height          =   285
         Left            =   -73590
         TabIndex        =   30
         Top             =   2650
         Width           =   6975
         VariousPropertyBits=   679493659
         MaxLength       =   40
         Size            =   "12303;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM05 
         Height          =   285
         Left            =   -73590
         TabIndex        =   28
         Top             =   2130
         Width           =   6975
         VariousPropertyBits=   679493659
         Size            =   "12303;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM05_1 
         Height          =   792
         Left            =   -73935
         TabIndex        =   27
         Top             =   2130
         Width           =   7580
         VariousPropertyBits=   679493659
         MaxLength       =   140
         ScrollBars      =   2
         Size            =   "13370;1397"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCP44_2 
         Height          =   264
         Left            =   -68700
         TabIndex        =   162
         TabStop         =   0   'False
         Top             =   900
         Width           =   2450
         VariousPropertyBits=   679493655
         Size            =   "8555;529"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM80_2 
         Height          =   264
         Left            =   -72990
         TabIndex        =   139
         TabStop         =   0   'False
         Top             =   3480
         Width           =   2175
         VariousPropertyBits=   679493663
         ForeColor       =   -2147483641
         Size            =   "3836;466"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM81_2 
         Height          =   264
         Left            =   -68550
         TabIndex        =   138
         TabStop         =   0   'False
         Top             =   3480
         Width           =   1935
         VariousPropertyBits=   679493663
         ForeColor       =   -2147483641
         Size            =   "3413;466"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textSP59_2 
         Height          =   264
         Left            =   -68550
         TabIndex        =   128
         TabStop         =   0   'False
         Top             =   3210
         Width           =   1935
         VariousPropertyBits=   679493663
         ForeColor       =   -2147483641
         Size            =   "3413;466"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textSP58_2 
         Height          =   264
         Left            =   -72990
         TabIndex        =   126
         TabStop         =   0   'False
         Top             =   3210
         Width           =   2175
         VariousPropertyBits=   679493663
         ForeColor       =   -2147483641
         Size            =   "3836;466"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM44_2 
         Height          =   264
         Left            =   -68550
         TabIndex        =   115
         TabStop         =   0   'False
         Top             =   2940
         Width           =   1935
         VariousPropertyBits=   679493663
         Size            =   "3413;466"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM23_2 
         Height          =   264
         Left            =   -72990
         TabIndex        =   113
         TabStop         =   0   'False
         Top             =   2940
         Width           =   2175
         VariousPropertyBits=   679493663
         ForeColor       =   -2147483641
         Size            =   "3836;466"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCP13_2 
         Height          =   264
         Left            =   -70140
         TabIndex        =   102
         TabStop         =   0   'False
         Top             =   540
         Width           =   1095
         VariousPropertyBits=   679493663
         MaxLength       =   20
         Size            =   "1931;466"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCP14_2 
         Height          =   285
         Left            =   -73110
         TabIndex        =   95
         TabStop         =   0   'False
         Top             =   270
         Width           =   1095
         VariousPropertyBits=   679493663
         MaxLength       =   20
         Size            =   "1931;503"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label70 
         Caption         =   "特殊商標 :"
         Height          =   255
         Left            =   -70110
         TabIndex        =   178
         Top             =   4710
         Width           =   855
      End
      Begin VB.Label Label63 
         AutoSize        =   -1  'True
         Caption         =   "聯絡人２(日)："
         Height          =   180
         Left            =   120
         TabIndex        =   175
         Top             =   1890
         Width           =   1200
      End
      Begin VB.Label Label62 
         AutoSize        =   -1  'True
         Caption         =   "聯絡人２(英)："
         Height          =   180
         Left            =   120
         TabIndex        =   174
         Top             =   1590
         Width           =   1200
      End
      Begin VB.Label Label61 
         AutoSize        =   -1  'True
         Caption         =   "聯絡人１(日)："
         Height          =   180
         Left            =   120
         TabIndex        =   173
         Top             =   990
         Width           =   1200
      End
      Begin VB.Label Label60 
         AutoSize        =   -1  'True
         Caption         =   "聯絡人１(英)："
         Height          =   180
         Left            =   120
         TabIndex        =   172
         Top             =   690
         Width           =   1200
      End
      Begin VB.Label Label59 
         AutoSize        =   -1  'True
         Caption         =   "聯絡人２(中)："
         Height          =   180
         Left            =   120
         TabIndex        =   171
         Top             =   1320
         Width           =   1200
      End
      Begin VB.Label Label53 
         AutoSize        =   -1  'True
         Caption         =   "聯絡人１(中)："
         Height          =   180
         Left            =   120
         TabIndex        =   170
         Top             =   390
         Width           =   1200
      End
      Begin VB.Label lblTM130 
         AutoSize        =   -1  'True
         Caption         =   "特殊出名公司:          (J:智權公司 空白:系統預設)"
         Height          =   180
         Left            =   -71580
         TabIndex        =   169
         Top             =   1305
         Width           =   3690
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "案件備註欄：不可銷卷案請加註 ""不銷卷"" 字樣！  與他案合併計算結餘請註明""與某案號合併計算結餘""！"
         ForeColor       =   &H000000FF&
         Height          =   180
         Index           =   115
         Left            =   -74880
         TabIndex        =   168
         Top             =   2480
         Width           =   8220
      End
      Begin VB.Label Label56 
         AutoSize        =   -1  'True
         Caption         =   "審定公告日 :"
         Height          =   180
         Left            =   -72100
         TabIndex        =   164
         Top             =   1656
         Width           =   996
      End
      Begin VB.Label Label54 
         Caption         =   "CF代理人 :"
         Height          =   255
         Left            =   -70740
         TabIndex        =   163
         Top             =   930
         Width           =   855
      End
      Begin VB.Label Label52 
         AutoSize        =   -1  'True
         Caption         =   "商品組群 :"
         Height          =   180
         Left            =   -74880
         TabIndex        =   161
         Top             =   630
         Width           =   810
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         Caption         =   "商品類別 :"
         Height          =   180
         Left            =   -74880
         TabIndex        =   160
         Top             =   330
         Width           =   810
      End
      Begin VB.Label Label51 
         AutoSize        =   -1  'True
         Caption         =   "申請地址5(中) :"
         Height          =   180
         Left            =   -74880
         TabIndex        =   159
         Top             =   3564
         Width           =   1200
      End
      Begin VB.Label Label50 
         AutoSize        =   -1  'True
         Caption         =   "申請地址5(英) :"
         Height          =   180
         Left            =   -74880
         TabIndex        =   158
         Top             =   3828
         Width           =   1200
      End
      Begin VB.Label Label49 
         AutoSize        =   -1  'True
         Caption         =   "申請地址5(日) :"
         Height          =   180
         Left            =   -74880
         TabIndex        =   157
         Top             =   4128
         Width           =   1200
      End
      Begin VB.Label Label48 
         AutoSize        =   -1  'True
         Caption         =   "申請地址4(中) :"
         Height          =   180
         Left            =   -74880
         TabIndex        =   156
         Top             =   2770
         Width           =   1200
      End
      Begin VB.Label Label47 
         AutoSize        =   -1  'True
         Caption         =   "申請地址4(英) :"
         Height          =   180
         Left            =   -74880
         TabIndex        =   155
         Top             =   3045
         Width           =   1200
      End
      Begin VB.Label Label46 
         AutoSize        =   -1  'True
         Caption         =   "申請地址4(日) :"
         Height          =   180
         Left            =   -74880
         TabIndex        =   154
         Top             =   3288
         Width           =   1200
      End
      Begin VB.Label Label45 
         AutoSize        =   -1  'True
         Caption         =   "申請地址3(中) :"
         Height          =   180
         Left            =   -74880
         TabIndex        =   153
         Top             =   1954
         Width           =   1200
      End
      Begin VB.Label Label44 
         AutoSize        =   -1  'True
         Caption         =   "申請地址3(英) :"
         Height          =   180
         Left            =   -74880
         TabIndex        =   152
         Top             =   2229
         Width           =   1200
      End
      Begin VB.Label Label43 
         AutoSize        =   -1  'True
         Caption         =   "申請地址3(日) :"
         Height          =   180
         Left            =   -74880
         TabIndex        =   151
         Top             =   2496
         Width           =   1200
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "申請地址2(中) :"
         Height          =   180
         Left            =   -74880
         TabIndex        =   150
         Top             =   1138
         Width           =   1200
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "申請地址2(英) :"
         Height          =   180
         Left            =   -74880
         TabIndex        =   149
         Top             =   1413
         Width           =   1200
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "申請地址2(日) :"
         Height          =   180
         Left            =   -74880
         TabIndex        =   148
         Top             =   1680
         Width           =   1200
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "查名本所案號 :"
         Height          =   180
         Left            =   -74880
         TabIndex        =   147
         Top             =   930
         Width           =   1170
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         Caption         =   "優先權資料 :"
         Height          =   180
         Left            =   -74880
         TabIndex        =   146
         Top             =   1260
         Width           =   990
      End
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         Caption         =   "案件備註 :"
         Height          =   180
         Left            =   -74880
         TabIndex        =   145
         Top             =   1590
         Width           =   810
      End
      Begin VB.Label Label32 
         AutoSize        =   -1  'True
         Caption         =   "進度備註 :"
         Height          =   180
         Left            =   -74880
         TabIndex        =   144
         Top             =   2760
         Width           =   810
      End
      Begin VB.Line Line1 
         X1              =   -72300
         X2              =   -72180
         Y1              =   4830
         Y2              =   4830
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         Caption         =   "商標名稱 :"
         Height          =   180
         Left            =   -74880
         TabIndex        =   143
         Top             =   3840
         Width           =   810
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         Caption         =   "商標專用期限 :"
         Height          =   180
         Left            =   -74880
         TabIndex        =   142
         Top             =   4710
         Width           =   1170
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "商標審定號 :"
         Height          =   180
         Left            =   -74880
         TabIndex        =   124
         Top             =   4653
         Width           =   990
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "分所案號 :"
         Height          =   180
         Left            =   -74880
         TabIndex        =   123
         Top             =   4395
         Width           =   810
      End
      Begin VB.Label Label38 
         Caption         =   "客戶案件案號："
         Height          =   180
         Left            =   -70980
         TabIndex        =   134
         Top             =   4395
         Width           =   1335
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "申請地址1(日) :"
         Height          =   180
         Left            =   -74880
         TabIndex        =   121
         Top             =   864
         Width           =   1200
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "申請地址1(英) :"
         Height          =   180
         Left            =   -74880
         TabIndex        =   122
         Top             =   597
         Width           =   1200
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         Caption         =   "申請人4 :"
         Height          =   180
         Left            =   -74910
         TabIndex        =   141
         Top             =   3525
         Width           =   720
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         Caption         =   "申請人5 :"
         Height          =   180
         Left            =   -70560
         TabIndex        =   140
         Top             =   3510
         Width           =   720
      End
      Begin VB.Label lblDivCase 
         AutoSize        =   -1  'True
         Caption         =   "分割母案本所案號:"
         Height          =   180
         Left            =   -69750
         TabIndex        =   137
         Top             =   1905
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "申請地址1(中) :"
         Height          =   180
         Left            =   -74880
         TabIndex        =   120
         Top             =   322
         Width           =   1200
      End
      Begin VB.Label Label42 
         Caption         =   "案件名稱 :"
         Height          =   255
         Left            =   -74910
         TabIndex        =   136
         Top             =   2160
         Width           =   1215
      End
      Begin VB.Label Label37 
         Caption         =   "收文日 :"
         Height          =   255
         Left            =   -70980
         TabIndex        =   132
         Top             =   1380
         Width           =   855
      End
      Begin VB.Label Label36 
         Caption         =   "轉本所案號 :"
         Height          =   255
         Left            =   -74910
         TabIndex        =   131
         Top             =   1380
         Width           =   1095
      End
      Begin VB.Label Label35 
         Caption         =   "取消收文日 :"
         Height          =   255
         Left            =   -68250
         TabIndex        =   129
         Top             =   1380
         Width           =   1095
      End
      Begin VB.Label Label34 
         Caption         =   "申請人3 :"
         Height          =   255
         Left            =   -70560
         TabIndex        =   127
         Top             =   3240
         Width           =   855
      End
      Begin VB.Label Label33 
         Caption         =   "申請人2 :"
         Height          =   255
         Left            =   -74910
         TabIndex        =   125
         Top             =   3255
         Width           =   855
      End
      Begin VB.Label Label40 
         Caption         =   "本案期限："
         Height          =   255
         Left            =   -74910
         TabIndex        =   118
         Top             =   4110
         Width           =   975
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "彼所案號："
         Height          =   180
         Left            =   -74910
         TabIndex        =   117
         Top             =   3780
         Width           =   900
      End
      Begin VB.Label Label55 
         Caption         =   "FC代理人 :"
         Height          =   255
         Left            =   -70560
         TabIndex        =   116
         Top             =   2940
         Width           =   855
      End
      Begin VB.Label Label14 
         Caption         =   "申請人1:"
         Height          =   255
         Left            =   -74910
         TabIndex        =   114
         Top             =   2970
         Width           =   735
      End
      Begin VB.Label Label12 
         Caption         =   "案件英文名稱 :"
         Height          =   255
         Left            =   -74910
         TabIndex        =   112
         Top             =   2460
         Width           =   1215
      End
      Begin VB.Label Label11 
         Caption         =   "案件日文名稱 :"
         Height          =   255
         Left            =   -74910
         TabIndex        =   111
         Top             =   2745
         Width           =   1215
      End
      Begin VB.Label Label13 
         Caption         =   "案件中文名稱 :"
         Height          =   255
         Left            =   -74910
         TabIndex        =   110
         Top             =   2175
         Width           =   1215
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "是否取消閉卷 :             (Y:取消)"
         Height          =   180
         Left            =   -74910
         TabIndex        =   109
         Top             =   1920
         Width           =   2400
      End
      Begin VB.Label Label10 
         Caption         =   "相關總收文號 :"
         Height          =   255
         Left            =   -69750
         TabIndex        =   108
         Top             =   1620
         Width           =   1335
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "是否算案件數 :             (N:不算)"
         Height          =   180
         Left            =   -74910
         TabIndex        =   107
         Top             =   1635
         Width           =   2400
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "卷宗性質 :          (1:申請 2:異議 3:評定 4:廢止)"
         Height          =   180
         Left            =   -71880
         TabIndex        =   106
         Top             =   1080
         Width           =   3615
      End
      Begin VB.Label Label6 
         Caption         =   "商標種類 :"
         Height          =   255
         Left            =   -74910
         TabIndex        =   103
         Top             =   1095
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "智權人員 :"
         Height          =   255
         Index           =   1
         Left            =   -71880
         TabIndex        =   101
         Top             =   540
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "申請國家 :"
         Height          =   255
         Index           =   8
         Left            =   -71880
         TabIndex        =   100
         Top             =   270
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "本所期限 :"
         Height          =   255
         Left            =   -74910
         TabIndex        =   99
         Top             =   810
         Width           =   975
      End
      Begin VB.Label Label25 
         Caption         =   "法定期限 :"
         Height          =   255
         Left            =   -71880
         TabIndex        =   98
         Top             =   810
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "案件性質 :"
         Height          =   255
         Left            =   -74910
         TabIndex        =   96
         Top             =   525
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "承辦人 :"
         Height          =   255
         Left            =   -74910
         TabIndex        =   94
         Top             =   270
         Width           =   855
      End
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "下一筆(&N)"
      Height          =   400
      Left            =   2880
      TabIndex        =   83
      Top             =   10
      Width           =   900
   End
   Begin VB.TextBox textTM29_2 
      BorderStyle     =   0  '沒有框線
      ForeColor       =   &H000000FF&
      Height          =   264
      Left            =   4650
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   133
      TabStop         =   0   'False
      Top             =   456
      Width           =   1000
   End
   Begin VB.CommandButton cmdCaseProgress 
      Caption         =   "案件進度(&C)"
      Height          =   400
      Left            =   3795
      TabIndex        =   84
      Top             =   10
      Width           =   1200
   End
   Begin VB.CommandButton cmdRelate 
      Caption         =   "相關卷號(&R)"
      Height          =   400
      Left            =   5040
      TabIndex        =   85
      Top             =   10
      Width           =   1200
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "回前畫面(&U)"
      Height          =   400
      Left            =   7080
      TabIndex        =   87
      Top             =   10
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Left            =   6240
      TabIndex        =   86
      Top             =   10
      Width           =   800
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   8295
      TabIndex        =   88
      Top             =   10
      Width           =   800
   End
   Begin VB.TextBox textTMKey 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   3390
      Locked          =   -1  'True
      TabIndex        =   93
      TabStop         =   0   'False
      Top             =   456
      Width           =   1200
   End
   Begin VB.TextBox textCP09 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   870
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   92
      TabStop         =   0   'False
      Top             =   456
      Width           =   1200
   End
   Begin VB.Label Label69 
      AutoSize        =   -1  'True
      Caption         =   "目前表單狀態："
      ForeColor       =   &H000000C0&
      Height          =   180
      Left            =   6120
      TabIndex        =   193
      Top             =   450
      Width           =   1230
   End
   Begin VB.Label Label39 
      Caption         =   "TS商品類別 不可 輸在""案件備註""欄!!!"
      ForeColor       =   &H000000FF&
      Height          =   405
      Left            =   30
      TabIndex        =   135
      Top             =   30
      Width           =   1750
   End
   Begin VB.Label Label2 
      Caption         =   "本所案號 :"
      Height          =   255
      Left            =   2490
      TabIndex        =   90
      Top             =   450
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "收文號 :"
      Height          =   252
      Index           =   0
      Left            =   120
      TabIndex        =   89
      Top             =   456
      Width           =   700
   End
End
Attribute VB_Name = "frm020101_02"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2025/10/23 TF基礎案號(TM06,TM07)改成可以輸入多筆(Table: TFBaseNo)，原本的輸入欄位直接刪除改成按鈕呼叫其他表單，若已有設定則按鈕設為綠色。
'Memo by Lydia 2022/05/30 (已檢查)整理frm880005改用寄信模組
'Memo by Amy 2021/12/21 Form2.0已修改 TextCP14_2(名).../TextTM05/TextTM07/TextTM05_1/TextTM58/TextCP64/textTm24(地址).../textTM38/textTM40/textTM41/textTM43/grdList
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/26 員工編號欄已修改
'Memo By Sindy 2010/8/5 日期欄已修改
Option Explicit

Dim m_TM01 As String
Dim m_TM02 As String
Dim m_TM03 As String
Dim m_TM04 As String
Dim m_TM77 As String '2011/6/10 add by sonia
Dim m_TM11 As String '2012/3/29 add by sonia
Dim m_TM56 As String 'Add By Sindy 2013/8/6

Dim m_CPKeyList() As String
Dim m_CPKeyCount As Integer
' 收文號
Dim m_CP09 As String
' 案件性質
Dim m_CP10 As String
' 是否新案件
Dim m_CP31 As String
' 國家代碼
Dim m_TM10 As String
' 卷宗性質
Dim m_TM28 As String
' 是否閉卷
Dim m_TM29 As String
Dim m_CP65 As String 'Add By Sindy 2010/8/6
'910626 Sieg 501
' 收據編號
Dim m_CP60 As String
' 申請人
Dim m_TM23 As String
'add by nickc 2006/12/14
Dim m_TM78 As String
Dim m_TM79 As String
Dim m_TM80 As String
Dim m_TM81 As String
Dim m_TM44 As String 'Added by Lydia 2024/06/13
'Added by Lydia 2023/03/15 檢查是否已經有商品及服務
Public ChkTG As Boolean

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
' 儲存國家的字串
Dim m_strCountry As String
Dim m_CurrSel As Integer
' 優先權畫面所使用的變數
Dim m_Pa(1 To 4) As String '本所案號
'Modify by Amy 2023/01/06 設為Public改變數
'Dim m_Priority(1 To 6) As String
Public m_Priority1 As String, m_Priority2 As String, m_Priority3 As String, m_Priority4 As String, m_Priority5 As String, m_Priority6 As String
'end 2023/01/06
'Add By Cheng 2002/06/12
Dim m_strCP06 As String '原本所期限
Dim m_strCP07 As String '原法定期限
Dim m_TM22 As String '專用期止日
'Add By Cheng 2002/08/22
'Mark by Lydia 2024/06/13
'Dim m_strCust1 As String '申請人1
'Dim m_strCust2 As String '申請人2
'Dim m_strCust3 As String '申請人3
''add by nickc 2006/12/14
'Dim m_strCust4 As String '申請人4
'Dim m_strCust5 As String '申請人5
'end  --- Mark by Lydia 2024/06/13
'控制是否讀過
'Dim Nick920224Bol As Boolean 'Removed by Morgan 2024/12/26 沒用了
'add by nickc 2005/03/17 加乘註記
Dim m_CP98 As String
Dim m_CP101 As String
Dim m_CP104 As String
'2008/11/12 add by sonia 相關總收文號的資料
Dim m_CP43CP08 As String
Dim m_CP43CP64 As String
Dim m_CP43CP110 As String
Dim t_CP43CP110 As String
'2008/11/12 END
Dim m_CP30 As String 'Add by Morgan 2011/4/22
Dim m_CP46 As String 'add by sonia  2018/11/20
Public rsAddrNotAlike As New ADODB.Recordset 'Add By Sindy 2011/7/8
Public m_AppAddr As String, m_Zipcode As String 'Add By Sindy 2011/7/8
Public m_AppAddrChange As Boolean
'Add By Sindy 2012/5/8
Dim m_EP06 As String, m_EP06DT As String, m_CP16 As String, m_CP122 As String
Dim m_CP48 As String
'2012/5/8 End
Dim m_CP27 As String 'Add By Sindy 2012/6/1
Dim m_CP31isYGetCP05 As String 'Add By Sindy 2014/1/29
Dim m_CP141 As String 'Added By Lydia 2015/11/24 收款後送件
Dim strChkCuAreaMail As String, strChkCuAreaMailTo As String 'Added by Lydia 2017/06/19 檢查收文智權人員和客戶智權人員不同業務區要發mail的內文和通知人員
Dim strUpdCusNo As String 'Add by Amy 2018/08/09 更新CU12/CU13之客戶編號
Public m_CP143 As String
'Modified by Lydia 2020/11/04 更名
'Dim p_cp143DT As String 'Added by Lydia 2018/12/10 查名是否齊備
Dim p_CP143DT As String '查名是否齊備 'Memo by Lydia 2022/11/04 改成取得查名齊備日(西元年月日)
Dim p_strCP143 As String 'Added by Lydia 2020/11/04 收文之查名齊備日(Y/N); 發現近日的商標申請案紙本有印查名齊備,但是到分案就不見了;是因為m_CP143的混淆,所以改變數名稱
'Added by Lydia 2020/05/20 法律所案源收文
Dim m_LOS15 As String '案源單號
Dim m_LOS01 As String '案源總收文號
Dim m_LOS02 As String 'Added by Lydia 2020/06/09 案源案件類型
Dim m_LOS07 As String '放棄日期
Public m_CP36 As String, m_CP21 As String 'Add By Sindy 2020/10/20
Dim cp() As String 'Add By Sindy 2021/3/23
Dim m_CP149 As String 'Add By Sindy 2022/4/27
'Add by Amy 2022/10/07
Public m_PrevForm As Form '前一畫面
Dim m_F0308 As String, m_F0309 As String, strUpdDate As String, strUpdTime As String, IsEConsultRec As Boolean
Dim stF0207_A6 As String 'Add by Amy 2022/10/19
Dim stF0309_Now As String 'Add by Amy 2022/10/20
Dim m_bolIsFirstKeyCP14 As Boolean 'Add by Amy 2022/11/03 北所第一次輸承辦人
Dim m_SalesST15 As String '畫面上智權人員的收文部門    'add by sonia 2023/11/7 從textCP13_Validate移上來
Dim strPTM As String, strSPT As String 'Added by Lydia 2023/11/16 暫存商標種類及特殊商標的Combo.ItemData
Dim strMsgCloseCancel As String 'Added by Lydia 2025/06/30 取消閉卷時，若下一程序有未過期且已上N之延展102、使用宣誓105期限，都自動取消NP06、NP11、NP12，並彈訊息提醒已還原XXX期限。

'Add by Amy 2022/10/18 補件完成
Private Sub CmdAddInfo_Click()
    
On Error GoTo ErrHand
    
    If txtNote = MsgText(601) Then
        MsgBox "呈報內容不可為空！", vbExclamation
        Exit Sub
    End If
    
    Screen.MousePointer = vbHourglass
    
    m_F0308 = "A6"
    m_F0309 = Flow_補件完成
    strUpdDate = strSrvDate(1)
    strUpdTime = Right("000000" & ServerTime, 6)
   
    cnnConnection.BeginTrans
    '簽核檔
    strSql = "update FLOW002 set " & _
               "F0205='" & strUpdDate & "'" & _
               ",F0206='" & strUpdTime & "'" & _
               ",F0207='5',F0204='" & strUserNum & "'" & _
               " where F0201='" & txtF0301 & "' and F0202='A7'  and F0207 is null "
    cnnConnection.Execute strSql
    
    '退回程序時再新增2筆待簽核的記錄
    Call SetConultRecPrePerson_Flow002(Me.Name, txtF0301, "A6") '商標主管
    Call SetConultRecPrePerson_Flow002(Me.Name, txtF0301, "A7") '商標程序
        
    '表單主檔
    'Modify by Amy 2022/10/19 原:F0307='" & strUserNum & "'
    strSql = "update FLOW003 set " & _
            "F0307='A7'" & _
            ",F0308='" & m_F0308 & "'" & _
            ",F0309='" & m_F0309 & "'" & _
            " where F0301='" & txtF0301 & "' "
    cnnConnection.Execute strSql
    
    strSql = GetInsertFLOW004Sql(txtF0301, strUserNum, strUpdDate, strUpdTime, m_F0309, ChgSQL(Trim(txtNote.Text)), "A7", "A6")
    cnnConnection.Execute strSql
    
    cnnConnection.CommitTrans
    
    Screen.MousePointer = vbDefault
    
    ClearAll
    If NextRecord = True Then
        QueryData
    Else
        Unload Me
        frm020101_01.QueryDB
        frm020101_01.Show
    End If
    Exit Sub
    
ErrHand:
    Screen.MousePointer = vbDefault
    cnnConnection.RollbackTrans
    MsgBox "補件失敗！" & vbCrLf & Err.Description
End Sub

Private Sub cmdCancel_Click()
   Unload Me
   frm020101_01.Show
End Sub

Private Sub cmdCaseProgress_Click()
   frm020101_03.SetData 0, m_TM01, True
   frm020101_03.SetData 1, m_TM02, False
   frm020101_03.SetData 2, m_TM03, False
   frm020101_03.SetData 3, m_TM04, False
   frm020101_03.SetData 4, m_CP09, False
   frm020101_03.SetParent Me
   Me.Hide
   frm020101_03.QueryData
   frm020101_03.Show
End Sub

Private Sub cmdExit_Click()
   Unload frm020101_01
   Unload Me
End Sub

'Add by Amy 2022/10/07 檢視接洽單
Private Sub cmdFile_Click()
    frm090801_Q.SetParent Me
    frm090801_Q.m_blnCallPrint = True
    frm090801_Q.Text5 = txtF0301
    Call frm090801_Q.cmdok_Click(4)
    frm090801_Q.Show 'Add by Amy 2022/11/17
End Sub

Private Sub cmdNation_Click()
'Removed by Morgan 2024/12/26 進畫面後m_strCountry就已設定，此處重讀國家會重復，若進子畫面又沒按確定的話，存檔後子案就會重複，Ex:TF-000940
'    '920224 nick 新增
'    If Nick920224Bol = False Then
'   Dim nick920224rs As New ADODB.Recordset
'   Set nick920224rs = New ADODB.Recordset
'   Dim nick920224str As String
'   nick920224str = "select DISTINCT(tm10) from trademark where tm01='" & m_TM01 & "' and tm02='" & m_TM02 & "' and tm03>'0'"
'   With nick920224rs
'        .CursorLocation = adUseClient
'        .Open nick920224str, cnnConnection, adOpenStatic, adLockReadOnly
'        If Not .EOF And Not .BOF Then
'            .MoveFirst
'            Do While Not .EOF
'                If Trim(m_strCountry) = "" Then
'                    m_strCountry = m_strCountry & CheckStr(.Fields(0).Value)
'                Else
'                    m_strCountry = m_strCountry & "," & CheckStr(.Fields(0).Value)
'                End If
'                .MoveNext
'            Loop
'        End If
'   End With
'    Nick920224Bol = True
'   End If
'end 2024/12/26
   
   '2012/3/29 MODIFY BY SONIA
   'ModifyAssignCountry m_strCountry  textCP10
   '2012/11/8 MODIFY BY SONIA 領土延伸案以收文日判斷會員國 TF-000122-0-00
   'ModifyAssignCountry m_strCountry, TransDate(m_TM11, 2)
   If textCP10 = "104" Then
      ModifyAssignCountry m_strCountry, TransDate(textCP05, 2)
   Else
      ModifyAssignCountry m_strCountry, TransDate(m_TM11, 2)
   End If
   '2012/11/8 END
End Sub

Private Sub cmdNext_Click()
   ClearAll
   If NextRecord = True Then
      'Add by Amy 2022/11/17
      If PUB_CheckFormExist("frm090801_Q") = True Then
            Unload frm090801_Q
      End If
      QueryData
   Else
      Unload Me
      frm020101_01.Show
   End If
End Sub

Private Sub cmdok_Click()
   Dim strTM01 As String
   Dim strTM02 As String
   Dim strTM03 As String
   Dim strTM04 As String
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   'Add By Sindy 2012/5/8
   Dim strSubject As String
   Dim strContent As String
   '2012/5/8 End
   Dim bolHadPoMsg As Boolean
   Dim m_StrTo As String, m_StrSub As String, m_StrCont As String 'Added by Lydia 2022/05/30 整理frm880005改用寄信模組
   
   'Added by Lydia 2019/07/29 若原本有文件齊備,拿掉Y彈提醒(因為T-213580延期AA8033478,發生文件齊備被拿掉,但是人員不記得操作)
   If textEP06.Tag = "Y" And textEP06.Tag <> textEP06.Text Then
       If MsgBox(Left(Label64.Caption, 2) & "齊備被拿掉，請問是否繼續？", vbCritical + vbYesNo + vbDefaultButton2) = vbNo Then
           textEP06.SetFocus
           Exit Sub
       End If
   End If
   
   If CheckDataValid = True Then
      'Add By Cheng 2002/05/23
      '重新檢查欄位有效性
      If TxtValidate = False Then Exit Sub
'      'Add By Cheng 2002/08/23
'      If Me.textTM01.Text <> "" And Me.textTM02.Text <> "" Then
'         MsgBox "本次只更新轉本所案號的資料，不更新其他欄位的修改!!!", vbCritical
'      End If
      '910626 Sieg
      If textTM01 <> "" And textTM02 <> "" Then
         strExc(1) = textTM01
         strExc(2) = textTM02
         If strExc(1) = "TF" Then
            If textTM02_2 = "" Then
               strExc(2) = strExc(2) & "0"
            Else
               strExc(2) = strExc(2) & textTM02_2
            End If
         End If
         strExc(3) = textTM03
         If strExc(3) = "" Then strExc(3) = "0"
         strExc(4) = textTM04
         If strExc(4) = "" Then strExc(4) = "00"
         strExc(5) = textCP10 '案件性質
         strExc(6) = textCP10_2 '案件性質名稱
         strExc(7) = textCP05 '收文日
         strExc(8) = textCP09 '總收文號
         '911118 nick 新增申請人
         strExc(9) = ChangeCustomerL(m_TM23)
         'edit by nickc 2007/02/06 不用 dll 了
         'If Not objLawDll.ChkSameCase(strExc) Then Exit Sub
         If Not ClsLawChkSameCase(strExc) Then Exit Sub
         'Added by Lydia 2020/08/18 更新相關卷號前,先檢查是否有重複
         If m_CP31 = "Y" Then
            If PUB_ChkUpdCR(m_TM01, m_TM02, m_TM03, m_TM04, strExc(1), strExc(2), strExc(3), strExc(4)) = False Then
                Exit Sub
            End If
         End If
         'end 2020/08/18
      End If
      strTM01 = m_TM01
      strTM02 = m_TM02
      strTM03 = m_TM03
      strTM04 = m_TM04
      'Add By Cheng 2002/11/18
      If Me.textTM01.Text <> "" And Me.textTM02.Text <> "" Then
         MsgBox "本次只更新轉本所案號的資料，不更新其他欄位的修改!!!", vbCritical
      'Added by Lydia 2023/12/14
      Else
        '檢查智財協作在分案時若未建立相關案號(caserelation1)時則跳提醒程序人員，但可選擇輸或不輸 !
         'Modified by Lydia 2023/12/15 PS及CPS之智財協作967，TT及S之智財協作737，L之智財協作7601，(也可用案件性質中文判斷)在分案時若未建立相關案號且為ACS且為TIPS的案件時，提醒文字：「案件性質為智財協作，請先依接洽單輸入相關卷號資料」。
'         If m_TM01 = "TT" And textCP10 = "737" Then
'            If PUB_IfCaseRelation1Exists(m_TM01, m_TM02, m_TM03, m_TM04) = False Then
'               If MsgBox("案件性質為" & textCP10_2 & "，請確認接洽單是否有相關案號，是否補輸入？", vbExclamation + vbYesNo + vbDefaultButton1) = vbYes Then
'                  Exit Sub
'               End If
'            End If
'         End If
         If m_TM01 = "TT" And InStr(textCP10_2, "智財協作") > 0 Then
            If PUB_ChkACSforTIPS(m_TM01 & m_TM02 & m_TM03 & m_TM04, , True) = False Then
               MsgBox "案件性質為" & textCP10_2 & "，請先依接洽單輸入相關卷號資料", vbExclamation
               Exit Sub
            End If
         End If
         'end 2023/12/15
      'end 2023/12/14
      End If
      
      'Added by Lydia 2017/06/19 (存檔前)檢查收文智權人員和客戶智權人員不同業務區要發mail的內文和通知人員
      'Modified by Lydia 2021/12/24  傳入分案畫面的智權人員; 因為分案要在存檔前才能檢查出要通知的範圍,所以另外傳入畫面的智權人員; ex.TS-001858
      'strChkCuAreaMail = PUB_ChkSameCustSales(m_TM01, m_TM02, m_TM03, m_TM04, textCP09, Trim(textTM23), Trim(textSP58), Trim(textSP59), Trim(textTM80), Trim(textTM81), strChkCuAreaMailTo)
      strChkCuAreaMail = PUB_ChkSameCustSales(m_TM01, m_TM02, m_TM03, m_TM04, textCP09, Trim(textTM23), Trim(textSP58), Trim(textSP59), Trim(textTM80), Trim(textTM81), strChkCuAreaMailTo, , textCP13.Text)
      'Added by Lydia 2020/01/21 所有系統別分案，若該收文號存在於下一程序(NP24)，若修改案件性質則彈訊息
      If textCP10 <> m_CP10 Then
          If Pub_CheckNP24Exists(textCP09.Text) = True Then
          End If
      End If
      'end 2020/01/21
      
      'Add By Sindy 2020/10/20
      'Modify By Sindy 2021/3/23 +
      If textCP10 = "210" And PUB_ChkCPExist(cp, "214") = False Then '214.陳述聲明
         Set frm020101_04.oNextForm = Me
         frm020101_04.m_TM01 = m_TM01
         frm020101_04.m_TM02 = m_TM02
         frm020101_04.m_TM03 = m_TM03
         frm020101_04.m_TM04 = m_TM04
         frm020101_04.m_TM10 = textCP10
         frm020101_04.textCP143 = m_CP143
         frm020101_04.textCP36 = m_CP36
         frm020101_04.textCP21 = m_CP21
         frm020101_04.Show vbModal
      End If
      '2020/10/20 END
      
     'Added by Lydia 2020/06/19 法律所案源收文：C類案源的案件性質若 "是否需要法律所配合"設定與來不同時提醒。
     'Modified by Lydia 2020/8/03 FCT商爭案由內商負責 +FCT
     If strSrvDate(1) >= 法律所案源收文啟用日 And (m_TM01 = "T" Or m_TM01 = "TC" Or m_TM01 = "FCT") Then
          strExc(1) = "" 'Added by Lydia 2021/07/12 清空預設
          'Modified by Lydia 2020/07/23 重新整理: 因為案源收文已設定不可變更案件性質和申請國家,所以只要判斷非案源收文
          If m_LOS02 = "" And m_CP10 <> textCP10 Then
             strExc(1) = PUB_GetLOSkind(m_TM01, textCP10, textTM10)
             strExc(1) = Replace(strExc(1), "T", "")
             '準備程序在輸入接洽單已決定是否為案源的補收款, 所以不用另外判斷
             If strExc(1) <> "" Then
                   MsgBox "收文不可修改為法務案源的案件性質！", vbCritical
                   Exit Sub
             End If
          End If
          'end 2020/07/23
          
          If m_LOS01 = "" And m_LOS07 = "" And FraLOS.Visible = True Then
              If (Left(strExc(1), 1) = "C" And m_LOS15 = "" And txtLOSagree = "Y") Or (Left(strExc(1), 1) = "C" And m_LOS15 <> "" And txtLOSagree <> "Y") _
                    Or (strExc(1) = "" And Left(m_LOS02, 1) = "C" And m_LOS15 <> "" And txtLOSagree <> "Y") Then
                   If MsgBox(" ""是否需要法律所配合"" 與接洽單不同，是否繼續作業？", vbExclamation + vbYesNo + vbDefaultButton2) = vbNo Then
                       txtLOSagree.SetFocus
                       txtLOSagree_GotFocus
                       Exit Sub
                   End If
              End If
          End If
     End If
     'end 2020/06/19
   
      'Add By Sindy 2021/3/9 增加控管轉本所案號時,不檢查接洽單PDF檔
      'Modify By Sindy 2022/3/33 增加判斷沒有輸入承辦人時,先不檢查接洽單PDF檔
      If (textTM01 <> "" And textTM02 <> "") Or _
         Trim(textCP14) = "" Then
         '不需檢查接洽單PDF檔
      Else
      '2021/3/9 END
         'Add By Sindy 2019/7/18
         If textCP09 < "B" Then
            'Modify By Sindy 2022/12/16 電子收文不用檢查
            If Not (txtF0301 <> "" And Len(txtF0301) = 10) Then
            '2022/12/16 END
               If PUB_CheckPDF2(textCP09, 0, True, strExc(0), textCP10, , bolHadPoMsg) = False Then
                  If bolHadPoMsg = False Then
                     MsgBox "無接洽單PDF檔,不可分案!", vbCritical
                  End If
                  Exit Sub
               End If
            End If
         End If
         '2019/7/18
      End If
   
      ' 設定滑鼠游標為等待狀態
      Screen.MousePointer = vbHourglass
      ' 儲存資料
      OnUpdateField
    'Modify By Cheng 2002/11/06
'      'OnSaveData
      If OnSaveData = False Then MsgBox "存檔失敗，請洽系統管理員 !", vbCritical: Screen.MousePointer = vbDefault: Exit Sub
      
      'Added by Lydia 2017/06/19 檢查收文智權人員和客戶智權人員不同業務區要發mail的內文和通知人員
      If strChkCuAreaMail <> "" Then
           'Modified by Lydia 2021/12/24 改主旨「案件收文通知」=>「分案通知」
           PUB_SendMail strUserNum, strChkCuAreaMailTo, "", "分案通知--此案收文非原智權人員(區)！", strChkCuAreaMail
      End If
      'end 2017/06/19
      
      'Add By Sindy 2012/5/8 台灣商標爭議案E-Mail通知承辦人承辦期限
      'Modified by Lydai 2018/12/10 判斷商爭案
      'If Frame21.Visible = True Then
      'Modify By Sindy 2025/7/30 FCT案727分析不屬於爭議案 + FCT_NotTMdebate
      If Frame21.Visible = True And InStr(TMdebate, textCP10) > 0 And Not (m_TM01 = "FCT" And InStr(FCT_NotTMdebate, textCP10) > 0) Then
         'Modify by Amy 2022/11/03 畫面上有承辦人,且為第一次分案或承辦人異動
         'If (textCP14.Tag <> textCP14 And Trim(textCP14) <> "") And
         If Trim(textCP14) <> "" And ((textCP14.Tag <> textCP14) Or m_bolIsFirstKeyCP14 = True) And _
            textEP06 = "Y" Then
            If m_TM03 = "0" And m_TM04 = "00" Then
               'Modified by Lydia 2022/07/15 去掉"台灣"
               strSubject = m_TM01 & "-" & m_TM02 & " 商標爭議案承辦期限通知"
            Else
               'Modified by Lydia 2022/07/15 去掉"台灣"
               strSubject = m_TM01 & "-" & m_TM02 & "-" & m_TM03 & "-" & m_TM04 & " 商標爭議案承辦期限通知"
            End If
            'Modify By Sindy 2023/12/11 +指定送件日期
            strContent = "本所案號：" + m_TM01 & "-" & m_TM02 & "-" & m_TM03 & "-" & m_TM04 + vbCrLf + _
                         "案件名稱：" + textTM05_1 + vbCrLf + _
                         "案件性質：" + textCP10_2 + vbCrLf + _
                         "收文日　：" + ChangeWStringToTDateString(DBDATE(textCP05)) + vbCrLf + _
                         "承辦期限：" + ChangeWStringToTDateString(DBDATE(m_CP48)) + vbCrLf + _
                         IIf(Trim(textCP142) <> "", "指定送件日期：" + ChangeWStringToTDateString(DBDATE(textCP142)) + IIf(Option1(0).Value = True, "當天", IIf(Option1(1).Value = True, "之後", IIf(Option1(2).Value = True, "之後", ""))) + vbCrLf, "") + _
                         "本所期限：" + ChangeWStringToTDateString(DBDATE(textCP06)) + vbCrLf + _
                         "法定期限：" + ChangeWStringToTDateString(DBDATE(textCP07)) + vbCrLf + _
                         "齊備日　：" + ChangeWStringToTDateString(DBDATE(m_EP06DT)) + vbCrLf + _
                         "是否急件：" + textCP122
            PUB_SendMail strUserNum, textCP14, "", strSubject, strContent, ""
         End If
      End If
      '2012/5/8 End
      
      'add by Toni 2008/10/27
      '2008/11/25 modify by sonia 無期限不發
      'If textCP10 = "204" Or textCP10 = "205" Then
      'Modify By Sindy 2023/3/28 控管台灣的才發Mail ex:TF-000870-1-06
      If (textCP10 = "204" Or textCP10 = "205") And textCP06.Text <> "" And textCP07.Text <> "" _
         And textTM10 = "000" Then
         'Modify by Amy 2022/11/03 +第一次分案
         If m_bolIsFirstKeyCP14 = True Or (textCP14.Text <> textCP14.Tag) Or (textCP06.Text <> textCP06.Tag) Or (Val(textCP07.Text) <> Val(textCP07.Tag)) Then
            '2008/11/12 ADD BY SONIA
            strSql = "select C1.CP08,C1.CP64,C2.CP110 from CASEPROGRESS C1,CASEPROGRESS C2 where C1.CP09='" & textCP43 & "' AND C1.CP43=C2.CP09(+)"
            CheckOC
            adoRecordset.CursorLocation = adUseClient
            adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
            If adoRecordset.RecordCount > 0 Then
               m_CP43CP08 = CheckStr(adoRecordset.Fields(0))
               m_CP43CP64 = CheckStr(adoRecordset.Fields(1))
               m_CP43CP110 = CheckStr(adoRecordset.Fields(2))
               If m_CP43CP110 <> "" Then
                  strSql = "select st01,st02,OA03 from staff,ouragent where instr('" & m_CP43CP110 & "',st01)>0 and oa01(+)='" & m_TM01 & "' and oa02(+)=st01 order by 3 , 1 "
                  CheckOC
                  adoRecordset.CursorLocation = adUseClient
                  adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
                  If adoRecordset.RecordCount > 0 Then
                     Do While Not adoRecordset.EOF
                        If InStr(m_CP43CP110, CheckStr(adoRecordset.Fields(0))) > 0 Then
                           t_CP43CP110 = t_CP43CP110 & CheckStr(adoRecordset.Fields(1)) & "、"
                        End If
                        adoRecordset.MoveNext
                     Loop
                  End If
                  CheckOC
                  If Right(t_CP43CP110, 1) = "、" Then
                     t_CP43CP110 = Mid(t_CP43CP110, 1, Len(t_CP43CP110) - 1)
                  End If
               End If
            End If
            '2008/11/12 END
            
            'Modified by Lydia 2022/05/30 整理frm880005改用寄信模組
            'Load frm880005
            ''2008/11/7 modify by sonia
            ''frm880005.txtEmail(0).Text = Pub_GetSpecMan("Q") & ";" & TextCP13
            ''Modify By Sindy 2012/8/16 開庭通知發mail對象,若為FCT案件再增加Pub_GetSpecMan("Q1")
            'frm880005.txtEmail(0).Text = Pub_GetSpecMan("Q") & ";" & _
                  IIf(m_TM01 = "FCT", PUB_GetFCTSalesNo(m_TM01, m_TM02, m_TM03, m_TM04), PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04)) & _
                  IIf(m_TM01 = "FCT", ";" & Pub_GetSpecMan("Q1"), "")
            '
            'frm880005.txtEmail(1).Text = "開庭通知--分案案號：" & m_TM01 & "-" & m_TM02 & "-" & m_TM03 & "-" & m_TM04
            'frm880005.txtEmail(2).Text = "本所案號：" & m_TM01 & "-" & m_TM02 & "-" & m_TM03 & "-" & m_TM04 & vbCrLf & _
                                          "案件名稱：" & textTM05_1 & vbCrLf & _
                                          "案件性質：" & textCP10_2.Text & vbCrLf & _
                                          "申請人　：" & textTM23_2.Text & vbCrLf & _
                                          "承辦人　：" & textCP14_2.Text & vbCrLf & _
                                          "智權人員：" & textCP13_2.Text & vbCrLf & _
                                          "法定期限：" & Val(Mid(DBDATE(textCP07.Text), 1, 4)) - 1911 & " 年 " & Mid(DBDATE(textCP07.Text), 5, 2) & " 月 " & Mid(DBDATE(textCP07.Text), 7, 2) & " 日 " & vbCrLf & _
                                          "時間地點：" & m_CP43CP64 & vbCrLf & _
                                          "法院案號：" & m_CP43CP08 & vbCrLf & _
                                          "律　　師：" & t_CP43CP110
            'frm880005.Form_Activate: DoEvents
            'frm880005.cmdOK_Click 0: DoEvents
            'Modified by Lydia 2022/08/15 加發承辦人 textCP14
            'Modify By Sindy 2023/12/8 法律所調整內專行政訴訟開庭通知之系統通知信也請一併轉陳亮之; 商標一併調整
            'Modified by Lydia 2024/10/30 串法律所案源資料，抓出法律所案件有承辦人且收文日最大的進度，抓承辦人及所有出庭律師。
            'm_StrTo = Pub_GetSpecMan("Q") & ";" & Pub_GetSpecMan("Q1") & ";" & _
            '      IIf(m_TM01 = "FCT", PUB_GetFCTSalesNo(m_TM01, m_TM02, m_TM03, m_TM04), PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04))
            '      'IIf(m_TM01 = "FCT", ";" & Pub_GetSpecMan("Q1"), "") & IIf(textCP14 <> "", ";" & textCP14, "")
            m_StrTo = PUB_GetLosCL02list(m_TM01, m_TM02, m_TM03, m_TM04)
            m_StrTo = IIf(m_StrTo <> "", m_StrTo & ";", "") & Pub_GetSpecMan("Q") & ";" & Pub_GetSpecMan("Q1") & ";" & _
                  IIf(m_TM01 = "FCT", PUB_GetFCTSalesNo(m_TM01, m_TM02, m_TM03, m_TM04), PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04))
            'end 2024/10/30
            
            m_StrSub = "開庭通知--分案案號：" & m_TM01 & "-" & m_TM02 & "-" & m_TM03 & "-" & m_TM04
            m_StrCont = "本所案號：" & m_TM01 & "-" & m_TM02 & "-" & m_TM03 & "-" & m_TM04 & vbCrLf & _
                                          "案件名稱：" & textTM05_1 & vbCrLf & _
                                          "案件性質：" & textCP10_2.Text & vbCrLf & _
                                          "申請人　：" & textTM23_2.Text & vbCrLf & _
                                          "承辦人　：" & textCP14_2.Text & vbCrLf & _
                                          "智權人員：" & textCP13_2.Text & vbCrLf & _
                                          "法定期限：" & Val(Mid(DBDATE(textCP07.Text), 1, 4)) - 1911 & " 年 " & Mid(DBDATE(textCP07.Text), 5, 2) & " 月 " & Mid(DBDATE(textCP07.Text), 7, 2) & " 日 " & vbCrLf & _
                                          "時間地點：" & m_CP43CP64 & vbCrLf & _
                                          "法院案號：" & m_CP43CP08 & vbCrLf & _
                                          "律　　師：" & t_CP43CP110
            PUB_SendMail strUserNum, m_StrTo, cp(9), m_StrSub, m_StrCont
            'end 2022/05/30
         End If
      End If
      'end 2008/10/27
      
      ' 設定滑鼠游標為預設
      Screen.MousePointer = vbDefault
      ' 當轉本所案號時檢查原本所案號是否還有案件進度的資料
      If IsEmptyText(textTM01) = False Then
         If IsCaseProgressExist(strTM01, strTM02, strTM03, strTM04) = False Then
            strTit = "檢核資料"
            strMsg = "原本所案號" & strTM01 & "-" & strTM02 & "-" & strTM03 & "-" & strTM04 & "已無案件進度資料，請通知收文人員刪號！"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         '2010/12/7 add by sonia
         Else
            MsgBox "原本所案號為 " & strTM01 & "-" & strTM02 & "-" & strTM03 & "-" & strTM04 & "，請自行更新原本所案號之下一程序資料 !", vbInformation
         '2010/12/7 end
         End If
      End If
      'Added by Lydia 2023/01/18 內外商分案之卷宗性質，由3或4改為1時，存檔後跳出商標基本資料頁面供使用者補資料
      If (m_TM01 = "T" Or m_TM01 = "TF" Or m_TM01 = "FCT") And InStr("3,4", m_TM28) > 0 And textTM28 = "1" Then
         ShowMaintainForm m_CP09, "N", "分案"
         MsgBox "請輸入專用期限！", vbInformation
      End If
      'end 2023/01/18
      'Added by Lydia 2025/06/30 取消閉卷時，若下一程序有未過期且已上N之延展102、使用宣誓105期限，都自動取消NP06、NP11、NP12，並彈訊息提醒已還原XXX期限。
      If strMsgCloseCancel <> "" Then
         MsgBox "已還原「" & strMsgCloseCancel & "」期限", vbInformation, "取消閉卷"
      End If
      If textTM29 = "Y" And m_TM01 = "TF" And m_TM10 = "238" Then
         MsgBox "請自行取消有效子案之閉卷欄位！", vbInformation, "取消閉卷"
      End If
      'end 2025/06/30
      
      'Added By Lydia 2023/03/15 檢查商品資料與基本檔商品類別是否一致
      If CheckTMGoodsErr(strTM01, strTM02, strTM03, strTM04, True) = True Then
         frm03010303_04.Hide
         Set frm03010303_04.UpForm = Me
         frm03010303_04.TGKey = strTM01 & "-" & strTM02 & "-" & strTM03 & "-" & strTM04
         frm03010303_04.AllClass = textTM09.Text
         frm03010303_04.cmdOK(2).Visible = True
         If textTM09 <> "" Then
            Me.Hide
            frm03010303_04.QueryData
            frm03010303_04.Show vbModal
         Else
            MsgBox ("無商品類別，不可使用此按鈕 !")
         End If
      End If
      'end 2023/03/15
      
      ClearAll
      If NextRecord = True Then
         QueryData
      Else
         Unload Me
         'Add by Morgan 2003/12/05
         frm020101_01.QueryDB
         'End 2003/12/05
         frm020101_01.Show
      End If
   End If
End Sub

Private Sub cmdPriority_Click()
   '修改優先權資料
   'Add by Amy 2023/01/06 此支進優先權表單改不是強制表單,故進入時畫面鎖住
   mdiMain.Enabled = False
   Me.Enabled = False
   'Modify by Amy 2014/04/17 +, m_Priority(4), m_Priority(5)
   'Modify by Sindy 2017/10/12 +, m_Priority(6)
   'Modify by Amy 2023/01/06 原m_Priority為陣列,加表單名
   'Modify By Sindy 2024/4/19 + m_TM01 & m_TM02 & m_TM03 & m_TM04
   ModifyPriority m_Priority1, m_Priority2, m_Priority3, , , m_TM01 & m_TM02 & m_TM03 & m_TM04, , , m_Priority4, m_Priority5, m_Priority6, Me
End Sub

Private Sub cmdRelate_Click()
   Where1103ComeFrom Me, m_TM01, m_TM02, m_TM03, m_TM04
End Sub

Private Sub Form_Load()
   ' 設定控制項的背景顏色
   textTMKey.BackColor = &H8000000F
   textTM05S.BackColor = &H8000000F
'   textTM06S.BackColor = &H8000000F
'   textTM07S.BackColor = &H8000000F
   textTM21S.BackColor = &H8000000F
   textTM22S.BackColor = &H8000000F
   textTM29_2.BackColor = &H8000000F
   textTM08_2.BackColor = &H8000000F
   textTM72_2.BackColor = &H8000000F 'Add By Sindy 2019/4/9
   textTM10_2.BackColor = &H8000000F
   textTM23_2.BackColor = &H8000000F
   textTM44_2.BackColor = &H8000000F
   'add by nickc 2008/02/01
   textCP44_2.BackColor = &H8000000F
   
   textSP58_2.BackColor = &H8000000F
   textSP59_2.BackColor = &H8000000F
   'add by nickc 2006/12/13
   textTM80_2.BackColor = &H8000000F
   textTM81_2.BackColor = &H8000000F
   
   textCP09.BackColor = &H8000000F
   textCP10_2.BackColor = &H8000000F
   textCP13_2.BackColor = &H8000000F
   textCP14_2.BackColor = &H8000000F
   textCP57.BackColor = &H8000000F
   
   MoveFormToCenter Me
   
   'Added by Lydia 2020/05/20 法律所案源收文
   FraLOS.Visible = False
   FraLOS.BackColor = &H8000000F
   txtLOSagree.Text = ""
   'end 2020/05/20
   
   '920224 nick
   'Nick920224Bol = False 'Removed by Morgan 2024/12/26 沒用了
   
   'Added by Lydia 2018/12/10
   Frame21.Height = 710 'Modify by Amy 2022/11/17 原:810
   Label65.Top = Label57.Top
   textCP143.Top = textEP34.Top
   'Add By Sindy 2023/12/11
   Frame2.BorderStyle = 0
   Frame3.BorderStyle = 0
   '2023/12/11 END
   
   'Add By Sindy 2023/12/11
   If strSrvDate(1) < 指定日期啟用日 Then
      Frame3.Visible = False
   End If
   '2023/12/11 END
End Sub

Private Sub InitialData()
   m_CPCount = 0
End Sub

Private Sub ClearAll()
    'ClearCPList
    ClearTMSPFieldList
    ClearCPFieldList
    'm_CP09 = Empty
    m_TM28 = Empty
    textTMKey = Empty
    textTM01 = Empty
    textTM02 = Empty
    textTM02_2 = Empty
    textTM03 = Empty
    textTM04 = Empty
    textTM05 = Empty
    textTM05_1 = Empty
    textTM05S = Empty
    textTM06 = Empty
    textTM07 = Empty
    textTM08 = Empty
    textTM08_2 = Empty
    'Add By Sindy 2019/4/9
    textTM72 = Empty
    textTM72_2 = Empty
    '2019/4/9 END
    textTM09 = Empty
    textTM10 = Empty
    textTM10_2 = Empty
    textTM14 = Empty 'Add By Sindy 2010/7/16
    textTM21S = Empty
    textTM22S = Empty
    textTM23 = Empty
    textTM23_2 = Empty
    textTM24 = Empty
    textTM25 = Empty
    textTM26 = Empty
    textTM28 = Empty
    textTM29 = Empty
    textTM29_2 = Empty
    textTM34 = Empty
    textTM35 = Empty
    textTM44 = Empty
    textTM44_2 = Empty
    'add by nickc 2008/02/01
    textCP44 = Empty
    textCP44_2 = Empty
    
    textTM45 = Empty
    'add by nickc 2008/01/31 新增聯絡人1(中)
    textTM38 = Empty
    textTM39 = Empty 'Add By Sindy 2015/2/26 新增聯絡人1(英)
    textTM40 = Empty 'Add By Sindy 2015/2/26 新增聯絡人1(日)
    'add by Sindy 2012/12/20 新增聯絡人2(中)
    textTM41 = Empty
    textTM42 = Empty 'Add By Sindy 2015/2/26 新增聯絡人2(英)
    textTM43 = Empty 'Add By Sindy 2015/2/26 新增聯絡人2(日)
    
    textTM58 = Empty
    textSP32 = Empty
    textSP58 = Empty
    textSP58_2 = Empty
    textSP59 = Empty
    textSP59_2 = Empty
    textCP05 = Empty
    textCP06 = Empty
    textCP07 = Empty
    textCP09 = Empty
    textCP10 = Empty
    textCP10_2 = Empty
    textCP13 = Empty
    textCP13_2 = Empty
    textCP14 = Empty
    textCP14_2 = Empty
    textCP26 = Empty
    textCP43 = Empty
    textCP57 = Empty
    textCP64 = Empty
    
    'add by nickc 2006/12/14
    textTM80 = Empty
    textTM80_2 = Empty
    textTM81 = Empty
    textTM81_2 = Empty
    textTM32 = Empty
    textTM82 = Empty
    textTM83 = Empty
    textTM84 = Empty
    textTM85 = Empty
    textTM86 = Empty
    textTM87 = Empty
    textTM88 = Empty
    textTM89 = Empty
    textTM90 = Empty
    textTM91 = Empty
    textTM92 = Empty
    textTM93 = Empty
    
    'Add By Cheng 2003/08/19
    'Begin
    Me.Text1.Text = Empty
    Me.Text2.Text = Empty
    Me.Text3.Text = Empty
    Me.Text4.Text = Empty
    'End
    'Add By Cheng 2004/04/14
    Me.txtDivCaseNo(0).Text = "": Me.txtDivCaseNo(1).Text = "": Me.txtDivCaseNo(2).Text = "": Me.txtDivCaseNo(3).Text = "": Me.txtDivCaseNo(4).Text = ""
    'End
    
    textTM130 = Empty 'Add By Sindy 2013/12/16
    
    m_strCountry = Empty
    
    'Added by Lydia 2019/04/29 T-213068,T-213069連續分案,未清空文件是否齊備欄位,令判斷為未變更,所以T-213069未能上文件齊備日
    textEP06 = Empty
    textEP34 = Empty
    textCP122 = Empty
    textCP143 = Empty
    
    Me.txtF0301 = Empty 'Add by Amy 2022/10/07
    
    'Added by Lydia 2023/01/31 T大陸查名
    'Modified by Lydia 2023/02/10 改為獨立欄位EP13=>EP43
    Frame22.Visible = False
    textEP43 = Empty
    textEP43.Tag = Empty
    ChkEP43.Value = 0
    lblEP43.Caption = ""
    
    cboTM08.Text = Empty: cboTM72.Text = "": cboTM08.Tag = Empty: cboTM72.Tag = "" 'Added by Lydia 2023/11/16
    
End Sub

Private Sub AddCPToList(ByVal strCP09 As String)
   Dim bFind As Boolean
   Dim nIndex As Integer
   bFind = False
   For nIndex = 0 To m_CPKeyCount - 1
      If m_CPKeyList(nIndex) = strCP09 Then
         bFind = True
         Exit For
      End If
   Next nIndex
   If bFind = False Then
      ReDim Preserve m_CPKeyList(m_CPKeyCount + 1)
      m_CPKeyList(m_CPKeyCount) = strCP09
      m_CPKeyCount = m_CPKeyCount + 1
   End If
End Sub

Private Sub ClearCPList()
   If m_CPKeyCount > 0 Then
      Erase m_CPKeyList
   End If
   m_CPKeyCount = 0
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

Public Sub SetData(ByVal nType As Integer, ByVal strData As String, Optional ByVal bClear As Boolean = False)
   ' 清除搜尋的Key
   If bClear = True Then
      ClearAll
      ClearCPList
   End If
   
   Select Case nType
      ' 收文號
      Case 0:
         AddCPToList strData
         ' 第一筆
         If m_CPKeyCount > 0 Then: m_CP09 = m_CPKeyList(0)
      ' 相關總收文號
      Case 99: textCP43 = strData
   End Select
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
      
'      ' 案件中文名稱
      ' 案件名稱
      If IsNull(rsTmp.Fields("TM05")) = False Then
'         textTM05 = rsTmp.Fields("TM05")
         Me.textTM05_1.Text = rsTmp.Fields("TM05")
      End If
      SetTMSPFieldOldData "TM05", textTM05_1, 0
'      ' 案件英文名稱
'      If IsNull(rsTmp.Fields("TM06")) = False Then
'         textTM06 = rsTmp.Fields("TM06")
'      End If
'      SetTMSPFieldOldData "TM06", textTM06, 0
'      ' 案件日文名稱
'      If IsNull(rsTmp.Fields("TM07")) = False Then
'         textTM07 = rsTmp.Fields("TM07")
'      End If
'      SetTMSPFieldOldData "TM07", textTM07, 0
      ' 商標種類
      If IsNull(rsTmp.Fields("TM08")) = False Then
         textTM08 = rsTmp.Fields("TM08")
         textTM08_Validate False
      End If
      SetTMSPFieldOldData "TM08", textTM08, 0
      'Add By Sindy 2019/4/9
      ' 特殊商標
      If IsNull(rsTmp.Fields("TM72")) = False Then
         textTM72 = rsTmp.Fields("TM72")
         textTM72_Validate False
      End If
      SetTMSPFieldOldData "TM72", textTM72, 0
      '2019/4/9 END
      ' 商品類別
      If IsNull(rsTmp.Fields("TM09")) = False Then
         textTM09 = rsTmp.Fields("TM09")
      End If
      SetTMSPFieldOldData "TM09", textTM09, 0
      ' 申請國家
      If IsNull(rsTmp.Fields("TM10")) = False Then
         textTM10 = rsTmp.Fields("TM10")
         textTM10_Validate False
         m_TM10 = rsTmp.Fields("TM10")
         'add by nickc 2005/11/21
         'edit by nickc 2006/06/05 應帶所有子案國家
         'm_strCountry = rsTmp.Fields("TM10")
         Dim nick920224rs As New ADODB.Recordset
         Set nick920224rs = New ADODB.Recordset
         Dim nick920224str As String
         nick920224str = "select DISTINCT(tm10) from trademark where tm01='" & m_TM01 & "' and tm02='" & m_TM02 & "' and tm03>'0'"
         With nick920224rs
             .CursorLocation = adUseClient
             .Open nick920224str, cnnConnection, adOpenStatic, adLockReadOnly
             If Not .EOF And Not .BOF Then
                 .MoveFirst
                 Do While Not .EOF
                     If Trim(m_strCountry) = "" Then
                         m_strCountry = m_strCountry & CheckStr(.Fields(0).Value)
                     Else
                         m_strCountry = m_strCountry & "," & CheckStr(.Fields(0).Value)
                     End If
                     .MoveNext
                 Loop
             End If
         End With
      End If
      SetTMSPFieldOldData "TM10", m_TM10, 0
      
      'Added by Lydia 2023/11/16 內外商之分案及商標基本資料維護之商標種類、特殊商標欄位增加下拉功能
      Pub_SetTMcombo "1", cboTM08, textTM08, IIf(textTM10 <> "000", True, False), strPTM '商標種類
      Pub_SetTMcombo "2", cboTM72, textTM72, IIf(textTM10 <> "000", True, False), strSPT '特殊商標種類
      'end 2023/11/16
      
      ' 申請人
      If IsNull(rsTmp.Fields("TM23")) = False Then
         textTM23 = ChangeCustomerL(rsTmp.Fields("TM23"))
         m_TM23 = ChangeCustomerL(textTM23)
         'Modify By Cheng 2002/09/20
'         textTM23_Validate False
         textTM23_2 = GetCustomerName(textTM23, 0)
      Else
         m_TM23 = ""
      End If
      SetTMSPFieldOldData "TM23", textTM23, 0 'Add By Sindy 2011/7/11
      
      '2011/6/10 add by sonia抓畫面上的定稿語文TM77
      m_TM77 = ""
      'Modify by Morgan 2011/6/13
      'If IsNull(rsTmp.Fields("TM27")) = False Then
      If IsNull(rsTmp.Fields("TM77")) = False Then
         m_TM77 = rsTmp.Fields("TM77")
      Else
         m_TM77 = GetTWordLng(m_TM01, m_TM02, m_TM03, m_TM04)
      End If
      '2011/6/10 END
      
      'Add By Sindy 2013/8/6
      m_TM56 = ""
      If IsNull(rsTmp.Fields("TM56")) = False Then
         m_TM56 = rsTmp.Fields("TM56")
      End If
      '2013/8/6 END
      
      'Add By Cheng 2002/08/22
      'Mark by Lydia 2024/06/13
      'm_strCust1 = "" & Me.textTM23.Text
      'm_strCust2 = ""
      'm_strCust3 = ""
      'end 2024/06/13
      
      'add by nickc 2006/12/14
      textSP58 = ChangeCustomerL(CheckStr(rsTmp.Fields("TM78")))
      SetTMSPFieldOldData "TM78", textSP58, 0
      m_TM78 = ChangeCustomerL(textSP58)
      textSP58_2 = GetCustomerName(textSP58, 0)
      'm_strCust2 = "" & Me.textSP58.Text 'Mark by Lydia 2024/06/13
      textSP59 = ChangeCustomerL(CheckStr(rsTmp.Fields("TM79")))
      SetTMSPFieldOldData "TM79", textSP59, 0
      m_TM79 = ChangeCustomerL(textSP59)
      textSP59_2 = GetCustomerName(textSP59, 0)
      'm_strCust3 = "" & Me.textSP59.Text 'Mark by Lydia 2024/06/13
      textTM80 = ChangeCustomerL(CheckStr(rsTmp.Fields("TM80")))
      SetTMSPFieldOldData "TM80", textTM80, 0
      m_TM80 = ChangeCustomerL(textTM80)
      textTM80_2 = GetCustomerName(textTM80, 0)
      'm_strCust4 = "" & Me.textTM80.Text 'Mark by Lydia 2024/06/13
      textTM81 = ChangeCustomerL(CheckStr(rsTmp.Fields("TM81")))
      SetTMSPFieldOldData "TM81", textTM81, 0
      m_TM81 = ChangeCustomerL(textTM81)
      textTM81_2 = GetCustomerName(textTM81, 0)
      'm_strCust5 = "" & Me.textTM81.Text 'Mark by Lydia 2024/06/13
      textTM82 = CheckStr(rsTmp.Fields("TM82"))
      SetTMSPFieldOldData "TM82", textTM82, 0
      textTM83 = CheckStr(rsTmp.Fields("TM83"))
      SetTMSPFieldOldData "TM83", textTM83, 0
      textTM84 = CheckStr(rsTmp.Fields("TM84"))
      SetTMSPFieldOldData "TM84", textTM84, 0
      textTM85 = CheckStr(rsTmp.Fields("TM85"))
      SetTMSPFieldOldData "TM85", textTM85, 0
      textTM86 = CheckStr(rsTmp.Fields("TM86"))
      SetTMSPFieldOldData "TM86", textTM86, 0
      textTM87 = CheckStr(rsTmp.Fields("TM87"))
      SetTMSPFieldOldData "TM87", textTM87, 0
      textTM88 = CheckStr(rsTmp.Fields("TM88"))
      SetTMSPFieldOldData "TM88", textTM88, 0
      textTM89 = CheckStr(rsTmp.Fields("TM89"))
      SetTMSPFieldOldData "TM89", textTM89, 0
      textTM90 = CheckStr(rsTmp.Fields("TM90"))
      SetTMSPFieldOldData "TM90", textTM90, 0
      textTM91 = CheckStr(rsTmp.Fields("TM91"))
      SetTMSPFieldOldData "TM91", textTM91, 0
      textTM92 = CheckStr(rsTmp.Fields("TM92"))
      SetTMSPFieldOldData "TM92", textTM92, 0
      textTM93 = CheckStr(rsTmp.Fields("TM93"))
      SetTMSPFieldOldData "TM93", textTM93, 0
      textTM32 = CheckStr(rsTmp.Fields("TM32"))
      SetTMSPFieldOldData "TM32", textTM32, 0
      
      ' 申請地址
      If IsNull(rsTmp.Fields("TM24")) = False Then
         textTM24 = rsTmp.Fields("TM24")
      End If
      SetTMSPFieldOldData "TM24", textTM24, 0
      ' 申請地址
      If IsNull(rsTmp.Fields("TM25")) = False Then
         textTM25 = rsTmp.Fields("TM25")
      End If
      SetTMSPFieldOldData "TM25", textTM25, 0
      ' 申請地址
      If IsNull(rsTmp.Fields("TM26")) = False Then
         textTM26 = rsTmp.Fields("TM26")
      End If
      SetTMSPFieldOldData "TM26", textTM26, 0
      ' 卷宗性質
      If IsNull(rsTmp.Fields("TM28")) = False Then
         m_TM28 = rsTmp.Fields("TM28")
         textTM28 = rsTmp.Fields("TM28")
      End If
      SetTMSPFieldOldData "TM28", textTM28, 0
      ' 是否閉卷
      If IsNull(rsTmp.Fields("TM29")) = False Then
         m_TM29 = rsTmp.Fields("TM29")
      End If
      SetTMSPFieldOldData "TM29", m_TM29, 0
      ' 分所案號
      If IsNull(rsTmp.Fields("TM34")) = False Then
         textTM34 = rsTmp.Fields("TM34")
      End If
      SetTMSPFieldOldData "TM34", textTM34, 0
      ' 客戶案件案號
      If IsNull(rsTmp.Fields("TM35")) = False Then
         textTM35 = rsTmp.Fields("TM35")
      End If
      SetTMSPFieldOldData "TM35", textTM35, 0
      ' FC代理人
      If IsNull(rsTmp.Fields("TM44")) = False Then
         textTM44 = ChangeCustomerL(rsTmp.Fields("TM44"))
         textTM44_Validate False
      End If
      SetTMSPFieldOldData "TM44", textTM44, 0
      m_TM44 = "" & Me.textTM44.Text 'Added by Lydia 2024/06/13
      ' 彼所案號
      If IsNull(rsTmp.Fields("TM45")) = False Then
         textTM45 = rsTmp.Fields("TM45")
      End If
      SetTMSPFieldOldData "TM45", textTM45, 0
      ' 案件備註
      If IsNull(rsTmp.Fields("TM58")) = False Then
         textTM58 = rsTmp.Fields("TM58")
      End If
      SetTMSPFieldOldData "TM58", textTM58, 0
      'Add By Cheng 2002/06/12
      '取得專用期止日
      m_TM22 = "" & rsTmp.Fields("TM22")
      
      'add by nickc 2008/01/31  新增聯絡人1(中)
      If IsNull(rsTmp.Fields("TM38")) = False Then
         textTM38 = rsTmp.Fields("TM38")
      End If
      SetTMSPFieldOldData "TM38", textTM38, 0
      'Add By Sindy 2015/2/26 新增聯絡人1(英)
      If IsNull(rsTmp.Fields("TM39")) = False Then
         textTM39 = rsTmp.Fields("TM39")
      End If
      SetTMSPFieldOldData "TM39", textTM39, 0
      'Add By Sindy 2015/2/26 新增聯絡人1(日)
      If IsNull(rsTmp.Fields("TM40")) = False Then
         textTM40 = rsTmp.Fields("TM40")
      End If
      SetTMSPFieldOldData "TM40", textTM40, 0
      
      'add by Sindy 2012/12/20  新增聯絡人2(中)
      If IsNull(rsTmp.Fields("TM41")) = False Then
         textTM41 = rsTmp.Fields("TM41")
      End If
      SetTMSPFieldOldData "TM41", textTM41, 0
      'Add By Sindy 2015/2/26 新增聯絡人2(英)
      If IsNull(rsTmp.Fields("TM42")) = False Then
         textTM42 = rsTmp.Fields("TM42")
      End If
      SetTMSPFieldOldData "TM42", textTM42, 0
      'Add By Sindy 2015/2/26 新增聯絡人2(日)
      If IsNull(rsTmp.Fields("TM43")) = False Then
         textTM43 = rsTmp.Fields("TM43")
      End If
      SetTMSPFieldOldData "TM43", textTM43, 0
      
      'Add By Sindy 2010/7/16 公告日
      If IsNull(rsTmp.Fields("TM14")) = False Then
         textTM14 = Val(rsTmp.Fields("TM14")) - 19110000
         textTM14_Validate False
      End If
      SetTMSPFieldOldData "TM14", textTM14, 0
      m_TM11 = "" & rsTmp.Fields("TM11")   '2012/3/29 ADD BY SONIA
      
      textTM130 = "" 'Add by Amy 2016/11/07
      'Add By Sindy 2013/12/16
      If IsNull(rsTmp.Fields("TM130")) = False Then
         textTM130 = rsTmp.Fields("TM130")
      End If
      textTM130.Tag = textTM130
      SetTMSPFieldOldData "TM130", textTM130, 0
      '2013/12/16 END
      'Add by Amy 2016/08/12 +客戶檔收據公司別
      'Modify by Amy 2016/08/29
      'Modify by Amy 2017/03/21 +新案才帶 (CFT-013586 舊案不應該帶)
      'Mark by Amy 2017/11/24 個案為空白會預設申請人出名公司,若個案改為空仍會一直預帶(CFP-029915)-秀玲:拿掉
'      If textTM130 = MsgText(601) And m_CP31 = "Y" Then
'         textTM130 = GetReceiptCmp(Left(textTM23, 8), Mid(textTM23, 9, 1), m_TM01, textTM10)
'      End If
      
      'Added by Morgan 2022/12/15
      textTM136 = "" & rsTmp.Fields("TM136")
      textTM136.Tag = textTM136
      'end 2022/12/15
   End If
   rsTmp.Close
   Set rsTmp = Nothing
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
      If IsNull(rsTmp.Fields("SP05")) = False Then
        'Modify By Cheng 2004/02/24
        '查名案件名稱合併至一欄
        Select Case m_TM01
        Case "TS"
            textTM05_1 = rsTmp.Fields("SP05")
        Case Else
            textTM05 = rsTmp.Fields("SP05")
        End Select
        'End
      End If
        'Modify By Cheng 2004/02/24
        '查名案件名稱合併至一欄
        Select Case m_TM01
        Case "TS"
            SetTMSPFieldOldData "SP05", textTM05_1, 0
        Case Else
            SetTMSPFieldOldData "SP05", textTM05, 0
        End Select
        'End
      ' 案件英文名稱
      If IsNull(rsTmp.Fields("SP06")) = False Then
         textTM06 = rsTmp.Fields("SP06")
      End If
      SetTMSPFieldOldData "SP06", textTM06, 0
      ' 案件日文名稱
      If IsNull(rsTmp.Fields("SP07")) = False Then
         textTM07 = rsTmp.Fields("SP07")
      End If
      SetTMSPFieldOldData "SP07", textTM07, 0
      ' 申請人
      If IsNull(rsTmp.Fields("SP08")) = False Then
         textTM23 = ChangeCustomerL(rsTmp.Fields("SP08"))
         m_TM23 = ChangeCustomerL(textTM23)
         'Modify By Cheng 2002/09/20
'         textTM23_Validate False
         textTM23_2 = GetCustomerName(textTM23, 0)
      Else
         m_TM23 = ""
      End If
      SetTMSPFieldOldData "SP08", textTM23, 0
      'Add By Cheng 2002/08/22
      'Mark by Lydia 2024/06/13
      'm_strCust1 = "" & Me.textTM23.Text
      ''add by nickc 2006/12/14
      'm_strCust2 = ""
      'm_strCust3 = ""
      'm_strCust4 = ""
      'm_strCust5 = ""
      'end 2024/06/13
      
      ' 第二申請人及第三申請人
'EDIT BY NICKC  2006/12/14
'      If m_TM01 = "TC" Then
         If IsNull(rsTmp.Fields("SP58")) = False Then
            'edit by nick 2004/07/21 修正
            'textSP58 = ChangeCustomerL(rsTmp.Fields("TM58"))
            textSP58 = ChangeCustomerL(rsTmp.Fields("SP58"))
            'Modify By Cheng 2002/09/23
'            textSP58_Validate False
            textSP58_2 = GetCustomerName(textSP58, 0)
         End If
         If IsNull(rsTmp.Fields("SP59")) = False Then
            'edit by nick  2004/07/21 修正
            'textSP59 = ChangeCustomerL(rsTmp.Fields("TM59"))
            textSP59 = ChangeCustomerL(rsTmp.Fields("SP59"))
            'Modify By Cheng 2002/09/23
'            textSP59_Validate False
            textSP59_2 = GetCustomerName(textSP59, 0)
         End If
         'add by nickc 2006/12/14
         textTM80 = ChangeCustomerL(CheckStr(rsTmp.Fields("SP65")))
         textTM80_2 = GetCustomerName(textTM80, 0)
         textTM81 = ChangeCustomerL(CheckStr(rsTmp.Fields("SP66")))
         textTM81_2 = GetCustomerName(textTM81, 0)
         m_TM78 = "" & Me.textSP58.Text
         m_TM79 = "" & Me.textSP59.Text
         m_TM80 = "" & Me.textTM80.Text
         m_TM81 = "" & Me.textTM81.Text
         m_TM44 = "" & Me.textTM44.Text 'Added by Lydia 2024/06/13
         
         'Add By Cheng 2002/08/22
         'Mark by Lydia 2024/06/13
         'm_strCust2 = "" & Me.textSP58.Text
         'm_strCust3 = "" & Me.textSP59.Text
         '
         'add by nickc 2006/12/14
        ' m_strCust4 = "" & Me.textTM80.Text
        ' m_strCust5 = "" & Me.textTM81.Text
        'end 2024/06/13
        
         textTM09 = CheckStr(rsTmp.Fields("sp73"))
         textTM32 = CheckStr(rsTmp.Fields("sp74"))
         'add by nickc 2007/08/28  修正錯誤
         SetTMSPFieldOldData "SP58", textSP58, 0
         SetTMSPFieldOldData "SP59", textSP59, 0
         SetTMSPFieldOldData "SP65", textTM80, 0
         SetTMSPFieldOldData "SP66", textTM81, 0
         SetTMSPFieldOldData "SP73", textTM09, 0
         SetTMSPFieldOldData "SP74", textTM32, 0
         
         
'edit by nickc 2006/12/14
'      End If
      ' 申請國家
      If IsNull(rsTmp.Fields("SP09")) = False Then
         m_TM10 = rsTmp.Fields("SP09")
         textTM10 = rsTmp.Fields("SP09")
         textTM10_Validate False
      End If
      SetTMSPFieldOldData "SP09", m_TM10, 0
      ' 是否閉卷
      If IsNull(rsTmp.Fields("SP15")) = False Then
         m_TM29 = rsTmp.Fields("SP15")
      End If
      ' FC代理人
      If IsNull(rsTmp.Fields("SP26")) = False Then
         textTM44 = ChangeCustomerL(rsTmp.Fields("SP26"))
         textTM44_Validate False
      End If
      SetTMSPFieldOldData "SP26", textTM44, 0
      '911120 nick
      SetTMSPFieldOldData "SP27", CheckStr(rsTmp.Fields("SP27")), 0
      textTM45 = CheckStr(rsTmp.Fields("SP27"))
      SetTMSPFieldOldData "SP28", CheckStr(rsTmp.Fields("SP28")), 0
      textTM34 = CheckStr(rsTmp.Fields("SP28"))
      
      'add by nickc 2008/01/31 新增聯絡人
      SetTMSPFieldOldData "SP30", CheckStr(rsTmp.Fields("SP30")), 0
      textTM38 = CheckStr(rsTmp.Fields("SP30"))
      textTM38.MaxLength = 60 'Add By Sindy 2016/10/27
      textTM39.Enabled = False 'Add By Sindy 2015/2/26
      textTM40.Enabled = False 'Add By Sindy 2015/2/26
      'add by Sindy 2012/12/20 新增聯絡人2
      SetTMSPFieldOldData "SP75", CheckStr(rsTmp.Fields("SP75")), 0
      textTM41 = CheckStr(rsTmp.Fields("SP75"))
      textTM41.MaxLength = 60 'Add By Sindy 2016/10/27
      textTM42.Enabled = False 'Add By Sindy 2015/2/26
      textTM43.Enabled = False 'Add By Sindy 2015/2/26
      
      ' 商標審定號
      If IsNull(rsTmp.Fields("SP32")) = False Then
         textSP32 = rsTmp.Fields("SP32")
      End If
      SetTMSPFieldOldData "SP32", textSP32, 0
        'Add By Cheng 2004/02/03
        '顯示此審定號的相關資訊
        textSP32_Validate False
        'End
      ' 案件備註
      If IsNull(rsTmp.Fields("SP18")) = False Then
         textTM58 = rsTmp.Fields("SP18")
      End If
      SetTMSPFieldOldData "SP18", textTM58, 0
      
      'Modify by Amy 2016/11/07
      textTM130 = ""
      'Add By Sindy 2013/12/16
      If IsNull(rsTmp.Fields("SP85")) = False Then
         textTM130 = rsTmp.Fields("SP85")
      End If
      textTM130.Tag = textTM130
      SetTMSPFieldOldData "SP85", textTM130, 0
      '2013/12/16 END
      '客戶檔收據公司別
      'Mark by Amy 2017/11/24 個案為空白會預設申請人出名公司,若個案改為空仍會一直預帶(CFP-029915)-秀玲:拿掉
'      If textTM130 = MsgText(601) Then
'         textTM130 = GetReceiptCmp(Left(textTM23, 8), Mid(textTM23, 9, 1), m_TM01, textTM10)
'      End If
      'end 2016/11/07
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Sub

' 取得案件進度檔的欄位內容
Private Sub QueryCaseProgress()
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   Dim strTemp As String
   Dim strSubSQL As String
   Dim rsSubTmp As ADODB.Recordset
   
   'add by nickc 2006/09/28
   m_CP31 = ""
   
   strSql = "SELECT * FROM CaseProgress " & _
            "WHERE CP09 = '" & m_CP09 & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst

      'Add By Sindy 2012/5/8
      '費用
      m_CP16 = ""
      If IsNull(rsTmp.Fields("CP16")) = False Then
         m_CP16 = rsTmp.Fields("CP16")
      End If
      '是否急件
      m_CP122 = ""
      If IsNull(rsTmp.Fields("CP122")) = False Then
         m_CP122 = rsTmp.Fields("CP122")
         textCP122 = rsTmp.Fields("CP122")
      End If
      SetCPFieldOldData "CP122", textCP122, 0
      '2012/5/8 End
      
      ' 案件性質
      If IsNull(rsTmp.Fields("CP10")) = False Then
         m_CP10 = rsTmp.Fields("CP10")
         textCP10 = rsTmp.Fields("CP10")
         textCP10_Validate False
         textTM14_Validate False 'Add By Sindy 2010/7/16 公告日
      End If
      
      m_CP149 = "" & rsTmp.Fields("CP149") 'Add By Sindy 2022/4/27
      
      'Add by Morgan 2011/4/22
      m_CP30 = "" & rsTmp.Fields("cp30")
      SetCPFieldOldData "CP30", m_CP30, 0
      'end 2011/4/22
      
      'add by sonia 2018/11/20
      m_CP46 = "" & rsTmp.Fields("cp46")
      SetCPFieldOldData "CP46", m_CP46, 0
      'end 2018/11/20
      
      'Add by Morgan 2009/12/25
      If textCP10 = "303" Then
         textCP10.Enabled = False
      Else
         textCP10.Enabled = True
      End If
      SetCPFieldOldData "CP10", textCP10, 0
      ' 收文日
      strTemp = Empty
      If IsNull(rsTmp.Fields("CP05")) = False Then
         strTemp = rsTmp.Fields("CP05")
         textCP05 = TAIWANDATE(rsTmp.Fields("CP05"))
      End If
      SetCPFieldOldData "CP05", strTemp, 1
      ' 本所期限
      If IsNull(rsTmp.Fields("CP06")) = False Then
         textCP06 = TAIWANDATE(rsTmp.Fields("CP06"))
         'add by Toni 2008/10/27
         textCP06.Tag = TAIWANDATE(rsTmp.Fields("CP06"))
         'end 2008/10/27
      End If
      SetCPFieldOldData "CP06", textCP06, 1
      'Add By Cheng 2002/06/12
      m_strCP06 = "" & rsTmp.Fields("CP06")
      
      'Add By Sindy 2012/5/8
      '承辦期限
      m_CP48 = Empty
      If IsNull(rsTmp.Fields("CP48")) = False Then
         m_CP48 = rsTmp.Fields("CP48")
      End If
      '2012/5/8 End
      
      ' 法定期限
      If IsNull(rsTmp.Fields("CP07")) = False Then
         textCP07 = TAIWANDATE(rsTmp.Fields("CP07"))
         'add by Toni 2008/10/27
         textCP07.Tag = TAIWANDATE(rsTmp.Fields("CP07"))
         'end 2008/10/27
      End If
      SetCPFieldOldData "CP07", textCP07, 1
      'Add By Cheng 2002/06/12
      m_strCP07 = "" & rsTmp.Fields("CP07")
      ' 業務區
      '911030 nick 解決 null 問題
      'SetCPFieldOldData "CP12", rsTmp.Fields("CP12"), 0
      SetCPFieldOldData "CP12", CheckStr(rsTmp.Fields("CP12")), 0
      
      ' 智權人員
      If IsNull(rsTmp.Fields("CP13")) = False Then
         textCP13 = rsTmp.Fields("CP13")
         textCP13_Validate False
      End If
      SetCPFieldOldData "CP13", textCP13, 0
      
      ' 承辦人員
      If IsNull(rsTmp.Fields("CP14")) = False Then
         textCP14 = rsTmp.Fields("CP14")
         'add by Toni 2008/10/27
         textCP14.Tag = rsTmp.Fields("CP14")
         'end 2008/10/27
         textCP14_Validate False
      End If
      SetCPFieldOldData "CP14", textCP14, 0
      'Modify by Amy 2022/11/03 +cp157為空=第一次分案
      'Add By Sindy 2022/8/19
      '北所分案日
'      If IsNull(rsTmp.Fields("CP157")) = True Then
'         textCP14.Tag = ""
'      End If
      If "" & rsTmp.Fields("CP157") = MsgText(601) Then
            m_bolIsFirstKeyCP14 = True
      End If
      SetCPFieldOldData "CP157", "" & rsTmp.Fields("CP157"), 1
      '2022/8/19 END
      
      ' 相關總收文號
      If IsNull(rsTmp.Fields("CP43")) = False Then
         textCP43 = rsTmp.Fields("CP43")
      End If
      SetCPFieldOldData "CP43", textCP43, 0
      ' 是否算案件數
      If IsNull(rsTmp.Fields("CP26")) = False Then
         textCP26 = rsTmp.Fields("CP26")
      End If
      SetCPFieldOldData "CP26", textCP26, 0
      
      'Add By Sindy 2012/6/1
      m_CP27 = Empty
      If IsNull(rsTmp.Fields("CP27")) = False Then
         m_CP27 = rsTmp.Fields("CP27")
      End If
      '2012/6/1 End
      
      ' 取消收文日期
      If IsNull(rsTmp.Fields("CP57")) = False Then
         textCP57 = TAIWANDATE(rsTmp.Fields("CP57"))
      End If
      ' 是否新案件
      If IsNull(rsTmp.Fields("CP31")) = False Then
         m_CP31 = rsTmp.Fields("CP31")
      End If
      ' CreateID Add By Sindy 2010/8/6
      m_CP65 = ""
      If IsNull(rsTmp.Fields("CP65")) = False Then
         m_CP65 = rsTmp.Fields("CP65")
      End If
      
      'Added By Lydia 2015/11/24 收款後送件
      m_CP141 = "" & rsTmp.Fields("CP141")
      SetCPFieldOldData "CP141", m_CP141, 0
      'end 2015/11/24
      'Add by Sindy 2023/4/21
      textCP142 = ""
      OptSendType(1).Caption = PUB_GetCP114Opt1Desc(m_TM01, m_CP10) 'Add By Sindy 2024/1/22
      Select Case m_CP141
         Case "1"
            OptSendType(1).Value = True
         Case "2"
            OptSendType(2).Value = True
         Case "3"
            OptSendType(3).Value = True
            textCP142.Text = TransDate("" & rsTmp.Fields("CP142"), 1)
         '要清除否則按下一筆分案會預設前筆狀態
         Case Else
            OptSendType(1).Value = False
            OptSendType(2).Value = False
            OptSendType(3).Value = False
      End Select
      SetCPFieldOldData "CP142", textCP142, 1
      '2023/4/21 END
      'Add By Sindy 2023/12/11 指定日期方式
      If Frame3.Visible = True Then
         If "" & rsTmp.Fields("CP164") = "1" Then
            Option1(0).Value = True
         ElseIf "" & rsTmp.Fields("CP164") = "2" Then
            Option1(1).Value = True
         ElseIf "" & rsTmp.Fields("CP164") = "3" Then
            Option1(2).Value = True
         End If
      End If
      '2023/12/11 END
      
      'm_CP143 = "" & rsTmp.Fields("CP143")  'Added by Lydia 2018/12/10 查名是否齊備 'Mark by Lydia 2020/11/04 debug-發現近日的商標申請案紙本有印查名齊備,但是到分案就不見了;是因為這一行不見
      p_strCP143 = "" & rsTmp.Fields("CP143") 'Added by Lydia 2020/11/04 收文之查名齊備日
      
      'Add By Sindy 2020/10/20
      If textCP10 = "210" Then '陳述意見書
         m_CP143 = "" & rsTmp.Fields("CP143")
         If m_CP143 <> "" Then m_CP143 = Val(m_CP143) - 19110000
         SetCPFieldOldData "CP143", "" & rsTmp.Fields("CP143"), 1
         m_CP36 = "" & rsTmp.Fields("CP36")
         SetCPFieldOldData "CP36", m_CP36, 0
         m_CP21 = "" & rsTmp.Fields("CP21")
         SetCPFieldOldData "CP21", m_CP21, 0
      End If
      '2020/10/20 END
      
      SetCPFieldOldData "CP44", textCP44, 0
      'add by nickc 2008/02/01 CF代理人，若是沒有，代一個月內或未發文的有輸過的代理人
      If IsNull(rsTmp.Fields("CP44")) = False Then
         textCP44 = ChangeCustomerL(rsTmp.Fields("CP44"))
         textCP44_Validate False
         SetCPFieldOldData "CP44", textCP44, 0
      Else
            '新案抓同一申請人一個月內，或是未發文的
            If m_CP31 = "Y" Then
                Set rsSubTmp = New ADODB.Recordset
                'edit by nickc 2008/03/19 鎖定相同國家
                'strSubSQL = "SELECT * FROM CaseProgress " & _
                                      "WHERE (cp27 is null or cp27>='" & DBDATE(DateAdd("m", -1, ChangeWStringToWDateString(strSrvDate(1)))) & "' ) and cp44 is not null and (cp01,cp02,cp03,cp04) in (select tm01,tm02,tm03,tm04 from trademark where tm23 in (select tm23 from trademark where tm01='" & m_TM01 & "' and tm02='" & m_TM02 & "' and tm03='" & m_TM03 & "' and tm04='" & m_TM04 & "' )) order by cp05 desc "
                'modify by sonia 2015/12/11 加入未取消收文條件,否則T-201479會抓到T-162444
                strSubSQL = "SELECT * FROM CaseProgress " & _
                                      "WHERE (cp27 is null or cp27>='" & DBDATE(DateAdd("m", -1, ChangeWStringToWDateString(strSrvDate(1)))) & "' ) and cp44 is not null and nvl(cp57,0)=0 and (cp01,cp02,cp03,cp04) in (select tm01,tm02,tm03,tm04 from trademark where (tm23,tm10) in (select tm23,tm10 from trademark where tm01='" & m_TM01 & "' and tm02='" & m_TM02 & "' and tm03='" & m_TM03 & "' and tm04='" & m_TM04 & "' )) order by cp05 desc "
                rsSubTmp.CursorLocation = adUseClient
                rsSubTmp.Open strSubSQL, cnnConnection, adOpenStatic, adLockReadOnly
                If rsSubTmp.RecordCount > 0 Then
                    rsSubTmp.MoveFirst
                    textCP44 = ChangeCustomerL(rsSubTmp.Fields("CP44"))
                    textCP44_Validate False
                End If
            Else
                'add by nickc 2008/03/21 加入台灣不管
                If m_TM10 <> "000" Then
                    Set rsSubTmp = New ADODB.Recordset             'copy from 發文
                    'modify by sonia 2015/12/11 加入未取消收文條件
                    strSubSQL = "SELECT CP44, Max(CP27||CP09) FROM CaseProgress " & _
                                "WHERE CP01 = '" & m_TM01 & "' AND " & _
                                      "CP02 = '" & m_TM02 & "' AND " & _
                                      "CP03 = '" & m_TM03 & "' AND " & _
                                      "CP04 = '" & m_TM04 & "' AND " & _
                                      "CP09 <> '" & m_CP09 & "' And CP09<'C' and nvl(cp57,0)=0 And CP44 Is Not Null Group By CP44 Order By 2 Desc, 1 "
                    rsSubTmp.CursorLocation = adUseClient
                    rsSubTmp.Open strSubSQL, cnnConnection, adOpenStatic, adLockReadOnly
                    If rsSubTmp.RecordCount > 0 Then
                        rsSubTmp.MoveFirst
                        textCP44 = ChangeCustomerL(rsSubTmp.Fields("CP44"))
                        textCP44_Validate False
                    End If
                End If
            End If
      End If
      
      'add by nickc 2006/10/20
      If m_CP31 = "" Then
         lblDivCase.Visible = False
         txtDivCaseNo(0).Visible = False
         txtDivCaseNo(1).Visible = False
         txtDivCaseNo(2).Visible = False
         txtDivCaseNo(3).Visible = False
         txtDivCaseNo(4).Visible = False
      Else
         lblDivCase.Visible = True
         txtDivCaseNo(0).Visible = True
         txtDivCaseNo(1).Visible = True
         txtDivCaseNo(2).Visible = True
         txtDivCaseNo(3).Visible = True
         txtDivCaseNo(4).Visible = True
      End If
      
      '910626 Sieg
      '收據編號
      If IsNull(rsTmp.Fields("CP60")) = False Then
         m_CP60 = rsTmp.Fields("CP60")
      Else
         m_CP60 = ""
      End If
      
      ' 進度備註
      If IsNull(rsTmp.Fields("CP64")) = False Then
         textCP64 = rsTmp.Fields("CP64")
      End If
      SetCPFieldOldData "CP64", textCP64, 0
      ' 對造案件中文名稱
      strTemp = Empty
      If IsNull(rsTmp.Fields("CP37")) = False Then
         strTemp = rsTmp.Fields("CP37")
      End If
      SetCPFieldOldData "CP37", strTemp, 0
      ' 對造案件英文名稱
      strTemp = Empty
      If IsNull(rsTmp.Fields("CP38")) = False Then
         strTemp = rsTmp.Fields("CP38")
      End If
      SetCPFieldOldData "CP38", strTemp, 0
      ' 對造案件日文名稱
      strTemp = Empty
      If IsNull(rsTmp.Fields("CP39")) = False Then
         strTemp = rsTmp.Fields("CP39")
      End If
      SetCPFieldOldData "CP39", strTemp, 0
                                       
        'Add By Cheng 2003/03/05
        '判斷系統類別
        Select Case m_TM01
        Case "T", "TF", "FCT"
            'Modify By Cheng 2004/02/12
            'Mark下段程式因為目前卷宗性質不是1的案件, 其案件名稱基本檔也有存
'            ' 卷宗性質不為1時, 案件中英日文名稱從案件進度檔中帶入
'            If IsEmptyText(m_CP10) = False Then
'               If m_TM28 <> "1" Then
'                  textTM05 = Empty
'                  Me.textTM05_1.Text = Empty
'                  textTM06 = Empty
'                  textTM07 = Empty
'                  Set rsSubTmp = New ADODB.Recordset
'                  strSubSQL = "SELECT * FROM CaseProgress " & _
'                              "WHERE CP01 = '" & m_TM01 & "' AND " & _
'                                    "CP02 = '" & m_TM02 & "' AND " & _
'                                    "CP03 = '" & m_TM03 & "' AND " & _
'                                    "CP04 = '" & m_TM04 & "' AND " & _
'                                    "CP31 = 'Y' "
'                  rsSubTmp.CursorLocation = adUseClient
'                  rsSubTmp.Open strSubSQL, cnnConnection, adOpenStatic, adLockReadOnly
'                  If rsSubTmp.RecordCount > 0 Then
'                     rsSubTmp.MoveFirst
'                    If m_TM01 = "T" Or m_TM01 = "FCT" Or m_TM01 = "CFT" Or m_TM01 = "TF" Then
'                        ' 對造案件名稱
'                        If IsNull(rsSubTmp.Fields("CP37")) = False Then
'                           Me.textTM05_1.Text = rsSubTmp.Fields("CP37")
'                        End If
'                        SetCPFieldOldData "CP37", Me.textTM05_1.Text, 0
'                    Else
'                        ' 對造案件中文名稱
'                        If IsNull(rsSubTmp.Fields("CP37")) = False Then
'                           textTM05 = rsSubTmp.Fields("CP37")
'                        End If
'                        SetCPFieldOldData "CP37", textTM05, 0
'                        ' 對造案件英文名稱
'                        If IsNull(rsSubTmp.Fields("CP38")) = False Then
'                           textTM06 = rsSubTmp.Fields("CP38")
'                        End If
'                        SetCPFieldOldData "CP38", textTM06, 0
'                        ' 對造案件日文名稱
'                        If IsNull(rsSubTmp.Fields("CP39")) = False Then
'                           textTM07 = rsSubTmp.Fields("CP39")
'                        End If
'                        SetCPFieldOldData "CP39", textTM07, 0
'                    End If
'                  End If
'                  rsSubTmp.Close
'                  Set rsSubTmp = Nothing
'               End If
'            End If
            'End
        Case Else
            'Do Nothing
        End Select
   End If
   rsTmp.Close
   Set rsTmp = Nothing
   
   'Add By Sindy 2015/5/13 案件性質為延展, 若收文日<法定期限,畫面上的本所期限及法定期限欄都鎖住,不可修改.
   If m_CP10 = "102" Then
      If Val(textCP05) < Val(textCP07) Then
         textCP06.Enabled = False
         textCP07.Enabled = False
      End If
   End If
   '2015/5/13 END
End Sub

' 讀取資料庫
Public Sub QueryData()
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   Dim arrTmp 'Add by Amy 2022/10/19
   Dim strSubSQL As String, rsSubTmp As ADODB.Recordset   'add by sonia 2023/6/20
   Dim bolTmp As Boolean 'Added by Lydia 2025/09/12
   
   ' 顯示畫面為第一頁
   SSTab1.Tab = 0
   
   textTM29_2 = Empty
   m_TM29 = Empty
   stF0309_Now = Empty 'Add by Amy 2022/10/20
   
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
      'Modify By Sindy 2022/11/23
      'If IsNull(rsTmp.Fields("CP140")) = False Then: txtF0301 = rsTmp.Fields("CP140") 'Add by Amy 2022/10/07
      If IsNull(rsTmp.Fields("CP140")) = False Then: txtF0301 = Pub_GetIsFlowCP140(m_CP09)
      '2022/11/23 END
   End If
   rsTmp.Close
    
   ' 本所案號
   textTMKey = m_TM01 & "-" & m_TM02 & "-" & m_TM03 & "-" & m_TM04
   ' 收文號
   textCP09 = m_CP09
   
   '2007/8/13 ADD BY SONIA銷卷提醒
   CheckCaseDestroy m_TM01, m_TM02, m_TM03, m_TM04
   '2007/8/13 END
      
'   Select Case m_TM01
'      ' 系統類別為CFT的為讀取商標基本檔
'      Case "T", "TF", "FCT":
'         '911120 nick
'         textTM24.Locked = False
'         textTM24.BackColor = &H80000005
'         textTM25.Locked = False
'         textTM25.BackColor = &H80000005
'         textTM26.Locked = False
'         textTM26.BackColor = &H80000005
'         textTM28.Locked = False
'         textTM28.BackColor = &H80000005
'         textTM08.Locked = False
'         textTM08.BackColor = &H80000005
'         textSP32.Locked = True
'         textSP32.BackColor = &H8000000F
'         QueryTradeMark
'         '92.10.31 ADD BY SONIA
'         textTM05.MaxLength = 40
'         textTM07.MaxLength = 40
'         '92.10.31 END
'        Me.Label13.Visible = False
'        Me.textTM05.Visible = False
'        Me.textTM05.Enabled = False
'        Me.Label12.Visible = False
'        Me.textTM06.Visible = False
'        Me.textTM06.Enabled = False
'        Me.Label11.Visible = False
'        Me.textTM07.Visible = False
'        Me.textTM07.Enabled = False
'        Me.Label42.Visible = True
'        Me.textTM05_1.Visible = True
'        Me.textTM05_1.Enabled = True
'      Case Else:
'         '911120 nick
'         textTM24.Locked = True
'         textTM24.BackColor = &H8000000F
'         textTM25.Locked = True
'         textTM25.BackColor = &H8000000F
'         textTM26.Locked = True
'         textTM26.BackColor = &H8000000F
'         textTM28.Locked = True
'         textTM28.BackColor = &H8000000F
'         textTM08.Locked = True
'         textTM08.BackColor = &H8000000F
'         If UCase(m_TM01) = "TM" Then
'             textSP32.Locked = False
'             textSP32.BackColor = &H80000005
'         Else
'             textSP32.Locked = True
'             textSP32.BackColor = &H8000000F
'         End If
'         QueryServicePractice
'         '92.10.31 ADD BY SONIA
'         textTM05.MaxLength = 60
'         textTM07.MaxLength = 60
'         '92.10.31 END
'        'Modify By Cheng 2004/02/24
'        '查名案件性質名稱合併成一欄
'        Select Case m_TM01
'        Case "TS"
'            Me.Label13.Visible = False
'            Me.textTM05.Visible = False
'            Me.textTM05.Enabled = False
'            Me.Label12.Visible = False
'            Me.textTM06.Visible = False
'            Me.textTM06.Enabled = False
'            Me.Label11.Visible = False
'            Me.textTM07.Visible = False
'            Me.textTM07.Enabled = False
'            Me.Label42.Visible = True
'            Me.textTM05_1.Visible = True
'            Me.textTM05_1.Enabled = True
'        Case Else
'            Me.Label13.Visible = True
'            Me.textTM05.Visible = True
'            Me.textTM05.Enabled = True
'            Me.Label12.Visible = True
'            Me.textTM06.Visible = True
'            Me.textTM06.Enabled = True
'            Me.Label11.Visible = True
'            Me.textTM07.Visible = True
'            Me.textTM07.Enabled = False
'            Me.Label42.Visible = False
'            Me.textTM05_1.Visible = False
'            Me.textTM05_1.Enabled = False
'        End Select
'        'End
'   End Select
   'Modify By Sindy 2012/6/1 把查詢基本檔的程式寫到此函數裡,共用之
   Call QueryMainFile
   
   ' 取得案件進度檔的欄位
   QueryCaseProgress
   'Add By Sindy 2021/3/23
   ReDim cp(TF_CP)
   cp(9) = m_CP09
   Call PUB_ReadCaseProgressDatabase(cp(), 1)
   '2021/3/23 END
   
   'Modify By Sindy 2014/1/29
   m_CP31isYGetCP05 = GetCP31isY_CP05(m_TM01, m_TM02, m_TM03, m_TM04) '取得本所案號新案件的收文日
   'Add By Sindy 2013/12/16
   textTM130.Visible = False
   lblTM130.Visible = False
   'If strSrvDate(1) >= InvoiceStartDate Then
   If Val(m_CP31isYGetCP05) >= Val(InvoiceStartDate) Then
      'Modify By Sindy 2014/2/10 改非台灣新案都可以收J智權公司
      'If m_TM01 = "T" And m_CP31 = "Y" And textTM10 = "020" Then
      If m_CP31 = "Y" And textTM10 <> "000" Then
      '2014/2/10 END
         textTM130.Visible = True
         lblTM130.Visible = True
         'Add by Amy 2018/07/03 收文日在一個月之內才可修改 特殊出名公司
         textTM130.Enabled = False
         If Val(strSrvDate(1)) >= Val(textCP05) + 19110000 And Val(strSrvDate(1)) <= Val(DBDATE(DateAdd("m", 1, Format(Val(textCP05) + 19110000, "####/##/##")))) Then
              textTM130.Enabled = True
         End If
         'end 2018/07/03
      End If
   End If
   '2013/12/16 END
   
   ' 是否閉卷
   If m_TM29 = "Y" Then
      EnableTextBox textTM29, True
      textTM29_2 = "已閉卷"
   Else
      EnableTextBox textTM29, False
      textTM29_2 = Empty
   End If
   
   ' 更新指定國家按鈕狀態
   If m_TM01 = "TF" And (m_CP10 = "101" Or m_CP10 = "104") Then
      cmdNation.Enabled = True
      cmdNation.Visible = True
      cmdNation.TabStop = True
      'add by sonia 2023/6/20 子案已有進度資料，不可再使用指定國家按鈕，否則若增減子案，會造成基本檔與進度檔不符
      Set rsSubTmp = New ADODB.Recordset
      strSubSQL = "SELECT * FROM CaseProgress " & _
                            "WHERE cp01='" & m_TM01 & "' and cp02='" & m_TM02 & "' and cp03||cp04<>'000'"
      rsSubTmp.CursorLocation = adUseClient
      rsSubTmp.Open strSubSQL, cnnConnection, adOpenStatic, adLockReadOnly
      If rsSubTmp.RecordCount > 0 Then
         MsgBox "子案已有進度資料，不得再增減指定國家!!!", vbExclamation + vbOKOnly
         cmdNation.Enabled = False
         cmdNation.TabStop = False
      End If
      'end 2023/6/20
   Else
      cmdNation.Enabled = False
      cmdNation.Visible = False
      cmdNation.TabStop = False
   End If
   
   
   ' 更新第二及第三申請人的狀態
'edit by nickc    2006/12/14
'   If m_TM01 = "TC" Then
      EnableTextBox textSP58, True
      EnableTextBox textSP59, True
      'add by nickc 2006/12/14
      EnableTextBox textTM80, True
      EnableTextBox textTM81, True
      
'edit by nickc 2006/12/14
'   Else
'      EnableTextBox textSP58, False
'      EnableTextBox textSP59, False
'   End If
   
   ' 依讀取的是商標基本檔還是服務業務基本檔來更新控制項的狀態
   'add by nickc 2006/12/14
   EnableTextBox textTM09, True
   EnableTextBox textTM32, True
   
   Select Case m_TM01
      Case "T", "TF", "FCT":
      'edit by nickc 2006/12/14
'         EnableTextBox textTM09, True
         EnableTextBox textSP32, False
      Case Else:
'edit by nickc 2006/12/14
'         EnableTextBox textTM09, False
         '911120 nick
         'EnableTextBox textSP32, True
   End Select
   ' 讀取優先權資料
   m_Pa(1) = m_TM01
   m_Pa(2) = m_TM02
   m_Pa(3) = m_TM03
   m_Pa(4) = m_TM04
   
   'Modify by Amy 2014/04/17 +, m_Priority(4), m_Priority(5)
   'edit by nickc 2007/02/06 不用 dll 了 objPublicData.ReadPriority m_Pa, m_Priority(1), m_Priority(2), m_Priority(3)
   'Modify by Sindy 2017/10/12 +, m_Priority(6)
   'Modify by Amy 2023/01/06 原m_Priority為陣列
   ClsPDReadPriority m_Pa, m_Priority1, m_Priority2, m_Priority3, m_Priority4, m_Priority5, m_Priority6
   ' 更新本案期限的資料
   UpdateGrdList m_TM01, m_TM02, m_TM03, m_TM04
    'Add By Cheng 2004/04/14
    '顯示此分割案的母案資料
    If m_CP10 = "308" Then
        Me.lblDivCase.Visible = True
        Me.txtDivCaseNo(0).Visible = True: Me.txtDivCaseNo(0).Enabled = True
        Me.txtDivCaseNo(1).Visible = True: Me.txtDivCaseNo(1).Enabled = True
        Me.txtDivCaseNo(2).Visible = False: Me.txtDivCaseNo(2).Enabled = False
        Me.txtDivCaseNo(3).Visible = True: Me.txtDivCaseNo(3).Enabled = True
        Me.txtDivCaseNo(4).Visible = True: Me.txtDivCaseNo(4).Enabled = True
        ShowOriginCaseNo m_TM01, m_TM02, m_TM03, m_TM04
    Else
        Me.lblDivCase.Visible = False
        Me.txtDivCaseNo(0).Visible = False: Me.txtDivCaseNo(0).Enabled = False
        Me.txtDivCaseNo(1).Visible = False: Me.txtDivCaseNo(1).Enabled = False
        Me.txtDivCaseNo(2).Visible = False: Me.txtDivCaseNo(2).Enabled = False
        Me.txtDivCaseNo(3).Visible = False: Me.txtDivCaseNo(3).Enabled = False
        Me.txtDivCaseNo(4).Visible = False: Me.txtDivCaseNo(4).Enabled = False
    End If
    'End
   ' 設定輸入的位置
   SetInputEntry
   ' 顯示商標基本資料的畫面
   'modify by sonia 90.11.14發文才加
   'Modify By Sindy 2010/10/15 將此段程式解開
   ''ShowMaintainForm m_CP09
   'Modify By Sindy 2012/6/1 +me
   ShowMaintainForm m_CP09, "Y", "分案", Me
   '2010/10/15 End
   
   'Add by Morgan 2003/12/07
   Call PUB_CheckSales(m_TM01, m_TM02, m_TM03, m_TM04, textCP05, textCP13, textCP13_2)
   'End 2003/12/07
   'Add By Cheng 2004/05/13
   '若非C類來函, Enable轉本所案號欄位
   'Modify By Sindy 2012/6/1 C類來函或已發文案件須鎖住轉本所案號欄位, 若為併號請以聯絡單通知電腦中心處理
   'If m_CP09 < "C" Then
   If m_CP09 < "C" And Val(m_CP27) = 0 Then
   '2012/6/1 End
       Me.textTM01.Enabled = True
       Me.textTM02.Enabled = True
       Me.textTM02_2.Enabled = True
       Me.textTM03.Enabled = True
       Me.textTM04.Enabled = True
   '若為C類來函, Disable轉本所案號欄位
   Else
       Me.textTM01.Enabled = False
       Me.textTM02.Enabled = False
       Me.textTM02_2.Enabled = False
       Me.textTM03.Enabled = False
       Me.textTM04.Enabled = False
   End If
   'End
   
   'Add By Sindy 2012/5/8
   '台灣商標Ｔ,FCT案若收文爭議案件性質時,開放Frame21欄位
   'Modified by Lydia 2022/07/15 移出為獨立函數
'   Frame21.Visible = False
'   m_EP06 = "": m_EP06DT = "": textEP34.Enabled = True
'   'Modified by Lydia 2018/12/10 T台灣案填寫接洽單管控文件及查名是否齊備
'   'If (m_TM01 = "T" Or m_TM01 = "FCT") And textTM10 = "000" And InStr(TMdebate, textCP10) > 0 And DBDATE(textCP05) >= TMdebateStarDT Then
'   Label57.Visible = False: textEP34.Visible = False '預設會稿、查名不顯示
'   Label65.Visible = False: textCP143.Visible = False
'   If ((m_TM01 = "T" Or m_TM01 = "FCT") And textTM10 = "000" And InStr(TMdebate, textCP10) > 0 And DBDATE(textCP05) >= TMdebateStarDT) _
'          Or (m_TM01 = "T" And textTM10 = "000" And DBDATE(textCP05) >= T案收文齊備啟用日) Then
'   'end 2018/12/10
'      Frame21.Visible = True
'      'Added by Lydia 2018/12/10 區分商爭和商申
'      If m_TM01 = "T" Then
'            If InStr(TMdebate, textCP10) > 0 Then   '商爭
'                Label57.Visible = True: textEP34.Visible = True
'                Label64.Caption = "資料是否齊備：       (Y/N)"
'            Else  '商申
'                 If textCP10 = 申請 Then  '商申
'                    Label65.Visible = True: textCP143.Visible = True
'                    'Added by Lydia 2019/01/30
'                    'Modified by Lydia 2020/11/04 收文之查名齊備日; m_CP143=>p_strCP143
'                    If Val(p_strCP143) = 0 Then
'                        textCP143.Text = "N"
'                    Else
'                        textCP143.Text = "Y"
'                    End If
'                    p_strCP143 = textCP143.Text
'                    Call textCP143_Validate(False) '檢查-查名是否齊備
'                    'end 2019/01/30
'                 End If
'                 Label64.Caption = "文件是否齊備：       (Y/N)"
'            End If
'      ElseIf m_TM01 = "FCT" Then
'            Label57.Visible = True: textEP34.Visible = True
'            Label64.Caption = "資料是否齊備：       (Y/N)"
'      End If
'      'end 2018/12/10
'      '讀取資料
'      strSql = "SELECT ep06,ep34 FROM engineerprogress WHERE ep02='" & Trim(textCP09) & "'"
'      intI = 1
'      Set RsTemp = ClsLawReadRstMsg(intI, strSql)
'      If intI = 1 Then
'         If Not IsNull(RsTemp.Fields(0)) Then
'            If RsTemp.Fields(0) > 0 Then
'               m_EP06DT = RsTemp.Fields(0)
'               textEP06.Text = "Y"
'            Else
'               textEP06.Text = "N"
'            End If
'         End If
'         If Not IsNull(RsTemp.Fields(1)) Then
'            textEP34.Text = RsTemp.Fields(1)
'         End If
'         '案件性質為613補充答辯或612補充理由時，則只可不會稿
'         If Trim(textCP10) = "613" Or _
'            Trim(textCP10) = "612" Then
'            If textEP34.Text = "N" Then 'Add By Sindy 2013/3/12 +if
'               textEP34.Enabled = False
'            End If
'         End If
'      End If
'      m_EP06 = textEP06
'   End If
'   '2012/5/8 End
'   'Added by Lydia 2019/04/11 非爭議案(A類)之T案收文齊備排除的案件性質,預設文件齊備=Y
'   If textEP06.Visible = True And Left(m_CP09, 1) = "A" And InStr(T案收文齊備排除, m_CP10) > 0 Then
'       textEP06 = "Y"
'   End If
'   'end 2019/04/11
'
'   textEP06.Tag = textEP06.Text 'Added by Lydia 2019/07/29 記錄預設(文件是否齊備)
  Call setFrame21
   'end 2022/07/15
   
   '2013/10/31 add by sonia 非台灣新申請案收費0,第一次分案時要提醒二案案件備註加註同時合併計算結餘" T-189182(T-188512)
   'Modified by Lydia 2022/03/09 'Modified by Lydia 2022/03/09 改判斷分案日;  ex.2021/06/18 CFT案件承辦人若空白時，預設為國家檔之CFT承辦人---統一判斷
   'If textTM10 <> "000" And textCP10 = "101" And textCP14 = "" And Val(m_CP16) = 0 Then
   If textTM10 <> "000" And textCP10 = "101" And Val(cp(149)) = 0 And Val(m_CP16) = 0 Then
      MsgBox "此新申請案未收費, 若有前案則請至第三頁頁籤之案件備註欄加註與前案號合併計算結餘(前案之案件備註也要加註)!!!", vbExclamation + vbOKOnly
      SSTab1.Tab = 2
      textTM58_GotFocus
      textTM58.SetFocus
   End If
   '2013/10/31 end
    
   'Added by Lydia 2020/05/20 法律所案源收文
   Call ReadLOS
   Call SetLOSagree
   
   'Modify by Amy 2022/10/20 +簽核頁籤,接洽單電子收文才顯示「檢視接洽單」鈕
   cmdFile.Visible = False
   SSTab1.TabVisible(4) = False
   Label69.Visible = False: txtF0309.Visible = False 'Add by Amy 2022/11/17 目前狀態
   'Modify by Amy 2023/01/03 8碼(結案單)不可開接洽單會錯: + And Len(txtF0301) = 10
   If strSrvDate(1) >= 接洽單電子收文啟用日 And txtF0301 <> MsgText(601) And Len(txtF0301) = 10 Then
        Label58.BackStyle = 0 'Add by Amy 2022/11/17 急件 : Label58.BackColor = &H8080FF
        cmdFile.Visible = True
        '補件完成 欄-案件表單流程備註檔屬於分案作業相關資訊
        SetFlow004TextBox txtF0407, txtF0301, " And F0408 in('A6','A7') And F0409 in('A6','A7') "
        '案件表單簽核檔
        strSql = "SELECT ST02||nvl(F0208,'') 簽核人員,decode(F0202," & ShowFlow簽核人員身份 & ") 身份,sqldateT(F0205) 日期,sqltime6(F0206) 時間,decode(F0207," & ShowFlow簽核結果 & ") 簽核結果,F0204 FROM FLOW002,Staff WHERE F0201='" & txtF0301 & "' and F0204=ST01(+) order by decode(F0205,null,2,1) asc,F0205||sqltime6(F0206) asc,F0202,F0203 asc"
        rsTmp.CursorLocation = adUseClient
        rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
        If rsTmp.RecordCount > 0 Then
           Set GRD1.Recordset = rsTmp
           SetGrd
        End If
        cmdOK.Enabled = False: CmdAddInfo.Enabled = False
        txtNote.Locked = True
        '北所分案人員已同意,才可按 確定 鈕
        If ChkConultRecFlow002(Me.Name, txtF0301, "A6", IsEConsultRec, stF0207_A6) = True Then
             cmdOK.Enabled = True
        End If
        arrTmp = Split(GetFlow003Data(txtF0301, , "F0308||';'||Nvl(F0309,'NULL')"), ";")
        stF0309_Now = arrTmp(1)
        'Add by Amy 2022/11/17 +表單狀態/急件
        'Modify by Sindy 2022/11/22
        txtF0309 = PUB_GetCP157forF0309(cp(9)) '表單狀態
        Label69.Visible = True: txtF0309.Visible = True
        'end 2022/11/17
        
        '補件完成 鈕
        '下一處理人員是A7(多筆案件性質已處理一筆)且目前表單狀態不是已分案
        If arrTmp(0) = "A7" And stF0309_Now <> "17" Then
            CmdAddInfo.Enabled = True
            txtNote.Locked = False
        End If
        '接洽單電子收文顯示簽核頁籤
        SSTab1.TabVisible(4) = True
        'Add by Amy 2022/11/17 狀態為 程序補件 時,切至 簽核 頁籤
        If stF0309_Now = "20" Then SSTab1.Tab = 4
        'Add by Amy 2023/01/07 直接開啟接洽單-桂英
        frm090801_Q.SetParent Me
        frm090801_Q.m_blnCallPrint = True
        frm090801_Q.Text5 = txtF0301
        Call frm090801_Q.cmdok_Click(4)
        frm090801_Q.Show
        'end 2023/01/07
   End If
   'end 2022/10/20
   
   'Add By Sindy 2024/1/30 各部門分案時，若本所期限與法定期限與接洽單的本所期限與法定期限不同時，要提醒
   Call PUB_ChkCRLdtCP06CP07(m_CP09)
   
   'Added by Lydia 2025/08/21
   If m_TM01 = "T" And m_bolIsFirstKeyCP14 = True And textTM10 = "000" And textCP10 = "707" Then
      MsgBox "台灣707調查新案號，請確認卷宗性質，目前預設為4廢止。", vbInformation
   End If
   'end 2025/08/21

   'Added by Lydia 2025/09/12 TF基礎案號設定：TF案未閉卷(無專用期 or 專用期未過)，卷宗性質為申請之母案案號，即TF-xxxxx0-0-00
   'Modified by Lydia 2025/10/23 TF基礎案號(TM06,TM07)改成可以輸入多筆(Table: TFBaseNo)，原本的輸入欄位直接刪除改成按鈕呼叫其他表單，若已有設定則按鈕設為綠色。
   cmdTFBaseNo.Visible = False
   cmdTFBaseNo.BackColor = &H8000000F
   'Modified by Lydia 2025/10/23 不要限製制新案才能設定，拿掉m_CP31 = "Y" And---雅雯
   If m_TM01 = "TF" And Mid(m_TM02, 6, 1) = "0" And m_TM03 = "0" And m_TM04 = "00" And textTM28 = "1" And _
       m_TM29 = "" And (Trim(textTM22S) = "" Or (textTM22S <> "" And DBDATE(textTM22S) >= strSrvDate(1))) Then
       cmdTFBaseNo.Visible = True
       strExc(0) = Pub_GetField("TFBaseNo", "TFBN01='" & m_TM01 & "' AND TFBN02='" & m_TM02 & "' AND TFBN03='" & m_TM03 & "' AND TFBN04='" & m_TM04 & "'", "TFBN05")
       If strExc(0) <> "" Then
          cmdTFBaseNo.BackColor = &HC0FFC0
       Else
          cmdTFBaseNo.BackColor = &H8000000F
       End If
   End If
   'end 2025/09/12
End Sub

'Add By Sindy 2012/6/1 為防使用者在前基本檔維護作業有修改資料, 因此基本檔資料再重新讀取一次
Public Sub QueryMainFile()
   Select Case m_TM01
      ' 系統類別為CFT的為讀取商標基本檔
      Case "T", "TF", "FCT":
         '911120 nick
         textTM24.Locked = False
         textTM24.BackColor = &H80000005
         textTM25.Locked = False
         textTM25.BackColor = &H80000005
         textTM26.Locked = False
         textTM26.BackColor = &H80000005
         textTM28.Locked = False
         textTM28.BackColor = &H80000005
         textTM08.Locked = False
         textTM08.BackColor = &H80000005
         'Add By Sindy 2019/4/9
         textTM72.Locked = False
         textTM72.BackColor = &H80000005
         '2019/4/9 END
         cboTM08.Locked = False: cboTM72.Locked = False  'Added by Lydia 2023/11/16
         textSP32.Locked = True
         textSP32.BackColor = &H8000000F
         QueryTradeMark
         '92.10.31 ADD BY SONIA
         textTM05.MaxLength = 40
         textTM07.MaxLength = 40
         '92.10.31 END
         Me.Label13.Visible = False
         Me.textTM05.Visible = False
         Me.textTM05.Enabled = False
         Me.Label12.Visible = False
         Me.textTM06.Visible = False
         Me.textTM06.Enabled = False
         Me.Label11.Visible = False
         Me.textTM07.Visible = False
         Me.textTM07.Enabled = False
         Me.Label42.Visible = True
         Me.textTM05_1.Visible = True
         Me.textTM05_1.Enabled = True
      Case Else:
         '911120 nick
         textTM24.Locked = True
         textTM24.BackColor = &H8000000F
         textTM25.Locked = True
         textTM25.BackColor = &H8000000F
         textTM26.Locked = True
         textTM26.BackColor = &H8000000F
         textTM28.Locked = True
         textTM28.BackColor = &H8000000F
         textTM08.Locked = True
         textTM08.BackColor = &H8000000F
         If UCase(m_TM01) = "TM" Then
             textSP32.Locked = False
             textSP32.BackColor = &H80000005
         Else
             textSP32.Locked = True
             textSP32.BackColor = &H8000000F
         End If
         cboTM08.Locked = True: cboTM72.Locked = True  'Added by Lydia 2023/11/16
         QueryServicePractice
         '92.10.31 ADD BY SONIA
         textTM05.MaxLength = 60
         textTM07.MaxLength = 60
         '92.10.31 END
        'Modify By Cheng 2004/02/24
        '查名案件性質名稱合併成一欄
        Select Case m_TM01
        Case "TS"
            Me.Label13.Visible = False
            Me.textTM05.Visible = False
            Me.textTM05.Enabled = False
            Me.Label12.Visible = False
            Me.textTM06.Visible = False
            Me.textTM06.Enabled = False
            Me.Label11.Visible = False
            Me.textTM07.Visible = False
            Me.textTM07.Enabled = False
            Me.Label42.Visible = True
            Me.textTM05_1.Visible = True
            Me.textTM05_1.Enabled = True
        Case Else
            Me.Label13.Visible = True
            Me.textTM05.Visible = True
            Me.textTM05.Enabled = True
            Me.Label12.Visible = True
            Me.textTM06.Visible = True
            Me.textTM06.Enabled = True
            Me.Label11.Visible = True
            Me.textTM07.Visible = True
            Me.textTM07.Enabled = False
            Me.Label42.Visible = False
            Me.textTM05_1.Visible = False
            Me.textTM05_1.Enabled = False
        End Select
        'End
   End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
   'Add by Amy 2022/11/17
   If PUB_CheckFormExist("frm090801_Q") = True Then
        Unload frm090801_Q
   End If
   PUB_SendMailCache 'Add by Sindy 2010/6/18
   'Add by Amy 2023/01/06 frm880002從此支開啟改不為強制表單,故需判斷存在時要關
   m_Priority1 = "": m_Priority2 = "": m_Priority3 = "": m_Priority4 = "": m_Priority5 = "": m_Priority6 = ""
   m_CP09 = Empty
   'Add By Cheng 2002/07/18
   Set frm020101_02 = Nothing
End Sub

Private Sub grdList_Click()
      If grdList.row > 0 Then
         grdList.col = 0
         If grdList.Text = "V" Then
            grdList.Text = Empty
         Else
            'Added by Lydia 2021/08/31 各系統之分案作業和內部收文作業：勾選下一程序的期限，且該收文的案件性質與下一程序的案件性質不同，請SHOW訊息提醒
            If Pub_CheckNpTheSameShow(m_TM01, textCP10, Trim("" & grdList.TextMatrix(grdList.row, 9))) = False Then
                Exit Sub
            End If
            'end 2021/08/31
            grdList.Text = "V"
         End If
         'Add By Cheng 2002/11/18
         PasteGridData
      End If
End Sub

'Add By Sindy 2023/4/21
Private Sub OptSendType_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
   Dim oOpt As OptionButton
   If OptSendType(Index).Tag = "1" Then
      OptSendType(Index).Value = False
      OptSendType(Index).Tag = "0"
      If Index = 3 Then
         textCP142.Text = ""
         textCP142.Enabled = False
         'Add By Sindy 2024/1/22
         If Frame3.Visible = True Then
            Option1(0).Value = False
            Option1(1).Value = False
            Option1(2).Value = False
            Option1(0).Enabled = False
            Option1(1).Enabled = False
            Option1(2).Enabled = False
         End If
         '2024/1/22 END
      End If
      
   Else
      For Each oOpt In OptSendType
         If oOpt.Index = Index Then
            oOpt.Tag = "1"
         Else
            oOpt.Tag = "0"
         End If
      Next
      'Add By Sindy 2024/1/22
      'If Index = 3 Then
      If Index = 3 And OptSendType(Index).Value Then
      '2024/1/22 END
         textCP142.Enabled = True
         textCP142.SetFocus
         'Add By Sindy 2024/1/22
         If Frame3.Visible Then
            Option1(0).Enabled = True
            Option1(1).Enabled = True
            Option1(2).Enabled = True
         End If
         '2024/1/22 END
      Else
         textCP142.Text = ""
         textCP142.Enabled = False
         'Add By Sindy 2024/1/22
         If Frame3.Visible Then
            Option1(0).Value = False
            Option1(1).Value = False
            Option1(2).Value = False
            Option1(0).Enabled = False
            Option1(1).Enabled = False
            Option1(2).Enabled = False
         End If
         '2024/1/22 END
      End If
   End If
End Sub

Private Sub Text1_GotFocus()
    TextInverse Me.Text1
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text2_GotFocus()
    TextInverse Me.Text2
End Sub

Private Sub Text3_GotFocus()
    TextInverse Me.Text3
End Sub

Private Sub Text4_GotFocus()
    TextInverse Me.Text4
End Sub

Private Sub Text4_LostFocus()
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
    
    'Add By Cheng 2003/06/16
    '若有輸入查名本所案號
    If Me.Text1.Text <> "" And Me.Text2.Text <> "" Then
        StrSQLa = "Select * From ServicePractice Where " & ChgService(Me.Text1.Text & Me.Text2.Text & Left(Me.Text3.Text & "0", 1) & Left(Me.Text4.Text & "00", 2))
        rsA.CursorLocation = adUseClient
        rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
        If rsA.RecordCount <= 0 Then
            MsgBox "您輸入的查名本所案號錯誤，請重新輸入!!!", vbExclamation + vbOKOnly
            Me.Text1.SetFocus
            Text1_GotFocus
        End If
        If rsA.State <> adStateClosed Then rsA.Close
        Set rsA = Nothing
    End If
End Sub

' 收文日
Private Sub textCP05_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textCP05) = False Then
      If CheckIsTaiwanDate(textCP05, False) = False Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "收文日日期格式不正確"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      End If
   End If
End Sub

' 本所期限
Private Sub textCP06_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textCP06) = False Then
      If CheckIsTaiwanDate(textCP06, False) = False Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "本所期限的日期格式不正確"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      'Added by Lydia 2020/07/07 本所期限檢查：若本所期限非工作天則直接調整至最近的工作天
      Else
          textCP06.Text = TransDate(PUB_GetWorkDay1(textCP06, True), 1)
      'end 2020/07/07
      End If
   End If
End Sub

' 法定期限
Private Sub textCP07_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textCP07) = False Then
      If CheckIsTaiwanDate(textCP07, False) = False Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "法定期限的日期格式不正確"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      End If
   End If
End Sub

' 案件性質
Private Sub textCP10_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   textCP10_2 = Empty
   Cancel = False
   If IsEmptyText(textCP10) = False Then
      If m_TM10 < "010" Then
         ' 取得國內的案件性質名稱
         textCP10_2 = GetCaseTypeName(m_TM01, textCP10, 0)
      Else
         textCP10_2 = GetCaseTypeName(m_TM01, textCP10, 1)
      End If
      If IsEmptyText(textCP10_2) = True Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "案件性質代號不存在"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      'Added by Lydia 2020/05/20 法律所案源收文
      ElseIf m_CP10 <> textCP10 Then
          SetLOSagree
         'Added by Lydia 2022/07/15 T案之齊備日管控; TC案之文件齊備日管控
         If m_CP10 <> "" And textCP05 <> "" Then 'Modfied by Lydia 2022/07/21 +收文日非空白；因為連續分案遇到更換國家時，尚未傳入
             Call setFrame21
         End If
         'end 2022/07/15
      'end 2020/05/20
      End If
      SetFrame1 'Added by Morgan 2022/12/15
   End If
End Sub

'Add By Sindy 2010/11/26
Private Sub textCP13_KeyPress(KeyAscii As Integer)
KeyAscii = UpperCase(KeyAscii)
End Sub

' 智權人員
Private Sub textCP13_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Dim rsTmp As New ADODB.Recordset
   'Added by Lydia 2019/02/14
   Dim m_Tuser As String '創新業務部預設收文人員
   
   Cancel = False
   textCP13_2 = Empty
   If IsEmptyText(textCP13) = False Then
      textCP13_2 = GetStaffName(textCP13)
      If IsEmptyText(textCP13_2) = True Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "智權人員代號不存在"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      'Added by Lydia 2019/02/14
         SSTab1.Tab = 0
         textCP13.SetFocus
         Call textCP13_GotFocus
         Exit Sub
      '創新業務部人員收文控管
      Else
         m_SalesST15 = GetST15(textCP13)
         If PUB_ChkIsT10T20("2", textCP13.Text, m_Tuser, strTit) = True Then
             SSTab1.Tab = 0
             textCP13.Text = m_Tuser
             textCP13_2.Text = strTit
             textCP13.SetFocus
             Call textCP13_GotFocus
             Cancel = True
             Exit Sub
         End If
      'end 2019/02/14
      End If
   End If
End Sub

'Add By Sindy 2010/11/26
Private Sub textCP14_KeyPress(KeyAscii As Integer)
KeyAscii = UpperCase(KeyAscii)
End Sub

' 承辦人
Private Sub textCP14_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Dim rsTmp As New ADODB.Recordset
   
   Cancel = False
   textCP14_2 = Empty
   If IsEmptyText(textCP14) = False Then
      textCP14_2 = GetStaffName(textCP14)
      If IsEmptyText(textCP14_2) = True Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "承辦人代號不存在"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      End If
   End If
End Sub

Private Sub textCP142_GotFocus()
   TextInverse textCP142
End Sub

Private Sub textCP142_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 8 And Not IsNumeric(Chr(KeyAscii)) Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub textCP142_Validate(Cancel As Boolean)
   If textCP142 <> "" Then
      If ChkDate(textCP142) = False Then
         Cancel = True
         textCP142.SetFocus
         textCP142_GotFocus
'      ElseIf Val(textCP06) > 0 And Val(textCP142) > Val(textCP06) Then
'         MsgBox "指定送件日期不可晚於本所期限！"
'         Cancel = True
      ElseIf Not ChkWorkDay(DBDATE(textCP142)) Then
         MsgBox "指定送件日期必須是工作天 !", vbExclamation, "輸入指定日期錯誤"
         Cancel = True
         textCP142.SetFocus
         textCP142_GotFocus
      'Add By Sindy 2023/12/11
      ElseIf Val(textCP142) < Val(strSrvDate(2)) Then
         MsgBox "指定日期不可小於系統日！", vbExclamation, "輸入指定日期錯誤"
         Cancel = True
         textCP142.SetFocus
         textCP142_GotFocus
      '指定日期不可大於法定期限
      ElseIf textCP142 <> "" And textCP07 <> "" And Val(textCP142) > Val(textCP07) Then
         MsgBox "指定日期不可大於法定期限！", vbExclamation, "輸入指定日期錯誤"
         Cancel = True
         textCP142.SetFocus
         textCP142_GotFocus
      '2023/12/11 END
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

' 相關總收文號
Private Sub textCP43_Validate(Cancel As Boolean)
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textCP43) = False Then
      If textCP43 = m_CP09 Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "相關總收文號不可為本收文號"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         GoTo EXITSUB
      End If
      
      '2008/11/18 ADD BY SONIA,T-156581
      If textCP10 = "313" And Mid(textCP43, 1, 1) <> "C" Then
         Cancel = True
         strTit = "資料檢核"
         'Modified by Lydia 2024/12/25 改成提醒不限制
         'strMsg = "減縮商品案, 請輸入C類相關總收文號"
         'nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         If MsgBox("減縮商品案, 是否要輸入C類相關總收文號？", vbInformation + vbYesNo + vbDefaultButton1) = vbYes Then
         'end 2024/12/25
            GoTo EXITSUB
         End If
      End If
      '2008/11/18 end
      
      strSql = "SELECT * FROM CaseProgress " & _
               "WHERE CP01 = '" & m_TM01 & "' AND " & _
                     "CP02 = '" & m_TM02 & "' AND " & _
                     "CP03 = '" & m_TM03 & "' AND " & _
                     "CP04 = '" & m_TM04 & "' AND " & _
                     "CP09 = '" & textCP43 & "' "
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount <= 0 Then
         rsTmp.Close
         Cancel = True
         strTit = "資料檢核"
         strMsg = "相關總收文號資料不存在"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         GoTo EXITSUB
      End If
      
      'add by sonia 2018/10/29 自請撤回306, 自請撤銷307之相關總收文號不可選延期FCT-042719
      If (textCP10 = "306" Or textCP10 = "307") And rsTmp.Fields("CP10") = "303" Then
         rsTmp.Close
         Cancel = True
         strTit = "資料檢核"
         strMsg = "自請撤回306, 自請撤銷307之相關總收文號不可選延期 !"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         Me.textCP43.SetFocus
         GoTo EXITSUB
      End If
      'end 2018/10/29
      rsTmp.Close
   End If
EXITSUB:
   Set rsTmp = Nothing
End Sub

' 進度備註
Private Sub textCP64_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If CheckLengthIsOK(textCP64, textCP64.MaxLength) = False Then
      Cancel = True
      strTit = "檢核資料"
      strMsg = "進度備註內容太長"
      'nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textCP64_GotFocus
   End If
End Sub

'Add By Sindy 2012/5/8
Private Sub textEP06_GotFocus()
   TextInverse textEP06
End Sub
Private Sub textEP06_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 8 And KeyAscii <> Asc("N") And KeyAscii <> Asc("Y") Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub textEP34_GotFocus()
   TextInverse textEP34
End Sub
Private Sub textEP34_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 8 And KeyAscii <> Asc("N") And KeyAscii <> Asc("Y") Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub textCP122_GotFocus()
   TextInverse textCP122
End Sub
Private Sub textCP122_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 8 And KeyAscii <> Asc("N") And KeyAscii <> Asc("Y") Then
      KeyAscii = 0
      Beep
   End If
End Sub
'2012/5/8 End

' 商標審定號
Private Sub textSP32_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   
   If IsEmptyText(textSP32) = False Then
      'Modify By Cheng 2002/07/15
'      strSQL = "SELECT * FROM TradeMark " & _
'               "WHERE TM15 = '" & textSP32 & "' "
      strSql = "SELECT * FROM TradeMark " & _
               "WHERE TM15 = '" & textSP32 & "' AND TM16='1' "
      
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         rsTmp.MoveFirst
'         ' 案件中文名稱
         ' 案件名稱
         If IsNull(rsTmp.Fields("TM05")) = False Then
            textTM05S = rsTmp.Fields("TM05")
         End If
'         ' 案件英文名稱
'         If IsNull(rsTmp.Fields("TM06")) = False Then
'            textTM06S = rsTmp.Fields("TM06")
'         End If
'         ' 案件日文名稱
'         If IsNull(rsTmp.Fields("TM07")) = False Then
'            textTM07S = rsTmp.Fields("TM07")
'         End If
         ' 專用期間
         If IsNull(rsTmp.Fields("TM21")) = False Then
            textTM21S = TAIWANDATE(rsTmp.Fields("TM21"))
         End If
         If IsNull(rsTmp.Fields("TM22")) = False Then
            textTM22S = TAIWANDATE(rsTmp.Fields("TM22"))
         End If
      Else
            strTit = "檢核資料"
            strMsg = "商標審定號不存在"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            Me.textTM05S.Text = ""
            Me.textTM21S.Text = ""
            Me.textTM22S.Text = ""
      End If
      rsTmp.Close
   Else
        Me.textTM05S.Text = ""
        Me.textTM21S.Text = ""
        Me.textTM22S.Text = ""
   End If
   Set rsTmp = Nothing
End Sub

Private Sub textSP58_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textSP59_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textTM01_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

' 本所案號的系統別
Private Sub textTM01_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   Cancel = False
   If IsEmptyText(textTM01) = False Then
      If textTM01 <> m_TM01 Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "轉本所案號中的系統別不正確"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
        textTM01_GotFocus
         GoTo EXITSUB
      End If
      
      Select Case textTM01
         Case "TF":
            textTM02_2.Visible = True
            textTM02_2.Locked = False
            textTM02_2.TabStop = True
            textTM02.MaxLength = 5
         Case Else
            textTM02_2.Visible = False
            textTM02_2.Locked = True
            textTM02_2.TabStop = False
            textTM02.MaxLength = 6
      End Select
   Else
      textTM02_2.Visible = False
      textTM02_2.Locked = True
      textTM02_2.TabStop = False
      textTM02.MaxLength = 6
   End If
EXITSUB:
End Sub

Private Sub UpdateGrdList(ByVal strTM01 As String, ByVal strTM02 As String, ByVal strTM03 As String, ByVal strTM04 As String)
   Dim nIndex As Integer
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   
   ' 本案期限
   InitialGrdList
   ' 取得下一程序檔案中的資料列表在 Grid List 中
   'Modify by Morgan 2009/12/25 下一程序要排除程序管制的案件性質
   '2010/3/23 MODIFY BY SONIA 剔除下一程序非智權人員掌控之案件性質改以strNpSqlOfNoSalesDuty控制
   strSql = "SELECT NP01,NP07,NP08,NP09,NP10,NP11,NP12,NP13,NP14,NP15,NP22 FROM NextProgress " & _
            "WHERE NP02 = '" & strTM01 & "' AND " & _
                  "NP03 = '" & strTM02 & "' AND " & _
                  "NP04 = '" & strTM03 & "' AND " & _
                  "NP05 = '" & strTM04 & "' AND " & _
                  "(NP06 IS NULL OR NP06 <> 'Y') " & strNpSqlOfNoSalesDuty
   
   'Add by Morgan 2009/12/25 延期+AB類未發文未取消收文的程序
   If textCP10 = "303" Then
      strSql = strSql & " UNION SELECT CP09,CP10,CP06,CP07,CP13,CP57,CP58,CP08,CP40,CP64,0" & _
         " FROM CASEPROGRESS WHERE CP01 = '" & strTM01 & "' AND CP02 = '" & strTM02 & "'" & _
         " AND CP03 = '" & strTM03 & "' AND CP04 = '" & strTM04 & "'" & _
         " AND CP09<'C' and cp10<>'303' and cp07>0 AND CP27 IS NULL AND CP57 IS NULL"
   End If
   strSql = strSql & " ORDER BY 3"   'add by sonia 2018/8/30 依本所期限由小至大排序
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      Do While rsTmp.EOF = False
         grdList.Rows = grdList.Rows + 1
         nIndex = grdList.Rows - 1
         
         ' 收文號
         If IsNull(rsTmp.Fields("NP01")) = False Then
            grdList.TextMatrix(nIndex, 8) = rsTmp.Fields("NP01")
         End If
         ' 下一程序
         If IsNull(rsTmp.Fields("NP07")) = False Then
            '2005/4/18 MODIFY BY SONIA
            'grdList.TextMatrix(nIndex, 1) = GetCaseTypeName(m_TM01, rsTmp.Fields("NP07"))
            If m_TM10 = "000" Then
               grdList.TextMatrix(nIndex, 1) = GetCaseTypeName(m_TM01, rsTmp.Fields("NP07"), 0)
            Else
               grdList.TextMatrix(nIndex, 1) = GetCaseTypeName(m_TM01, rsTmp.Fields("NP07"), 1)
            End If
            '2005/4/18 END
            grdList.TextMatrix(nIndex, 9) = rsTmp.Fields("NP07")
         End If
         ' 本所期限
         If IsNull(rsTmp.Fields("NP08")) = False Then
            If IsEmptyText(rsTmp.Fields("NP08")) = False Then
               grdList.TextMatrix(nIndex, 2) = ChangeWStringToTString(rsTmp.Fields("NP08"))
            End If
         End If
         ' 法定期限
         If IsNull(rsTmp.Fields("NP09")) = False Then
            If IsEmptyText(rsTmp.Fields("NP09")) = False Then
               grdList.TextMatrix(nIndex, 3) = ChangeWStringToTString(rsTmp.Fields("NP09"))
            End If
         End If
         ' 機關文號
         If IsNull(rsTmp.Fields("NP13")) = False Then
            grdList.TextMatrix(nIndex, 4) = rsTmp.Fields("NP13")
         End If
         ' 相關人
         If IsNull(rsTmp.Fields("NP14")) = False Then
            grdList.TextMatrix(nIndex, 5) = rsTmp.Fields("NP14")
         End If
         ' 備註
         If IsNull(rsTmp.Fields("NP15")) = False Then
            grdList.TextMatrix(nIndex, 6) = rsTmp.Fields("NP15")
         End If
         ' 解除期限日期
         If IsNull(rsTmp.Fields("NP11")) = False Then
            grdList.TextMatrix(nIndex, 7) = ChangeWStringToTString(rsTmp.Fields("NP11"))
         End If
         ' 序號
         If IsNull(rsTmp.Fields("NP22")) = False Then
            grdList.TextMatrix(nIndex, 10) = rsTmp.Fields("NP22")
         End If
NextRecord:
         rsTmp.MoveNext
      Loop
   End If
   rsTmp.Close
   
   Set rsTmp = Nothing
End Sub

Private Function IsCaseProgressExist(ByVal strTM01 As String, ByVal strTM02 As String, ByVal strTM03 As String, ByVal strTM04 As String) As Boolean
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   IsCaseProgressExist = False
   strSql = "SELECT * FROM CaseProgress " & _
            "WHERE CP01 = '" & strTM01 & "' AND " & _
                  "CP02 = '" & strTM02 & "' AND " & _
                  "CP03 = '" & strTM03 & "' AND " & _
                  "CP04 = '" & strTM04 & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      IsCaseProgressExist = True
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Function

Private Function IsDataRecordExist(ByVal strTM01 As String, ByVal strTM02 As String, ByVal strTM03 As String, ByVal strTM04 As String) As Boolean
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   IsDataRecordExist = False
   Select Case strTM01
      Case "T", "TF", "FCT":
         strSql = "SELECT * FROM TradeMark " & _
                  "WHERE TM01 = '" & strTM01 & "' AND " & _
                        "TM02 = '" & strTM02 & "' AND " & _
                        "TM03 = '" & strTM03 & "' AND " & _
                        "TM04 = '" & strTM04 & "' "
      Case Else
         strSql = "SELECT * FROM ServicePractice " & _
                  "WHERE SP01 = '" & strTM01 & "' AND " & _
                        "SP02 = '" & strTM02 & "' AND " & _
                        "SP03 = '" & strTM03 & "' AND " & _
                        "SP04 = '" & strTM04 & "' "
   End Select
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      IsDataRecordExist = True
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Function

' 初始化 GridList
Private Sub InitialGrdList()
Dim nCol As Integer 'Add By Sindy 2023/6/29

   grdList.Clear
   grdList.Rows = 1
   grdList.Cols = 11
   grdList.ColWidth(0) = 300
   grdList.row = 0
   grdList.col = 1
   grdList.Text = "下一程序"
   grdList.ColWidth(1) = 1200
   grdList.col = 2
   grdList.Text = "本所期限"
   grdList.ColWidth(2) = 800
   grdList.col = 3
   grdList.Text = "法定期限"
   grdList.ColWidth(3) = 800
   grdList.col = 4
   grdList.Text = "機關文號"
   grdList.ColWidth(4) = 1000
   grdList.col = 5
   grdList.Text = "相關人"
   grdList.ColWidth(5) = 1200
   grdList.col = 6
   grdList.Text = "備註"
   grdList.ColWidth(6) = 1200
   grdList.col = 7
   grdList.Text = "解除期限日"
   grdList.ColWidth(7) = 800
   grdList.col = 8
   grdList.Text = "收文號"
   grdList.ColWidth(8) = 0
   grdList.col = 9
   grdList.Text = "下一程序代號"
   grdList.ColWidth(9) = 0
   grdList.col = 10
   grdList.Text = "序號"
   grdList.ColWidth(10) = 0
   'Add By Sindy 2023/6/29
   For nCol = 0 To grdList.Cols - 1
      grdList.col = nCol
      grdList.CellBackColor = &H8000000F
   Next nCol
   '2023/6/29 END
End Sub

Private Sub grdList_SelChange()
grdList_ShowSelection
End Sub

' 將GridList所選取的列反白, 並將未選取的列設成一般顏色
Private Sub grdList_ShowSelection()
   Dim nCurrSel As Integer
   Dim nCol As Integer
   
   nCurrSel = grdList.row
   
   ' 與前一選擇的列位置相同則不處理
   If m_CurrSel = grdList.row Then
      GoTo EXITSUB
   End If
   
   ' 將原先選取的列回復到正常的顏色
   If m_CurrSel > 0 And m_CurrSel < grdList.Rows Then
      grdList.row = m_CurrSel
      grdList.col = 1
      If grdList.CellBackColor <> &H80000005 Then
         For nCol = 1 To grdList.Cols - 1
            grdList.col = nCol
            If grdList.CellBackColor <> &H80000005 Then: grdList.CellBackColor = &H80000005
            If grdList.CellForeColor <> &H80000008 Then: grdList.CellForeColor = &H80000008
         Next nCol
      End If
      grdList.col = 0
   End If
   ' 設定成所選取的列
   m_CurrSel = nCurrSel
   ' 將所選取的列反白
   If m_CurrSel > 0 And m_CurrSel < grdList.Rows Then
      grdList.row = m_CurrSel
      grdList.col = 1
      For nCol = 1 To grdList.Cols - 1
         grdList.col = nCol
         grdList.CellBackColor = &H8000000D
         grdList.CellForeColor = &H80000005
      Next nCol
      grdList.col = 0
   End If
EXITSUB:
End Sub

Private Sub textTM03_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textTM04_LostFocus()
   Dim strTM01 As String
   Dim strTM02 As String
   Dim strTM03 As String
   Dim strTM04 As String
'Add By Cheng 2002/12/31
If (Me.textTM01.Text <> "" And Me.textTM02.Text = "") Or (Me.textTM01.Text = "" And Me.textTM02.Text <> "") Then
    MsgBox "轉本所案號輸入不完整!!!", vbExclamation + vbOKOnly
    Me.textTM01.SetFocus
    textTM01_GotFocus
    Exit Sub
End If

If textTM01 <> "" And textTM02 <> "" Then
   strTM01 = textTM01
   strTM02 = textTM02
   If strTM01 = "TF" Then: strTM02 = strTM02 & textTM02_2
   strTM03 = textTM03
   If IsEmptyText(strTM03) = True Then: strTM03 = "0"
   strTM04 = textTM04
   If IsEmptyText(strTM04) = True Then: strTM04 = "00"
   
   ' 更新本案期限的資料
   If IsDataRecordExist(strTM01, strTM02, strTM03, strTM04) Then
      UpdateGrdList strTM01, strTM02, strTM03, strTM04
   
   Else
      
      '910722 Sieg
      InitialGrdList
      strExc(1) = strTM01
      strExc(2) = strTM02
      strExc(3) = strTM03
      strExc(4) = strTM04
      If Not chkNewTMNo(strExc, intI) Then
         Select Case intI
            Case 1
               textTM01.SetFocus
            Case 2
               textTM02.SetFocus
         End Select
      'Add By Cheng 2002/09/09
      Else
         If Me.textTM01.Text <> "" And Me.textTM02.Text <> "" Then
            MsgBox "不可同時做" & """轉本所案號""" & "及" & """修改其他欄位""" & "，否則資料可能有不一致的情形產生!!!", vbExclamation + vbOKOnly
         End If
      End If
   End If
End If
End Sub

'910722 Sieg
Private Function chkNewTMNo(strNo() As String, iChk As Integer) As Boolean
Dim strTmp As String
   chkNewTMNo = True
   'edit by nickc 2007/02/06 不用 dll 了
   'If objPublicData.GetMaxNumber(strNo(1), strExc(0)) Then
   If ClsPDGetMaxNumber(strNo(1), strExc(0)) Then
      '92.3.24 modify by sonia
      'If strNo(1) & strNo(2) & strNo(3) & strNo(4) > strNo(1) & String(6 - Len(strExc(0)), "0") & strExc(0) Then
      If strNo(1) = "TF" Then
         strTmp = strNo(1) & String(5 - Len(strExc(0)), "0") & strExc(0)
      Else
         strTmp = strNo(1) & String(6 - Len(strExc(0)), "0") & strExc(0)
      End If
      If strNo(1) & strNo(2) > strTmp Then
      '92.3.24 end
         MsgBox "新本所案號不可大於自動編號，請重新輸入 !", vbCritical
         iChk = 2
         chkNewTMNo = False
      Else
         If MsgBox("此本所案號不存在 ( " & strNo(1) & strNo(2) & strNo(3) & strNo(4) & " ) ，請確認 ?", vbQuestion + vbYesNo) = vbNo Then
            iChk = 1
            chkNewTMNo = False
         End If
      End If
   End If
End Function

' 更新欄位的內容
Private Sub OnUpdateField()
Dim strTM01 As String
Dim strTM02 As String
Dim strTM03 As String
Dim strTM04 As String
Dim strCP64 As String
   
   If IsEmptyText(textTM01) = False And IsEmptyText(textTM02) = False Then
      strTM01 = textTM01
      strTM02 = textTM02
      If strTM02 = "TF" Then: strTM02 = strTM02 & textTM02_2
      strTM03 = textTM03
      If IsEmptyText(strTM03) Then: strTM03 = "0"
      strTM04 = textTM04
      If IsEmptyText(strTM04) Then: strTM04 = "00"
      ' 更新案件進度檔的本所案號
      SetCPFieldOldData "CP01", m_TM01, 0
      SetCPFieldOldData "CP02", m_TM02, 0
      SetCPFieldOldData "CP03", m_TM03, 0
      SetCPFieldOldData "CP04", m_TM04, 0
      SetCPFieldNewData "CP01", strTM01
      SetCPFieldNewData "CP02", strTM02
      SetCPFieldNewData "CP03", strTM03
      SetCPFieldNewData "CP04", strTM04
   End If
   
   ' 案件性質
   SetCPFieldNewData "CP10", textCP10
   ' 收文日
   If IsEmptyText(textCP05) = False Then
      SetCPFieldNewData "CP05", DBDATE(textCP05)
   Else
      SetCPFieldNewData "CP05", Empty
   End If
   ' 本所期限
   If IsEmptyText(textCP06) = False Then
      SetCPFieldNewData "CP06", DBDATE(textCP06)
   Else
      SetCPFieldNewData "CP06", Empty
   End If
   ' 法定期限
   If IsEmptyText(textCP07) = False Then
      SetCPFieldNewData "CP07", DBDATE(textCP07)
   Else
      SetCPFieldNewData "CP07", Empty
   End If
   
   'Add By Cheng 2002/06/12
   Select Case m_TM01
      ' 更新商標基本檔
      Case "T", "TF", "FCT", "CFT":
         '若案件性質為"延展"(102)
         If Me.textCP10.Text = "102" Then
            '93.3.4 ADD BY SONIA 中間接延展案無專用期止日
            If m_TM22 = "" Then
               m_TM22 = DBDATE(Val(textCP07))
            End If
            '93.3.4 END
            
            '若系統日小於等於法定期限
            '91.12.22 MODIFY BY SONIA 改判斷收文日
            'If ServerDate <= Val(m_strCP07) Then
            If Val(DBDATE(Val(textCP05))) <= Val(m_strCP07) Then
            '91.12.22 END
               '本所期限及法定期限不可修改
               SetCPFieldNewData "CP06", DBDATE(m_strCP06)
               SetCPFieldNewData "CP07", DBDATE(m_strCP07)
               '2006/4/27 ADD BY SONIA
               If (Val(DBDATE(Val(textCP06))) <> Val(m_strCP06)) Or (Val(DBDATE(Val(textCP07))) <> Val(m_strCP07)) Then
                  MsgBox "注意!!延展案收文日小於等於法定期限，不可在分案修改本所期限或法定期限，其他欄位仍會更新!!!", vbCritical
               End If
               '2006/4/27 END
            '若系統日大於法定期限
            Else
               '91.12.22 ADD BY SOhttps://www.samsung.com/tw/NIA
               If Val(textCP05) > Val(textCP07) Then
               '91.12.22 END
                  '91.11.3 MODIFY BY SONIA 應抓 CF12 或 CF28, 不可只抓 CF12
                  '法定期限為商標基本檔的"專用期止日"+案件國家收費表的"下次管制期限"
                  'm_strCP07 = DBDATE(Format(DateSerial(Val(DBYEAR(m_TM22)), Val(DBMONTH(m_TM22)), Val(DBDAY(m_TM22)) + GetCF12(m_TM01, m_TM10, Me.textCP10.Text))))
                  If GetCF12(m_TM01, m_TM10, Me.textCP10.Text) <> 0 Then
                     'modify by sonia 2023/4/14 改用textCP07計算,以免m_TM22與textCP07不同而未注意到(T-131010)
                     'm_strCP07 = DBDATE(CompDate(2, (GetCF12(m_TM01, m_TM10, Me.textCP10.Text)), Format(m_TM22)))
                     m_strCP07 = DBDATE(CompDate(2, (GetCF12(m_TM01, m_TM10, Me.textCP10.Text)), Format(textCP07)))
                  Else
                     'modify by sonia 2023/4/14 改用textCP07計算,以免m_TM22與textCP07不同而未注意到(T-131010)
                     'm_strCP07 = DBDATE(CompDate(1, (GetCF28(m_TM01, m_TM10, Me.textCP10.Text)), Format(m_TM22)))
                     m_strCP07 = DBDATE(CompDate(1, (GetCF28(m_TM01, m_TM10, Me.textCP10.Text)), Format(textCP07)))
                  End If
                  '91.11.3 END
                  SetCPFieldNewData "CP07", DBDATE(m_strCP07)
                  '本所期限 = 法定期限 - 2天
                  'Modify By Cheng 2003/09/01
                  'm_strCP06 = DBDATE(Format(DateSerial(Val(DBYEAR(m_strCP07)), Val(DBMONTH(m_strCP07)), Val(DBDAY(m_strCP07)) - 2)))
                  'Modify By Sindy 2014/10/6 台灣案之本所期限設定
                  If m_TM10 = "000" And Val(strSrvDate(1)) >= 台灣案所限新規則啟用日 Then
                     m_strCP06 = PUB_GetOurDeadline(DBDATE(m_strCP07))
                  Else
                  '2014/10/6 END
                     m_strCP06 = DBDATE(DateAdd("d", -2, ChangeWStringToWDateString(DBDATE(m_strCP07))))
                  End If
                  m_strCP06 = PUB_GetWorkDay1(m_strCP06, True) 'Added by Lydia 2020/07/07 本所期限檢查：若本所期限非工作天則直接調整至最近的工作天
                  SetCPFieldNewData "CP06", DBDATE(m_strCP06)
               '91.12.22 ADD BY SONIA
               Else
                  SetCPFieldNewData "CP06", DBDATE(textCP06)
                  SetCPFieldNewData "CP07", DBDATE(textCP07)
               End If
               '91.12.22 END
            End If
            ' 91.10.25 modify by sonia
            '規費
            'SetCPFieldNewData "CP07", (Val(GetCF08(m_TM01, m_TM10, Me.textCP10.Text)) * 2)
            '91.11.3 CANCEL BY SONIA
            'SetCPFieldNewData "CP17", (Val(GetCF08(m_TM01, m_TM10, Me.textCP10.Text)) * 2)
            '91.11.3 END
            ' 91.10.25 end
         
         'Add By Sindy 2020/10/20 陳述意見書
         ElseIf Me.textCP10.Text = "210" And _
            (m_TM01 = "T" Or m_TM01 = "FCT") Then
            '法定期限為空白
            textCP07 = ""
            SetCPFieldNewData "CP07", textCP07
            '本所期限規則如下：
            '為快軌案件為申請日+2個月
            If m_CP21 = "Y" Then
               m_strCP06 = DBDATE(DateAdd("m", 2, ChangeWStringToWDateString(DBDATE(m_CP143))))
            '一般案件期限管控為申請日+3個月
            Else
               m_strCP06 = DBDATE(DateAdd("m", 3, ChangeWStringToWDateString(DBDATE(m_CP143))))
            End If
            '若A或B計算出來的期限<=收文日隔日+10個工作天時，則一律設定為收文日隔日+10個工作天
            If m_strCP06 <= CompWorkDay(11, DBDATE(textCP05)) Then
               m_strCP06 = CompWorkDay(11, DBDATE(textCP05))
            End If
            '以上期限若為假日則再推算為前一工作天
            m_strCP06 = PUB_GetWorkDay1(m_strCP06, True)
            textCP06 = DBDATE(m_strCP06) - 19110000
            SetCPFieldNewData "CP06", DBDATE(textCP06)
         '2020/10/20 END
         End If
   End Select
   
   ' 業務區
   SetCPFieldNewData "CP12", GetST15(textCP13)
   ' 智權人員
   SetCPFieldNewData "CP13", textCP13
   ' 承辦人員
   SetCPFieldNewData "CP14", textCP14
   'Add By Sindy 2022/8/19
   If textCP14 <> "" Then
      SetCPFieldNewData "CP157", strSrvDate(1)
   Else
      SetCPFieldNewData "CP157", Empty
   End If
   '2022/8/19 END
   
   ' 相關總收文號
   SetCPFieldNewData "CP43", textCP43
   ' 是否算案件數
   SetCPFieldNewData "CP26", textCP26
   
   'Add By Sindy 2012/5/8
   ' 是否急件
   SetCPFieldNewData "CP122", textCP122
   '2012/5/8 End
   
'   'Added By Lydia 2015/11/24 收款後送件
'   If chkCP141.Value = 1 Then
'      SetCPFieldNewData "CP141", "2"
'   Else
'      SetCPFieldNewData "CP141", Empty
'   End If
'   'end 2015/11/24
   'Add By Sindy 2023/4/21
   If Val(m_CP27) = 0 Then
      If OptSendType(2).Value = True Then
         SetCPFieldNewData "CP141", "2"
         SetCPFieldNewData "CP142", Empty 'Add By Sindy 2024/4/10
      ElseIf OptSendType(3).Value = True Then
         SetCPFieldNewData "CP141", "3"
         SetCPFieldNewData "CP142", DBDATE(textCP142)
      'Modify By Sindy 2024/3/18
      ElseIf OptSendType(1).Value = True Then
      '2024/3/18 END
         SetCPFieldNewData "CP141", "1"
         SetCPFieldNewData "CP142", Empty
      'Add By Sindy 2024/4/10
      Else
         SetCPFieldNewData "CP141", Empty
         SetCPFieldNewData "CP142", Empty
      '2024/4/10 END
      End If
   End If
   '2023/4/21 END
   
   'Add By Sindy 2020/10/20
   If textCP10 = "210" Then '陳述意見書
      SetCPFieldNewData "CP143", IIf(m_CP143 <> "", DBDATE(m_CP143), "")
      SetCPFieldNewData "CP36", m_CP36
      SetCPFieldNewData "CP21", m_CP21
   End If
   '2020/10/20 END
   
   'Add by Morgan 2011/4/22 延期要紀錄NP22
   If textCP10 = "303" Then
      SetCPFieldNewData "CP30", m_CP30
   End If
      
    'add by nickc 2008/02/01 CF代理人
    If IsEmptyText(textCP44) = False Then
       SetCPFieldNewData "CP44", textCP44 & String(9 - Len(textCP44), "0")
    Else
       SetCPFieldNewData "CP44", textCP44
    End If
   
   ' 進度備註
   strCP64 = Me.textCP64.Text
   
    'Modify By Cheng 2003/09/05
    '取消
    'Begin
'    'Add By Cheng 2003/06/16
'    '若有輸入查名本所案號
'    'Modify By Cheng 2003/09/05
''    If Me.textTM01.Text <> "" And Me.textTM02.Text <> "" Then
'    If Me.Text1.Text <> "" And Me.Text2.Text <> "" Then
'        strCP64 = strCP64 & IIf(strCP64 <> "", ",", "") & "原查名本所案號：" & Me.Text1.Text & "-" & Me.Text2.Text & "-" & Left(Me.Text3.Text & "0", 1) & "-" & Left(Me.Text4.Text & "00", 2)
'    End If
    'End
   SetCPFieldNewData "CP64", strCP64
   ' 卷宗性質為非申請時, 更新案件進度檔的對造案件名稱
   If textTM28 <> "1" Then
      If m_TM01 = "T" Or m_TM01 = "TF" Or m_TM01 = "FCT" Then
'         ' 對造案件名稱(中)
'         SetCPFieldNewData "CP37", textTM05
         ' 對造案件名稱(中)
         SetCPFieldNewData "CP37", Me.textTM05_1.Text
'         ' 對造案件名稱(英)
'         SetCPFieldNewData "CP38", textTM06
'         ' 對造案件名稱(日)
'         SetCPFieldNewData "CP39", textTM07
      End If
   End If
   
   'add by sonia 2018/11/19   '廢止案要管制期限前不可發文,管制期限先存CP46,發文時再清除
   m_CP46 = ""
   'modify by sonia 2018/12/3 +623部分廢止
   If (textCP10 = "605" Or textCP10 = "623") Then
      If m_TM10 = "000" Then      '台灣案公告日起算三年為管制日
         m_CP46 = DBDATE(DateAdd("yyyy", 3, ChangeWStringToWDateString(DBDATE(Me.textTM14.Text))))
      ElseIf m_TM10 = "020" Then  '大陸案公告三個月期滿日起算三年為管制日
         m_CP46 = DBDATE(DateAdd("m", 3, DateAdd("yyyy", 3, ChangeWStringToWDateString(DBDATE(Me.textTM14.Text)))))
      End If
   End If
   If Val(m_CP27) = 0 Then SetCPFieldNewData "CP46", m_CP46
   'end 2018/11/19
   
   Select Case m_TM01
      ' 系統類別為CFT的為更新商標基本檔
      Case "T", "TF", "FCT":
         '92.8.28 MODIFY BY SONIA 取消卷宗性質條件
         '' 卷宗性質為非申請時, 不更新基本檔
         'If textTM28 = "1" Then
         '   ' 案件中文名稱
         '   SetTMSPFieldNewData "TM05", textTM05
         '   ' 案件英文名稱
         '   SetTMSPFieldNewData "TM06", textTM06
         '   ' 案件日文名稱
         '   SetTMSPFieldNewData "TM07", textTM07
         'End If
'         ' 案件中文名稱
'         SetTMSPFieldNewData "TM05", textTM05
         ' 案件名稱
         SetTMSPFieldNewData "TM05", Me.textTM05_1.Text
'         ' 案件英文名稱
'         SetTMSPFieldNewData "TM06", textTM06
'         ' 案件日文名稱
'         SetTMSPFieldNewData "TM07", textTM07
         '92.8.28 END
         ' 商標種類
         SetTMSPFieldNewData "TM08", textTM08
         
         'Add By Sindy 2019/4/9
         ' 特殊商標
         SetTMSPFieldNewData "TM72", textTM72
         '2019/4/9 END
         
         ' 商品類別
         SetTMSPFieldNewData "TM09", textTM09
         ' 申請國家
         SetTMSPFieldNewData "TM10", textTM10
         'Add By Sindy 2010/7/16 公告日
         If textTM14 = "" Then
            SetTMSPFieldNewData "TM14", ""
         Else
            SetTMSPFieldNewData "TM14", DBDATE(textTM14)
         End If
         ' 申請人
         SetTMSPFieldNewData "TM23", textTM23
         ' 申請地址(中)
         SetTMSPFieldNewData "TM24", textTM24
         ' 申請地址(英)
         SetTMSPFieldNewData "TM25", textTM25
         ' 申請地址(日)
         SetTMSPFieldNewData "TM26", textTM26
         ' 卷宗性質
         SetTMSPFieldNewData "TM28", textTM28
         ' 分所案號
         SetTMSPFieldNewData "TM34", textTM34
         ' 客戶案件案號
         SetTMSPFieldNewData "TM35", textTM35
         ' FC代理人
         If IsEmptyText(textTM44) = False Then
            SetTMSPFieldNewData "TM44", textTM44 & String(9 - Len(textTM44), "0")
         Else
            SetTMSPFieldNewData "TM44", textTM44
         End If
         ' 彼所案號
         SetTMSPFieldNewData "TM45", textTM45
         ' 案件備註
         SetTMSPFieldNewData "TM58", textTM58
         'add by nickc 2006/12/15
         SetTMSPFieldNewData "TM32", textTM32
         SetTMSPFieldNewData "TM78", textSP58
         SetTMSPFieldNewData "TM79", textSP59
         SetTMSPFieldNewData "TM80", textTM80
         SetTMSPFieldNewData "TM81", textTM81
         SetTMSPFieldNewData "TM82", textTM82
         SetTMSPFieldNewData "TM83", textTM83
         SetTMSPFieldNewData "TM84", textTM84
         SetTMSPFieldNewData "TM85", textTM85
         SetTMSPFieldNewData "TM86", textTM86
         SetTMSPFieldNewData "TM87", textTM87
         SetTMSPFieldNewData "TM88", textTM88
         SetTMSPFieldNewData "TM89", textTM89
         SetTMSPFieldNewData "TM90", textTM90
         SetTMSPFieldNewData "TM91", textTM91
         SetTMSPFieldNewData "TM92", textTM92
         SetTMSPFieldNewData "TM93", textTM93
         'add by nickc 2008/01/31
         SetTMSPFieldNewData "TM38", textTM38
         SetTMSPFieldNewData "TM39", textTM39 'Add By Sindy 2015/2/26
         SetTMSPFieldNewData "TM40", textTM40 'Add By Sindy 2015/2/26
         'add by Sindy 2012/12/20
         SetTMSPFieldNewData "TM41", textTM41
         SetTMSPFieldNewData "TM42", textTM42 'Add By Sindy 2015/2/26
         SetTMSPFieldNewData "TM43", textTM43 'Add By Sindy 2015/2/26
         'Add By Sindy 2013/12/16
         ' 特殊出名公司
         SetTMSPFieldNewData "TM130", textTM130
         '2013/12/16 END
      Case Else:
        'Modify By Cheng 2004/02/24
        '查名案件名稱合併至一欄
        Select Case m_TM01
        Case "TS"
            ' 案件中文名稱
            SetTMSPFieldNewData "SP05", textTM05_1
        Case Else
            ' 案件中文名稱
            SetTMSPFieldNewData "SP05", textTM05
        End Select
        'End
         ' 案件英文名稱
         SetTMSPFieldNewData "SP06", textTM06
         ' 案件日文名稱
         SetTMSPFieldNewData "SP07", textTM07
         ' 案件備註
         SetTMSPFieldNewData "SP18", textTM58
         ' 申請人
         SetTMSPFieldNewData "SP08", textTM23
'edit by nickc           2006/12/15
'         If m_TM01 = "TC" Then
            ' 申請人2
            SetTMSPFieldNewData "SP58", textSP58
            ' 申請人3
            SetTMSPFieldNewData "SP59", textSP59
'         End If
         ' 申請國家
         '911120 nick
         'SetTMSPFieldNewData "SP09", m_TM10
         SetTMSPFieldNewData "SP09", textTM10
         ' FC代理人
         If IsEmptyText(textTM44) = False Then
            SetTMSPFieldNewData "SP26", textTM44 & String(9 - Len(textTM44), "0")
         Else
            SetTMSPFieldNewData "SP26", textTM44
         End If
         '911120 nick
         SetTMSPFieldNewData "SP27", textTM45
         SetTMSPFieldNewData "SP28", textTM34
         ' 商標審定號
         SetTMSPFieldNewData "SP32", textSP32
         ' 案件備註
         SetTMSPFieldNewData "SP18", textTM58
         'add by nickc 2006/12/15
         SetTMSPFieldNewData "SP74", textTM32
         SetTMSPFieldNewData "SP58", textSP58
         SetTMSPFieldNewData "SP59", textSP59
         SetTMSPFieldNewData "SP65", textTM80
         SetTMSPFieldNewData "SP66", textTM81
         SetTMSPFieldNewData "SP73", textTM09
         'add by nickc 2008/01/31
         SetTMSPFieldNewData "SP30", textTM38
         'add by Sindy 2012/12/20
         SetTMSPFieldNewData "SP75", textTM41
         'Add By Sindy 2013/12/16
         ' 特殊出名公司
         SetTMSPFieldNewData "SP85", textTM130
         '2013/12/16 END
   End Select
End Sub

Private Function NextRecord() As Boolean
   Dim nIndex As Integer
   NextRecord = False
   
   For nIndex = 0 To m_CPKeyCount - 1
      If m_CP09 = m_CPKeyList(nIndex) Then
         If nIndex < m_CPKeyCount - 1 Then
            m_CP09 = m_CPKeyList(nIndex + 1)
            NextRecord = True
            Exit For
         End If
      End If
   Next nIndex
End Function

' 更新商標基本檔的相關欄位
'Modify By Cheng 2002/11/06
'Private Sub OnUpdateTradeMark()
Private Function OnUpdateTradeMark() As Boolean
   Dim strTmp As String
   Dim strSql As String
   Dim nIndex As Integer
   Dim bFirst As Boolean
   Dim bDifference As Boolean
      
On Error GoTo ErrorHandler
OnUpdateTradeMark = True

   '910702 Sieg 先檢查是否有修改申請人1，參照 501
   Dim strTmp1(1 To 3) As String
   If textTM23 <> "" Then
      'edit by nickc 2007/02/06 不用 dll 了
      'If objPublicData.GetCustomerNameAndAddress(textTM23, strExc(0), strTmp1(1), strTmp1(2), strTmp1(3)) Then
      If ClsPDGetCustomerNameAndAddress(textTM23, strExc(0), strTmp1(1), strTmp1(2), strTmp1(3)) Then
         '修改申請人時
         If InStr(ChangeCustomerL(m_TM23), ChangeCustomerL(textTM23)) = 0 Then
            If m_CP60 <> "" Then
               strExc(1) = m_TM01
               strExc(2) = m_TM02
               strExc(3) = m_TM03
               strExc(4) = m_TM04
               strExc(5) = m_CP60
               strExc(6) = ChangeCustomerL(textTM23)
               strExc(7) = strExc(0)
               '911118 nick 新增申請人
               strExc(8) = ChangeCustomerL(m_TM23)
               'edit by nickc 2007/02/06 不用 dll 了
               'If Not objLawDll.UpdAcc0k0(strExc(), True) Then
               If Not ClsLawUpdAcc0k0(strExc(), True) Then
                  textTM23.SetFocus
               End If
            End If
            SetTMSPFieldNewData "TM24", strTmp1(1)
            SetTMSPFieldNewData "TM25", strTmp1(2)
            SetTMSPFieldNewData "TM26", strTmp1(3)
            SetTMSPFieldNewData "TM47", ""
            SetTMSPFieldNewData "TM48", ""
            SetTMSPFieldNewData "TM49", ""
            SetTMSPFieldNewData "TM50", ""
            SetTMSPFieldNewData "TM51", ""
            SetTMSPFieldNewData "TM52", ""
         End If
      End If
      Call CheckAppAddr(1) 'Add By Sindy 2011/7/8
   Else
      SetTMSPFieldNewData "TM24", ""
      SetTMSPFieldNewData "TM25", ""
      SetTMSPFieldNewData "TM26", ""
      SetTMSPFieldNewData "TM47", ""
      SetTMSPFieldNewData "TM48", ""
      SetTMSPFieldNewData "TM49", ""
      SetTMSPFieldNewData "TM50", ""
      SetTMSPFieldNewData "TM51", ""
      SetTMSPFieldNewData "TM52", ""
   End If
   
'add by nickc 2006/12/15
   If textSP58 <> "" Then
      'edit by nickc 2007/02/06 不用 dll 了
      'If objPublicData.GetCustomerNameAndAddress(textSP58, strExc(0), strTmp1(1), strTmp1(2), strTmp1(3)) Then
      If ClsPDGetCustomerNameAndAddress(textSP58, strExc(0), strTmp1(1), strTmp1(2), strTmp1(3)) Then
         '修改申請人時
         If InStr(ChangeCustomerL(m_TM78), ChangeCustomerL(textSP58)) = 0 Then
            SetTMSPFieldNewData "TM82", strTmp1(1)
            SetTMSPFieldNewData "TM86", strTmp1(2)
            SetTMSPFieldNewData "TM90", strTmp1(3)
            SetTMSPFieldNewData "TM94", ""
            SetTMSPFieldNewData "TM95", ""
            SetTMSPFieldNewData "TM96", ""
            SetTMSPFieldNewData "TM97", ""
            SetTMSPFieldNewData "TM98", ""
            SetTMSPFieldNewData "TM99", ""
         End If
      End If
      Call CheckAppAddr(2) 'Add By Sindy 2011/7/8
   Else
      SetTMSPFieldNewData "TM82", ""
      SetTMSPFieldNewData "TM86", ""
      SetTMSPFieldNewData "TM90", ""
      SetTMSPFieldNewData "TM94", ""
      SetTMSPFieldNewData "TM95", ""
      SetTMSPFieldNewData "TM96", ""
      SetTMSPFieldNewData "TM97", ""
      SetTMSPFieldNewData "TM98", ""
      SetTMSPFieldNewData "TM99", ""
   End If
   
   If textSP59 <> "" Then
      'edit by nickc 2007/02/06 不用 dll 了
      'If objPublicData.GetCustomerNameAndAddress(textSP59, strExc(0), strTmp1(1), strTmp1(2), strTmp1(3)) Then
      If ClsPDGetCustomerNameAndAddress(textSP59, strExc(0), strTmp1(1), strTmp1(2), strTmp1(3)) Then
         '修改申請人時
         If InStr(ChangeCustomerL(m_TM79), ChangeCustomerL(textSP59)) = 0 Then
            SetTMSPFieldNewData "TM83", strTmp1(1)
            SetTMSPFieldNewData "TM87", strTmp1(2)
            SetTMSPFieldNewData "TM91", strTmp1(3)
            SetTMSPFieldNewData "TM100", ""
            SetTMSPFieldNewData "TM101", ""
            SetTMSPFieldNewData "TM102", ""
            SetTMSPFieldNewData "TM103", ""
            SetTMSPFieldNewData "TM104", ""
            SetTMSPFieldNewData "TM105", ""
         End If
      End If
      Call CheckAppAddr(3) 'Add By Sindy 2011/7/8
   Else
      SetTMSPFieldNewData "TM83", ""
      SetTMSPFieldNewData "TM87", ""
      SetTMSPFieldNewData "TM91", ""
      SetTMSPFieldNewData "TM100", ""
      SetTMSPFieldNewData "TM101", ""
      SetTMSPFieldNewData "TM102", ""
      SetTMSPFieldNewData "TM103", ""
      SetTMSPFieldNewData "TM104", ""
      SetTMSPFieldNewData "TM105", ""
   End If
   
   If textTM80 <> "" Then
      'edit by nickc 2007/02/06 不用 dll 了
      'If objPublicData.GetCustomerNameAndAddress(textTM80, strExc(0), strTmp1(1), strTmp1(2), strTmp1(3)) Then
      If ClsPDGetCustomerNameAndAddress(textTM80, strExc(0), strTmp1(1), strTmp1(2), strTmp1(3)) Then
         '修改申請人時
         If InStr(ChangeCustomerL(m_TM80), ChangeCustomerL(textTM80)) = 0 Then
            SetTMSPFieldNewData "TM84", strTmp1(1)
            SetTMSPFieldNewData "TM88", strTmp1(2)
            SetTMSPFieldNewData "TM92", strTmp1(3)
            SetTMSPFieldNewData "TM106", ""
            SetTMSPFieldNewData "TM107", ""
            SetTMSPFieldNewData "TM108", ""
            SetTMSPFieldNewData "TM109", ""
            SetTMSPFieldNewData "TM110", ""
            SetTMSPFieldNewData "TM111", ""
         End If
      End If
      Call CheckAppAddr(4) 'Add By Sindy 2011/7/8
   Else
      SetTMSPFieldNewData "TM84", ""
      SetTMSPFieldNewData "TM88", ""
      SetTMSPFieldNewData "TM92", ""
      SetTMSPFieldNewData "TM106", ""
      SetTMSPFieldNewData "TM107", ""
      SetTMSPFieldNewData "TM108", ""
      SetTMSPFieldNewData "TM109", ""
      SetTMSPFieldNewData "TM110", ""
      SetTMSPFieldNewData "TM111", ""
   End If
   
   If textTM81 <> "" Then
      'edit by nickc 2007/02/06 不用 dll 了
      'If objPublicData.GetCustomerNameAndAddress(textTM81, strExc(0), strTmp1(1), strTmp1(2), strTmp1(3)) Then
      If ClsPDGetCustomerNameAndAddress(textTM81, strExc(0), strTmp1(1), strTmp1(2), strTmp1(3)) Then
         '修改申請人時
         If InStr(ChangeCustomerL(m_TM81), ChangeCustomerL(textTM81)) = 0 Then
            SetTMSPFieldNewData "TM85", strTmp1(1)
            SetTMSPFieldNewData "TM89", strTmp1(2)
            SetTMSPFieldNewData "TM93", strTmp1(3)
            SetTMSPFieldNewData "TM112", ""
            SetTMSPFieldNewData "TM113", ""
            SetTMSPFieldNewData "TM114", ""
            SetTMSPFieldNewData "TM115", ""
            SetTMSPFieldNewData "TM116", ""
            SetTMSPFieldNewData "TM117", ""
         End If
      End If
      Call CheckAppAddr(5) 'Add By Sindy 2011/7/8
   Else
      SetTMSPFieldNewData "TM85", ""
      SetTMSPFieldNewData "TM89", ""
      SetTMSPFieldNewData "TM93", ""
      SetTMSPFieldNewData "TM112", ""
      SetTMSPFieldNewData "TM113", ""
      SetTMSPFieldNewData "TM114", ""
      SetTMSPFieldNewData "TM115", ""
      SetTMSPFieldNewData "TM116", ""
      SetTMSPFieldNewData "TM117", ""
   End If
   
   'Add By Sindy 2012/12/20
   SetTMSPFieldNewData "TM38", textTM38
   SetTMSPFieldNewData "TM39", textTM39 'Add By Sindy 2015/2/26
   SetTMSPFieldNewData "TM40", textTM40 'Add By Sindy 2015/2/26
   SetTMSPFieldNewData "TM41", textTM41
   SetTMSPFieldNewData "TM42", textTM42 'Add By Sindy 2015/2/26
   SetTMSPFieldNewData "TM43", textTM43 'Add By Sindy 2015/2/26
   '2012/12/20 End
   
   ' 更新案件進度檔
   strSql = "UPDATE TradeMark SET "
   bFirst = True
   bDifference = False
   For nIndex = 0 To m_TMSPCount - 1
      strTmp = Empty
      If m_TMSPList(nIndex).fiOldData <> m_TMSPList(nIndex).fiNewData Then
         bDifference = True
         If m_TMSPList(nIndex).fiType = 0 Then
            If m_TMSPList(nIndex).fiNewData = Empty Then
               strTmp = m_TMSPList(nIndex).fiName & " = " & "NULL"
            Else
               ' 91.03.25 modify by louis (單引號)
               'strTmp = m_TMSPList(nIndex).fiName & " = '" & m_TMSPList(nIndex).fiNewData & "'"
               strTmp = m_TMSPList(nIndex).fiName & " = '" & ChgSQL(m_TMSPList(nIndex).fiNewData) & "'"
            End If
         Else
            If m_TMSPList(nIndex).fiNewData = Empty Then
               strTmp = m_TMSPList(nIndex).fiName & " = " & "NULL"
            Else
               strTmp = m_TMSPList(nIndex).fiName & " = " & m_TMSPList(nIndex).fiNewData
            End If
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
   ' 設定SQL語法更新的條件
   strSql = strSql & " " & _
                  "WHERE TM01 = '" & m_TM01 & "' AND " & _
                        "TM02 = '" & m_TM02 & "' AND " & _
                        "TM03 = '" & m_TM03 & "' AND " & _
                        "TM04 = '" & m_TM04 & "' "
   ' 執行SQL指令
   If bDifference = True Then: cnnConnection.Execute strSql
   
   'Added by Lydia 2019/04/08 內商分案為FCT案時，判斷代理人國籍若為「日本」並且案件之定稿語文欄空白，則自動設定為「3」日文。
   If m_TM01 = "FCT" And textTM44 <> "" Then
         strSql = "SELECT fa01,fa02,fa10 FROM fagent" & _
               " WHERE fa01=" & CNULL(Left(Me.textTM44.Text, 8)) & _
               " and fa02=" & CNULL(Mid(Me.textTM44.Text, 9, 1))
        intI = 1
        Set RsTemp = ClsLawReadRstMsg(intI, strSql)
        If intI = 1 Then
            If Left("" & RsTemp.Fields("fa10"), 3) = "011" Then
              strSql = "UPDATE TradeMark SET TM53='3'" & _
                       " WHERE TM01 = '" & m_TM01 & "' AND " & _
                          "TM02 = '" & m_TM02 & "' AND " & _
                          "TM03 = '" & m_TM03 & "' AND " & _
                          "TM04 = '" & m_TM04 & "' AND " & _
                          "TM53 is null"
              cnnConnection.Execute strSql
            End If
        End If
   End If
   
'Add By Cheng 2002/11/06
Exit Function
ErrorHandler:
    OnUpdateTradeMark = False
End Function

' 更新服務業務基本檔的相關欄位
'Modify By Cheng 2002/11/06
'Private Sub OnUpdateServicePractice()
Private Function OnUpdateServicePractice() As Boolean
   Dim strTmp As String
   Dim strSql As String
   Dim nIndex As Integer
   Dim bFirst As Boolean
   Dim bDifference As Boolean
      
'Add By Cheng 2002/11/06
On Error GoTo ErrorHandler
OnUpdateServicePractice = True
   
   ' 更新案件進度檔
   strSql = "UPDATE ServicePractice SET "
   bFirst = True
   bDifference = False
   For nIndex = 0 To m_TMSPCount - 1
      strTmp = Empty
      If m_TMSPList(nIndex).fiOldData <> m_TMSPList(nIndex).fiNewData Then
         If m_TMSPList(nIndex).fiType = 0 Then
            If m_TMSPList(nIndex).fiNewData = Empty Then
               strTmp = m_TMSPList(nIndex).fiName & " = " & "NULL"
            Else
               ' 91.03.25 modify by louis (單引號)
               'strTmp = m_TMSPList(nIndex).fiName & " = '" & m_TMSPList(nIndex).fiNewData & "'"
               strTmp = m_TMSPList(nIndex).fiName & " = '" & ChgSQL(m_TMSPList(nIndex).fiNewData) & "'"
            End If
         Else
            If m_TMSPList(nIndex).fiNewData = Empty Then
               strTmp = m_TMSPList(nIndex).fiName & " = " & "NULL"
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
'Add By Cheng 2002/11/06
Exit Function
ErrorHandler:
    OnUpdateServicePractice = False
End Function

' 更新案件進度檔
Private Sub OnUpdateCaseProgress()
   Dim strTmp As String
   Dim strSql As String
   Dim nIndex As Integer
   Dim bFirst As Boolean
   Dim bDifference As Boolean
   
   ' 更新案件進度檔
   strSql = "UPDATE CaseProgress SET "
   bFirst = True
   bDifference = False
   For nIndex = 0 To m_CPCount - 1
      strTmp = Empty
      If m_CPList(nIndex).fiOldData <> m_CPList(nIndex).fiNewData Then
         If m_CPList(nIndex).fiType = 0 Then
            If m_CPList(nIndex).fiNewData = Empty Then
               strTmp = m_CPList(nIndex).fiName & " = " & "NULL"
            Else
               ' 91.03.25 modify by louis (單引號)
               'strTmp = m_CPList(nIndex).fiName & " = '" & m_CPList(nIndex).fiNewData & "'"
               strTmp = m_CPList(nIndex).fiName & " = '" & ChgSQL(m_CPList(nIndex).fiNewData) & "'"
            End If
         Else
            If m_CPList(nIndex).fiNewData = Empty Then
               strTmp = m_CPList(nIndex).fiName & " = " & "NULL"
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
   
End Sub

'Modify By Cheng 2002/11/06
'Private Function OnSaveData()
Public Function OnSaveData() As Boolean
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
'   Dim strCF13 As String
'   Dim strCF14 As String
   Dim strDay As String
   Dim strDate As String
   Dim strNP07 As String
   Dim strNP22 As String
   Dim strTemp As String
   Dim nIndex As Integer
   Dim nSubIndex As Integer
   Dim strCountry As String
   Dim strProduct As String
   Dim objCopyTM As ClsCopyTM
   Dim strTM01 As String
   Dim strTM02 As String
   Dim strTM03 As String
   Dim strTM04 As String
   'Add By Cheng 2002/09/09
   Dim StrSQLa As String
   Dim rsA As New ADODB.Recordset
   'edit by nickc 2006/09/18
   'Dim tm(1 To T_TM) As String
   'Dim sp(1 To T_SP) As String
   Dim tm() As String
   Dim sp() As String
   ReDim tm(1 To TF_TM) As String
   ReDim sp(1 To tf_SP) As String
   Dim i As Integer, strCUNo As String  ', strFaData(1) As String, strCuData(1) As String 'Add by Amy 2017/01/03
   Dim strApply As Variant, strAllApp As String 'Add by Amy 2017/03/14
   Dim strTran As Variant 'Add by Amy 2018/08/09
   Dim strCP09B  As String 'Add by Amy 2020/10/20
   Dim douStPrice As Double, douLowPrice As Double
   
'Add By Cheng 2002/11/06
On Error GoTo ErrorHandler
OnSaveData = True
cnnConnection.BeginTrans
   
   'Add By Sindy 2010/8/17 智權人員是96030(巨京)提醒請款對象是Y52269
   'Modify By Sindy 2011/3/3 fa30改抓fa107
   If textCP13 = "96030" Then
'      strSql = "SELECT fa107 FROM fagent WHERE fa01='" & Left(textTM44, 8) & "' and fa02='" & Mid(textTM44, 9, 1) & "' "
'      intI = 1
'      Set RsTemp = ClsLawReadRstMsg(intI, strSql)
'      If intI = 1 Then
'         If IsNull(RsTemp("fa107")) Or RsTemp("fa107") = "" Then
'            If MsgBox("FC代理人" & textTM44 & "，未設定固定請款對象為巨京，是否同時設定？", vbQuestion + vbYesNo) = vbYes Then
'               strSql = "Update fagent Set fa107='Y52269000' WHERE fa01='" & Left(textTM44, 8) & "' and fa02='" & Mid(textTM44, 9, 1) & "' "
'               cnnConnection.Execute strSql, intI
'            End If
'         End If
'      End If
      
      'Modify By Sindy 2013/8/6 請改為檢查若為新案件（cp31='Y')且為業務為96030收文,
      '且案件之固定請款對象(TM56)未設定時,會出現 "業務96030收文，案件未設定固定請款對象為巨京，是否同時設定？"
      If m_CP31 = "Y" And m_TM56 = "" Then
         If MsgBox("業務96030收文，案件未設定固定請款對象為巨京，是否同時設定？", vbQuestion + vbYesNo) = vbYes Then
            strSql = "Update trademark Set tm56='Y52269000' " & _
                     "WHERE tm01='" & m_TM01 & "' and tm02='" & m_TM02 & "'" & _
                      " and tm03='" & m_TM03 & "' and tm04='" & m_TM04 & "'"
            cnnConnection.Execute strSql
         End If
      End If
   End If
   
   'Modify By Cheng 2002/08/22
   '若有輸入轉本所案號
   If Me.textTM01.Text <> "" And Me.textTM02.Text <> "" Then
      'Add By Cheng 2002/09/09
      '判斷是否新增商標或服務業務基本案
      Select Case m_TM01
         Case "T", "TF", "FCT":
            StrSQLa = "SELECT * FROM TRADEMARK WHERE " & ChgTradeMark(Me.textTM01.Text & Me.textTM02.Text & IIf(Me.textTM02_2.Visible, Me.textTM02_2.Text, "") & Me.textTM03.Text & Me.textTM04.Text)
         Case Else:
            StrSQLa = "SELECT * FROM SERVICEPRACTICE WHERE " & ChgService(Me.textTM01.Text & Me.textTM02.Text & IIf(Me.textTM02_2.Visible, Me.textTM02_2.Text, "") & Me.textTM03.Text & Me.textTM04.Text)
      End Select
      rsA.CursorLocation = adUseClient
      rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
      If rsA.RecordCount <= 0 Then
         Select Case m_TM01
            Case "T", "TF", "FCT":
               If PUB_ReadTradeMarkData(tm(), m_TM01, m_TM02, m_TM03, m_TM04) Then
                  tm(1) = Me.textTM01.Text
                  tm(2) = Me.textTM02.Text & IIf(Me.textTM02_2.Visible, Me.textTM02_2.Text, "")
'                  tm(3) = Me.textTM03.Text
'                  tm(4) = Me.textTM04.Text
                  tm(3) = Left(Me.textTM03.Text & "0", 1)
                  tm(4) = Left(Me.textTM04.Text & "00", 2)
                  If PUB_AddNewTradeMark(tm()) Then
                  'Add By Cheng 2002/11/06
                  Else
                    GoTo ErrorHandler
                  End If
               End If
            Case Else:
               If PUB_ReadServicePracticeData(sp(), m_TM01, m_TM02, m_TM03, m_TM04) Then
                  sp(1) = Me.textTM01.Text
                  sp(2) = Me.textTM02.Text & IIf(Me.textTM02_2.Visible, Me.textTM02_2.Text, "")
'                  sp(3) = Me.textTM03.Text
'                  sp(4) = Me.textTM04.Text
                  sp(3) = Left(Me.textTM03.Text & "0", 1)
                  sp(4) = Left(Me.textTM04.Text & "00", 2)
                  If PUB_AddNewServicePractice(sp()) Then
                  'Add By Cheng 2002/11/06
                  Else
                    GoTo ErrorHandler
                  End If
               End If
         End Select
      'Add By Cheng 2002/12/06
      '若基本檔有資料, 若是否續辦欄為'Y'更新為Null
      Else
            'Modify by Morgan 2007/5/30 要用收文號更新
            'strSQL = " Update CaseProgress Set CP31=DECODE(CP31,'Y',NULL,CP31) WHERE " & ChgCaseprogress(Me.textTM01.Text & Me.textTM02.Text & IIf(Me.textTM02_2.Visible, Me.textTM02_2.Text, "") & Me.textTM03.Text & Me.textTM04.Text)
            strSql = " Update CaseProgress Set CP31=NULL WHERE CP09='" & textCP09 & "'"
            'end 2007/5/30
            cnnConnection.Execute strSql, intI
      End If
      If rsA.State <> adStateClosed Then rsA.Close
      Set rsA = Nothing
      
      strSql = "Update CASEPROGRESS SET CP01='" & Me.textTM01 & "',CP02='" & Left(Me.textTM02.Text & Me.textTM02_2.Text & "000000", 6) & "',CP03='" & Left(Me.textTM03.Text & "0", 1) & "',CP04='" & Left(Me.textTM04.Text & "00", 2) & "',CP43='' WHERE CP09='" & m_CP09 & "'"
      cnnConnection.Execute strSql
      
      'Added by Lydia 2020/08/18 更新CaseRelation1和DivisionCase
      If m_CP31 = "Y" Then
          Call PUB_UpdateCaseRelation1(m_TM01, m_TM02, m_TM03, m_TM04, Me.textTM01, Left(Me.textTM02.Text & Me.textTM02_2.Text & "000000", 6), Left(Me.textTM03.Text & "0", 1), Left(Me.textTM04.Text & "00", 2))
      End If
      'end 2020/08/18
      
      'Add by Sindy 2010/8/12
      '更正財務相關資料
      PUB_UpdateAccData textCP09, m_TM01 & m_TM02 & m_TM03 & m_TM04
      
'cancel by sonia 2024/11/26 已不立卷不必再通知分所收文人員
'      'Add by Sindy 2010/8/6 若為分所收文案件則發Mail通知收文人員
'      strExc(0) = PUB_GetST06(m_CP65)
'      If strExc(0) > "1" Then
'         strExc(1) = "原本所案號 " & m_TM01 & "-" & m_TM02 & IIf(m_TM03 & m_TM04 = "000", "", "-" & m_TM03 & "-" & m_TM04)
'         strExc(1) = strExc(1) & " 已更改為 " & Me.textTM01.Text & "-" & Me.textTM02.Text & Me.textTM02_2.Text & IIf(Me.textTM03.Text & Me.textTM04.Text = "000", "", "-" & Me.textTM03.Text & "-" & Me.textTM04.Text) & " 。"
'         'Modify By Sindy 2010/12/3
''         strSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08)" & _
''            " values ('" & strUserNum & "','" & m_CP65 & "',to_char(sysdate,'yyyymmdd'),to_char(sysdate,'hh24miss')" & _
''            ",'" & ChgSQL(strExc(1)) & "','如旨' )"
'         strSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08)" & _
'            " values ('" & strUserNum & "','" & m_CP65 & "',to_char(sysdate,'yyyymmdd'),to_char(sysdate,'hh24miss')" & _
'            ",'" & ChgSQL(strExc(1)) & "','總收文號：" & textCP09 & " 改本所案號如主旨')"
'         cnnConnection.Execute strSql, intI
'      End If
'      'end 2010/8/6
'end 2024/11/26
      
   '若未輸入轉本所案號
   Else
      ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
      ' 更新案件進度檔
      OnUpdateCaseProgress
      'Add By Sindy 2023/12/11
      strSql = ""
      If Frame3.Visible = True Then
         If Option1(0).Value = 0 And Option1(1).Value = 0 And Option1(2).Value = 0 Then
            strSql = "cp164=null"
         Else
            If Option1(0).Value = True Then
               strSql = "cp164='1'"
            ElseIf Option1(1).Value = True Then
               strSql = "cp164='2'"
            Else
               strSql = "cp164='3'"
            End If
         End If
      End If
      If strSql <> "" Then
         strSql = "UPDATE CaseProgress SET " & strSql & _
                  " WHERE CP09 = '" & m_CP09 & "'"
         cnnConnection.Execute strSql
      End If
      '2023/12/11 END
      ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
      
      ' 更新基本檔
      If IsEmptyText(textTM01) = False And IsEmptyText(textTM02) = False Then
         ' 本所案號
         strTM01 = textTM01
         strTM02 = textTM02
         If strTM01 = "TF" Then
            If textTM02_2 = "" Then
               strTM02 = strTM02 & "0"
            Else
               strTM02 = strTM02 & textTM02_2
            End If
         End If
         strTM03 = textTM03
         If IsEmptyText(strTM03) = True Then: strTM03 = "0"
         strTM04 = textTM04
         If IsEmptyText(strTM04) = True Then: strTM04 = "00"
         ' 檢查原始檔是否存在
         If IsDataRecordExist(strTM01, strTM02, strTM03, strTM04) = False Then
            Set objCopyTM = New ClsCopyTM
            objCopyTM.SetSrc m_TM01, m_TM02, m_TM03, m_TM04
            objCopyTM.SetDes strTM01, strTM02, strTM03, strTM04
            objCopyTM.SetExtraField "TM45", textTM45
            'add by nickc 2008/01/31 新增聯絡人
            objCopyTM.SetExtraField "TM38", textTM38
            objCopyTM.SetExtraField "TM39", textTM39 'Add By Sindy 2015/2/26
            objCopyTM.SetExtraField "TM40", textTM40 'Add By Sindy 2015/2/26
            'add by Sindy 2012/12/20 新增聯絡人2
            objCopyTM.SetExtraField "TM41", textTM41
            objCopyTM.SetExtraField "TM42", textTM42 'Add By Sindy 2015/2/26
            objCopyTM.SetExtraField "TM43", textTM43 'Add By Sindy 2015/2/26
            
            objCopyTM.CopyTradeMark
            Set objCopyTM = Nothing
            m_TM01 = strTM01
            m_TM02 = strTM02
            m_TM03 = strTM03
            m_TM04 = strTM04
            Select Case m_TM01
               ' 更新商標基本檔
               Case "T", "TF", "FCT":
                    'Modify By  Cheng 2002/11/06
'                  OnUpdateTradeMark
                  If OnUpdateTradeMark = False Then GoTo ErrorHandler
               ' 更新服務業務基本檔
               Case Else:
                    'Modify By Cheng 2002/11/06
'                  OnUpdateServicePractice
                  If OnUpdateServicePractice = False Then GoTo ErrorHandler
            End Select
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            ' 儲存優先權資料
            m_Pa(1) = m_TM01
            m_Pa(2) = m_TM02
            m_Pa(3) = m_TM03
            m_Pa(4) = m_TM04
            'Modify  By Cheng 2002/11/06
'            objPublicData.SavePriority m_Pa, m_Priority(1), m_Priority(2), m_Priority(3)
            'edit by nickc 2007/02/06 不用 dll 了
            'If objPublicData.SavePriority(m_Pa, m_Priority(1), m_Priority(2), m_Priority(3)) = False Then GoTo ErrorHandler
            'Modify by Amy 2014/04/17 +, m_Priority(4), m_Priority(5)
            'Modify by Sindy 2017/10/12 +, m_Priority(6)
            'Modify by Amy 2023/01/06 原m_Priority為陣列
            If ClsPDSavePriority(m_Pa, m_Priority1, m_Priority2, m_Priority3, m_Priority4, m_Priority5, m_Priority6) = False Then GoTo ErrorHandler
         Else
            m_TM01 = strTM01
            m_TM02 = strTM02
            m_TM03 = strTM03
            m_TM04 = strTM04
         End If
      Else
         Select Case m_TM01
            ' 更新商標基本檔
            Case "T", "TF", "FCT":
                'Modify By Cheng 2002/11/06
'               OnUpdateTradeMark
               If OnUpdateTradeMark = False Then GoTo ErrorHandler
            ' 更新服務業務基本檔
            Case Else:
                'Modify By Cheng 2002/11/06
'               OnUpdateServicePractice
               If OnUpdateServicePractice = False Then GoTo ErrorHandler
         End Select
         ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
         ' 儲存優先權資料
        'Modify By Cheng 2002/11/06
'         objPublicData.SavePriority m_Pa, m_Priority(1), m_Priority(2), m_Priority(3)
         'edit by nickc 2007/02/06 不用 dll 了
         'If objPublicData.SavePriority(m_Pa, m_Priority(1), m_Priority(2), m_Priority(3)) = False Then GoTo ErrorHandler
         'Modify by Amy 2014/04/17 +, m_Priority(4), m_Priority(5)
         'Modify by Sindy 2017/10/12 +, m_Priority(6)
         'Modify by Amy 2023/01/06 原m_Priority為陣列
         If ClsPDSavePriority(m_Pa, m_Priority1, m_Priority2, m_Priority3, m_Priority4, m_Priority5, m_Priority6) = False Then GoTo ErrorHandler
      End If
      ' 更新基本檔是否閉卷, 閉卷日期, 閉卷原因
      If textTM29 = "Y" Then
         Select Case m_TM01
         ' 更新商標基本檔
            Case "T", "TF", "CFT", "FCT":
               strSql = "UPDATE TRADEMARK SET TM29=NULL, TM30=NULL,TM31=NULL " & _
                        "WHERE TM01 = '" & m_TM01 & "' AND " & _
                              "TM02 = '" & m_TM02 & "' AND " & _
                              "TM03 = '" & m_TM03 & "' AND " & _
                              "TM04 = '" & m_TM04 & "' "
            Case Else:
               strSql = "UPDATE SERVICEPRACTICE SET SP15=NULL, SP16=NULL,SP17=NULL " & _
                        "WHERE SP01 = '" & m_TM01 & "' AND " & _
                              "SP02 = '" & m_TM02 & "' AND " & _
                              "SP03 = '" & m_TM03 & "' AND " & _
                              "SP04 = '" & m_TM04 & "' "
         End Select
         cnnConnection.Execute strSql
         'Added by Lydia 2025/06/30 取消閉卷時，若下一程序有未過期且已上N之延展102、使用宣誓105期限，都自動取消NP06、NP11、NP12，並彈訊息提醒已還原XXX期限。
         strMsgCloseCancel = PUB_GetCaseCloseCancel(m_TM01, m_TM02, m_TM03, m_TM04, m_TM10)
      End If
      ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
      ' 若有修改案件性質時,
      If textCP10 <> m_CP10 Then
'         strCF13 = "0"
'         strCF14 = "0"
'         strSql = "SELECT * FROM CaseFee " & _
'                  "WHERE CF01 = '" & m_TM01 & "' AND " & _
'                        "CF02 = '" & textTM10 & "' AND " & _
'                        "CF03 = '" & textCP10 & "' "
'         rsTmp.CursorLocation = adUseClient
'         rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
'         If rsTmp.RecordCount > 0 Then
'            rsTmp.MoveFirst
'            If IsNull(rsTmp.Fields("CF13")) = False Then
'               strCF13 = rsTmp.Fields("CF13")
'            End If
'            If IsNull(rsTmp.Fields("CF14")) = False Then
'               strCF14 = rsTmp.Fields("CF14")
'            End If
'         End If
'         rsTmp.Close
         
         'Add By Sindy 2022/12/15 有修改案件性質
         'Modify By Sindy 2023/9/22 + , m_CP10
         If PUB_ModCrLCRCData(m_CP09, txtF0301, textCP10, m_CP10, textTM10, textCP64) = False Then
            GoTo ErrorHandler
         End If
         '2022/12/15 END
         
         'Modify By Sindy 2022/12/16
         Call ClsPDGetCaseLowPrice(m_TM01, textTM10, textCP10, douStPrice, douLowPrice, textTM08, "", txtF0301)
         ' 更新案件進度檔的標準價及底價欄位
         strSql = "UPDATE CaseProgress SET CP33 = " & douStPrice & ", " & _
                                          "CP34 = " & douLowPrice & " " & _
                  "WHERE CP09 = '" & m_CP09 & "' "
         cnnConnection.Execute strSql
         '2022/12/16 END
      End If
      ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
      ' 案件性質為查名, 補收款, 後金時更新發文日為系統日
      Select Case textCP10
         '2012/12/21 MODFIY BY SONIA 取消查名,因為誤寫為"01"
         'Case "01", "705", "909":
         Case "705", "909":
            strSql = "UPDATE CaseProgress SET CP27 = " & DBDATE(SystemDate()) & " " & _
                     "WHERE CP09 = '" & m_CP09 & "' "
            cnnConnection.Execute strSql
         'add by sonia 2015/10/7 大陸商標為自動發證,故業務若收文領證則上發文日為系統日
         Case "701":
            If m_TM10 = "020" Then
               strSql = "UPDATE CaseProgress SET CP27 = " & DBDATE(SystemDate()) & " " & _
                        "WHERE CP09 = '" & m_CP09 & "' "
               cnnConnection.Execute strSql
            End If
         'end 2015/10/7
         'Add By Sindy 2020/10/20 陳述意見書
         Case "210":
            strExc(0) = "select cp09,cp10 from caseprogress where cp01='" & m_TM01 & "' and cp02='" & m_TM02 & "'" & _
                        " and cp03='" & m_TM03 & "' and cp04='" & m_TM04 & "' and cp158=0 and cp159=0" & _
                        " and cp10='214'"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 0 Then
               strCP09B = AutoNo("B", 6)
               '計算承辦期限
               strDate = PUB_TMdebateCountCP48(textCP06, "Y", strSrvDate(1), m_CP09, textCP13)
               'm_CP48 = strDate
               strSql = "insert into caseprogress(cp01,cp02,cp03,cp04,cp05,cp06,cp09,cp10,cp12,cp13,cp14,cp20,cp26,cp32,cp43,CP48) " & _
                              "values (" & CNULL(m_TM01) & "," & CNULL(m_TM02) & "," & CNULL(m_TM03) & _
                              "," & CNULL(m_TM04) & "," & CNULL(strSrvDate(1)) & "," & CompWorkDay(4, DBDATE(textCP05)) & "," & CNULL(strCP09B) & ",'214'," & _
                              CNULL(GetSalesArea(IIf(m_TM01 = "FCT", PUB_GetFCTSalesNo(m_TM01, m_TM02, m_TM03, m_TM04), PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04)))) & "," & _
                              CNULL(IIf(m_TM01 = "FCT", PUB_GetFCTSalesNo(m_TM01, m_TM02, m_TM03, m_TM04), PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04))) & "," & _
                              CNULL(textCP14) & ",'N','N','N'," & CNULL(m_CP09) & "," & strDate & ")"
               cnnConnection.Execute strSql
               'Add By Sindy 2020/11/12 系統自動上齊備日=系統日
               strSql = "update engineerprogress set ep06=" & strSrvDate(1) & ",ep34='N' where ep02='" & strCP09B & "'"
               cnnConnection.Execute strSql
               '2020/11/12 END
            End If
      End Select
      
'******************************************************************************
      ' 計算承辦期限
'******************************************************************************
      'Modify By Sindy 2022/9/29 改成共用函數,因自動收文且有承辦人時需計算承辦期限
      If textCP10 = "102" Then '延展
         'Added by Lydia 2019/08/27 T-223492分案已確定文件齊備
         If Frame21.Visible = True Then
              Call SaveFrame21(m_CP09)
         End If
      Else
         'Modified by Lydia 2018/12/10 判斷商爭案
         'If Frame21.Visible = True Then
         'Modified by Lydia 2022/07/15 限制T,FCT案 => And (m_TM01 = "T" Or m_TM01 = "FCT")
         'Modify By Sindy 2025/7/30 FCT案727分析不屬於爭議案 + FCT_NotTMdebate
         If Frame21.Visible = True And (m_TM01 = "T" Or m_TM01 = "FCT") And InStr(TMdebate, textCP10) > 0 And Not (m_TM01 = "FCT" And InStr(FCT_NotTMdebate, textCP10) > 0) Then
            Call SaveFrame21(m_CP09)
         'Added by Lydia 2018/12/10 非爭議案收文日在T案收文齊備啟用日之後
         'Memo by Lydia 2019/04/11 非爭議案(A類)T案收文齊備排除的案件性質皆不用管制齊備日(預設文件齊備=Y)
         'Modified by Lydia 2022/07/15  T大陸案之齊備日管控; TC案之文件齊備日管控;
         'ElseIf Frame21.Visible = True And InStr(TMdebate, textCP10) = 0 And DBDATE(textCP05) >= T案收文齊備啟用日 Then
         ElseIf Frame21.Visible = True And (m_TM01 = "TC" Or (m_TM01 = "T" And InStr(TMdebate, textCP10) = 0 And DBDATE(textCP05) >= T案收文齊備啟用日)) Then
            Call SaveFrame21(m_CP09)
         Else
            'Added by Lydia 2019/04/11 儲存文件齊備日 (ex.非爭議案的內部收文(B類)和T案收文齊備排除的案件性質)
            Call SaveFrame21(m_CP09)
         End If
      End If
      '計算承辦期限
      'Modify by Sindy 2023/2/10 畫面上有承辦人,且為第一次分案或承辦人異動
      'If textCP14.Tag = "" And textCP14 <> "" Then '承辦人由無－＞有時，才計算承辦期限
      If Trim(textCP14) <> "" And ((textCP14.Tag <> textCP14) Or m_bolIsFirstKeyCP14 = True) Then
      '2023/2/10 END
         If PUB_CountUpdTxCP48(m_CP09, textCP10, p_CP143DT, textCP05, textCP06, textCP07, textCP13, textCP122, m_TM01, m_TM10, strExc(10)) = True Then
            m_CP48 = strExc(10)
         End If
      End If
      '2022/9/29 END
'原程式Mark:
'      'Add By Sindy 2022/4/27 取得分案日
'      strExc(0) = "select cp09,cp149 from caseprogress WHERE CP09 = '" & m_CP09 & "' "
'      intI = 1
'      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'      If intI = 1 Then
'         m_CP149 = "" & RsTemp.Fields("cp149")
'      End If
'      '2022/4/27 END
'      'Modify by Morgan 2003/12/05
''      strDay = GetWorkDays(m_TM01, m_TM10, textCP10)
''      If IsEmptyText(strDay) = False Then
''         ' 90.07.03 承辦期限以工作天計算
''         'strDate = DBDATE(DateSerial(Val(DBYEAR(textCP05)), Val(DBMONTH(textCP05)), Val(DBDAY(textCP05)) + Val(strDay)))
''         strDate = DBDATE(CompWorkDay(Val(strDay), DBDATE(textCP05), 0))
''
''         strSQL = "UPDATE CaseProgress SET CP48 = " & strDate & " " & _
''                  "WHERE CP09 = '" & m_CP09 & "' "
''         cnnConnection.Execute strSQL
''      End If
'      'edit by nick 2004/12/08
'      'If m_CP10 = "102" Then
'      If textCP10 = "102" Then '延展
'         'Added by Lydia 2019/08/27 T-223492分案已確定文件齊備
'         If Frame21.Visible = True Then
'              Call SaveFrame21(m_CP09)
'         End If
'
'         Dim tmpDate As Date
'         'MODIFY BY SONIA 2013/6/10 馬德里續展改法定期限前三個月 TF-000110
'         'tmpDate = DateAdd("M", -6, ChangeTStringToWDateString(textCP07))
'         If m_TM01 = "TF" Then
'            'Modified by Lydia 2019/04/12
'            tmpDate = DateAdd("M", -3, ChangeTStringToWDateString(textCP07))
'         Else
'            'Modify By Sindy 2014/5/8
'            If m_TM10 = "020" Then '因T大陸修法，承辦期限改為法定期限減一年 ex.T-1731721
'               tmpDate = DateAdd("M", -12, ChangeTStringToWDateString(textCP07))
'            Else
'            '2014/5/8 END
'               tmpDate = DateAdd("M", -6, ChangeTStringToWDateString(textCP07))
'            End If
'         End If
'         '2013/6/10 END
'         '收文日大於法定期限：收文日期+3天
'         If (Val(textCP05) > Val(textCP07)) Then
'            'Modified by Lydia 2019/04/12 改成+3工作天(不含當天)
'            'strDate = Format(DateAdd("d", 3, ChangeTStringToWDateString(textCP05)), "YYYYMMDD")
'            strDate = CompWorkDay(4, DBDATE(textCP05))
'         '收文日小於承辦期限：承辦期限（法定期限減一年或半年或3個月）+3天
'         ElseIf Val(textCP05) < Val(Format(tmpDate, "YYYYMMDD") - 19110000) Then
'            'Modified by Lydia 2019/04/12 改成+3工作天(不含當天)
'            'strDate = Format(tmpDate + 3, "YYYYMMDD")
'            strDate = CompWorkDay(4, Format(tmpDate, "YYYYMMDD"))
'         '否則，收文日期+3天，若大於法定期限則承辦期限=法定期限
'         Else
'            'Modified by Lydia 2019/04/12 改成+3工作天(不含當天)
'            'strDate = Format(DateAdd("d", 3, ChangeTStringToWDateString(textCP05)), "YYYYMMDD")
'            strDate = CompWorkDay(4, DBDATE(textCP05))
'            If strDate > ChangeTStringToWString(textCP07) Then
'               strDate = ChangeTStringToWString(textCP07)
'            End If
'         End If
'         m_CP48 = strDate 'Add By Sindy 2012/5/8
'         strSql = "UPDATE CaseProgress SET CP48 = " & strDate & " " & _
'                  "WHERE CP09 = '" & m_CP09 & "' "
'         cnnConnection.Execute strSql
'      Else
'''''edit by nickc 2007/10/11 改抓有時效性的
'''''         strDay = GetWorkDays(m_TM01, m_TM10, textCP10)
'''''         If IsEmptyText(strDay) = False Then
'''''            strDate = DBDATE(CompWorkDay(Val(strDay), DBDATE(textCP05), 0))
'         'Add By Sindy 2012/5/8
'         'Modified by Lydia 2018/12/10 判斷商爭案
'         'If Frame21.Visible = True Then
'         'Modified by Lydia 2022/07/15 限制T,FCT案 => And (m_TM01 = "T" Or m_TM01 = "FCT")
'         If Frame21.Visible = True And (m_TM01 = "T" Or m_TM01 = "FCT") And InStr(TMdebate, textCP10) > 0 Then
'            Call SaveFrame21(m_CP09)
'            '承辦人欄由無－＞有時，且已輸入資料齊備時則計算承辦期限
'            'Modify By Sindy 2022/4/27 + And Val(m_CP149) > 0
'            If (textCP14.Tag = "" And textCP14 <> "") And textEP06 = "Y" And Val(m_CP149) > 0 Then
'               'Modify By Sindy 2022/4/27
'               '無齊備日時,不計算承辦期限, 計算承辦期限以分案日為起算日
'               'strDate = PUB_TMdebateCountCP48(textCP06, textCP122, m_EP06DT, m_CP09, textCP13)
'               strDate = PUB_TMdebateCountCP48(textCP06, textCP122, m_CP149, m_CP09, textCP13)
'               '2022/4/27 END
'               m_CP48 = strDate
'               strSql = "UPDATE CaseProgress SET CP48 = " & strDate & " " & _
'                        "WHERE CP09 = '" & m_CP09 & "' "
'               cnnConnection.Execute strSql
'            End If
'         'Added by Lydia 2018/12/10 非爭議案收文日在T案收文齊備啟用日之後
'         'Memo by Lydia 2019/04/11 非爭議案(A類)T案收文齊備排除的案件性質皆不用管制齊備日(預設文件齊備=Y)
'         'Modified by Lydia 2022/07/15  T大陸案之齊備日管控; TC案之文件齊備日管控;
'         'ElseIf Frame21.Visible = True And InStr(TMdebate, textCP10) = 0 And DBDATE(textCP05) >= T案收文齊備啟用日 Then
'         ElseIf Frame21.Visible = True And (m_TM01 = "TC" Or (m_TM01 = "T" And InStr(TMdebate, textCP10) = 0 And DBDATE(textCP05) >= T案收文齊備啟用日)) Then
'            Call SaveFrame21(m_CP09)
'            'Added by Lydia 2019/01/30 承辦人欄由無－＞有時，且已輸入文件和查名齊備時則計算承辦期限
'            'Modified by Lydia 2022/07/15 T大陸案之齊備日管控; TC案之文件齊備日管控;
'            'If (textCP14.Tag = "" And textCP14 <> "") And ((textCP10 = 申請 And textEP06 = "Y" And textCP143 = "Y") _
'                                 Or (textCP10 <> 申請 And textEP06 = "Y")) Then
'            If (textCP14.Tag = "" And textCP14 <> "") And ((m_TM01 = "T" And textCP10 = 申請 And textEP06 = "Y" And textCP143 = "Y") _
'                                 Or (m_TM01 = "T" And textCP10 <> 申請 And textEP06 = "Y") Or m_TM01 = "TC") Then
'               'Modified by Lydia 2019/04/11 承辦期限以齊備日+案件性質所設工作天數
'               'strDate = PUB_TMdebateCountCP48(textCP06, textCP122, m_EP06DT, m_CP09, textCP13)
'               strExc(1) = ""
'               If textEP06.Visible = True Then
'                  strExc(1) = m_EP06DT
'                  If textCP143.Visible = True Then
'                     If Val(strExc(1)) < Val(p_CP143DT) Then
'                          strExc(1) = p_CP143DT
'                     End If
'                  End If
'               End If
'               'Modify By Sindy 2022/4/27
'               '無齊備日時,不計算承辦期限
'               '而計算承辦期限以分案日為起算日
'               If Val(strExc(1)) = 0 Or Val(m_CP149) = 0 Then '無齊備日或無分案日
'                  'strExc(1) = DBDATE(textCP05) '無齊備日,就用收文日
'                  m_CP48 = Empty
'                  strSql = "UPDATE CaseProgress SET CP48 = null " & _
'                           "WHERE CP09 = '" & m_CP09 & "' "
'                  cnnConnection.Execute strSql
'               Else
'                  'strDate = Pub_GetHandleDay(m_TM01, m_TM10, textCP10, strExc(1), DBDATE(textCP06), textCP09)
'                  'Memo by Lydia 2022/07/15 TC案之文件齊備日管控: 自文件齊備日起算，五個工作天；與秀玲討論決定直接修改CaseFee，有設定的性質3天改為臺灣案5天/大陸案6天
'                                                                 '臺灣案的5個工作天含當天(by 嘉雯)；大陸案的5個工作天不含當天(by 承慧)。
'                  strDate = Pub_GetHandleDay(m_TM01, m_TM10, textCP10, m_CP149, DBDATE(textCP06), textCP09)
'                  If strDate <> "" Then
'                  'end 2019/04/11
'                       m_CP48 = strDate
'                       strSql = "UPDATE CaseProgress SET CP48 = " & strDate & " " & _
'                                "WHERE CP09 = '" & m_CP09 & "' "
'                       cnnConnection.Execute strSql
'                  End If
'               End If
'               '2022/4/27 END
'            End If
'         'end 2018/12/10
'         Else
'         '2012/5/8 End
'            'Added by Lydia 2019/04/11 儲存文件齊備日 (ex.非爭議案的內部收文(B類)和T案收文齊備排除的案件性質)
'            Call SaveFrame21(m_CP09)
'            'Modify By Sindy 2022/4/27
'            '無齊備日時,不計算承辦期限, 計算承辦期限以分案日為起算日
'            If Trim(textEP06.Text) = "Y" And Val(m_CP149) > 0 Then
'               'strDate = Pub_GetHandleDay(m_TM01, m_TM10, textCP10, DBDATE(textCP05), DBDATE(textCP06), textCP09)
'               strDate = Pub_GetHandleDay(m_TM01, m_TM10, textCP10, m_CP149, DBDATE(textCP06), textCP09)
'               If IsEmptyText(strDate) = False Then
'                  m_CP48 = strDate 'Add By Sindy 2012/5/8
'                  strSql = "UPDATE CaseProgress SET CP48 = " & strDate & " " & _
'                           "WHERE CP09 = '" & m_CP09 & "' "
'                  cnnConnection.Execute strSql
'               End If
'            End If
'         End If
'      End If
'      'End 2003/12/05
'******************************************************************************

      ' 若案件性質為救濟程序時或爭議程序更新基本檔的欄位
      Select Case Mid(textCP10, 1, 1)
         ' 救濟程序
         Case "4":
            Select Case m_TM01:
               Case "T", "TF", "FCT":
                  strSql = "UPDATE TradeMark SET TM18 = 'Y' " & _
                           "WHERE TM01 = '" & m_TM01 & "' AND " & _
                                 "TM02 = '" & m_TM02 & "' AND " & _
                                 "TM03 = '" & m_TM03 & "' AND " & _
                                 "TM04 = '" & m_TM04 & "' "
                  cnnConnection.Execute strSql
               Case Else:
            End Select
         ' 爭議程序
         Case "6":
            Select Case m_TM01:
               Case "T", "TF", "FCT":
                  strSql = "UPDATE TradeMark SET TM19 = 'Y' " & _
                           "WHERE TM01 = '" & m_TM01 & "' AND " & _
                                 "TM02 = '" & m_TM02 & "' AND " & _
                                 "TM03 = '" & m_TM03 & "' AND " & _
                                 "TM04 = '" & m_TM04 & "' "
                  cnnConnection.Execute strSql
               Case Else:
            End Select
      End Select
      
      '92.3.27 ADD BY SONIA
      'Modify By Sindy 2023/2/20 桂英提:恢復C類由承辦人判斷發文或不發文上"111111",不再收A類便一併發文。
      '因為發生智權人員急著收”補正”沖掉了審查報告(上了發文日),但承辦人還在操作歷程中~
'      If textCP43 <> "" Then
          '更新C類相關總收文號之發文日為收文日
'         strSql = "UPDATE CaseProgress SET CP27 = " & DBDATE(textCP05) & " " & _
'                  "WHERE CP09 = '" & textCP43 & "' AND CP09>'C' AND CP27 IS NULL"
'         cnnConnection.Execute strSql
'      End If
      '2023/2/20 END
      
      '92.3.27 END
      ' 若有輸入查名本所案號時, 更新該本所案號所有的案件進度檔其本所案號為本案之本所案號
      If IsEmptyText(Text1) = False And IsEmptyText(Me.Text2.Text) = False Then
         'add by nickc 2005/10/28 清未結餘的可結餘日期
         strSql = "UPDATE CaseProgress SET cp109=null " & _
                  "WHERE " & ChgCaseprogress(Me.Text1.Text & Me.Text2.Text & Me.Text3.Text & Me.Text4.Text) & " and cp59 is null "
         cnnConnection.Execute strSql
         'edit by nickc 2006/07/18 加入 cp31=null
         strSql = "UPDATE CaseProgress SET cp31=null " & _
                  "WHERE " & ChgCaseprogress(Me.Text1.Text & Me.Text2.Text & Me.Text3.Text & Me.Text4.Text) & " "
         cnnConnection.Execute strSql
         
         ' 組SQL語法
         strSql = "UPDATE CaseProgress SET CP01 = '" & m_TM01 & "', CP02 = '" & m_TM02 & "', " & _
                        "CP03 = '" & m_TM03 & "', CP04 = '" & m_TM04 & "', " & _
                       "CP64=CP64||Decode(CP64,Null,'','，')||'" & "原查名本所案號：" & Me.Text1.Text & "-" & Me.Text2.Text & "-" & Left(Me.Text3.Text & "0", 1) & "-" & Left(Me.Text4.Text & "00", 2) & "' " & _
                  "WHERE " & ChgCaseprogress(Me.Text1.Text & Me.Text2.Text & Me.Text3.Text & Me.Text4.Text)
         ' 執行更新的SQL指令
         cnnConnection.Execute strSql
         'Add By Cheng 2003/06/16
         strSql = "Update ServicePractice Set SP18=SP18||Decode(SP18,Null,'','，')||'轉入商標：" & m_TM01 & "-" & m_TM02 & "-" & m_TM03 & "-" & m_TM04 & "' Where " & ChgService(Me.Text1.Text & Me.Text2.Text & Me.Text3.Text & Me.Text4.Text)
         cnnConnection.Execute strSql
         '2005/4/18 ADD BY SONIA 1~4欄原查名本所案號,5~8欄新商標本所案號
         If PUB_UpdOther(Me.Text1.Text, Me.Text2.Text, Left(Me.Text3.Text & "0", 1), Left(Me.Text4.Text & "00", 2), m_TM01, m_TM02, m_TM03, m_TM04) = False Then
            GoTo ErrorHandler
         End If
         '2005/4/18 END
      End If
      '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
      ' 更新使用者所選取的本案期限資料
      For nIndex = 1 To grdList.Rows - 1
         ' 判斷該列是否有被選取
         If grdList.TextMatrix(nIndex, 0) = "V" Then
            strNP07 = grdList.TextMatrix(nIndex, 9)
            strNP22 = grdList.TextMatrix(nIndex, 10)
            'Modified by Lydia 2021/08/31 +更新NP24
            strSql = "UPDATE NextProgress SET NP06 = 'Y', NP24='" & cp(9) & "' " & _
                     "WHERE NP02 = '" & m_TM01 & "' AND " & _
                           "NP03 = '" & m_TM02 & "' AND " & _
                           "NP04 = '" & m_TM03 & "' AND " & _
                           "NP05 = '" & m_TM04 & "' AND " & _
                           "NP07 = " & strNP07 & " AND " & _
                           "NP22 = " & strNP22 & " "
            Pub_SeekTbLog strSql 'Added by Lydia 2021/08/31 各系統之分案作業和內部收文作業，若畫面勾選下一程序期限且存檔有上續辦Y的都寫Log以便事後能追蹤
            cnnConnection.Execute strSql
         End If
      Next nIndex
      '''''''''''''''''''''''''''''''''''''''''''''''''''''''
      ' 若此案為母案時
      '2014/3/5 MODIFY BY SONIA 加入cmdNation.Enabled = True的條件
      If m_TM03 = "0" And m_TM04 = "00" And (IsEmptyText(textTM01) = True Or IsEmptyText(textTM02) = True) And cmdNation.Enabled = True Then
        '920224 nick 先刪除所有分案
        'Modify By Cheng 2004/02/10
'        cnnConnection.Execute "delete from trademark where tm01='" & m_TM01 & "' and tm02='" & m_TM02 & "' and tm03>'0' "
        'Modify By Cheng 2004/02/27
'        cnnConnection.Execute "Delete From Trademark Where TM01='" & m_TM01 & "' And TM02='" & m_TM02 & "' And (TM03>'0' And TM03<='9') "
'2014/3/5 cancel by sonia 非TF不刪
'        Select Case m_TM01
'        Case "T"
'            If m_TM02 > "137100" Then
'                cnnConnection.Execute "Delete From Trademark Where TM01='" & m_TM01 & "' And TM02='" & m_TM02 & "' And (TM03>'0' And TM03<='9') "
'            End If
'        Case "FCT"
'            If m_TM02 > "017237" Then
'                cnnConnection.Execute "Delete From Trademark Where TM01='" & m_TM01 & "' And TM02='" & m_TM02 & "' And (TM03>'0' And TM03<='9') "
'            End If
'        Case "CFT"
'            If m_TM02 > "009179" Then
'                cnnConnection.Execute "Delete From Trademark Where TM01='" & m_TM01 & "' And TM02='" & m_TM02 & "' And (TM03>'0' And TM03<='9') "
'            End If
'        Case Else
            cnnConnection.Execute "Delete From Trademark Where TM01='" & m_TM01 & "' And TM02='" & m_TM02 & "' And (TM03>'0' And TM03<='9') "
'        End Select
'2014/3/5 end
        'End
        'End
         ' 系統類別為TF類
         Select Case m_TM01
            '2005/8/10 CANCEL BY SONIA
            'Case "T", "FCT":
            '   If GetSubStringCount(textTM09) > 1 Then
            '      For nIndex = 1 To GetSubStringCount(textTM09)
            '         strProduct = GetSubString(textTM09, nIndex)
            '         Set objCopyTM = New ClsCopyTM
            '         objCopyTM.SetSrc m_TM01, m_TM02, m_TM03, m_TM04
            '         objCopyTM.SetDes m_TM01, m_TM02, CStr(Val(m_TM03 + nIndex)), m_TM04
            '         objCopyTM.SetExtraField "TM09", strProduct
            '         objCopyTM.CopyTradeMark
            '         Set objCopyTM = Nothing
            '      Next nIndex
            '   End If
            '2005/8/10 END
            Case "TF":
               If IsEmptyText(m_strCountry) = False Then
                  If IsEmptyText(textTM09) = False Then
                     For nIndex = 1 To GetSubStringCount(textTM09)
                        strProduct = GetSubString(textTM09, nIndex)
                        For nSubIndex = 1 To GetSubStringCount(m_strCountry)
                           strCountry = GetSubString(m_strCountry, nSubIndex)
                           Set objCopyTM = New ClsCopyTM
                           objCopyTM.SetSrc m_TM01, m_TM02, m_TM03, m_TM04
                           objCopyTM.SetDes m_TM01, m_TM02, CStr(Val(m_TM03 + nIndex)), Format(CStr(Val(m_TM04) + nSubIndex), "00")
                           objCopyTM.SetExtraField "TM09", strProduct
                           objCopyTM.SetExtraField "TM10", strCountry
                           objCopyTM.CopyTradeMark
                           Set objCopyTM = Nothing
                        Next nSubIndex
                     Next nIndex
                  Else
                     For nSubIndex = 1 To GetSubStringCount(m_strCountry)
                        strCountry = GetSubString(m_strCountry, nSubIndex)
                        Set objCopyTM = New ClsCopyTM
                        objCopyTM.SetSrc m_TM01, m_TM02, m_TM03, m_TM04
                        objCopyTM.SetDes m_TM01, m_TM02, m_TM03, Format(CStr(Val(m_TM04) + nSubIndex), "00")
                        objCopyTM.SetExtraField "TM10", strCountry
                        objCopyTM.CopyTradeMark
                        Set objCopyTM = Nothing
                     Next nSubIndex
                  End If
               Else
                  '92.02.24 nick 又發現的問題，本來邱小姐說拿掉，後來又叫我恢復，但是加條件
                  If GetSubStringCount(textTM09) > 1 Then
                    For nIndex = 1 To GetSubStringCount(textTM09)
                    strProduct = GetSubString(textTM09, nIndex)
                    Set objCopyTM = New ClsCopyTM
                    objCopyTM.SetSrc m_TM01, m_TM02, m_TM03, m_TM04
                    objCopyTM.SetDes m_TM01, m_TM02, CStr(Val(m_TM03 + nIndex)), m_TM04
                    objCopyTM.SetExtraField "TM09", strProduct
                    objCopyTM.CopyTradeMark
                    Set objCopyTM = Nothing
                    Next nIndex
                   End If
               End If
         End Select
      End If
        
      'Add By Cheng 2004/04/14
      '更新分割案件關係資料
      If m_CP10 = "308" Then
          If PUB_UpdateDivisionCase(m_TM01, m_TM02, m_TM03, m_TM04, Me.txtDivCaseNo(0).Text, Me.txtDivCaseNo(1).Text & Me.txtDivCaseNo(2).Text, Me.txtDivCaseNo(3).Text, Me.txtDivCaseNo(4).Text) = False Then
              GoTo ErrorHandler
          End If
      End If
      'End
      
      'Add By Sindy 2013/12/16
      '若設公司別與已開收據不同時發Mail通知財務處及智權人員
      'If textTM130.Visible = True And textTM130.Tag <> textTM130 And Left(m_CP60, 1) = "E" Then
      'Remove by Lydia 2020/04/01 配合智慧所更名日，不再檢查
'      If textTM130.Visible = True And Left(m_CP60, 1) = "E" Then
'         strExc(0) = "select a0k01||DECODE(A0K32,'N','(暫不印)','Y','(待列印)') from acc0j0,acc0k0 where a0j01='" & m_CP09 & "' and a0k01(+)=a0j13 and a0k11<>'" & IIf(textTM130 = "J", "J", "1") & "'"
'         strExc(1) = ""
'         intI = 1
'         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'         If intI = 1 Then
'            Do While Not RsTemp.EOF
'               If strExc(1) = "" Then
'                  strExc(1) = RsTemp(0)
'               Else
'                  strExc(1) = strExc(1) & "," & RsTemp(0)
'               End If
'               RsTemp.MoveNext
'            Loop
'            strExc(1) = "商標案 " & m_TM01 & "-" & m_TM02 & IIf(m_TM03 & m_TM04 = "000", "", "-" & m_TM03 & "-" & m_TM04) & " 設定為以" & IIf(textTM130 = "J", "台一智權", "專利商標") & "出名與收據 " & strExc(1) & " 的公司別不同，請更正！"
'            strSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08)" & _
'               " values ('" & strUserNum & "','" & Pub_GetSpecMan("財務處總帳人員") & ";" & textCP13 & "',to_char(sysdate,'yyyymmdd'),to_char(sysdate,'hh24miss')" & _
'               ",'" & ChgSQL(strExc(1)) & "','如旨')"
'            cnnConnection.Execute strSql, intI
'         End If
'      End If
      '2013/12/16 END
      'end 2020/04/01
   End If  '------------'若未輸入轉本所案號
   
   'add by nickc 2005/03/17 加入加乘註記及寄件值
   m_CP98 = "": m_CP101 = "": m_CP104 = ""
   If PUB_GetFlagValue(m_CP09, m_CP98, m_CP101, m_CP104) = True Then
      strSql = "update caseprogress set cp98=" & m_CP98 & ",cp101=" & m_CP101 & ",cp104=" & m_CP104 & " WHERE CP09 = '" & m_CP09 & "' "
      cnnConnection.Execute strSql
   End If
   'PUB_UpdateCaseValue m_CP09 'Remove by Morgan 2005/4/13 改由 trigger 更新
    
   'Added by Lydia 2023/11/29 只要是FCT案內商承辦，請於內商分案時自動改為紙本送件
   If m_TM01 = "FCT" And "" & cp(118) <> "" Then
      strSql = "update caseprogress set cp118=null WHERE CP09 = '" & m_CP09 & "' "
      cnnConnection.Execute strSql
   End If
   'end 2023/11/29
   
   'Add by Amy 2017/01/17 MCTF控管 T字頭新案且有輸FC代理人且收文業務區為 P2字頭且申請國家是台灣時,依FC代理人之管控智權人員更新客戶檔之CU12,CU13
   'Modify by Amy 2017/03/14 原程式改寫至Function
   'Memo by Amy 2017/03/22 修改UpdMCTF_Cu13 拿掉申請國家是台灣的判斷
    If m_CP31 = "Y" And textTM44 <> MsgText(601) And Trim(textTM23) <> MsgText(601) Then
        For i = 0 To 4
            strCUNo = ""
            Select Case i
                Case 0
                    strCUNo = textTM23
                Case 1
                    strCUNo = textSP58
                Case 2
                    strCUNo = textSP59
                Case 3
                    strCUNo = textTM80
                Case 4
                    strCUNo = textTM81
            End Select
            If Len(Trim(strCUNo)) > 0 Then
                 strAllApp = strAllApp & "," & ChangeCustomerL(strCUNo)
            Else
                Exit For
            End If
        Next i
        If strAllApp <> MsgText(601) Then
            strApply = Split(Mid(strAllApp, 2), ",")
            If UpdMCTF_Cu13(m_TM01, textTM44, Trim(textTM10), strApply, PUB_GetST03(textCP13)) = False Then
                  GoTo ErrorHandler
            End If
        End If
    End If
    'end 2016/01/17
    'Add by Amy 2018/08/09
    If textTM44 <> MsgText(601) And strUpdCusNo <> MsgText(601) Then
        strTran = Split(strUpdCusNo, ",")
        If UpdMCTF_Cu13(m_TM01, textTM44, Trim(textTM10), strTran) = False Then
            GoTo ErrorHandler
        End If
    End If
    'end 201808/09
    
    'Add by Amy 2017/11/24 TF母案修改出名公司,子案一併更新 ex:CFP-029915
    If m_TM01 = "TF" And Right(m_TM02, 1) = "0" And m_TM03 = "0" And m_TM04 = "00" And textTM130.Tag <> textTM130 Then
        strSql = "Update Trademark Set tm130=" & CNULL(textTM130) & " " & _
                "WHERE tm01='" & m_TM01 & "' and tm02='" & m_TM02 & "' And tm03<>'0' and tm04<>'00'"
        cnnConnection.Execute strSql
    End If
    
    'Added by Morgan 2022/12/15
    '註冊證形式
    If textTM136.Visible And textTM136.Tag <> textTM136 Then
      strSql = "Update trademark Set tm136='" & textTM136 & "' " & _
                  "WHERE tm01='" & m_TM01 & "' and tm02='" & m_TM02 & "'" & _
                   " and tm03='" & m_TM03 & "' and tm04='" & m_TM04 & "'"
      cnnConnection.Execute strSql
    End If
    'end 2022/12/15
   
    'Added by Lydia 2020/05/20 法律所案源收文：如果案件性質或申請國家有變化,則需要對應分案; 5/28 +配合開庭
    'Modified by Lydia 2020/8/03 FCT商爭案由內商負責 +FCT
    If strSrvDate(1) >= 法律所案源收文啟用日 And (m_TM01 = "T" Or m_TM01 = "TC" Or m_TM01 = "FCT") And m_LOS07 = "" Then  '排除已放棄的案源
        'Modified by Lydia 2020/07/23 重新整理: 因為案源收文已設定不可變更案件性質和申請國家,所以只要判斷有案源
        'If textTM10.Text <> m_TM10 Or m_CP10 <> textCP10.Text Or (m_LOS15 = "" And txtLOSagree = "Y") Then
        '    Call PUB_UpdateCP10toPT(m_TM01, m_TM02, m_TM03, m_TM04, m_CP09, m_CP10, m_TM10, textCP10.Text, textTM10.Text, textCP06.Text, textCP13, textTM23, IIf(m_LOS15 = "" And txtLOSagree = "Y", True, False))
        'End If
        '
        'If textCP14.Tag = "" And textCP14.Text <> "" Then
        '  strSql = PUB_GetLOSkind(m_TM01, textCP10.Text, textTM10.Text)
        '  'Modified by Lydia 2020/06/09 判斷是否為補收文
        '  'If strSql <> "" Then
        '  strExc(1) = ""
        '  If m_LOS15 <> "" And strSql = "" Then strExc(1) = PUB_GetLOSplus(m_TM01, m_TM02, m_TM03, m_TM04, textCP10.Text, textTM10.Text, m_LOS02)
        '  If strSql <> "" Or strExc(1) <> "" Then
        '  'end 2020/06/09
        '      Call PUB_UpdateLOS01(m_TM01, m_TM02, m_TM03, m_TM04, m_CP09, textTM23 & "," & textSP58 & "," & textSP59 & "," & textTM80 & "," & textTM81, txtLOSagree)
        '  End If
        'End If
        'Modify by Sindy 2023/2/10 加判斷第一次分案
        'If m_LOS15 <> "" And textCP14.Tag = "" And textCP14.Text <> "" Then
        If m_LOS15 <> "" And (textCP14.Tag = "" Or m_bolIsFirstKeyCP14 = True) And textCP14.Text <> "" Then
        '2023/2/10 END
            Call PUB_UpdateLOS01(m_TM01, m_TM02, m_TM03, m_TM04, m_CP09, textTM23 & "," & textSP58 & "," & textSP59 & "," & textTM80 & "," & textTM81, txtLOSagree)
        End If
        'end 2020/07/23
    End If
    'end 2020/05/20
    
   'Add by Amy 2022/10/07 +接洽單電子化
   'Modify by Amy 2022/11/17 +急件
   If strSrvDate(1) >= 接洽單電子收文啟用日 Then
      If cp(122) <> textCP122 Then
          strSql = "Update CaseProgress Set CP122=" & CNULL(textCP122) & " Where cp09='" & cp(9) & "' "
          cnnConnection.Execute strSql
      End If
      'Modify by Amy 2022/10/21 +if 程序已分案不需再更新Flow002及Flow003
      'Modify by Amy 2022/11/18 同樣cp140且cp157都要有值才要更新(改最後一筆才上)
      If IsEConsultRec = True And stF0309_Now <> Flow_已分案 Then
          'Add By Sindy 2022/11/22 檢查接洽單全部案件性質是否全部分案完成
          If PUB_GetCP140CP157IsOK(cp(9)) = True Then
          '2022/11/22 END
              m_F0309 = Flow_已分案
              strUpdDate = strSrvDate(1)
              strUpdTime = Right("000000" & ServerTime, 6)
          
              '簽核檔(已處理)
              strSql = "update FLOW002 set " & _
                     "F0205='" & strUpdDate & "'" & _
                     ",F0206='" & strUpdTime & "'" & _
                     ",F0207='3',F0204='" & strUserNum & "'" & _
                     " where F0201='" & txtF0301 & "' and F0202='A7'  and F0207 is null "
              cnnConnection.Execute strSql
              '表單主檔
              strSql = "update FLOW003 set " & _
                      "F0309=" & CNULL(m_F0309) & _
                      " where F0301='" & txtF0301 & "' "
              cnnConnection.Execute strSql
          End If
      End If
      
      'Add By Sindy 2023/8/22 分案後要發通知信
      '有關TF之分案,當輸入之承辦人為外商CFT人員,請比照CFT案,由系統自動通知承辦人員
      If Trim(textCP14) <> "" And (textCP14.Tag <> textCP14 Or m_bolIsFirstKeyCP14 = True) Then
         If (m_TM01 = "TF" And Left(PUB_GetStaffST15(textCP14, 1), 2) = "F1") Or _
            m_TM01 = "TD" Then
            strExc(1) = "分案通知 " & textTM05_1 & "（本所案號：" & m_TM01 & "-" & m_TM02 & IIf(m_TM03 & m_TM04 = "000", "", "-" & m_TM03 & "-" & m_TM04) & "）"
            'Modify By Sindy 2023/12/11 +指定送件日期
            strExc(10) = "本所案號：" + m_TM01 & "-" & m_TM02 & "-" & m_TM03 & "-" & m_TM04 + vbCrLf + _
                         "申請國家：" + textTM10 + " " + textTM10_2 + vbCrLf + _
                         "案件名稱：" + textTM05_1 + vbCrLf + _
                         "案件性質：" + textCP10_2 + vbCrLf + _
                         "收文日　：" + ChangeWStringToTDateString(DBDATE(textCP05)) + vbCrLf + _
                         "智權人員：" + textCP13 + textCP13_2 + vbCrLf + _
                         "承辦人　：" + textCP14 + textCP14_2 + vbCrLf + _
                         "承辦期限：" + ChangeWStringToTDateString(DBDATE(m_CP48)) + vbCrLf + _
                         IIf(Trim(textCP142) <> "", "指定送件日期：" + ChangeWStringToTDateString(DBDATE(textCP142)) + IIf(Option1(0).Value = True, "當天", IIf(Option1(1).Value = True, "之後", IIf(Option1(2).Value = True, "之後", ""))) + vbCrLf, "") + _
                         "本所期限：" + ChangeWStringToTDateString(DBDATE(textCP06)) + vbCrLf + _
                         "法定期限：" + ChangeWStringToTDateString(DBDATE(textCP07)) + vbCrLf + _
                         IIf(textCP122.Visible = True And textCP122.Enabled = True And textCP122.Text = "Y", "是否急件：是" & vbCrLf, "")
            strSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08)" & _
               " values ('" & strUserNum & "','" & textCP14.Text & "',to_char(sysdate,'yyyymmdd'),to_char(sysdate,'hh24miss')" & _
               ",'" & ChgSQL(strExc(1)) & "','" & ChgSQL(strExc(10)) & "')"
            cnnConnection.Execute strSql, intI
         End If
      End If
   End If
   
   'Added by Morgan 2023/1/13
   '若延展收文日>=通知期限-延展的D類發文日，則MAIL通知通知期限-延展的D類的承辦人--林桂英
   If cp(27) = "" And textCP10 = "102" Then
      strExc(0) = "select cp14 from caseprogress,nextprogress where cp01='" & m_TM01 & "' and cp02='" & m_TM02 & "'" & _
                   " and cp03='" & m_TM03 & "' and cp04='" & m_TM04 & "' and cp27>=" & DBDATE(cp(5)) & _
                   " and cp10='1725' and np22(+)=cp30 and np24='" & cp(9) & "'"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         strExc(1) = cp(1) & "-" & cp(2) & IIf(cp(3) & cp(4) = "000", "", "-" & cp(3) & "-" & cp(4))
         strExc(2) = strExc(1) & "案已收文延展,無須再寄發紙本通知!"
         strExc(3) = ChangeCustomerL(textTM23)
         strSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08)" & _
            " select '" & strUserNum & "','" & RsTemp(0) & "',to_char(sysdate,'yyyymmdd')" & _
            ",to_char(sysdate,'hh24miss'),'" & ChgSQL(strExc(2)) & "'" & _
            ",'本所案號: " & strExc(1) & vbCrLf & "'" & _
            "||'申請人　: '||nvl(cu04,rtrim(cu05||' '||cu88||' '||cu89||' '||cu90))||'" & vbCrLf & "'" & _
            "||'聯絡地址: '||cu31||'" & vbCrLf & "'" & _
            "||'智權人員: " & textCP13_2 & vbCrLf & "'" & _
            "||'抽回原因: 本案已收文延展程序,無須再寄發紙本延展通知請抽回紙本!?'" & _
            " from customer where cu01='" & Left(strExc(3), 8) & "' and cu02='" & Mid(strExc(3), 9) & "'" & _
            " and not exists(select * from mailcache where mc03>=" & DBDATE(cp(5)) & " and mc07='" & ChgSQL(strExc(2)) & "')"
         cnnConnection.Execute strSql, intI
      End If
   End If
   'end 2023/1/13
   
   'Added by Lydia 2023/01/31 T大陸查名
   '1. 同一時間只有一位會進行分案
   '2.1/16 討論：79041若早上請假排給A2017，下午照順序給A7019，79041請假時間就跳過去。
   'Modified by Lydia 2023/02/10 改為獨立欄位EP13=>EP43
   If Frame22.Visible = True And textEP43.Tag <> textEP43.Text Then
      strSql = "Update engineerprogress set ep43='" & textEP43 & "' where ep02='" & textCP09 & "' "
      cnnConnection.Execute strSql
      If Trim(textEP43) <> "99997" Then
          strSql = "Update addressa4list set aal03='1' where aal01='大陸查名' and aal04='" & textEP43 & "' "
          cnnConnection.Execute strSql, intI
          If intI > 0 Then
             strSql = "Update addressa4list set aal03='0' where aal01='大陸查名' and aal04<>'" & textEP43 & "' and aal03='1' "
             cnnConnection.Execute strSql, intI
          End If
         strExc(1) = cp(1) & "-" & cp(2) & IIf(cp(3) & cp(4) = "000", "", "-" & cp(3) & "-" & cp(4))
         strExc(2) = strExc(1) & "大陸商標新申請案-待查名通知"
         strExc(3) = "本所案號：" & strExc(1) & vbCrLf & _
                          "案件名稱：" & Trim(textTM05_1) & vbCrLf & _
                          "案件性質：" & Trim(textCP10_2) & vbCrLf & _
                          "流程狀態：待查名"
         strSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08,mc09)" & _
            " values ('" & strUserNum & "','" & textEP43 & "',to_char(sysdate,'yyyymmdd')" & _
            ",to_char(sysdate,'hh24miss'),'" & ChgSQL(strExc(2)) & "','" & ChgSQL(strExc(3)) & "','" & textCP14 & "') "
         cnnConnection.Execute strSql, intI
      End If
   End If
   'end 2023/01/31
   
   'Add By Cheng 2002/11/06
   cnnConnection.CommitTrans
   ' 通知前畫面該筆收文資料已存檔
   frm020101_01.SetDataComplete m_CP09
'Add By Cheng 2002/11/06

   Exit Function
   
ErrorHandler:
    'Resume
    cnnConnection.RollbackTrans
    OnSaveData = False
    
End Function

Private Function CheckDataValid() As Boolean
Dim strTit As String
Dim strMsg As String
Dim strTemp As String
Dim nResponse
Dim ii As Integer
Dim strCode(0 To 7) As String
   
   CheckDataValid = False
   'Add By Cheng 2002/11/22
   '若非執行轉本所案號功能才要檢查
   If Me.textTM01.Text = "" Or Me.textTM02.Text = "" Then
        'Add By Cheng 2003/03/05
        '檢查案件名稱
'        If m_TM01 = "T" Or m_TM01 = "FCT" Or m_TM01 = "CFT" Or m_TM01 = "TF" Then
        If m_TM01 = "T" Or m_TM01 = "FCT" Or m_TM01 = "CFT" Or m_TM01 = "TF" Or m_TM01 = "TS" Then
            If Me.textTM05_1.Text = "" Then
                MsgBox "請輸入案件名稱!!!", vbExclamation + vbOKOnly
                Me.textTM05_1.SetFocus
                textTM05_1_GotFocus
                Exit Function
            End If
        Else
            If Me.textTM05.Text = "" And Me.textTM06.Text = "" And Me.textTM07.Text = "" Then
                MsgBox "案件名稱至少須輸入一項!!!", vbExclamation + vbOKOnly
                Me.textTM05.SetFocus
                textTM05_GotFocus
                Exit Function
            End If
        End If
       '92.02.24 nick 邱小姐後來說要檢查一定要輸入
       If cmdNation.Visible = True Then
            If Trim(m_strCountry) = "" Then
                strTit = "資料檢核"
                strMsg = "指定國家一定要輸入！"
                nResponse = MsgBox(strMsg, vbOKOnly, strTit)
                cmdNation.SetFocus
                GoTo EXITSUB
            End If
       End If
       ' 90.06.21
       strTemp = Empty
       strTemp = GetStaffName(textCP14)
       '92.3.20 CANCEL BY SONIA
       'If IsEmptyText(strTemp) = True Then
       '   strTit = "資料檢核"
       '   strMsg = "承辦人代號不存在"
       '   nResponse = MsgBox(strMsg, vbOKOnly, strTit)
       '   textCP14.SetFocus
       '   GoTo EXITSUB
       'End If
       '92.3.20 END
       ' 案件性質不可為空白
       If IsEmptyText(textCP10) = True Then
          strTit = "檢核資料"
          strMsg = "案件性質不可為空白"
          nResponse = MsgBox(strMsg, vbOKOnly, strTit)
          textCP10.SetFocus
          GoTo EXITSUB
       End If
       ' 申請國家不可空白
       If IsEmptyText(textTM10) = True Then
          strTit = "檢核資料"
          strMsg = "申請國家不可為空白"
          nResponse = MsgBox(strMsg, vbOKOnly, strTit)
          textTM10.SetFocus
          GoTo EXITSUB
       End If
       
       'Add By Sindy 2010/7/16
       ' 若為601異議案則檢查公告日不可為空白
       'modify by sonia 2018/11/19 +605廢止案
       'modify by sonia 2018/12/3 +623部分廢止
       'modify by sonia 2023/10/13 +627部分異議
       If (textCP10 = "601" Or textCP10 = "605" Or textCP10 = "623" Or textCP10 = "627") And Val(textTM14) = 0 Then
          strTit = "檢核資料"
          strMsg = "公告日不可為空白"
          nResponse = MsgBox(strMsg, vbOKOnly, strTit)
          textTM14.SetFocus
          GoTo EXITSUB
       End If
       
       ' 案件性質為延展或延期時本所期限及法定期限不可為空白
       'If m_CP10 = "102" Or m_CP10 = "303" Then
       'Modify By Sindy 2010/7/16 增加202,602,604,606
       '2013/10/7 MODIFY BY SONIA 加729復權T-183459
       'Modify By Sindy 2014/12/23 +內商大陸案案件性質401,403,408的期限不可空白,否則發文會down掉! T-195543(中間接進來所以未輸期限)
       'modify by sonia 2020/5/15 +410 (T-214110)
       'modify by sonia 2023/10/13 +624,628,630
       If textCP10 = "102" Or textCP10 = "303" Or textCP10 = "202" Or textCP10 = "602" Or _
          textCP10 = "604" Or textCP10 = "606" Or textCP10 = "729" Or textCP10 = "410" Or _
          textCP10 = "624" Or textCP10 = "628" Or textCP10 = "630" Or _
          (m_TM01 = "T" And m_TM10 = "020" And (textCP10 = "401" Or textCP10 = "403" Or textCP10 = "408")) Then
          If IsEmptyText(textCP06) = True Then
             strTit = "檢核資料"
            'Modify By Cheng 2002/10/30
    '         strMsg = "案件性質為延展, 本所期限不可為空白"
             'strMsg = "案件性質為延展或延期時, 本所期限不可為空白"
             strMsg = "本所期限不可為空白"
             nResponse = MsgBox(strMsg, vbOKOnly, strTit)
             textCP06.SetFocus
             GoTo EXITSUB
          End If
          If IsEmptyText(textCP07) = True Then
             strTit = "檢核資料"
            'Modify By Cheng 2002/10/30
    '         strMsg = "案件性質為延展, 法定期限不可為空白"
             'strMsg = "案件性質為延展或延期時, 法定期限不可為空白"
             strMsg = "法定期限不可為空白"
             nResponse = MsgBox(strMsg, vbOKOnly, strTit)
             textCP07.SetFocus
             GoTo EXITSUB
          End If
       End If
       If IsEmptyText(textCP06) = False And IsEmptyText(textCP07) = False Then
         'Modify By Sindy 2023/3/27 T.102.法定期限逾期,不檢查 ex:T-243230
         If Not (m_TM01 = "T" And m_CP10 = "102" And DBDATE(textCP07) < DBDATE(textCP05)) Then
         '2023/3/27 END
            If Val(textCP06) > Val(textCP07) Then
               strTit = "檢核資料"
               strMsg = "本所期限與法定期限範圍不正確"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textCP06.SetFocus
               GoTo EXITSUB
            End If
         End If
       End If
       ' 收文日
       If IsEmptyText(textCP05) = True Then
          strTit = "檢核資料"
          strMsg = "收文日不可空白"
          nResponse = MsgBox(strMsg, vbOKOnly, strTit)
          textCP05.SetFocus
          GoTo EXITSUB
       End If
       
       'Add By Sindy 2022/11/22
       If strSrvDate(1) >= 接洽單電子收文啟用日 Then
         'Modify By Sindy 2023/3/27 T.102.法定期限逾期,不檢查
         If Not (m_TM01 = "T" And m_CP10 = "102" And DBDATE(textCP07) < DBDATE(textCP05)) Then
         '2023/3/27 END
            'Modify By Sindy 2023/4/12 + , , , , textCP09
            If PUB_CRLUseCP07CheckCP06(m_CP31, textTM10, m_TM01, textCP10, textCP06, textCP07, , , , textCP09) = False Then
               textCP06.SetFocus
               GoTo EXITSUB
            End If
         End If
       End If
       '2022/11/22 END
       
       'add by nickc 2006/04/28 商標審定號不可空白
       Select Case m_TM01
          Case "TM":
             If textCP10 = "801" Then
                If IsEmptyText(textSP32) = True Then
                   strTit = "檢核資料"
                   strMsg = "商標審定號不可空白"
                   nResponse = MsgBox(strMsg, vbOKOnly, strTit)
                   'Me.SSTab1.Tab = 1
                   textSP32.SetFocus
                   GoTo EXITSUB
                End If
             End If
          Case Else
       End Select
       ' 商標種類不可空白
       Select Case m_TM01
          Case "T", "TF", "CFT", "FCT":
             If IsEmptyText(textTM08) = True Then
                strTit = "檢核資料"
                strMsg = "商標種類不可空白"
                nResponse = MsgBox(strMsg, vbOKOnly, strTit)
                'Modified by Lydia 2023/11/16
                'textTM08.SetFocus
                cboTM08.SetFocus
                GoTo EXITSUB
             End If
          Case Else
       End Select
       ' 卷宗性質不可空白
       Select Case m_TM01
          Case "T", "TF", "FCT":
             If IsEmptyText(textTM28) = True Then
                strTit = "檢核資料"
                strMsg = "卷宗性質不可空白"
                nResponse = MsgBox(strMsg, vbOKOnly, strTit)
                textTM28.SetFocus
                GoTo EXITSUB
             End If
       End Select
       
       'Add By Sindy 2011/01/06
       '內商(TS)申請人1或FC代理人至少要輸入一個
       '其他的一定要輸入申請人1
       '2011/9/16 MODIFY BY SONIA 加入TT,TR(TT-000119)
       If m_TM01 = "TS" Or m_TM01 = "TR" Or m_TM01 = "TT" Then
            If textTM23 = "" And textTM44 = "" Then
                MsgBox "申請人1或FC代理人至少要輸入一個!!!", vbExclamation + vbOKOnly
                Me.textTM23.SetFocus
                textTM23_GotFocus
                Exit Function
            End If
       Else
            If textTM23 = "" Then
                MsgBox "申請人1不可空白!!!", vbExclamation + vbOKOnly
                Me.textTM23.SetFocus
                textTM23_GotFocus
                Exit Function
            End If
       End If
       
        'Add By Cheng 2003/08/13
        '若案件性質為延期, 則不可點選本案期限
        If Me.textCP10.Text = "303" Then
            For ii = 1 To Me.grdList.Rows - 1
                If Me.grdList.TextMatrix(ii, 0) <> "" Then
                    MsgBox "此案僅收文<延期>，不可點選下一程序期限資料，" & vbCrLf & "否則無法管制下一程序的期限!!!", vbExclamation + vbOKOnly
                    GoTo EXITSUB
                End If
            Next ii
        End If
        'Add By Cheng 2004/04/14
        '若案件性質為分割案
        If m_CP10 = "308" Then
            'Add by Amy 2014/10/16 +T大陸案控制
            If m_TM01 = "T" And Me.textTM10 = "020" Then
                If Trim(Me.textCP06) = "" Then
                    MsgBox "大陸分割案本所期限不可為空!!", vbExclamation + vbOKOnly
                    textCP06_GotFocus
                    Me.textCP06.SetFocus
                    GoTo EXITSUB
                End If
                If Trim(Me.textCP07) = "" Then
                    MsgBox "大陸分割案法定期限不可為空!!", vbExclamation + vbOKOnly
                    textCP07_GotFocus
                    Me.textCP07.SetFocus
                    GoTo EXITSUB
                End If
            End If
            '若有輸分割案母案資料
            If Me.txtDivCaseNo(0).Text & Me.txtDivCaseNo(1).Text & Me.txtDivCaseNo(2).Text & Me.txtDivCaseNo(3).Text & Me.txtDivCaseNo(4).Text <> "" And m_CP31 = "Y" Then
                If Me.txtDivCaseNo(3).Text = "" Then Me.txtDivCaseNo(3).Text = "0"
                If Me.txtDivCaseNo(4).Text = "" Then Me.txtDivCaseNo(4).Text = "00"
                If ChkCaseExist(Me.txtDivCaseNo(0).Text, Me.txtDivCaseNo(1).Text & Me.txtDivCaseNo(2).Text, Me.txtDivCaseNo(3).Text, Me.txtDivCaseNo(4).Text) = False Then
                    MsgBox "分割案母案案號輸入錯誤!!!", vbExclamation + vbOKOnly
                    txtDivCaseNo_GotFocus 0
                    Me.txtDivCaseNo(0).SetFocus
                    GoTo EXITSUB
                End If
                If m_TM01 = Me.txtDivCaseNo(0).Text And m_TM02 = Me.txtDivCaseNo(1).Text & IIf(Me.txtDivCaseNo(0).Text = "TF", Me.txtDivCaseNo(2).Text, "") And m_TM03 = Me.txtDivCaseNo(3).Text And m_TM04 = Me.txtDivCaseNo(4).Text Then
                    MsgBox "分割案母案案號不可與分割案案號相同!!!", vbExclamation + vbOKOnly
                    txtDivCaseNo_GotFocus 0
                    Me.txtDivCaseNo(0).SetFocus
                    GoTo EXITSUB
                End If
                strCode(0) = m_TM01: strCode(1) = m_TM02: strCode(2) = m_TM03: strCode(3) = m_TM04
                strCode(4) = Me.txtDivCaseNo(0).Text: strCode(5) = Me.txtDivCaseNo(1).Text & IIf(Me.txtDivCaseNo(0).Text = "TF", Me.txtDivCaseNo(2).Text, ""): strCode(6) = Me.txtDivCaseNo(3).Text: strCode(7) = Me.txtDivCaseNo(4).Text
                If ChkCaseReleate(strCode()) = False Then
                    txtDivCaseNo_GotFocus 0
                    Me.txtDivCaseNo(0).SetFocus
                    GoTo EXITSUB
                End If
                Call txtDivCaseNo_Validate(0, False)
            '若未輸分割案母案資料
            Else
                'edit by nickc 2006/07/19
                'If MsgBox("是否輸入分割案母案資料???", vbExclamation + vbYesNo) = vbYes Then
                '   Me.txtDivCaseNo(0).SetFocus
                '    GoTo EXITSUB
                'End If
                If m_CP31 = "Y" Then    '新案分割一定要輸
                    MsgBox "分割母案案號一定要輸入！", vbExclamation
                    Me.txtDivCaseNo(0).SetFocus
                    GoTo EXITSUB
                End If
            End If
        End If
        'End
        'add by nickc 2008/02/01 申請國家非台灣的申請案(101)   CF 代理人不可以空白
        'edit by nickc 2008/02/04 改成不管是不是新案
        'If textTM10 <> "000" And textCP10 = "101" And Trim(textCP44) = "" Then
        'Modify By Sindy 2012/3/30 TD申請時皆在台灣申請不須控管CF代理人
        If textTM10 <> "000" And Trim(textCP44) = "" And m_TM01 <> "TD" Then
            MsgBox "申請國家非台灣的案件，CF 代理人不可以空白！", vbExclamation
            SSTab1.Tab = 2 'Added by Lydia 2023/06/27
            textCP44.SetFocus
            GoTo EXITSUB
        End If
        '2011/7/1 ADD BY SONIA
        If textTM10 = "000" And Trim(textCP44) <> "" Then
            MsgBox "申請國家為台灣的案件，CF 代理人不可以輸入！", vbExclamation
            SSTab1.Tab = 2 'Added by Lydia 2023/06/27
            textCP44.SetFocus
            GoTo EXITSUB
        End If
        '2011/7/1 END
        
        'Add By Sindy 2010/12/27
        'Modify By Sindy 2025/8/5 控管案件性質727分析一定要掛相關總收文號
        If (textCP10 = "303" Or textCP10 = "727") And IsEmptyText(textCP43) = True Then
            MsgBox textCP10_2 & "案, 一定要輸入相關總收文號！", vbExclamation
            SSTab1.Tab = 0 'Added by Lydia 2023/06/27
            textCP43.SetFocus
            GoTo EXITSUB
        End If
        '2010/12/27 end
        
      'Add By Sindy 2012/5/8
      '台灣商標Ｔ,FCT案若收文爭議案件性質時，若未填寫則不可存檔
      If Frame21.Visible = True Then
         'Modified by Lydia 2018/12/10 +判斷顯示; 延期303、放棄專用權206、暫緩審理310不必填「文件是否齊備」
         'If textEP06 = ""  Then
         '   MsgBox "資料是否齊備不可空白!!!", vbExclamation + vbOKOnly
         'Modified by Lydia 2022/07/15 T大陸案之齊備日管控=>T大陸案不要「T案收文齊備排除」
         'If textEP06 = "" And textEP06.Visible = True And InStr(T案收文齊備排除, textCP10) = 0 Then
         If textEP06 = "" And textEP06.Visible = True And (m_TM01 = "FCT" Or m_TM01 = "TC" Or _
                  ((m_TM01 = "T" And textTM10 = "000" And InStr(T案收文齊備排除, textCP10) = 0) Or _
                  (m_TM01 = "T" And textTM10 = "020"))) Then
            MsgBox Left(Label64.Caption, 2) & "是否齊備不可空白!!!", vbExclamation + vbOKOnly
         'end 2018/12/10
            Me.textEP06.SetFocus
            GoTo EXITSUB
         End If
         If textEP34 = "" And textEP34.Visible = True Then  'Modified by Lydia 2018/12/10 +判斷顯示
            MsgBox "是否會稿不可空白!!!", vbExclamation + vbOKOnly
            Me.textEP34.SetFocus
            GoTo EXITSUB
         End If
         'Modify By Sindy 2022/11/23
         'If textCP122 = "" And textCP122.Visible = True Then  'Modified by Lydia 2018/12/10 +判斷顯示
         If textCP122 = "" And textCP122.Enabled = True Then  'Modified by Lydia 2018/12/10 +判斷顯示
            MsgBox "是否急件不可空白!!!", vbExclamation + vbOKOnly
            Me.textCP122.SetFocus
            GoTo EXITSUB
         End If
         '若本所期限在7個日曆天內將到期或急件，要會稿且費用<8000元者，彈訊息且讓使用者可選擇繼續存檔
         'Modified by Lydia 2018/12/10 +判斷會稿顯示 textEP34.Visible = True
         'Modified by Lydia 2022/07/15 TC案之文件齊備日管控: 限制T,FCT案 And (m_TM01 = "T" Or m_TM01 = "FCT")
         If ((Val(DBDATE(textCP06)) > 0 And Val(DBDATE(textCP06)) <= Val(CompDate(2, 7, strSrvDate(1)))) Or _
             textCP122 = "Y") And _
            textEP34 = "Y" And textEP34.Visible = True And _
            Val(m_CP16) < 8000 And (m_TM01 = "T" Or m_TM01 = "FCT") Then
            If MsgBox("本所期限在7天內或急件且收費低於8,000元且要會稿，此為特殊案件，請注意有主管核可才可分案！是否要存檔？", vbExclamation + vbYesNo + vbDefaultButton2, "重要訊息！") = vbNo Then
               GoTo EXITSUB
            End If
         End If
      End If
      '2012/5/8 End
      
   '若執行轉本所案號功能才要檢查
    Else
        ' 轉本所案號
        If IsEmptyText(textTM01) = False Then
           If IsEmptyText(textTM02) = True Then
              strTit = "檢核資料"
              strMsg = "轉本所案號輸入不完整"
              nResponse = MsgBox(strMsg, vbOKOnly, strTit)
              textTM02.SetFocus
              GoTo EXITSUB
           End If
           If textTM01 = "TF" And IsEmptyText(textTM02_2) = True Then
              strTit = "檢核資料"
              strMsg = "轉本所案號輸入不完整"
              nResponse = MsgBox(strMsg, vbOKOnly, strTit)
              textTM02.SetFocus
              GoTo EXITSUB
           End If
        Else
           If IsEmptyText(textTM02) = False Or IsEmptyText(textTM03) = False Or IsEmptyText(textTM04) = False Then
              strTit = "檢核資料"
              strMsg = "轉本所案號輸入不完整"
              nResponse = MsgBox(strMsg, vbOKOnly, strTit)
              textTM02.SetFocus
              GoTo EXITSUB
           End If
        End If
        'Add By Sindy 2010/12/3
        If textTM01 = m_TM01 And textTM02 = m_TM02 And textTM03 = m_TM03 And textTM04 = m_TM04 Then
            strTit = "檢核資料"
            strMsg = "轉本所案號不可與原本所案號相同 !"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textTM02.SetFocus
            GoTo EXITSUB
        End If
        '2010/12/3 End
    End If
    
   CheckDataValid = True
EXITSUB:
End Function

Private Sub textTM04_Validate(Cancel As Boolean)
   'Modify By Cheng 2002/09/09
'   'Add By Cheng 2002/08/23
'   If Me.textTM01.Text <> "" And Me.textTM02.Text <> "" Then
'      MsgBox "不可同時做" & """轉本所案號""" & "及" & """修改其他欄位""" & "，否則資料可能有不一致的情形產生!!!", vbExclamation + vbOKOnly
'   End If
End Sub

Private Sub textTM05_1_GotFocus()
    TextInverse Me.textTM05_1
End Sub

Private Sub textTM05_1_Validate(Cancel As Boolean)
Dim strTit As String
Dim strMsg As String
Dim nResponse
    
    Cancel = False
    If CheckLengthIsOK(Me.textTM05_1.Text, textTM05_1.MaxLength) = False Then
        Cancel = True
        strTit = "檢核資料"
        strMsg = "案件名稱內容太長"
        textTM05_1_GotFocus
    End If
End Sub

' 案件中文名稱
Private Sub textTM05_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   '92.10.31 MODIFY BY SONIA
   'If CheckLengthIsOK(textTM05, 40) = False Then
   '   Cancel = True
   '   strTit = "檢核資料"
   '   strMsg = "案件中文名稱內容太長"
   '   'nResponse = MsgBox(strMsg, vbOKOnly, strTit)
   '   textTM05_GotFocus
   'End If
   If m_TM01 = "T" Or m_TM01 = "FCT" Or m_TM01 = "TF" Or m_TM01 = "CFT" Then
      If CheckLengthIsOK(textTM05, 140) = False Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "案件中文名稱內容太長"
         textTM05_GotFocus
      End If
   Else '服務業務
      If CheckLengthIsOK(textTM05, 140) = False Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "案件中文名稱內容太長"
         textTM05_GotFocus
      End If
   End If
   '92.10.31 END
'   If Cancel = False Then: textTM05.IMEMode = 2
'edit by nickc 2007/06/06 切換輸入法改用API
If Cancel = False Then CloseIme
End Sub

' 案件英文名稱
Private Sub textTM06_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If CheckLengthIsOK(textTM06, textTM06.MaxLength) = False Then
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
   '92.10.31 MODIFY BY SONIA
   'If CheckLengthIsOK(textTM07, 40) = False Then
   '   Cancel = True
   '   strTit = "檢核資料"
   '   strMsg = "案件日文名稱內容太長"
   '   'nResponse = MsgBox(strMsg, vbOKOnly, strTit)
   '   textTM07_GotFocus
   'End If
   If m_TM01 = "T" Or m_TM01 = "FCT" Or m_TM01 = "TF" Or m_TM01 = "CFT" Then
      If CheckLengthIsOK(textTM07, 40) = False Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "案件日文名稱內容太長"
         textTM07_GotFocus
      End If
   Else '服務業務
      If CheckLengthIsOK(textTM07, 60) = False Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "案件日文名稱內容太長"
         textTM07_GotFocus
      End If
   End If
   '92.10.31 END
   'edit by nickc 2007/06/06 切換輸入法改用API
   'If Cancel = False Then: textTM07.IMEMode = 2
   If Cancel = False Then CloseIme
End Sub

' 商標種類
Private Sub textTM08_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   
   textTM08_2 = Empty
   If IsEmptyText(textTM08) = False Then
      textTM08_2 = GetTradeMarkName(textTM08, 0)
      If IsEmptyText(textTM08_2) = True Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "商標種類不存在"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textTM08_GotFocus
         Exit Sub
      End If
      'Add By Sindy 2013/7/17
      '台灣案時, 檢查商標種類不可輸入2,4,5,6
      If Trim(textTM10) = "000" And _
         (Trim(textTM08) = "2" Or Trim(textTM08) = "4" Or Trim(textTM08) = "5" Or Trim(textTM08) = "6") Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "台灣案時, 商標種類不可輸入2,4,5,6！"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textTM08_GotFocus
         Exit Sub
      End If
      '2013/7/17 End
      'Add By Sindy 2015/6/30 證明標章時商品類別為證
      'modify by sonia 2020/12/28 +textTM09 = ""條件
      If textTM08 = "7" And (m_TM01 = "FCT" Or m_TM01 = "T") And textTM09 = "" Then
         textTM09 = "證"
      'add by sonia 2020/12/28
      ElseIf textTM08 = "8" And (m_TM01 = "T" Or m_TM01 = "FCT") And textTM09 = "" Then '8團體標章
         textTM09 = "團"
      'end 2020/12/28
      End If
      '2015/6/30 END
   End If
End Sub

' 商品類別
Private Sub textTM09_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim strTemp As String
   Dim nResponse
   Dim nCount As Integer
   Dim nIndex As Integer
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   Dim arrTM09, ii As Integer 'Add By Sindy 2011/3/18
   
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
   
   'Add By Sindy 2011/3/18
   arrTM09 = Split(textTM09, ",")
   For ii = LBound(arrTM09) To UBound(arrTM09)
      If Len(arrTM09(ii)) < 2 Or Len(arrTM09(ii)) > 3 Then
         If textTM08 <> "7" And textTM08 <> "8" Then  'Add By Sindy 2015/6/30 +if  'modify by sonia 2020/12/28 +textTM08 <> "8"
            MsgBox "商品類別 <" & arrTM09(ii) & "> 不可小於二碼且不可大於三碼!!!", vbExclamation + vbOKOnly
            Cancel = True
            textTM09_GotFocus
            GoTo EXITSUB
         End If
      End If
   Next ii
   
'add by nickc 2005/06/03
textTM09 = Replace(textTM09, " ", "")
EXITSUB:
   Set rsTmp = Nothing
End Sub

' 申請國家
Private Sub textTM10_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   
   textTM10_2 = Empty
   If IsEmptyText(textTM10) = False Then
      textTM10_2 = GetNationName(textTM10, 0)
      If IsEmptyText(textTM10_2) = True Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "申請國家代號不正確"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textTM10_GotFocus
      'Added by Lydia 2020/05/20 法律所案源收文
      ElseIf m_TM10 <> textTM10 Then
         SetLOSagree
         'Added by Lydia 2022/07/15 T案之齊備日管控; TC案之文件齊備日管控
         If m_TM10 <> "" And textCP05 <> "" Then  'Modfied by Lydia 2022/07/21 +收文日非空白；因為連續分案遇到更換國家時，尚未傳入
             Call setFrame21
         End If
         'end 2022/07/15
      'end 2020/05/20
      End If
   End If
   SetFrame1 'Added by Morgan 2022/12/15
End Sub

'Add By Sindy 2013/12/16
Private Sub textTM130_GotFocus()
   TextInverse textTM130
   CloseIme
End Sub
Private Sub textTM130_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 8 And KeyAscii <> Asc("J") Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub textTM136_GotFocus()
   TextInverse textTM136
End Sub

Private Sub textTM136_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 8 And KeyAscii <> Asc("1") And KeyAscii <> Asc("2") Then
      KeyAscii = 0
   End If
End Sub

'Add By Sindy 2010/7/16 公告日
Private Sub textTM14_GotFocus()
   InverseTextBox textTM14
End Sub

'Add By Sindy 2010/7/16 公告日
Private Sub textTM14_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textTM14) = False Then
      If CheckIsTaiwanDate(textTM14, False) = False Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "公告日日期格式不正確"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      End If
      'modify by sonia 2023/10/13 +627部分異議
      If textCP10 = "601" Or textCP10 = "627" Then '公告日起算三個月為管制日
         '減1天為法定
         'modify by sonia 2023/6/9 改為公告日起算三個月為法定，法定-2工作天為本所
         'textCP07 = DBDATE(DateAdd("d", -1, DateAdd("m", 3, ChangeWStringToWDateString(DBDATE(Me.textTM14.Text))))) - 19110000
         textCP07 = DBDATE(DateAdd("m", 3, ChangeWStringToWDateString(DBDATE(Me.textTM14.Text)))) - 19110000
         'Modify By Sindy 2014/10/6 台灣案之本所期限設定
         If m_TM10 = "000" And Val(strSrvDate(1)) >= 台灣案所限新規則啟用日 Then
            textCP06 = DBDATE(PUB_GetOurDeadline(DBDATE(textCP07))) - 19110000
         Else
         '2014/10/6 END
            '減3天為本所
            textCP06 = DBDATE(DateAdd("d", -3, DateAdd("m", 3, ChangeWStringToWDateString(DBDATE(Me.textTM14.Text))))) - 19110000
            textCP06.Text = TransDate(PUB_GetWorkDay1(textCP06.Text, True), 1) 'Added by Lydia 2020/07/07 本所期限檢查：若本所期限非工作天則直接調整至最近的工作天
         End If
      End If
   End If
End Sub

Private Sub textTM23_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

' 申請人
Private Sub textTM23_Validate(Cancel As Boolean)
Dim strTit As String
Dim strMsg As String
Dim nResponse
   
    Cancel = False
    textTM23_2 = Empty
    If IsEmptyText(textTM23) = False Then
        'Add By Cheng 2004/04/20
        '申請人代號補滿9碼
        Me.textTM23.Text = ChangeCustomerL(Me.textTM23.Text)
        'End
        textTM23_2 = GetCustomerName(textTM23, 0)
        If textTM23_2 = Empty Then
            Cancel = True
            strTit = "檢核資料"
            strMsg = "申請人代碼<" & textTM23 & ">不正確"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
        Else
            'Add By Cheng 2002/08/22
            'Modified by Lydia 2024/06/13
            'If Me.textTM23.Text <> m_strCust1 Then
            If ChangeCustomerL(Me.textTM23.Text) <> m_TM23 Then
                If Not PUB_EditCustOk(m_CP09, m_TM01, m_TM02, m_TM03, m_TM04) Then Cancel = True
            End If
            '在OnUpdateTrademark下執行
'            If Cancel = False Then
'                '910626 Sieg 501
'                If m_CP60 <> "" And InStr(ChangeCustomerL(m_TM23), ChangeCustomerL(textTM23)) = 0 Then
'                    strExc(1) = m_TM01
'                    strExc(2) = m_TM02
'                    strExc(3) = m_TM03
'                    strExc(4) = m_TM04
'                    strExc(5) = m_CP60
'                    strExc(6) = textTM23
'                    strExc(7) = textTM23_2
'                    '911118 nick 新增申請人
'                    strExc(8) = m_TM23
'                    If Not objLawDll.UpdAcc0k0(strExc()) Then
'                        textTM23_2 = ""
'                        Cancel = True
'                    End If
'                End If
'            End If
            ' 91.01.22 modify by louis (更新申請人地址)
            '2005/11/18 MODIFY BY SONIA 修改申請人時才更新地址
            'UpdateCustomerAddress
            If InStr(ChangeCustomerL(m_TM23), ChangeCustomerL(textTM23)) = 0 Then
               UpdateCustomerAddress
            End If
            '2005/11/18 END
        End If
    End If
    If Cancel = True Then textTM23_GotFocus
End Sub

' 申請人
Private Sub textSP58_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   Cancel = False
   textSP58_2 = Empty
   If IsEmptyText(textSP58) = False Then
        'Add By Cheng 2004/04/20
        '申請人代號補滿9碼
        Me.textSP58.Text = ChangeCustomerL(Me.textSP58.Text)
        'End
      textSP58_2 = GetCustomerName(textSP58, 0)
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
      'Modified by Lydia 2024/06/13
      'If Me.textSP58.Text <> m_strCust2 Then
      If ChangeCustomerL(Me.textSP58.Text) <> m_TM78 Then
         If Not PUB_EditCustOk(m_CP09, m_TM01, m_TM02, m_TM03, m_TM04) Then Cancel = True
      End If
      'add by nickc 2006/12/14
      If m_TM01 = "T" Or m_TM01 = "TF" Or m_TM01 = "FCT" Then
        If InStr(ChangeCustomerL(m_TM78), ChangeCustomerL(textSP58)) = 0 Then
           UpdateCustomerAddress2
        End If
      End If
   End If
   If Cancel = True Then textSP58_GotFocus
End Sub

' 申請人
Private Sub textSP59_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   Cancel = False
   textSP59_2 = Empty
   If IsEmptyText(textSP59) = False Then
        'Add By Cheng 2004/04/20
        '申請人代碼補滿9碼
        Me.textSP59.Text = ChangeCustomerL(Me.textSP59.Text)
        'End
      textSP59_2 = GetCustomerName(textSP59, 0)
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
      'Modified by Lydia 2024/06/13
      'If Me.textSP59.Text <> m_strCust3 Then
      If ChangeCustomerL(Me.textSP59.Text) <> m_TM79 Then
         If Not PUB_EditCustOk(m_CP09, m_TM01, m_TM02, m_TM03, m_TM04) Then Cancel = True
      End If
      'add by nickc 2006/12/14
      If m_TM01 = "T" Or m_TM01 = "TF" Or m_TM01 = "FCT" Then
        If InStr(ChangeCustomerL(m_TM79), ChangeCustomerL(textSP59)) = 0 Then
           UpdateCustomerAddress3
        End If
      End If
      
   End If
   If Cancel = True Then textSP59_GotFocus
End Sub

' 申請地址(中)
Private Sub textTM24_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textTM24) = False Then
      'edit by nickc 2007/05/03 長度不符
      'If CheckLengthIsOK(textTM24, 70) = False Then
      If CheckLengthIsOK(textTM24, textTM24.MaxLength) = False Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "申請地址(中)內容太長"
         'nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      End If
   End If
   If Cancel Then TextInverse textTM24
End Sub

' 申請地址(英)
Private Sub textTM25_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textTM25) = False Then
      If CheckLengthIsOK(textTM25, textTM25.MaxLength) = False Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "申請地址(英)內容太長"
         'nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textTM25_GotFocus
      End If
   End If
End Sub

' 申請地址(日)
Private Sub textTM26_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textTM26) = False Then
     If CheckLengthIsOK(textTM26, textTM26.MaxLength) = False Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "申請地址(日)內容太長"
         'nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textTM26_GotFocus
      End If
   End If
End Sub

' 卷宗性質
Private Sub textTM28_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   Cancel = False
   If IsEmptyText(textTM28) = False Then
      If IsEmptyText(textCP10) = False Then
         Select Case textCP10
            ' 異議   'modify by sonia 2019/10/4 +627部分異議
            Case "601", "627":
               If textTM28 <> "2" Then
                  Cancel = True
                  strTit = "檢核資料"
                  strMsg = "卷宗性質不正確"
                  nResponse = MsgBox(strMsg, vbOKOnly, strTit)
                  textTM28_GotFocus
               End If
            ' 評定   'modify by sonia 2019/10/4 +629部分評定
            Case "603", "629":
               If textTM28 <> "3" Then
                  Cancel = True
                  strTit = "檢核資料"
                  strMsg = "卷宗性質不正確"
                  nResponse = MsgBox(strMsg, vbOKOnly, strTit)
                  textTM28_GotFocus
               End If
            ' 廢止   'modify by sonia 2019/10/4 +623部分廢止
            Case "605", "623":
               If textTM28 <> "4" Then
                  Cancel = True
                  strTit = "檢核資料"
                  strMsg = "卷宗性質不正確"
                  nResponse = MsgBox(strMsg, vbOKOnly, strTit)
                  textTM28_GotFocus
               End If
            '2011/11/4 add by sonia
            '2012/12/19 MODIFY BY SONIA 加501(FCT-032571)
            Case "202", "501"
               If textTM28 <> "1" Then
                  Cancel = True
                  strTit = "檢核資料"
                  strMsg = "卷宗性質不正確"
                  nResponse = MsgBox(strMsg, vbOKOnly, strTit)
                  textTM28_GotFocus
               End If
            Case "210"
               If textTM28 = "1" Then
                  Cancel = True
                  strTit = "檢核資料"
                  strMsg = "卷宗性質不正確"
                  nResponse = MsgBox(strMsg, vbOKOnly, strTit)
                  textTM28_GotFocus
               End If
            '2011/11/4 end
            Case Else:
               '91.11.10 MODIFY BY SONIA
               'If textTM28 <> "1" Then
               '   Cancel = True
               '   strTit = "檢核資料"
               '   strMsg = "卷宗性質不正確"
               '   nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               '   textTM28_GotFocus
               'End If
               '91.11.10 END
         End Select
      End If
   End If
End Sub

Private Sub textTM29_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

' 是否閉卷
Private Sub textTM29_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   Cancel = False
   If IsEmptyText(textTM29) = False Then
      Select Case textTM29
         Case "Y", " ":
         Case Else
            Cancel = True
            strTit = "檢核資料"
            strMsg = "請輸入空白或Y"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textTM29_GotFocus
      End Select
   '2006/5/12 ADD BY SONIA
   Else
      '2012/5/3 MODIFY BY SONIA 退費不控制 T-158649
      'If m_TM29 = "Y" Then
      If m_TM29 = "Y" And textCP10 <> "725" Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "此案已閉卷, 應該要取消閉卷 ?"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textTM29_GotFocus
      End If
   '2006/5/12 END
   End If
End Sub

'add by nickc 2006/12/15
Private Sub textTM32_GotFocus()
InverseTextBox textTM32
End Sub
Private Sub textTM32_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim strTemp As String
   Dim nResponse
   Dim nCount As Integer
   Dim nIndex As Integer
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   
   Cancel = False
   If IsEmptyText(textTM32) = True Then
      GoTo EXITSUB
   End If
   
   'Modify By Sindy 2024/4/18 商品組群欄人員貼上資料後將全形或半形的「；」分號，轉為半形的逗號存入TM32。
   textTM32 = Replace(Replace(textTM32, ";", ","), "；", ",")
   '2024/4/18 END
   nCount = GetSubStringCount(textTM32)
   For nIndex = 1 To nCount
      strTemp = GetSubString(textTM32, nIndex)
      If Len(strTemp) > 6 Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "商品組群<" & strTemp & ">不正確"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textTM32_GotFocus
         GoTo EXITSUB
      End If
   Next nIndex
   
   For nIndex = 1 To nCount
      strTemp = GetSubString(textTM32, nIndex)
      For nCount = 1 To nCount
         If nIndex <> nCount Then
            If strTemp = GetSubString(textTM32, nCount) Then
               Cancel = True
               strTit = "檢核資料"
               strMsg = "商品組群<" & strTemp & ">不可重覆"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textTM32_GotFocus
               GoTo EXITSUB
            End If
         End If
      Next nCount
   Next nIndex
textTM32 = Replace(textTM32, " ", "")
EXITSUB:
   Set rsTmp = Nothing
End Sub

Private Sub textTM34_GotFocus()
   InverseTextBox textTM34
End Sub

Private Sub textTM34_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If CheckLengthIsOK(textTM34, textTM34.MaxLength) = False Then
      Cancel = True
      strTit = "檢核資料"
      strMsg = "分所號內容太長"
      'nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textTM34_GotFocus
   End If
End Sub

'add by nickc 2008/01/31 新增加聯絡人1(中)
Private Sub textTM38_GotFocus()
   InverseTextBox textTM38
End Sub
Private Sub textTM38_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   'Memo by Lydia 2017/06/14 聯絡人(中)改為30字;若為服務業務會改為60字
   If CheckLengthIsOK(textTM38, textTM38.MaxLength) = False Then
      Cancel = True
      strTit = "檢核資料"
      strMsg = "聯絡人1(中)太長"
      textTM38_GotFocus
   End If
End Sub

'Add By Sindy 2015/2/26 新增加聯絡人1(英)
Private Sub textTM39_GotFocus()
   InverseTextBox textTM39
End Sub
Private Sub textTM39_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   'Modified by Lydia 2017/06/14
   'If CheckLengthIsOK(textTM39, textTM39.MaxLength) = False Then
   If CheckLengthIsOK(textTM39, 35) = False Then
      Cancel = True
      strTit = "檢核資料"
      strMsg = "聯絡人1(英)太長"
      textTM39_GotFocus
   End If
End Sub

'Add By Sindy 2015/2/26 新增加聯絡人1(日)
Private Sub textTM40_GotFocus()
   InverseTextBox textTM40
End Sub
Private Sub textTM40_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   'Modified by Lydia 2017/06/14
   'If CheckLengthIsOK(textTM40, textTM40.MaxLength) = False Then
   If CheckLengthIsOK(textTM40, 60) = False Then
      Cancel = True
      strTit = "檢核資料"
      strMsg = "聯絡人1(日)太長"
      textTM40_GotFocus
   End If
End Sub

'add by Sindy 2012/12/20 新增加聯絡人2(中)
Private Sub textTM41_GotFocus()
   InverseTextBox textTM41
End Sub
Private Sub textTM41_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   'Modified by Lydia 2017/06/14 聯絡人(中)改為30字
   'If CheckLengthIsOK(textTM41, textTM41.MaxLength) = False Then
   If CheckLengthIsOK(textTM41, 30) = False Then
      Cancel = True
      strTit = "檢核資料"
      strMsg = "聯絡人2(中)太長"
      textTM41_GotFocus
   End If
End Sub

'Add By Sindy 2015/2/26 新增加聯絡人2(英)
Private Sub textTM42_GotFocus()
   InverseTextBox textTM42
End Sub
Private Sub textTM42_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   'Modified by Lydia 2017/06/14
   'If CheckLengthIsOK(textTM42, textTM42.MaxLength) = False Then
   If CheckLengthIsOK(textTM42, 35) = False Then
      Cancel = True
      strTit = "檢核資料"
      strMsg = "聯絡人2(英)太長"
      textTM42_GotFocus
   End If
End Sub

'Add By Sindy 2015/2/26 新增加聯絡人2(日)
Private Sub textTM43_GotFocus()
   InverseTextBox textTM43
End Sub
Private Sub textTM43_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   'Modified by Lydia 2017/06/14
   'If CheckLengthIsOK(textTM43, textTM43.MaxLength) = False Then
   If CheckLengthIsOK(textTM43, 60) = False Then
      Cancel = True
      strTit = "檢核資料"
      strMsg = "聯絡人2(日)太長"
      textTM43_GotFocus
   End If
End Sub

Private Sub textTM44_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

' FC代理人
Private Sub textTM44_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   'Add By Cheng 2002/07/09
   Dim strTempName As String
   
   Cancel = False
   textTM44_2 = Empty
   If IsEmptyText(textTM44) = False Then
        'Add By Cheng 2004/04/20
        '代理人代號補滿9碼
        Me.textTM44.Text = ChangeCustomerL(Me.textTM44.Text)
        'End
      'Modify By Cheng 2002/07/09
'      textTM44_2 = GetFAgentName(textTM44)
      If PUB_GetAgentName(m_TM01, textTM44.Text, strTempName) Then
         textTM44_2.Text = strTempName
      Else
         textTM44_2.Text = ""
      End If
      If IsEmptyText(textTM44_2) = True Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "FC代理人<" & textTM44 & ">不正確"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textTM44_GotFocus
      End If
   End If
End Sub
'add by nickc 2008/02/01 CF 代理人
Private Sub textCP44_GotFocus()
   InverseTextBox textCP44
   CloseIme
End Sub
Private Sub textCP44_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub
Private Sub textCP44_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Dim strTempName As String
   
   Cancel = False
   textCP44_2 = Empty
   If IsEmptyText(textCP44) = False Then
        Me.textCP44.Text = ChangeCustomerL(Me.textCP44.Text)
      If PUB_GetAgentName(m_TM01, textCP44.Text, strTempName) Then
         textCP44_2.Text = strTempName
      Else
         textCP44_2.Text = ""
      End If
      If IsEmptyText(textCP44_2) = True Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "CF代理人<" & textCP44 & ">不正確"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP44_GotFocus
      End If
   End If
End Sub
Private Sub textTM45_GotFocus()
   InverseTextBox textTM45
End Sub

' 案件備註
Private Sub textTM58_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If CheckLengthIsOK(textTM58, textTM58.MaxLength) = False Then
      Cancel = True
      strTit = "檢核資料"
      strMsg = "案件備註內容太長"
      'nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textTM58_GotFocus
   End If
End Sub

Private Sub SetInputEntry()
   textCP14.SetFocus
End Sub

Private Sub textTM10_GotFocus()
   InverseTextBox textTM10
End Sub

Private Sub textCP05_GotFocus()
   InverseTextBox textCP05
End Sub

Private Sub textCP06_GotFocus()
   InverseTextBox textCP06
End Sub

Private Sub textCP07_GotFocus()
   InverseTextBox textCP07
End Sub

Private Sub textCP10_GotFocus()
   InverseTextBox textCP10
End Sub

Private Sub textCP13_GotFocus()
   InverseTextBox textCP13
End Sub

Private Sub textCP14_GotFocus()
   InverseTextBox textCP14
   CloseIme
End Sub

Private Sub textCP26_GotFocus()
   InverseTextBox textCP26
End Sub

Private Sub textCP64_GotFocus()
   InverseTextBox textCP64
End Sub

Private Sub textTM01_GotFocus()
   InverseTextBox textTM01
End Sub

Private Sub textTM02_GotFocus()
   InverseTextBox textTM02
End Sub

Private Sub textTM02_2_GotFocus()
   InverseTextBox textTM02_2
End Sub

Private Sub textTM03_GotFocus()
   InverseTextBox textTM03
End Sub

Private Sub textTM04_GotFocus()
   InverseTextBox textTM04
End Sub

Private Sub textTM08_GotFocus()
   InverseTextBox textTM08
End Sub

Private Sub textTM28_GotFocus()
   InverseTextBox textTM28
End Sub

Private Sub textCP43_GotFocus()
   InverseTextBox textCP43
End Sub

Private Sub textTM23_GotFocus()
   InverseTextBox textTM23
End Sub

Private Sub textTM24_GotFocus()
   InverseTextBox textTM24
End Sub

Private Sub textTM25_GotFocus()
   InverseTextBox textTM25
End Sub

Private Sub textTM26_GotFocus()
   InverseTextBox textTM26
End Sub

Private Sub textTM29_GotFocus()
   InverseTextBox textTM29
End Sub

Private Sub textTM44_GotFocus()
   InverseTextBox textTM44
End Sub

Private Sub textSP32_GotFocus()
   InverseTextBox textSP32
End Sub

Private Sub textSP58_GotFocus()
   InverseTextBox textSP58
End Sub

Private Sub textSP59_GotFocus()
   InverseTextBox textSP59
End Sub

Private Sub textTM09_GotFocus()
   InverseTextBox textTM09
End Sub

Private Sub textTM58_GotFocus()
   InverseTextBox textTM58
End Sub

Private Sub textTM05_GotFocus()
   InverseTextBox textTM05
   'edit by nickc 2007/06/06 切換輸入法改用API
   OpenIme
End Sub

Private Sub textTM06_GotFocus()
   InverseTextBox textTM06
End Sub

Private Sub textTM07_GotFocus()
   InverseTextBox textTM07
   'edit by nickc 2007/06/06 切換輸入法改用API
   OpenIme
End Sub

'Add By Cheng 2002/05/23
Private Function TxtValidate() As Boolean
Dim objTxt As Object
Dim ii As Integer
Dim Cancel As Boolean
Dim strCU01 As String, strCU02 As String 'Add By Sindy 2013/12/16
'Add by Amy 2017/03/14
Dim strTmp(0) As String, strMCTF(0) As String, strMsg As String, strApply As String
Dim bolData As Boolean

TxtValidate = False

'Add by Amy 2021/12/21檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
If PUB_ChkUniText(Me, , True) = False Then
    Exit Function
End If

'Modify By Cheng 2002/11/22
'若執行轉本所案號功能才要檢查
If Me.textTM01.Text <> "" And Me.textTM02.Text <> "" Then

    '910722 Sieg
    If textTM01 <> "" And textTM02 <> "" Then
       Dim strTM01 As String
       Dim strTM02 As String
       Dim strTM03 As String
       Dim strTM04 As String
       strTM01 = textTM01
       strTM02 = textTM02
       If strTM01 = "TF" Then: strTM02 = strTM02 & textTM02_2
       If IsEmptyText(textTM03) = True Then: textTM03 = "0"
       If IsEmptyText(textTM04) = True Then: textTM04 = "00"
       strTM03 = textTM03
       strTM04 = textTM04
    
       If Not IsDataRecordExist(strTM01, strTM02, strTM03, strTM04) Then
          strExc(1) = strTM01
          strExc(2) = strTM02
          strExc(3) = strTM03
          strExc(4) = strTM04
          If Not chkNewTMNo(strExc, intI) Then
             Select Case intI
                Case 1
                   textTM01.SetFocus
                Case 2
                   textTM02.SetFocus
             End Select
             Exit Function
          End If
       End If
    End If

'若非執行轉本所案號功能才要檢查
Else

    If Me.textCP05.Enabled = True Then
       Cancel = False
       textCP05_Validate Cancel
       If Cancel = True Then
          Exit Function
       End If
    End If
    
    If Me.textCP06.Enabled = True Then
       Cancel = False
       textCP06_Validate Cancel
       If Cancel = True Then
          Exit Function
       End If
    End If
    
    If Me.textCP07.Enabled = True Then
       Cancel = False
       textCP07_Validate Cancel
       If Cancel = True Then
          Exit Function
       End If
    End If
    
    If Me.textCP10.Enabled = True Then
       Cancel = False
       textCP10_Validate Cancel
       If Cancel = True Then
          Exit Function
       End If
    End If
    
    If Me.textCP13.Enabled = True Then
       Cancel = False
       textCP13_Validate Cancel
       If Cancel = True Then
          Exit Function
       End If
    End If
    
    If Me.textCP14.Enabled = True Then
       Cancel = False
       textCP14_Validate Cancel
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
    
    If Me.textCP43.Enabled = True Then
       Cancel = False
       textCP43_Validate Cancel
       If Cancel = True Then
          Exit Function
       End If
      '2008/11/18 add by sonia
      '2011/6/10 modify by sonia 大->台才一定要輸,因為發文定稿要抓 T-173388
      If textCP10 = "313" And textCP43 = "" And m_TM77 = "2" Then
         'Modified by Lydia 2024/12/25 改成提醒不限制
         'MsgBox "減縮商品案, 請輸入C類相關總收文號！", vbExclamation
         If MsgBox("減縮商品案, 是否要輸入C類相關總收文號？", vbInformation + vbYesNo + vbDefaultButton1) = vbYes Then
            Cancel = True
            Exit Function
         End If 'Added by Lydia 2024/12/25
      End If
      '2008/11/18 end
      '2009/10/14 add by sonia
      If textCP10 = "725" And textCP43 = "" Then
         MsgBox "退費, 請輸入相關總收文號！", vbExclamation
         Cancel = True
         Exit Function
      End If
      '2009/10/14 end
      '2012/7/2 ADD BY SONIA
      If textCP10 = "310" And textCP43 = "" Then
         MsgBox "暫緩審理案件, 請輸入相關總收文號!!! 可按 案件進度 按鈕 點選 !!", vbExclamation
         Cancel = True
         Exit Function
      End If
      '2012/7/2 END
    End If
    
   'Add by Sindy 2023/4/21
   If OptSendType(3).Value = True And textCP142.Enabled = True Then
      If textCP142.Text = "" Then
         MsgBox "送件方式選指定日期送件時，指定日期不可空白！"
         textCP142.SetFocus
         Exit Function
      Else
         Cancel = False
         textCP142_Validate Cancel
         If Cancel = True Then
            textCP142.SetFocus
            Exit Function
         End If
         'Add By Sindy 2023/12/11 檢查指定送件日相關欄位
         If Frame3.Visible = True Then
            If Option1(0).Value = False And Option1(1).Value = False And Option1(2).Value = False Then
               MsgBox "有輸入指定送件日，當天或之前或之後請擇一。", vbExclamation
               Exit Function
            'Add By Sindy 2025/4/25
            'Modify By Sindy 2025/5/19 有顯示出齊備日欄位時,才需要檢查此條件 Ex:TS就不用檢查。 + and Frame21.Visible = True
            ElseIf ((Trim(textEP06) = "" Or textEP06 = "N")) And _
                   (Option1(0).Value = True Or Option1(1).Value = True) And _
                   Frame21.Visible = True Then
               MsgBox "若「文件未齊備」時,不得指定日期「當日」及「之前」送件。", vbExclamation
               Exit Function
            '2025/4/25 END
            End If
         End If
         '2023/12/11 END
      End If
   'Add By Sindy 2023/12/11
   Else
      Option1(0).Value = False
      Option1(1).Value = False
      Option1(2).Value = False
   '2023/12/11 END
   End If
    
    If Me.textCP64.Enabled = True Then
       Cancel = False
       textCP64_Validate Cancel
       If Cancel = True Then
          Exit Function
       End If
    End If
    'add by nickc 2008/02/01
    If Me.textCP44.Enabled = True Then
       Cancel = False
       'SSTab1.Tab = 2
       textCP44_Validate Cancel
       If Cancel = True Then
          Exit Function
       End If
    End If
    If Me.textSP32.Enabled = True Then
       Cancel = False
       textSP32_Validate Cancel
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
    
    If Me.textTM01.Enabled = True Then
       Cancel = False
       textTM01_Validate Cancel
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
    
    If Me.textTM08.Enabled = True Then
       Cancel = False
       textTM08_Validate Cancel
       If Cancel = True Then
          Exit Function
       End If
    End If
    
    'Add By Sindy 2019/4/9
    If Me.textTM72.Enabled = True Then
       Cancel = False
       textTM72_Validate Cancel
       If Cancel = True Then
          Exit Function
       End If
    End If

   'Added by Lydia 2023/11/16
   If Me.cboTM08.Enabled = True Then
      Cancel = False
      cboTM08_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If Me.cboTM72.Enabled = True Then
      Cancel = False
      cboTM72_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   'end 2023/11/16
   
    If Me.textTM09.Enabled = True Then
       Cancel = False
       textTM09_Validate Cancel
       If Cancel = True Then
          Exit Function
       End If
    End If
    
    If Me.textTM10.Enabled = True Then
       Cancel = False
       textTM10_Validate Cancel
       If Cancel = True Then
          Exit Function
       End If
    End If
    
    'Add By Sindy 2010/7/16 公告日
    If Me.textTM14.Enabled = True Then
       Cancel = False
       textTM14_Validate Cancel
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
    
    If Me.textTM24.Enabled = True Then
       Cancel = False
       textTM24_Validate Cancel
       If Cancel = True Then
          Exit Function
       End If
    End If
    
    If Me.textTM25.Enabled = True Then
       Cancel = False
       textTM25_Validate Cancel
       If Cancel = True Then
          Exit Function
       End If
    End If
    
    If Me.textTM26.Enabled = True Then
       Cancel = False
       textTM26_Validate Cancel
       If Cancel = True Then
          Exit Function
       End If
    End If
    
    If Me.textTM28.Enabled = True Then
       Cancel = False
       textTM28_Validate Cancel
       If Cancel = True Then
          Exit Function
       End If
    End If
    
    If Me.textTM29.Enabled = True Then
       Cancel = False
       textTM29_Validate Cancel
       If Cancel = True Then
          Exit Function
       End If
    End If
    
    If Me.textTM44.Enabled = True Then
       Cancel = False
       textTM44_Validate Cancel
       If Cancel = True Then
          Exit Function
       End If
    End If
    
    If Me.textTM58.Enabled = True Then
       Cancel = False
       textTM58_Validate Cancel
       If Cancel = True Then
          Exit Function
       End If
    End If

'add by nickc 2006/12/15
    If Me.textTM80.Enabled = True Then
       Cancel = False
       textTM80_Validate Cancel
       If Cancel = True Then
          Exit Function
       End If
    End If
    If Me.textTM81.Enabled = True Then
       Cancel = False
       textTM81_Validate Cancel
       If Cancel = True Then
          Exit Function
       End If
    End If
    If Me.textTM82.Enabled = True Then
       Cancel = False
       textTM82_Validate Cancel
       If Cancel = True Then
          Exit Function
       End If
    End If
    If Me.textTM83.Enabled = True Then
       Cancel = False
       textTM83_Validate Cancel
       If Cancel = True Then
          Exit Function
       End If
    End If
    If Me.textTM84.Enabled = True Then
       Cancel = False
       textTM84_Validate Cancel
       If Cancel = True Then
          Exit Function
       End If
    End If
    If Me.textTM85.Enabled = True Then
       Cancel = False
       textTM85_Validate Cancel
       If Cancel = True Then
          Exit Function
       End If
    End If
    If Me.textTM86.Enabled = True Then
       Cancel = False
       textTM86_Validate Cancel
       If Cancel = True Then
          Exit Function
       End If
    End If
    If Me.textTM87.Enabled = True Then
       Cancel = False
       textTM87_Validate Cancel
       If Cancel = True Then
          Exit Function
       End If
    End If
    If Me.textTM88.Enabled = True Then
       Cancel = False
       textTM88_Validate Cancel
       If Cancel = True Then
          Exit Function
       End If
    End If
    If Me.textTM89.Enabled = True Then
       Cancel = False
       textTM89_Validate Cancel
       If Cancel = True Then
          Exit Function
       End If
    End If
    If Me.textTM90.Enabled = True Then
       Cancel = False
       textTM90_Validate Cancel
       If Cancel = True Then
          Exit Function
       End If
    End If
    If Me.textTM91.Enabled = True Then
       Cancel = False
       textTM91_Validate Cancel
       If Cancel = True Then
          Exit Function
       End If
    End If
    If Me.textTM92.Enabled = True Then
       Cancel = False
       textTM92_Validate Cancel
       If Cancel = True Then
          Exit Function
       End If
    End If
    If Me.textTM93.Enabled = True Then
       Cancel = False
       textTM93_Validate Cancel
       If Cancel = True Then
          Exit Function
       End If
    End If
    If Me.textTM32.Enabled = True Then
       Cancel = False
       textTM32_Validate Cancel
       If Cancel = True Then
          Exit Function
       End If
    End If
    
    If m_TM01 = "TF" Then
        If Trim(textTM09) = "" Then
            MsgBox "TF 必須有類別！", vbExclamation
            Cancel = True
            Exit Function
        End If
    End If
    
   'Add By Sindy 2013/12/16
   '若未設定特殊出名公司則提醒
   '2014/1/10 MODIFY BY SONIA 加入CP12 條件
   If textTM130.Visible = True And textTM130 = "" And Mid(GetST15(textCP13), 1, 1) <> "F" Then
      For ii = 1 To 5
         strCU01 = "": strCU02 = ""
         If ii = 1 Then
            strCU01 = Left(ChangeCustomerL(textTM23), 8)
            strCU02 = Right(ChangeCustomerL(textTM23), 1)
         ElseIf ii = 2 Then
            strCU01 = Left(ChangeCustomerL(textSP58), 8)
            strCU02 = Right(ChangeCustomerL(textSP58), 1)
         ElseIf ii = 3 Then
            strCU01 = Left(ChangeCustomerL(textSP59), 8)
            strCU02 = Right(ChangeCustomerL(textSP59), 1)
         ElseIf ii = 4 Then
            strCU01 = Left(ChangeCustomerL(textTM80), 8)
            strCU02 = Right(ChangeCustomerL(textTM80), 1)
         Else
            strCU01 = Left(ChangeCustomerL(textTM81), 8)
            strCU02 = Right(ChangeCustomerL(textTM81), 1)
         End If
      Next ii
   End If
   '檢查是否有客戶不開發票
   If textTM130.Visible = True And textTM130 = "J" Then
      For ii = 1 To 5
         strCU01 = "": strCU02 = ""
         If ii = 1 Then
            strCU01 = Left(ChangeCustomerL(textTM23), 8)
            strCU02 = Right(ChangeCustomerL(textTM23), 1)
         ElseIf ii = 2 Then
            strCU01 = Left(ChangeCustomerL(textSP58), 8)
            strCU02 = Right(ChangeCustomerL(textSP58), 1)
         ElseIf ii = 3 Then
            strCU01 = Left(ChangeCustomerL(textSP59), 8)
            strCU02 = Right(ChangeCustomerL(textSP59), 1)
         ElseIf ii = 4 Then
            strCU01 = Left(ChangeCustomerL(textTM80), 8)
            strCU02 = Right(ChangeCustomerL(textTM80), 1)
         Else
            strCU01 = Left(ChangeCustomerL(textTM81), 8)
            strCU02 = Right(ChangeCustomerL(textTM81), 1)
         End If
         If strCU01 <> "" Then
            'Modify By Sindy 2024/9/25 增加傳入公司別做判斷
            If PUB_ChkCU144isN(strCU01, strCU02, "", textTM130, False) = True Then
               MsgBox strCU01 & strCU02 & "此客戶為不開發票，因此特殊出名公司不可選智權公司 !", vbCritical
               Me.SSTab1.Tab = 2
               textTM130.SetFocus
               Exit Function
            End If
         End If
      Next ii
   End If
   '2013/12/16 END
   'Added by Lydia 2025/09/12 TF基礎案號設定
   'Modified by Lydia 2025/10/23
   If m_TM01 = "TF" And cmdTFBaseNo.Visible = True Then
       strExc(0) = Pub_GetField("TFBaseNo", "TFBN01='" & m_TM01 & "' AND TFBN02='" & m_TM02 & "' AND TFBN03='" & m_TM03 & "' AND TFBN04='" & m_TM04 & "'", "TFBN05")
       If strExc(0) <> "" Then
          cmdTFBaseNo.BackColor = &HC0FFC0
       Else
          cmdTFBaseNo.BackColor = &H8000000F
          If MsgBox("是否要設定TF基礎案號？", vbInformation + vbYesNo + vbDefaultButton1) = vbYes Then
             Me.SSTab1.Tab = 3
             Exit Function
          End If
       End If
   End If
   'end 2025/09/12
End If  '非執行轉本所案號功能

'Add by Amy 2017/03/14 MCTF組別控制(有輸代理人且為MCTF,判斷申請人若與代理人的MCTF組別不同不可收文)
strMsg = "": strUpdCusNo = ""
'modify by sonia 2023/11/7 96029及96030之P29部門，不是MCTF所以不檢查故加入m_SalesST15<>"P29"條件
If Len(Trim(textTM44)) > 0 And m_SalesST15 <> "P29" Then
    bolData = GetCusORFagentData(ChangeCustomerL(textTM44), "FA120", strMCTF())
    If Left(strMCTF(0), 4) = "MCTF" Then
        'modify by sonia 2021/11/29 不再判斷申請人之智權人員，改判斷收文智權人員之組別與FC代理人的MCTF組別
        'For ii = 0 To 4
        '    strApply = ""
        '    Select Case ii
        '        Case 0
        '            strApply = textTM23
        '        Case 1
        '            strApply = textSP58
        '        Case 2
        '            strApply = textSP59
        '        Case 3
        '            strApply = textTM80
        '        Case 4
        '            strApply = textTM81
        '    End Select
        '    If strApply = MsgText(601) Then Exit For
        '    bolData = GetCusORFagentData(ChangeCustomerL(strApply), "CU13", strTmp())
        '    If strMCTF(0) <> strTmp(0) And Left(strTmp(0), 4) = "MCTF" Then
        '        strMsg = strMsg & "申請人" & ii + 1 & "：" & strApply & " (" & strTmp(0) & ")" & "及"
        '    End If
        'Next ii
        If ChkMCTF0XSales(strMCTF(0), textCP13) = False Then
           MsgBox strMsg & "智權人員組別與代理人" & textTM44 & "商標管控智權人員(" & strMCTF(0) & ")不同！"
           Exit Function
        End If
        'end 2021/11/29
        'cancel by sonia 2021/11/29
        'If strMsg <> MsgText(601) Then
        '    MsgBox Left(strMsg, Len(strMsg) - 1) & vbCrLf & "與代理人" & textTM44 & _
        '                "商標管控智權人員(" & strMCTF(0) & ")不同！"
        '    Exit Function
        'End If
        'Add by Amy 2018/08/09 移轉(501)案之移轉申請人之客戶MCTF檢查
        'If textCP10 = "501" Then
        '    If ChkMCTF_Tran(strMsg, strMCTF(0)) = False And strMsg <> MsgText(601) Then
        '        MsgBox strMsg & "與代理人" & textTM44 & _
        '                    "商標管控智權人員(" & strMCTF(0) & ")不同！"
        '        Exit Function
        '    End If
        'End If
        'end 2021/11/29
        'end 2018/08/09
    End If
End If
'end 2017/03/14
    
'Added by Lydia 2019/01/30 檢查-查名是否齊備
If Me.textCP143.Visible = True Then
    Cancel = False
    textCP143_Validate Cancel
    If Cancel = True Then
       Exit Function
    End If
End If

'Added by Morgan 2022/12/15
If textTM136.Visible Then
   If strSrvDate(1) > "20230000" Then
      If textTM136 = "" Then
         MsgBox "請輸入註冊證形式！", vbExclamation
         textTM136.SetFocus
         Exit Function
      End If
   End If
End If
'end 2022/12/15

'Added by Lydia 2023/01/31 T大陸查名
'Modified by Lydia 2023/02/10 改為獨立欄位EP13=>EP43
If Frame22.Visible = True And Trim(textCP143) <> "Y" Then
   If Trim(textEP43.Text) = "" Then
       MsgBox "請輸入查名人員！", vbExclamation
       textEP43.SetFocus
       Exit Function
   Else
       textEP43_Validate Cancel
       If Cancel = True Then
           Exit Function
       End If
   End If
End If
'end 2023/01/31

'Added by Lydia 2024/06/14 對申請人1~5的重複輸入檢查
If Pub_ChkAppList(strExc(0), textTM23 & "," & textSP58 & "," & textSP59 & "," & textTM80 & "," & textTM81) = False Then
   Me.SSTab1.Tab = 0
   Select Case Val(strExc(0))
      Case 1
         textTM23.SetFocus
         textTM23_GotFocus
      Case 2
         textSP58.SetFocus
         textSP58_GotFocus
      Case 3
         textSP59.SetFocus
         textSP59_GotFocus
      Case 4
         textTM80.SetFocus
         textTM80_GotFocus
      Case 5
         textTM81.SetFocus
         textTM81_GotFocus
   End Select
   Exit Function
End If
'end 2024/06/14

'Added by Lydia 2024/06/13 檢查更新代理人／申請人狀態排除「不得代理」
For ii = 1 To 6
   strExc(1) = ""
   Select Case ii
      Case 1 '申請人1
         strExc(1) = ChangeCustomerL(textTM23)
         strExc(2) = ChangeCustomerL(m_TM23)
      Case 2 '申請人2
         strExc(1) = ChangeCustomerL(textSP58)
         strExc(2) = ChangeCustomerL(m_TM78)
      Case 3 '申請人3
         strExc(1) = ChangeCustomerL(textSP59)
         strExc(2) = ChangeCustomerL(m_TM79)
      Case 4 '申請人4
         strExc(1) = ChangeCustomerL(textTM80)
         strExc(2) = ChangeCustomerL(m_TM80)
      Case 5 '申請人5
         strExc(1) = ChangeCustomerL(textTM81)
         strExc(2) = ChangeCustomerL(m_TM81)
      Case 6 '代理人
         strExc(1) = ChangeCustomerL(textTM44)
         strExc(2) = ChangeCustomerL(m_TM44)
   End Select
   If strExc(1) <> "" And strExc(1) <> strExc(2) Then
      If Left(strExc(1), 1) = "X" Then
         If GetCustomerAndState(strExc(1), strExc(3), , , , m_TM01, strExc(8), False, Me.Name, m_TM02, m_TM03, m_TM04) = False Then
            Me.SSTab1.Tab = 0
            If ii = 1 Then
               textTM23.SetFocus
               textTM23_GotFocus
               Exit Function
            ElseIf ii = 2 Then
               textSP58.SetFocus
               textSP58_GotFocus
               Exit Function
            ElseIf ii = 3 Then
               textSP59.SetFocus
               textSP59_GotFocus
               Exit Function
            ElseIf ii = 4 Then
               textTM80.SetFocus
               textTM80_GotFocus
               Exit Function
            ElseIf ii = 5 Then
               textTM81.SetFocus
               textTM81_GotFocus
               Exit Function
            End If
         End If
      Else
         If GetAgentAndState(strExc(1), strExc(3), , , , m_TM01, strExc(2), False) = False Then
            Me.SSTab1.Tab = 0
            textTM44.SetFocus
            textTM44_GotFocus
            Exit Function
         End If
      End If
   End If
Next
'end 2024/06/13
   
TxtValidate = True
End Function

'Add By Cheng 2002/06/12
'取得案件收費表的下次期限天數
Private Function GetCF12(strCF01 As String, strCF02 As String, strCF03 As String) As String
Dim rsA As New ADODB.Recordset
Dim StrSQLa As String

GetCF12 = "0"
If rsA.State <> adStateClosed Then rsA.Close
Set rsA = Nothing
'91.11.3 MODIFY BY SONIA
'strSQLA = "Select CF12 From CaseFee Where CF01='" & strCF01 & "' AND CF02='" & strCF02 & "' AND CF03='" & strCF03 & "' AND CF12 IS NOT NULL"
StrSQLa = "Select CF12 From CaseFee Where CF01='" & strCF01 & "' AND CF02='" & strCF02 & "' AND CF03='" & strCF03 & "'"
'91.11.3 END
rsA.CursorLocation = adUseClient
rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
If rsA.RecordCount <> 0 Then
   If Not IsNull(rsA.Fields(0).Value) Then
      GetCF12 = rsA.Fields(0).Value
   End If
End If
If rsA.State <> adStateClosed Then rsA.Close
Set rsA = Nothing

End Function

'Add By SONIA 91.11.3
'取得案件收費表的下次期限月數
Private Function GetCF28(strCF01 As String, strCF02 As String, strCF03 As String) As String
Dim rsA As New ADODB.Recordset
Dim StrSQLa As String

GetCF28 = "0"
If rsA.State <> adStateClosed Then rsA.Close
Set rsA = Nothing
StrSQLa = "Select CF28 From CaseFee Where CF01='" & strCF01 & "' AND CF02='" & strCF02 & "' AND CF03='" & strCF03 & "'"
rsA.CursorLocation = adUseClient
rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
If rsA.RecordCount <> 0 Then
   If Not IsNull(rsA.Fields(0).Value) Then
      GetCF28 = rsA.Fields(0).Value
   End If
End If
If rsA.State <> adStateClosed Then rsA.Close
Set rsA = Nothing

End Function

'Add By Cheng 2002/06/12
'取得案件收費表的規費
Private Function GetCF08(strCF01 As String, strCF02 As String, strCF03 As String) As String
Dim rsA As New ADODB.Recordset
Dim StrSQLa As String

GetCF08 = "0"
If rsA.State <> adStateClosed Then rsA.Close
Set rsA = Nothing
StrSQLa = "Select CF08 From CaseFee Where CF01='" & strCF01 & "' AND CF02='" & strCF02 & "' AND CF03='" & strCF03 & "' AND CF08 IS NOT NULL"
rsA.CursorLocation = adUseClient
rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
If rsA.RecordCount > 0 Then
   GetCF08 = rsA.Fields(0).Value
End If
If rsA.State <> adStateClosed Then rsA.Close
Set rsA = Nothing

End Function

'Add By Cheng 2002/11/18
'將本案期限最上面一筆的資料帶到畫面欄位上
Private Sub PasteGridData()
Dim ii As Integer
    With Me.grdList
        For ii = 1 To Me.grdList.Rows - 1
            '若有勾選資料
            If .TextMatrix(ii, 0) <> "" Then
                '本所期限
                Me.textCP06.Text = "" & .TextMatrix(ii, 2)
                '法定期限
                Me.textCP07.Text = "" & .TextMatrix(ii, 3)
                '相關總收文號
                Me.textCP43.Text = "" & .TextMatrix(ii, 8)
                '進度備註
                Me.textCP64.Text = "" & .TextMatrix(ii, 6)
                
               'Add by Morgan 2009/12/25 延期只更新期限不可點選
               If textCP10 = "303" Then
                  .TextMatrix(ii, 0) = ""
               End If
               
               'Add by Morgan 2011/4/22
               If .TextMatrix(ii, 10) = "0" Then
                  m_CP30 = ""
               Else
                  m_CP30 = .TextMatrix(ii, 10)
               End If
               
               '2013/10/7 ADD BY SONIA 註冊費717過期收復權729,T-183459
               If textCP10 = "729" And ("" & .TextMatrix(ii, 9) = "717") Then
                  '本所期限
                  Me.textCP06.Text = strSrvDate(2)
                  '法定期限
                  Me.textCP07.Text = ChangeWDateStringToTString(DateAdd("M", 6, ChangeTStringToWDateString("" & .TextMatrix(ii, 3))))
               End If
               '2013/10/7 END
               
                Exit For
            '若取消勾選資料
            Else
                '本所期限
                Me.textCP06.Text = ""
                '法定期限
                Me.textCP07.Text = ""
                '相關總收文號
                Me.textCP43.Text = ""
                '進度備註
                Me.textCP64.Text = ""
                m_CP30 = "" 'Add by Morgan 2011/4/22
            End If
        Next ii
    End With
End Sub

Private Sub UpdateCustomerAddress()
   Dim strSql As String
   Dim rsTmp As ADODB.Recordset
   Dim strCU01 As String
   Dim strCU02 As String
   Dim strTM24 As String
   Dim strTM25 As String
   Dim strTM26 As String
   
   If IsEmptyText(textTM23) Then
      textTM24 = Empty
      textTM25 = Empty
      textTM26 = Empty
      Exit Sub
   End If
   
   If Len(textTM23) > 8 Then
      strCU01 = Mid(textTM23, 1, 8)
      strCU02 = Mid(textTM23, 9, 1)
   Else
      strCU01 = textTM23 & String(8 - Len(textTM23), "0")
      strCU02 = "0"
   End If
   Set rsTmp = New ADODB.Recordset
   strSql = "SELECT * FROM CUSTOMER " & _
            "WHERE CU01 = '" & strCU01 & "' AND " & _
                  "CU02 = '" & strCU02 & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If Not IsNull(rsTmp.Fields("CU23")) Then
         If Not IsEmptyText(rsTmp.Fields("CU23")) Then
            strTM24 = rsTmp.Fields("CU23")
         End If
      End If
      If Not IsNull(rsTmp.Fields("CU24")) Then
         If Not IsEmptyText(rsTmp.Fields("CU24")) Then
            If Not IsEmptyText(strTM25) Then strTM25 = strTM25 & " "
            strTM25 = strTM25 & rsTmp.Fields("CU24")
         End If
      End If
      If Not IsNull(rsTmp.Fields("CU25")) Then
         If Not IsEmptyText(rsTmp.Fields("CU25")) Then
            If Not IsEmptyText(strTM25) Then strTM25 = strTM25 & " "
            strTM25 = strTM25 & rsTmp.Fields("CU25")
         End If
      End If
      If Not IsNull(rsTmp.Fields("CU26")) Then
         If Not IsEmptyText(rsTmp.Fields("CU26")) Then
            If Not IsEmptyText(strTM25) Then strTM25 = strTM25 & " "
            strTM25 = strTM25 & rsTmp.Fields("CU26")
         End If
      End If
      If Not IsNull(rsTmp.Fields("CU27")) Then
         If Not IsEmptyText(rsTmp.Fields("CU27")) Then
            If Not IsEmptyText(strTM25) Then strTM25 = strTM25 & " "
            strTM25 = strTM25 & rsTmp.Fields("CU27")
         End If
      End If
      If Not IsNull(rsTmp.Fields("CU28")) Then
         If Not IsEmptyText(rsTmp.Fields("CU28")) Then
            If Not IsEmptyText(strTM25) Then strTM25 = strTM25 & " "
            strTM25 = strTM25 & rsTmp.Fields("CU28")
         End If
      End If
      If Not IsNull(rsTmp.Fields("CU29")) Then
         If Not IsEmptyText(rsTmp.Fields("CU29")) Then
            'Modify By Cheng 2004/02/03
            '日文地址變數應為strTM26
'            strTM25 = rsTmp.Fields("CU29")
            strTM26 = rsTmp.Fields("CU29")
            'End
         End If
      End If
   End If
   rsTmp.Close
   Set rsTmp = Nothing
   
   textTM24 = strTM24
   textTM25 = strTM25
   textTM26 = strTM26
   
End Sub

'add by nickc 2006/12/14
Private Sub UpdateCustomerAddress2()
   Dim strSql As String
   Dim rsTmp As ADODB.Recordset
   Dim strCU01 As String
   Dim strCU02 As String
   Dim strTM82 As String
   Dim strTM86 As String
   Dim strTM90 As String
   
   If IsEmptyText(textSP58) Then
      textTM82 = Empty
      textTM86 = Empty
      textTM90 = Empty
      Exit Sub
   End If
   
   If Len(textSP58) > 8 Then
      strCU01 = Mid(textSP58, 1, 8)
      strCU02 = Mid(textSP58, 9, 1)
   Else
      strCU01 = textSP58 & String(8 - Len(textSP58), "0")
      strCU02 = "0"
   End If
   Set rsTmp = New ADODB.Recordset
   strSql = "SELECT * FROM CUSTOMER " & _
            "WHERE CU01 = '" & strCU01 & "' AND " & _
                  "CU02 = '" & strCU02 & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If Not IsNull(rsTmp.Fields("CU23")) Then
         If Not IsEmptyText(rsTmp.Fields("CU23")) Then
            strTM82 = rsTmp.Fields("CU23")
         End If
      End If
      If Not IsNull(rsTmp.Fields("CU24")) Then
         If Not IsEmptyText(rsTmp.Fields("CU24")) Then
            If Not IsEmptyText(strTM86) Then strTM86 = strTM86 & " "
            strTM86 = strTM86 & rsTmp.Fields("CU24")
         End If
      End If
      If Not IsNull(rsTmp.Fields("CU25")) Then
         If Not IsEmptyText(rsTmp.Fields("CU25")) Then
            If Not IsEmptyText(strTM86) Then strTM86 = strTM86 & " "
            strTM86 = strTM86 & rsTmp.Fields("CU25")
         End If
      End If
      If Not IsNull(rsTmp.Fields("CU26")) Then
         If Not IsEmptyText(rsTmp.Fields("CU26")) Then
            If Not IsEmptyText(strTM86) Then strTM86 = strTM86 & " "
            strTM86 = strTM86 & rsTmp.Fields("CU26")
         End If
      End If
      If Not IsNull(rsTmp.Fields("CU27")) Then
         If Not IsEmptyText(rsTmp.Fields("CU27")) Then
            If Not IsEmptyText(strTM86) Then strTM86 = strTM86 & " "
            strTM86 = strTM86 & rsTmp.Fields("CU27")
         End If
      End If
      If Not IsNull(rsTmp.Fields("CU28")) Then
         If Not IsEmptyText(rsTmp.Fields("CU28")) Then
            If Not IsEmptyText(strTM86) Then strTM86 = strTM86 & " "
            strTM86 = strTM86 & rsTmp.Fields("CU28")
         End If
      End If
      If Not IsNull(rsTmp.Fields("CU29")) Then
         If Not IsEmptyText(rsTmp.Fields("CU29")) Then
            'Modify By Cheng 2004/02/03
            '日文地址變數應為strTM26
'            strTM25 = rsTmp.Fields("CU29")
            strTM90 = rsTmp.Fields("CU29")
            'End
         End If
      End If
   End If
   rsTmp.Close
   Set rsTmp = Nothing
   
   textTM82 = strTM82
   textTM86 = strTM86
   textTM90 = strTM90
   
End Sub
Private Sub UpdateCustomerAddress3()
   Dim strSql As String
   Dim rsTmp As ADODB.Recordset
   Dim strCU01 As String
   Dim strCU02 As String
   Dim strTM83 As String
   Dim strTM87 As String
   Dim strTM91 As String
   
   If IsEmptyText(textSP59) Then
      textTM83 = Empty
      textTM87 = Empty
      textTM91 = Empty
      Exit Sub
   End If
   
   If Len(textSP59) > 8 Then
      strCU01 = Mid(textSP59, 1, 8)
      strCU02 = Mid(textSP59, 9, 1)
   Else
      strCU01 = textSP59 & String(8 - Len(textSP59), "0")
      strCU02 = "0"
   End If
   Set rsTmp = New ADODB.Recordset
   strSql = "SELECT * FROM CUSTOMER " & _
            "WHERE CU01 = '" & strCU01 & "' AND " & _
                  "CU02 = '" & strCU02 & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If Not IsNull(rsTmp.Fields("CU23")) Then
         If Not IsEmptyText(rsTmp.Fields("CU23")) Then
            strTM83 = rsTmp.Fields("CU23")
         End If
      End If
      If Not IsNull(rsTmp.Fields("CU24")) Then
         If Not IsEmptyText(rsTmp.Fields("CU24")) Then
            If Not IsEmptyText(strTM87) Then strTM87 = strTM87 & " "
            strTM87 = strTM87 & rsTmp.Fields("CU24")
         End If
      End If
      If Not IsNull(rsTmp.Fields("CU25")) Then
         If Not IsEmptyText(rsTmp.Fields("CU25")) Then
            If Not IsEmptyText(strTM87) Then strTM87 = strTM87 & " "
            strTM87 = strTM87 & rsTmp.Fields("CU25")
         End If
      End If
      If Not IsNull(rsTmp.Fields("CU26")) Then
         If Not IsEmptyText(rsTmp.Fields("CU26")) Then
            If Not IsEmptyText(strTM87) Then strTM87 = strTM87 & " "
            strTM87 = strTM87 & rsTmp.Fields("CU26")
         End If
      End If
      If Not IsNull(rsTmp.Fields("CU27")) Then
         If Not IsEmptyText(rsTmp.Fields("CU27")) Then
            If Not IsEmptyText(strTM87) Then strTM87 = strTM87 & " "
            strTM87 = strTM87 & rsTmp.Fields("CU27")
         End If
      End If
      If Not IsNull(rsTmp.Fields("CU28")) Then
         If Not IsEmptyText(rsTmp.Fields("CU28")) Then
            If Not IsEmptyText(strTM87) Then strTM87 = strTM87 & " "
            strTM87 = strTM87 & rsTmp.Fields("CU28")
         End If
      End If
      If Not IsNull(rsTmp.Fields("CU29")) Then
         If Not IsEmptyText(rsTmp.Fields("CU29")) Then
            'Modify By Cheng 2004/02/03
            '日文地址變數應為strTM26
'            strTM25 = rsTmp.Fields("CU29")
            strTM91 = rsTmp.Fields("CU29")
            'End
         End If
      End If
   End If
   rsTmp.Close
   Set rsTmp = Nothing
   
   textTM83 = strTM83
   textTM87 = strTM87
   textTM91 = strTM91
   
End Sub
Private Sub UpdateCustomerAddress4()
   Dim strSql As String
   Dim rsTmp As ADODB.Recordset
   Dim strCU01 As String
   Dim strCU02 As String
   Dim strTM84 As String
   Dim strTM88 As String
   Dim strTM92 As String
   
   If IsEmptyText(textTM80) Then
      textTM84 = Empty
      textTM88 = Empty
      textTM92 = Empty
      Exit Sub
   End If
   
   If Len(textTM80) > 8 Then
      strCU01 = Mid(textTM80, 1, 8)
      strCU02 = Mid(textTM80, 9, 1)
   Else
      strCU01 = textTM80 & String(8 - Len(textTM80), "0")
      strCU02 = "0"
   End If
   Set rsTmp = New ADODB.Recordset
   strSql = "SELECT * FROM CUSTOMER " & _
            "WHERE CU01 = '" & strCU01 & "' AND " & _
                  "CU02 = '" & strCU02 & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If Not IsNull(rsTmp.Fields("CU23")) Then
         If Not IsEmptyText(rsTmp.Fields("CU23")) Then
            strTM84 = rsTmp.Fields("CU23")
         End If
      End If
      If Not IsNull(rsTmp.Fields("CU24")) Then
         If Not IsEmptyText(rsTmp.Fields("CU24")) Then
            If Not IsEmptyText(strTM88) Then strTM88 = strTM88 & " "
            strTM88 = strTM88 & rsTmp.Fields("CU24")
         End If
      End If
      If Not IsNull(rsTmp.Fields("CU25")) Then
         If Not IsEmptyText(rsTmp.Fields("CU25")) Then
            If Not IsEmptyText(strTM88) Then strTM88 = strTM88 & " "
            strTM88 = strTM88 & rsTmp.Fields("CU25")
         End If
      End If
      If Not IsNull(rsTmp.Fields("CU26")) Then
         If Not IsEmptyText(rsTmp.Fields("CU26")) Then
            If Not IsEmptyText(strTM88) Then strTM88 = strTM88 & " "
            strTM88 = strTM88 & rsTmp.Fields("CU26")
         End If
      End If
      If Not IsNull(rsTmp.Fields("CU27")) Then
         If Not IsEmptyText(rsTmp.Fields("CU27")) Then
            If Not IsEmptyText(strTM88) Then strTM88 = strTM88 & " "
            strTM88 = strTM88 & rsTmp.Fields("CU27")
         End If
      End If
      If Not IsNull(rsTmp.Fields("CU28")) Then
         If Not IsEmptyText(rsTmp.Fields("CU28")) Then
            If Not IsEmptyText(strTM88) Then strTM88 = strTM88 & " "
            strTM88 = strTM88 & rsTmp.Fields("CU28")
         End If
      End If
      If Not IsNull(rsTmp.Fields("CU29")) Then
         If Not IsEmptyText(rsTmp.Fields("CU29")) Then
            'Modify By Cheng 2004/02/03
            '日文地址變數應為strTM92
'            strTM88 = rsTmp.Fields("CU29")
            strTM92 = rsTmp.Fields("CU29")
            'End
         End If
      End If
   End If
   rsTmp.Close
   Set rsTmp = Nothing
   
   textTM84 = strTM84
   textTM88 = strTM88
   textTM92 = strTM92
   
End Sub
Private Sub UpdateCustomerAddress5()
   Dim strSql As String
   Dim rsTmp As ADODB.Recordset
   Dim strCU01 As String
   Dim strCU02 As String
   Dim strTM85 As String
   Dim strTM89 As String
   Dim strTM93 As String
   
   If IsEmptyText(textTM81) Then
      textTM85 = Empty
      textTM89 = Empty
      textTM93 = Empty
      Exit Sub
   End If
   
   If Len(textTM81) > 8 Then
      strCU01 = Mid(textTM81, 1, 8)
      strCU02 = Mid(textTM81, 9, 1)
   Else
      strCU01 = textTM81 & String(8 - Len(textTM81), "0")
      strCU02 = "0"
   End If
   Set rsTmp = New ADODB.Recordset
   strSql = "SELECT * FROM CUSTOMER " & _
            "WHERE CU01 = '" & strCU01 & "' AND " & _
                  "CU02 = '" & strCU02 & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If Not IsNull(rsTmp.Fields("CU23")) Then
         If Not IsEmptyText(rsTmp.Fields("CU23")) Then
            strTM85 = rsTmp.Fields("CU23")
         End If
      End If
      If Not IsNull(rsTmp.Fields("CU24")) Then
         If Not IsEmptyText(rsTmp.Fields("CU24")) Then
            If Not IsEmptyText(strTM89) Then strTM89 = strTM89 & " "
            strTM89 = strTM89 & rsTmp.Fields("CU24")
         End If
      End If
      If Not IsNull(rsTmp.Fields("CU25")) Then
         If Not IsEmptyText(rsTmp.Fields("CU25")) Then
            If Not IsEmptyText(strTM89) Then strTM89 = strTM89 & " "
            strTM89 = strTM89 & rsTmp.Fields("CU25")
         End If
      End If
      If Not IsNull(rsTmp.Fields("CU26")) Then
         If Not IsEmptyText(rsTmp.Fields("CU26")) Then
            If Not IsEmptyText(strTM89) Then strTM89 = strTM89 & " "
            strTM89 = strTM89 & rsTmp.Fields("CU26")
         End If
      End If
      If Not IsNull(rsTmp.Fields("CU27")) Then
         If Not IsEmptyText(rsTmp.Fields("CU27")) Then
            If Not IsEmptyText(strTM89) Then strTM89 = strTM89 & " "
            strTM89 = strTM89 & rsTmp.Fields("CU27")
         End If
      End If
      If Not IsNull(rsTmp.Fields("CU28")) Then
         If Not IsEmptyText(rsTmp.Fields("CU28")) Then
            If Not IsEmptyText(strTM89) Then strTM89 = strTM89 & " "
            strTM89 = strTM89 & rsTmp.Fields("CU28")
         End If
      End If
      If Not IsNull(rsTmp.Fields("CU29")) Then
         If Not IsEmptyText(rsTmp.Fields("CU29")) Then
            'Modify By Cheng 2004/02/03
            '日文地址變數應為strTM93
'            strTM89 = rsTmp.Fields("CU29")
            strTM93 = rsTmp.Fields("CU29")
            'End
         End If
      End If
   End If
   rsTmp.Close
   Set rsTmp = Nothing
   
   textTM85 = strTM85
   textTM89 = strTM89
   textTM93 = strTM93
End Sub

'Add By Sindy 2019/4/9
Private Sub textTM72_GotFocus()
    TextInverse Me.textTM72
End Sub

Private Sub textTM72_Validate(Cancel As Boolean)
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
    
    If Me.textTM72.Text <> "" Then
        Me.textTM72_2.Text = PUB_GetSpecialPTName("2", Me.textTM72.Text)
        If Me.textTM72_2.Text = "" Then
           MsgBox "特殊商標代碼輸入錯誤!!!", vbExclamation + vbOKOnly
           Cancel = True
        End If
    Else
        Me.textTM72.Text = "" 'Added by Lydia 2023/11/16
        Me.textTM72_2.Text = ""
    End If
    If Cancel = True Then textTM72_GotFocus
End Sub
'2019/4/9 END

'Added by Lydia 2023/11/14
Private Sub textTM72_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textTM80_GotFocus()
InverseTextBox textTM80
End Sub
Private Sub textTM80_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub
Private Sub textTM80_Validate(Cancel As Boolean)
Dim strTit As String
Dim strMsg As String
Dim nResponse
   
    Cancel = False
    textTM80_2 = Empty
    If IsEmptyText(textTM80) = False Then
        Me.textTM80.Text = ChangeCustomerL(Me.textTM80.Text)
        textTM80_2 = GetCustomerName(textTM80, 0)
        If textTM80_2 = Empty Then
            Cancel = True
            strTit = "檢核資料"
            strMsg = "申請人代碼<" & textTM80 & ">不正確"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
        Else
            'Modified by Lydia 2024/06/13
            'If Me.textTM80.Text <> m_strCust4 Then
            If ChangeCustomerL(Me.textTM80.Text) <> m_TM80 Then
                If Not PUB_EditCustOk(m_CP09, m_TM01, m_TM02, m_TM03, m_TM04) Then Cancel = True
            End If
            If InStr(ChangeCustomerL(m_TM80), ChangeCustomerL(textTM80)) = 0 Then
               UpdateCustomerAddress4
            End If
        End If
    End If
    If Cancel = True Then textTM80_GotFocus
End Sub

Private Sub textTM81_GotFocus()
InverseTextBox textTM81
End Sub
Private Sub textTM81_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub
Private Sub textTM81_Validate(Cancel As Boolean)
Dim strTit As String
Dim strMsg As String
Dim nResponse
   
    Cancel = False
    textTM81_2 = Empty
    If IsEmptyText(textTM81) = False Then
        Me.textTM81.Text = ChangeCustomerL(Me.textTM81.Text)
        textTM81_2 = GetCustomerName(textTM81, 0)
        If textTM81_2 = Empty Then
            Cancel = True
            strTit = "檢核資料"
            strMsg = "申請人代碼<" & textTM81 & ">不正確"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
        Else
            'Modified by Lydia 2024/06/13
            'If Me.textTM81.Text <> m_strCust5 Then
            If ChangeCustomerL(Me.textTM81.Text) <> m_TM81 Then
                If Not PUB_EditCustOk(m_CP09, m_TM01, m_TM02, m_TM03, m_TM04) Then Cancel = True
            End If
            If InStr(ChangeCustomerL(m_TM81), ChangeCustomerL(textTM81)) = 0 Then
               UpdateCustomerAddress5
            End If
        End If
    End If
    If Cancel = True Then textTM81_GotFocus
End Sub
'add by nickc 2006/12/18
Private Sub textTM82_GotFocus()
InverseTextBox textTM82
End Sub
Private Sub textTM82_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textTM82) = False Then
      If CheckLengthIsOK(textTM82, textTM82.MaxLength) = False Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "申請地址(中)內容太長"
      End If
   End If
   If Cancel Then TextInverse textTM82
End Sub
Private Sub textTM83_GotFocus()
InverseTextBox textTM83
End Sub
Private Sub textTM83_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textTM83) = False Then
      If CheckLengthIsOK(textTM83, textTM83.MaxLength) = False Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "申請地址(中)內容太長"
      End If
   End If
   If Cancel Then TextInverse textTM83
End Sub
Private Sub textTM84_GotFocus()
InverseTextBox textTM84
End Sub
Private Sub textTM84_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textTM84) = False Then
      If CheckLengthIsOK(textTM84, textTM84.MaxLength) = False Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "申請地址(中)內容太長"
      End If
   End If
   If Cancel Then TextInverse textTM84
End Sub
Private Sub textTM85_GotFocus()
InverseTextBox textTM85
End Sub
Private Sub textTM85_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textTM85) = False Then
      If CheckLengthIsOK(textTM85, textTM85.MaxLength) = False Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "申請地址(中)內容太長"
      End If
   End If
   If Cancel Then TextInverse textTM85
End Sub
Private Sub textTM86_GotFocus()
InverseTextBox textTM86
End Sub
Private Sub textTM86_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textTM86) = False Then
      If CheckLengthIsOK(textTM86, textTM86.MaxLength) = False Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "申請地址(英)內容太長"
         textTM86_GotFocus
      End If
   End If
End Sub
Private Sub textTM87_GotFocus()
InverseTextBox textTM87
End Sub
Private Sub textTM87_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textTM87) = False Then
      If CheckLengthIsOK(textTM87, textTM87.MaxLength) = False Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "申請地址(英)內容太長"
         textTM87_GotFocus
      End If
   End If
End Sub
Private Sub textTM88_GotFocus()
InverseTextBox textTM88
End Sub
Private Sub textTM88_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textTM88) = False Then
      If CheckLengthIsOK(textTM88, textTM88.MaxLength) = False Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "申請地址(英)內容太長"
         textTM88_GotFocus
      End If
   End If
End Sub
Private Sub textTM89_GotFocus()
InverseTextBox textTM89
End Sub
Private Sub textTM89_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textTM89) = False Then
      If CheckLengthIsOK(textTM89, textTM89.MaxLength) = False Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "申請地址(英)內容太長"
         textTM89_GotFocus
      End If
   End If
End Sub
Private Sub textTM90_GotFocus()
InverseTextBox textTM90
End Sub
Private Sub textTM90_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textTM90) = False Then
     If CheckLengthIsOK(textTM90, textTM90.MaxLength) = False Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "申請地址(日)內容太長"
         textTM90_GotFocus
      End If
   End If
End Sub
Private Sub textTM91_GotFocus()
InverseTextBox textTM91
End Sub
Private Sub textTM91_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textTM91) = False Then
     If CheckLengthIsOK(textTM91, textTM91.MaxLength) = False Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "申請地址(日)內容太長"
         textTM91_GotFocus
      End If
   End If
End Sub
Private Sub textTM92_GotFocus()
InverseTextBox textTM92
End Sub
Private Sub textTM92_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textTM92) = False Then
     If CheckLengthIsOK(textTM92, textTM92.MaxLength) = False Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "申請地址(日)內容太長"
         textTM92_GotFocus
      End If
   End If
End Sub
Private Sub textTM93_GotFocus()
InverseTextBox textTM93
End Sub
Private Sub textTM93_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textTM93) = False Then
     If CheckLengthIsOK(textTM93, textTM93.MaxLength) = False Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "申請地址(日)內容太長"
         textTM93_GotFocus
      End If
   End If
End Sub

Private Sub txtDivCaseNo_Change(Index As Integer)
    Select Case Index
    Case 0 '分割案母案系統類別
        If Me.txtDivCaseNo(0).Text = "TF" Then
            Me.txtDivCaseNo(2).Visible = True
            Me.txtDivCaseNo(2).Enabled = True
            Me.txtDivCaseNo(1).MaxLength = 5
            Me.txtDivCaseNo(2).Text = ""
        Else
            Me.txtDivCaseNo(2).Visible = False
            Me.txtDivCaseNo(2).Enabled = False
            Me.txtDivCaseNo(1).MaxLength = 6
            Me.txtDivCaseNo(2).Text = ""
        End If
    End Select
End Sub

Private Sub txtDivCaseNo_GotFocus(Index As Integer)
    TextInverse Me.txtDivCaseNo(Index)
End Sub

Private Sub txtDivCaseNo_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyAscii = UpperCase(KeyAscii)
End Sub

'Add By Cheng 2004/04/14
'顯示分割案的母案案號
Private Sub ShowOriginCaseNo(strDC01 As String, strDC02 As String, strDC03 As String, strDC04 As String)
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset

StrSQLa = "Select * From DivisionCase Where DC01='" & strDC01 & "' And DC02='" & strDC02 & "' And DC03='" & strDC03 & "' And DC04='" & strDC04 & "'"
rsA.CursorLocation = adUseClient
rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
If rsA.RecordCount > 0 Then
   Me.txtDivCaseNo(0).Text = "" & rsA("DC05").Value
   Me.txtDivCaseNo(1).Text = IIf("" & rsA("DC05").Value = "TF", Mid("" & rsA("DC06").Value, 1, 5), "" & rsA("DC06").Value)
   Me.txtDivCaseNo(2).Text = IIf("" & rsA("DC05").Value = "TF", Mid("" & rsA("DC06").Value, 6, 1), "")
   Me.txtDivCaseNo(3).Text = "" & rsA("DC07").Value
   Me.txtDivCaseNo(4).Text = "" & rsA("DC08").Value
Else
   'Add By Sindy 2023/11/13 於分案時自動帶入母案案號,前因接洽記錄單未電子化,須透過櫃台收文才產生子案本所案號,
   '                        今接洽記錄單已電子化,希望系統自動產生子案案號時同時帶入母案案號,避免人員輸入錯誤
   StrSQLa = "select CRL01,CRL07,CRL08,CRL09,CRL10 from consultrecordlist where CRL01='" & txtF0301 & "'"
   If rsA.State <> adStateClosed Then rsA.Close
   rsA.CursorLocation = adUseClient
   rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
   If rsA.RecordCount > 0 Then
      'Modify By Sindy 2023/11/27 只有子案(T-246482、T-246483)才能帶母案案號，母案不能帶自己，例：T-246481。
      If m_TM02 <> "" & rsA("CRL08").Value Then
      '2023/11/27 END
         Me.txtDivCaseNo(0).Text = "" & rsA("CRL07").Value
         Me.txtDivCaseNo(1).Text = IIf("" & rsA("CRL07").Value = "TF", Mid("" & rsA("CRL08").Value, 1, 5), "" & rsA("CRL08").Value)
         Me.txtDivCaseNo(2).Text = IIf("" & rsA("CRL07").Value = "TF", Mid("" & rsA("CRL08").Value, 6, 1), "")
         Me.txtDivCaseNo(3).Text = "" & rsA("CRL09").Value
         Me.txtDivCaseNo(4).Text = "" & rsA("CRL10").Value
      Else
         Me.txtDivCaseNo(0).Text = ""
         Me.txtDivCaseNo(1).Text = ""
         Me.txtDivCaseNo(2).Text = ""
         Me.txtDivCaseNo(3).Text = ""
         Me.txtDivCaseNo(4).Text = ""
      End If
   Else
   '2023/11/13 END
      Me.txtDivCaseNo(0).Text = ""
      Me.txtDivCaseNo(1).Text = ""
      Me.txtDivCaseNo(2).Text = ""
      Me.txtDivCaseNo(3).Text = ""
      Me.txtDivCaseNo(4).Text = ""
   End If
End If
If rsA.State <> adStateClosed Then rsA.Close
Set rsA = Nothing
End Sub

Private Sub txtDivCaseNo_LostFocus(Index As Integer)
    Select Case Index
    Case 2
        If Me.txtDivCaseNo(0).Text <> "" And Me.txtDivCaseNo(1).Text <> "" And Me.txtDivCaseNo(Index).Text = "" Then Me.txtDivCaseNo(Index).Text = "0"
    Case 3
        If Me.txtDivCaseNo(0).Text <> "" And Me.txtDivCaseNo(1).Text <> "" And Me.txtDivCaseNo(Index).Text = "" Then Me.txtDivCaseNo(Index).Text = "0"
    Case 4
        If Me.txtDivCaseNo(0).Text <> "" And Me.txtDivCaseNo(1).Text <> "" And Me.txtDivCaseNo(Index).Text = "" Then Me.txtDivCaseNo(Index).Text = "00"
    End Select
End Sub

'Add By Cheng 2004/04/14
'檢查案號是否存在
Private Function ChkCaseExist(strCN01 As String, strCN02 As String, strCN03 As String, strCN04 As String) As Boolean
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset

StrSQLa = "Select PA01 From Patent Where PA01='" & strCN01 & "' And PA02='" & strCN02 & "' And PA03='" & strCN03 & "' And PA04='" & strCN04 & "' "
StrSQLa = StrSQLa & " Union Select TM01 From Trademark Where TM01='" & strCN01 & "' And TM02='" & strCN02 & "' And TM03='" & strCN03 & "' And TM04='" & strCN04 & "' "
StrSQLa = StrSQLa & " Union Select LC01 From Lawcase Where LC01='" & strCN01 & "' And LC02='" & strCN02 & "' And LC03='" & strCN03 & "' And LC04='" & strCN04 & "' "
StrSQLa = StrSQLa & " Union Select HC01 From Hirecase Where HC01='" & strCN01 & "' And HC02='" & strCN02 & "' And HC03='" & strCN03 & "' And HC04='" & strCN04 & "' "
StrSQLa = StrSQLa & " Union Select SP01 From Servicepractice Where SP01='" & strCN01 & "' And SP02='" & strCN02 & "' And SP03='" & strCN03 & "' And SP04='" & strCN04 & "' "
rsA.CursorLocation = adUseClient
rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
If rsA.RecordCount > 0 Then
    ChkCaseExist = True
Else
    ChkCaseExist = False
End If
If rsA.State <> adStateClosed Then rsA.Close
Set rsA = Nothing
End Function

Private Function ChkCaseReleate(ByRef strCode() As String) As Boolean
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
Dim strNation As String
Dim strNation_1 As String
Dim strCust(0 To 4) As String
Dim strCust_1(0 To 4) As String
Dim strPA08 As String
Dim strPA08_1 As String
Dim ii As Integer
Dim jj As Integer

    ChkCaseReleate = True
    StrSQLa = "Select PA09, PA26, PA27, PA28, PA29, PA30, PA08 From Patent Where PA01='" & strCode(0) & "' And PA02='" & strCode(1) & "' And PA03='" & strCode(2) & "' And PA04='" & strCode(3) & "' "
'edit by nickc 2006/12/15
'    StrSQLa = StrSQLa & " Union Select TM10, TM23, '', '', '', '', '' From Trademark Where TM01='" & strCode(0) & "' And TM02='" & strCode(1) & "' And TM03='" & strCode(2) & "' And TM04='" & strCode(3) & "' "
    StrSQLa = StrSQLa & " Union Select TM10, TM23,tm78,tm79,tm80,tm81, '' From Trademark Where TM01='" & strCode(0) & "' And TM02='" & strCode(1) & "' And TM03='" & strCode(2) & "' And TM04='" & strCode(3) & "' "
    StrSQLa = StrSQLa & " Union Select LC15, LC11,'', '', '', '', '' From Lawcase Where LC01='" & strCode(0) & "' And LC02='" & strCode(1) & "' And LC03='" & strCode(2) & "' And LC04='" & strCode(3) & "' "
    StrSQLa = StrSQLa & " Union Select '000', HC05, '', '', '', '', '' From Hirecase Where HC01='" & strCode(0) & "' And HC02='" & strCode(1) & "' And HC03='" & strCode(2) & "' And HC04='" & strCode(3) & "' "
'edit by nickc 2006/12/15
'    StrSQLa = StrSQLa & " Union Select SP09, SP08, '', '', '', '', '' From Servicepractice Where SP01='" & strCode(0) & "' And SP02='" & strCode(1) & "' And SP03='" & strCode(2) & "' And SP04='" & strCode(3) & "' "
    StrSQLa = StrSQLa & " Union Select SP09, SP08, sp58, sp59, sp65, sp66, '' From Servicepractice Where SP01='" & strCode(0) & "' And SP02='" & strCode(1) & "' And SP03='" & strCode(2) & "' And SP04='" & strCode(3) & "' "
    rsA.CursorLocation = adUseClient
    rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
    If rsA.RecordCount > 0 Then
        strNation = "" & rsA.Fields(0).Value
        For ii = 0 To 4
            strCust(ii) = "" & rsA.Fields(ii + 1).Value
        Next ii
        strPA08 = "" & rsA.Fields(6).Value
    Else
        MsgBox "查無此分割案號資料!!!", vbExclamation + vbOKOnly
        ChkCaseReleate = False
        GoTo ExitFunction
    End If
    If rsA.State <> adStateClosed Then rsA.Close
    Set rsA = Nothing
        
    StrSQLa = "Select PA09, PA26, PA27, PA28, PA29, PA30, PA08 From Patent Where PA01='" & strCode(4) & "' And PA02='" & strCode(5) & "' And PA03='" & strCode(6) & "' And PA04='" & strCode(7) & "' "
'edit by nickc 2006/12/15
'    StrSQLa = StrSQLa & " Union Select TM10, TM23, '', '', '', '', '' From Trademark Where TM01='" & strCode(4) & "' And TM02='" & strCode(5) & "' And TM03='" & strCode(6) & "' And TM04='" & strCode(7) & "' "
    StrSQLa = StrSQLa & " Union Select TM10, TM23, tm78, tm79, tm80, tm81, '' From Trademark Where TM01='" & strCode(4) & "' And TM02='" & strCode(5) & "' And TM03='" & strCode(6) & "' And TM04='" & strCode(7) & "' "
    StrSQLa = StrSQLa & " Union Select LC15, LC11, '', '', '', '', '' From Lawcase Where LC01='" & strCode(4) & "' And LC02='" & strCode(5) & "' And LC03='" & strCode(6) & "' And LC04='" & strCode(7) & "' "
    StrSQLa = StrSQLa & " Union Select '000', HC05, '', '', '', '', '' From Hirecase Where HC01='" & strCode(4) & "' And HC02='" & strCode(5) & "' And HC03='" & strCode(6) & "' And HC04='" & strCode(7) & "' "
'edit by nickc 2006/12/15
'    StrSQLa = StrSQLa & " Union Select SP09, SP08, '', '', '' ,'', '' From Servicepractice Where SP01='" & strCode(4) & "' And SP02='" & strCode(5) & "' And SP03='" & strCode(6) & "' And SP04='" & strCode(7) & "' "
    StrSQLa = StrSQLa & " Union Select SP09, SP08, sp58, sp59, sp65 ,sp66, '' From Servicepractice Where SP01='" & strCode(4) & "' And SP02='" & strCode(5) & "' And SP03='" & strCode(6) & "' And SP04='" & strCode(7) & "' "
    rsA.CursorLocation = adUseClient
    rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
    If rsA.RecordCount > 0 Then
        strNation_1 = "" & rsA.Fields(0).Value
        For ii = 0 To 4
            strCust_1(ii) = "" & rsA.Fields(ii + 1).Value
        Next ii
        strPA08_1 = "" & rsA.Fields(6).Value
    Else
        MsgBox "查無此母案案號資料!!!", vbExclamation + vbOKOnly
        ChkCaseReleate = False
        GoTo ExitFunction
    End If
    If strNation <> strNation_1 Then
        MsgBox "您輸入的分割案及母案申請國家不同!!!", vbExclamation + vbOKOnly
        ChkCaseReleate = False
        GoTo ExitFunction
    End If
    ChkCaseReleate = False
    For ii = 0 To 4
        For jj = 0 To 4
            If strCust(ii) <> "" And strCust_1(jj) <> "" Then
                If Left(strCust(ii), 6) = Left(strCust_1(jj), 6) Then
                    ChkCaseReleate = True
                End If
            End If
            If ChkCaseReleate = True Then Exit For
        Next jj
        If ChkCaseReleate = True Then Exit For
    Next ii
    If ChkCaseReleate = False Then
        MsgBox "您輸入的分割案及母案申請人非關係企業!!!", vbExclamation + vbOKOnly
        GoTo ExitFunction
    End If
    If InStr(strCode(0), "P") > 0 And InStr(strCode(4), "P") > 0 And strPA08 <> strPA08_1 Then
        MsgBox "您輸入的分割案及母案專利種類不同!!!", vbExclamation + vbOKOnly
        ChkCaseReleate = False
        GoTo ExitFunction
    End If
            
ExitFunction:
    If rsA.State <> adStateClosed Then rsA.Close
    Set rsA = Nothing
End Function

'Add By Sindy 2011/7/8
'若有舊案申請地址與接洽紀錄單上新址提申不同者
Private Sub CheckAppAddr(Index As Integer)
Dim strApplID As String, strName As String, strAddr As String
   
   If Trim(textTM10) = "" Then Exit Sub
   
   '新案時,系統別為T或TF,並且非台灣案
   If m_CP31 = "Y" And (m_TM01 = "T" Or m_TM01 = "TF") And Trim(textTM10) <> "000" Then
      strApplID = "": strName = "": strAddr = ""
      Select Case Index
         Case 1
            If Trim(textTM23) <> "" Then '申請人1
               strApplID = Trim(textTM23)
               strName = Trim(textTM23_2)
               strAddr = Trim(textTM24)
            End If
         Case 2
            If Trim(textSP58) <> "" Then '申請人2
               strApplID = Trim(textSP58)
               strName = Trim(textSP58_2)
               strAddr = Trim(textTM82)
            End If
         Case 3
            If Trim(textSP59) <> "" Then '申請人3
               strApplID = Trim(textSP59)
               strName = Trim(textSP59_2)
               strAddr = Trim(textTM83)
            End If
         Case 4
            If Trim(textTM80) <> "" Then '申請人4
               strApplID = Trim(textTM80)
               strName = Trim(textTM80_2)
               strAddr = Trim(textTM84)
            End If
         Case 5
            If Trim(textTM81) <> "" Then '申請人5
               strApplID = Trim(textTM81)
               strName = Trim(textTM81_2)
               strAddr = Trim(textTM85)
            End If
      End Select
      If strApplID <> "" Then
         If ChkOCaseAndCAddrNotAlikeChoose(strApplID, Trim(textTM10), m_TM01, "", rsAddrNotAlike, False) = True Then
            '***** 抓出目前客戶的申請地址
            strSql = "SELECT cu23,cu112 FROM customer WHERE CU01='" & Left(strApplID, 8) & "' AND CU02='" & Mid(strApplID, 9, 1) & "'"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strSql)
            If intI = 1 Then
               If (strAddr = "" & RsTemp("cu112") & "" & RsTemp("cu23")) Or _
                  (strAddr = "" & RsTemp("cu23")) Then
                  strAddr = "" & RsTemp("cu23")
            '*****
                  Me.m_AppAddr = ""
                  Me.m_Zipcode = ""
                  Set frm090801_4.UpForm = Me
                  frm090801_4.Label2 = strApplID
                  frm090801_4.lblName = strName
                  frm090801_4.strAddr = strAddr
                  If Index = 1 Then frm090801_4.Label1(1) = "申請人1："
                  If Index = 2 Then frm090801_4.Label1(1) = "申請人2："
                  If Index = 3 Then frm090801_4.Label1(1) = "申請人3："
                  If Index = 4 Then frm090801_4.Label1(1) = "申請人4："
                  If Index = 5 Then frm090801_4.Label1(1) = "申請人5："
                  If frm090801_4.QueryData = True Then
                     Me.Hide
                     frm090801_4.Show vbModal
                     If Me.m_AppAddrChange = True Then
                        Me.m_AppAddr = Me.m_Zipcode & Me.m_AppAddr
                        Select Case Index
                           Case 1
                              Me.textTM24 = Me.m_AppAddr
                              Me.textTM25 = ""
                              Me.textTM26 = ""
                              SetTMSPFieldNewData "TM24", Me.m_AppAddr
                              SetTMSPFieldNewData "TM25", ""
                              SetTMSPFieldNewData "TM26", ""
                           Case 2
                              Me.textTM82 = Me.m_AppAddr
                              Me.textTM86 = ""
                              Me.textTM90 = ""
                              SetTMSPFieldNewData "TM82", Me.m_AppAddr
                              SetTMSPFieldNewData "TM86", ""
                              SetTMSPFieldNewData "TM90", ""
                           Case 3
                              Me.textTM83 = Me.m_AppAddr
                              Me.textTM87 = ""
                              Me.textTM91 = ""
                              SetTMSPFieldNewData "TM83", Me.m_AppAddr
                              SetTMSPFieldNewData "TM87", ""
                              SetTMSPFieldNewData "TM91", ""
                           Case 4
                              Me.textTM84 = Me.m_AppAddr
                              Me.textTM88 = ""
                              Me.textTM92 = ""
                              SetTMSPFieldNewData "TM84", Me.m_AppAddr
                              SetTMSPFieldNewData "TM88", ""
                              SetTMSPFieldNewData "TM92", ""
                           Case 5
                              Me.textTM85 = Me.m_AppAddr
                              Me.textTM89 = ""
                              Me.textTM93 = ""
                              SetTMSPFieldNewData "TM85", Me.m_AppAddr
                              SetTMSPFieldNewData "TM89", ""
                              SetTMSPFieldNewData "TM93", ""
                        End Select
                     End If
                  End If
               End If
            End If
         End If
      End If
   End If
End Sub

'Add By Sindy 2012/5/8
Private Sub SaveFrame21(strCP09 As String)
   If Frame21.Visible = True Then
      '資料是否齊備
      'Modify By Sindy 2012/11/19 收文已上齊備, 分案時不可更新齊備日
      'If Trim(textEP06) = "Y" Then
      If textEP06.Visible = True Then  'Added by Lydia 2018/12/10 +判斷顯示
            If (m_EP06 = "" Or m_EP06 = "N") And Trim(textEP06) = "Y" Then
            '2012/11/19 End
               strSql = "update engineerprogress set ep06=" & strSrvDate(1) & ",ep36=" & strSrvDate(1) & " where ep02='" & strCP09 & "'"
               cnnConnection.Execute strSql
               m_EP06DT = strSrvDate(1) 'Added by Lydia 2019/04/11
            ElseIf Trim(textEP06) = "N" Then
               strSql = "update engineerprogress set ep06=0,ep36=0 where ep02='" & strCP09 & "'"
               cnnConnection.Execute strSql
               m_EP06DT = "0" 'Added by Lydia 2019/04/11
            ElseIf Trim(textEP06) = "" Then
               strSql = "update engineerprogress set ep06=null,ep36=null where ep02='" & strCP09 & "'"
               cnnConnection.Execute strSql
               m_EP06DT = "" 'Added by Lydia 2019/04/11
            End If
            '資料齊備
            If textCP143.Visible = False Then  'Added by Lydia 2019/01/30 +判斷顯示
                If (m_EP06 = "" Or m_EP06 = "N") And Trim(textEP06) = "Y" Then
                   m_EP06DT = strSrvDate(1)
                   strSql = "insert into tmctldate(tcd01,tcd02,tcd03,tcd04,tcd05,tcd06,tcd07)" & _
                            " values('" & strCP09 & "','1','" & strUserNum & "'," & strSrvDate(1) & "," & Right("000000" & ServerTime, 6) & "," & strSrvDate(1) & ",'分案')"
                   cnnConnection.Execute strSql
                '分案取消齊備
                ElseIf m_EP06 = "Y" And (Trim(textEP06) = "N" Or Trim(textEP06) = "") Then
                   m_EP06DT = ""
                   strSql = "insert into tmctldate(tcd01,tcd02,tcd03,tcd04,tcd05,tcd06,tcd07)" & _
                            " values('" & strCP09 & "','1','" & strUserNum & "'," & strSrvDate(1) & "," & Right("000000" & ServerTime, 6) & ",null,'分案取消齊備')"
                   cnnConnection.Execute strSql
                   'Add By Sindy 2012/7/13
                   strSql = "update CaseProgress SET CP48=null WHERE CP09='" & strCP09 & "'"
                   cnnConnection.Execute strSql
                   '2012/7/13 End
                End If
            End If 'end 2019/01/30
      End If
      
      '是否會稿
      If textEP34.Visible = True Then  'Added by Lydia 2018/12/10 +判斷顯示
            strSql = "update engineerprogress set ep34='" & textEP34 & "' where ep02='" & strCP09 & "'"
            cnnConnection.Execute strSql
      End If
      
      'Added by Lydia 2019/01/30 查名是否齊備
      'Modified by Lydia 2020/11/04 收文之查名齊備日 m_CP143=> p_strCP143
      If textCP143.Visible = True Then
        If (p_strCP143 = "" Or p_strCP143 = "N") And Trim(textCP143) = "Y" Then
           strSql = "update caseprogress set cp143=" & strSrvDate(1) & " where cp09='" & strCP09 & "'"
           cnnConnection.Execute strSql
           'Added by Lydia 2019/04/11 如果沒有勾選查名單,又設查名已齊備
           If Val(p_CP143DT) = 0 Then
               p_CP143DT = strSrvDate(1)
           End If
           'end 2019/04/11
        ElseIf Trim(textCP143) = "N" Or Trim(textCP143) = "" Then
           strSql = "update caseprogress set cp143=0 where cp09='" & strCP09 & "'"
           cnnConnection.Execute strSql
           p_CP143DT = "0" 'Added by Lydia 2019/04/11
        End If
        '文字+查名=>資料齊備
        If Trim(textEP06) = "Y" And Trim(textCP143) = "Y" And (m_EP06 <> Trim(textEP06) Or p_strCP143 <> Trim(textCP143)) Then
           m_EP06DT = strSrvDate(1)
           strSql = "insert into tmctldate(tcd01,tcd02,tcd03,tcd04,tcd05,tcd06,tcd07)" & _
                    " values('" & strCP09 & "','1','" & strUserNum & "'," & strSrvDate(1) & "," & Right("000000" & ServerTime, 6) & "," & strSrvDate(1) & ",'分案')"
           cnnConnection.Execute strSql
        '分案取消齊備
        ElseIf m_EP06 = "Y" And p_strCP143 = "Y" And (m_EP06 <> Trim(textEP06) Or p_strCP143 <> Trim(textCP143)) Then
           m_EP06DT = ""
           strSql = "insert into tmctldate(tcd01,tcd02,tcd03,tcd04,tcd05,tcd06,tcd07)" & _
                    " values('" & strCP09 & "','1','" & strUserNum & "'," & strSrvDate(1) & "," & Right("000000" & ServerTime, 6) & ",null,'分案取消齊備')"
           cnnConnection.Execute strSql
           strSql = "update CaseProgress SET CP48=null WHERE CP09='" & strCP09 & "'"
           cnnConnection.Execute strSql
        End If
      End If
      'end 2019/01/30
      
   End If
End Sub

'Add by Amy 2014/10/24
Private Sub txtDivCaseNo_Validate(Index As Integer, Cancel As Boolean)
    If Me.txtDivCaseNo(0).Text = "" Or Me.txtDivCaseNo(1).Text & Me.txtDivCaseNo(2).Text = "" Then
        Exit Sub
    End If
    
    If m_TM01 = "T" And Me.textTM10 = "020" And m_CP10 = "308" And m_CP31 = "Y" And textTM130 = "" Then
        '分割新案之特殊出名公司預帶母案的
        textTM130 = GetTM130_308Ma(txtDivCaseNo(0), txtDivCaseNo(1) & txtDivCaseNo(2), IIf(txtDivCaseNo(3) = "", "0", txtDivCaseNo(3)), IIf(txtDivCaseNo(4) = "", "00", txtDivCaseNo(4)))
    End If
End Sub

'取得母案特殊出名公司
Private Function GetTM130_308Ma(stTM01 As String, stTM02 As String, stTM03 As String, stTM04 As String) As String
    Dim RsQ As New ADODB.Recordset
    Dim strQuery As String
    Dim intQ As Integer
    
    GetTM130_308Ma = ""
    
    strQuery = "Select TM130 From TradeMark Where TM01='" & stTM01 & "' And TM02='" & stTM02 & "'And TM03='" & stTM03 & "'And TM04='" & stTM04 & "' "
    intQ = 1
    Set RsQ = ClsLawReadRstMsg(intQ, strQuery)
    If intQ = 1 Then
        GetTM130_308Ma = "" & RsQ.Fields("TM130")
    End If
    RsQ.Close
End Function
'end 2014/10/24

'Add by Amy 2018/08/09 加移轉申請人MCTF檢查
Private Function ChkMCTF_Tran(ByRef strMsg As String, ByVal strMCTF As String) As Boolean
    Dim Rs As New ADODB.Recordset
    Dim strQ As String, strCU(0) As String, strCU13 As String, strCU82 As String
    Dim intQ As Integer, ii As Integer
    Dim HasData As Boolean
    
    ChkMCTF_Tran = False: strMsg = ""
    strQ = "Select cp56,cp89,cp90,cp91,cp92 From CaseProgress Where cp09='" & m_CP09 & "' "
    intQ = 1
    Set Rs = ClsLawReadRstMsg(intQ, strQ)
    If intQ = 1 Then
        For ii = 0 To 4
            If IsNull(Rs.Fields(ii)) Then Exit For
            HasData = GetCusORFagentData(ChangeCustomerL(Rs.Fields(ii)), "CU13||';'||CU82", strCU())
            strCU13 = Mid(strCU(0), 1, InStr(strCU(0), ";") - 1)
            strCU82 = Replace(strCU(0), strCU13 & ";", "")
            'Modify by Amy 2019/03/15 新客戶可能不是當天分案 原:strCU82 = strSrvDate(1)
            If Val(strCU82) >= PUB_GetWorkDayAfterSysDate(CDbl(strSrvDate(1)), -3) + 19110000 Then
                '若客戶智權人員為MCTF,判斷是否與FC代理人不同組不可存檔
                If Left(strCU13, 4) = "MCTF" Then
                    If strCU13 <> strMCTF Then
                        strMsg = strMsg & "移轉人" & ii + 1 & "編號 " & Rs.Fields(ii) & " 智權人員為" & strCU13 & vbCrLf
                    End If
                '若客戶智權人員非MCTF員,判斷人員若為MCTF小組,與FC代理人不同組或非MCTF者不可存檔
                Else
                    'modify by sonia 2019/4/18 改新模組
                    'If GetMCTF0XCode(strCU13) = strMCTF Then
                    If ChkMCTF0XSales(strMCTF, strCU13) = True Then
                        strUpdCusNo = strUpdCusNo & "," & Rs.Fields(ii)
                    Else
                        strMsg = strMsg & "移轉人" & ii + 1 & "編號 " & Rs.Fields(ii) & " 智權人員為" & strCU13 & vbCrLf
                    End If
                End If
            End If
        Next ii
        '不可存檔
        If strMsg <> MsgText(601) Then
            strUpdCusNo = ""
        '可存檔
        ElseIf strUpdCusNo <> MsgText(601) Then
            strUpdCusNo = Mid(strUpdCusNo, 2)
            ChkMCTF_Tran = True
        End If
    End If
    Rs.Close
End Function

'Added by Lydia 2018/12/10
Private Sub textCP143_GotFocus()
   TextInverse textCP143
End Sub

Private Sub textCP143_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 8 And KeyAscii <> Asc("N") And KeyAscii <> Asc("Y") Then
      KeyAscii = 0
      Beep
   End If
End Sub

'Added by Lydia 2019/01/30 檢查-查名是否齊備
Private Sub textCP143_Validate(Cancel As Boolean)
Dim strTmp As String
     
   'Added by Lydia 2022/07/15 T大陸案之齊備日管控
   If m_TM01 = "T" And m_TM10 = "020" Then
       p_CP143DT = cp(143)
       'Added by Lydia 2023/06/27 T大陸查名;原本查名齊備改成要查名
       If Trim(textCP10) = "101" And Val(p_CP143DT) > 0 And textCP143.Text <> "Y" And m_CP27 = "" And textCP57 = "" _
          And Frame22.Visible = True And (textEP43.Text = "99997" Or textEP43.Text = "") Then
          If MsgBox("請問是否要預設查名人員？", vbInformation + vbYesNo + vbDefaultButton1, "T大陸查名") = vbYes Then
             ChkEP43.Value = 0
            If GetDefEP43(strExc(1), strExc(2)) = True Then
               textEP43 = strExc(1)
               lblEP43 = strExc(2)
               SSTab1.Tab = 3
            End If
          End If
       End If
       'end 2023/06/27
   Else
   'end 2022/07/15
      p_CP143DT = PUB_TMQchkCP143(m_CP09, strTmp)
      If p_CP143DT <> "" And p_CP143DT <> textCP143.Text Then
          If p_CP143DT = "N" Then
               MsgBox "查名尚未齊備 !", vbInformation
               'Remark by Lycia 2019/10/07
               'Cancel = True
          ElseIf p_CP143DT = "Y" Then
               MsgBox "查名已齊備 !", vbInformation
          End If
          'Added by Lydia 2019/10/07 桂英:不用再人工輸入; 參考最初的說明文件"預設檢查查名單是否已齊備，若不一致則彈訊息「查名已/未齊備！」，自動更改畫面上的查名是否齊備=Y/N；"
          textCP143.Text = p_CP143DT
      End If
      p_CP143DT = strTmp 'Added by Lydia 2019/04/11 取得查名齊備日
   End If 'Added by Lydia 2022/07/15
End Sub

'Added by Lydia 2020/05/20 法律所案源收文：案件性質=>案源案件類型
Private Sub SetLOSagree()
Dim m_LOSkind As String
     
    'Modified by Lydia 2020/8/03 FCT商爭案由內商負責 +FCT
    If strSrvDate(1) >= 法律所案源收文啟用日 And (m_TM01 = "T" Or m_TM01 = "TC" Or m_TM01 = "FCT") Then
        'Modified by Lydia 2020/06/29 直接用案源檔的類型
        'm_LOSkind = PUB_GetLOSkind(m_TM01, textCP10, textTM10)
        m_LOSkind = m_LOS02
        txtLOSagree.Text = ""
        FraLOS.Visible = False
        If textTM10 = "000" Then
            'Modified by Lydia 2020/06/29 直接用案源檔的類型
            'If (Left(m_LOSkind, 1) = "C" Or Left(m_LOS02, 1) = "C") And m_LOS01 = "" Then  'C類-未分案通知
            If Left(m_LOSkind, 1) = "C" And m_LOS01 = "" Then
                 FraLOS.Visible = True
                 'Modified by Lydia 2020/07/14 內商林經理：商標案C類案源之言詞辯論，分案時不預設法律所配合，由程序個案輸入
                 'txtLOSagree.Text = "Y"
                 txtLOSagree.Text = ""
            End If
        End If
    End If

End Sub

Private Sub txtLOSagree_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperCase(KeyAscii)
    If KeyAscii <> 89 And KeyAscii <> 8 Then
       KeyAscii = 0
       Beep
    End If
End Sub

Private Sub txtLOSagree_GotFocus()
   TextInverse txtLOSagree 'Added by Lydia 2020/05/29
End Sub

'Added by Lydia 2020/05/20 法律所案源收文：讀取法務案源檔
Private Sub ReadLOS()
Dim stSQL As String, intQ As Integer
Dim RsQ As ADODB.Recordset
   
   m_LOS01 = ""
   m_LOS07 = ""
   m_LOS15 = ""
   textCP10.Locked = False
   textTM10.Locked = False
   'Modified by Lydia 2020/8/03 FCT商爭案由內商負責 +FCT
   If strSrvDate(1) >= 法律所案源收文啟用日 And (m_TM01 = "T" Or m_TM01 = "TC" Or m_TM01 = "FCT") Then
        stSQL = "select X.* from CaseProgress, LawOfficeSource X where CP09='" & m_CP09 & "'  and CP162=LOS15(+) and cp162 is not null "
        intQ = 1
        Set RsQ = ClsLawReadRstMsg(intQ, stSQL)
        If intQ = 1 Then
           '案源總收文號
           m_LOS01 = "" & RsQ.Fields("los01")
           'Added by Lydia 2020/06/09 案源案件類型
           m_LOS02 = "" & RsQ.Fields("los02")
           '放棄日期
           m_LOS07 = "" & RsQ.Fields("los07")
           '案源單號
           m_LOS15 = "" & RsQ.Fields("los15")
           '已分案通知: 不可變更案件性質和申請國家
           'If m_LOS01 <> "" Then  'Mark by Lydia 2020/07/14 都不可以變更
               textCP10.Locked = True
               textTM10.Locked = True
           'End If
        End If
        Set RsQ = Nothing
   End If
End Sub

'Added by Morgan 2022/12/15
'台灣112年以後繳註冊費需輸入形式
Private Sub SetFrame1()
   Frame1.Visible = False
   If textTM10 = "000" And Val(m_CP27) = 0 Then
      If PUB_TWCertPty(m_TM01, textCP10, m_TM02, m_TM03, m_TM04) = True Then
         Frame1.Visible = True
      End If
   End If
End Sub

'Added by Lydia 2022/07/15 移出為獨立函數
Private Sub setFrame21()

   Frame21.Visible = False
   m_EP06 = "": m_EP06DT = "": textEP34.Enabled = True
   'Modified by Lydia 2018/12/10 T台灣案填寫接洽單管控文件及查名是否齊備
   Label57.Visible = False: textEP34.Visible = False '預設會稿、查名不顯示
   Label65.Visible = False: textCP143.Visible = False
   '增加T大陸案之齊備日管控; TC案之齊備日管控
   'Modify By Sindy 2025/7/30 FCT案727分析不屬於爭議案 + FCT_NotTMdebate
   If (m_TM01 = "FCT" And textTM10 = "000" And InStr(TMdebate, textCP10) > 0 And InStr(FCT_NotTMdebate, textCP10) = 0 And DBDATE(textCP05) >= TMdebateStarDT) _
          Or (m_TM01 = "T" And InStr("000,020", textTM10) > 0) Or m_TM01 = "TC" Then
      Frame21.Visible = True
      'Added by Lydia 2018/12/10 區分商爭和商申
      If m_TM01 = "T" Then
            If InStr(TMdebate, textCP10) > 0 Then   '商爭
                Label57.Visible = True: textEP34.Visible = True
                Label64.Caption = "資料是否齊備：       (Y/N)"
            Else  '商申
                 If textCP10 = 申請 Then  '商申
                    Label65.Visible = True: textCP143.Visible = True
                    'Added by Lydia 2019/01/30
                    'Modified by Lydia 2020/11/04 收文之查名齊備日; m_CP143=>p_strCP143
                    If Val(p_strCP143) = 0 Then
                        textCP143.Text = "N"
                    Else
                        textCP143.Text = "Y"
                    End If
                    p_strCP143 = textCP143.Text
                    Call textCP143_Validate(False) '檢查-查名是否齊備
                    'end 2019/01/30
                 End If
                 Label64.Caption = "文件是否齊備：       (Y/N)"
            End If
      ElseIf m_TM01 = "FCT" Then
            Label57.Visible = True: textEP34.Visible = True
            Label64.Caption = "資料是否齊備：       (Y/N)"
      ElseIf m_TM01 = "TC" Then
           'TC案之齊備日管控: 臺灣TC案不會稿，但大陸TC案要會稿
           If textTM10 = "020" Then
              Label57.Visible = True: textEP34.Visible = True
           End If
           Label64.Caption = "文件是否齊備：       (Y/N)"
           'Modify by Amy 2022/11/17 急件搬出 Frame21,故改都顯示
           If strSrvDate(1) < 接洽單電子收文啟用日 Then
                'Label58.Visible = False: textCP122.Visible = False  '急件：不顯示
                textCP122.Enabled = False
           End If
      End If
      'end 2018/12/10
      '讀取資料
      'Modified by Lydia 2023/01/31 +EP13
      'Modified by Lydia 2023/02/10 改為獨立欄位EP13=>EP43
      strSql = "SELECT ep06,ep34,EP43 FROM engineerprogress WHERE ep02='" & Trim(textCP09) & "'"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strSql)
      If intI = 1 Then
         If Not IsNull(RsTemp.Fields(0)) Then
            If RsTemp.Fields(0) > 0 Then
               m_EP06DT = RsTemp.Fields(0)
               textEP06.Text = "Y"
            Else
               textEP06.Text = "N"
            End If
         End If
         If Not IsNull(RsTemp.Fields(1)) Then
            textEP34.Text = RsTemp.Fields(1)
         End If
         '案件性質為613補充答辯或612補充理由時，則只可不會稿
         If Trim(textCP10) = "613" Or _
            Trim(textCP10) = "612" Then
            If textEP34.Text = "N" Then 'Add By Sindy 2013/3/12 +if
               textEP34.Enabled = False
            End If
         End If
         'Added by Lydia 2023/01/31 T大陸查名; 判斷收文日在2/1以後啟用
         'Modified by Lydia 2023/02/10 改為獨立欄位EP13=>EP43
         If m_TM01 = "T" And m_TM10 = "020" And Trim(textCP10) = "101" And DBDATE(textCP05) >= "20230201" Then
            Frame22.Visible = True
            textEP43.Text = "" & RsTemp.Fields("EP43")
            textEP43.Tag = textEP43.Text
            If Trim(textEP43) = "" Then
               If m_CP27 = "" And textCP57 = "" And textCP143 <> "Y" Then
                  If GetDefEP43(strExc(1), strExc(2)) = True Then
                     textEP43 = strExc(1)
                     lblEP43 = strExc(2)
                  End If
               End If
            Else
               If textEP43.Text = "99997" Then
                  ChkEP43.Value = 1
               End If
               Call textEP43_Validate(False)
            End If
         End If
         'end 2023/01/31
      End If
      m_EP06 = textEP06
   End If
   '2012/5/8 End
   'Added by Lydia 2019/04/11 T台灣案: 非爭議案(A類)之T案收文齊備排除的案件性質,預設文件齊備=Y
   If textEP06.Visible = True And Left(m_CP09, 1) = "A" And textTM10 = "000" And InStr(T案收文齊備排除, textCP10) > 0 Then
       textEP06 = "Y"
   End If
   'end 2019/04/11

   'Add By Sindy 2023/3/13
   If m_bolIsFirstKeyCP14 = True And txtF0301 <> "" Then
      strSql = "SELECT CRL01,CRL88,CRL89,CRL90,CRL137 FROM ConsultRecordList " & _
               "WHERE CRL01 = '" & txtF0301 & "' "
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strSql)
      If intI = 1 Then
         If textEP06.Text = "" And "" & RsTemp.Fields("CRL88") = "N" Then textEP06.Text = RsTemp.Fields("CRL88")
         If textEP34.Text = "" And "" & RsTemp.Fields("CRL89") = "N" Then textEP34.Text = RsTemp.Fields("CRL89")
         If textCP143.Text = "" And "" & RsTemp.Fields("CRL137") = "N" Then textCP143.Text = RsTemp.Fields("CRL137")
      End If
   End If
   '2023/3/13 END

   textEP06.Tag = textEP06.Text 'Added by Lydia 2019/07/29 記錄預設(文件是否齊備)
End Sub

'Add by Amy 2022/10/07
Public Sub SetParent(ByRef fm As Form)
    Set m_PrevForm = fm
End Sub

Private Sub SetGrd()
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

'Added by Lydia 2023/01/31 取得預設T大陸查名人員
'Modified by Lydia 2023/02/10 改為獨立欄位EP13=>EP43
Private Function GetDefEP43(ByRef pUserNo As String, ByRef pUserName As String) As Boolean
Dim intQ As Integer
Dim strQ1 As String, rsQD As New ADODB.Recordset
   
   pUserNo = "": pUserName = ""
   GetDefEP43 = False
   strQ1 = "select '1' ord1, aal04, st02 from addressa4list,staff where aal01='大陸查名' and aal04=st01(+) and aal02> (select aal02 from addressa4list where aal01='大陸查名' and aal03='1') " & _
              "union all select '2' ord1, aal04, st02 from addressa4list,staff where aal01='大陸查名' and aal04=st01(+) and aal02< (select aal02 from addressa4list where aal01='大陸查名' and aal03='1') " & _
              "union all select '3' ord1, aal04, st02 from addressa4list,staff where aal01='大陸查名' and aal04=st01(+) and aal03='1' " & _
              "order by 1 asc ,2 asc "
   intQ = 1
   Set rsQD = ClsLawReadRstMsg(intQ, strQ1)
   If intQ = 1 Then
      rsQD.MoveFirst
      Do While Not rsQD.EOF
         '單純判斷是否請假
         strQ1 = GetCaseDutyAgent("" & rsQD.Fields("aal04"), "", False)
         If strQ1 = "" Then
            pUserNo = "" & rsQD.Fields("aal04")
            pUserName = "" & rsQD.Fields("st02")
            GetDefEP43 = True
            Exit Do
         End If
         rsQD.MoveNext
      Loop
   End If
   Set rsQD = Nothing
   
End Function

'Added by Lydia 2023/01/31
'Modified by Lydia 2023/02/10 改為獨立欄位EP13=>EP43
Private Sub textEP43_GotFocus()
   TextInverse textEP43
End Sub

Private Sub textEP43_KeyPress(KeyAscii As Integer)
KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textEP43_Validate(Cancel As Boolean)
   Cancel = False
   lblEP43.Caption = Empty
   If IsEmptyText(textEP43) = False Then
      lblEP43.Caption = GetStaffName(textEP43)
      If IsEmptyText(lblEP43.Caption) = True Then
         Cancel = True
         MsgBox "承辦人代號不存在", vbOKOnly, "資料檢核"
      End If
      'Added by Lydia 2023/03/20 收文接洽記錄單時若為不查名, 查名是否齊備應自動設定為 "Y", 以利承辦人員進行後續作業!
       If Val(m_CP27) = 0 And textEP43.Text = "99997" And textEP43.Tag <> textEP43.Text Then
           textCP143.Text = "Y"
       End If
       'end 2023/03/120
   End If
End Sub

Private Sub ChkEP43_Click()
   If ChkEP43.Value = 1 Then
       textEP43.Text = "99997"
       Call textEP43_Validate(False)
   Else
      If m_CP27 = "" And textCP57 = "" And textCP143 <> "Y" Then
         If GetDefEP43(strExc(1), strExc(2)) = True Then
            textEP43 = strExc(1)
            lblEP43 = strExc(2)
         End If
      End If
   End If
End Sub

'Added by Lydia 2023/11/16
Private Sub cboTM72_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub cboTM72_Validate(Cancel As Boolean)
Dim intX As Integer, intQ As Integer
   
   intX = -1
   If Trim(cboTM72.Text) <> "" And cboTM72.Tag <> cboTM72.Text Then
        For intQ = 0 To cboTM72.ListCount - 1
           If InStr(cboTM72.List(intQ), Trim(cboTM72.Text)) > 0 Then
              intX = intQ
              Exit For
           End If
        Next intQ
        If intX = -1 Then
           cboTM72.SetFocus
           cboTM72.Tag = cboTM72.Text
           Cancel = True
           Exit Sub
        Else
           cboTM72.ListIndex = intX
        End If
   End If
   textTM72 = Trim(Left(cboTM72, 1))
   cboTM72.Tag = cboTM72.Text
End Sub

Private Sub cboTM08_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub cboTM08_Validate(Cancel As Boolean)
Dim intX As Integer, intQ As Integer
   
   intX = -1
   If Trim(cboTM08.Text) <> "" And cboTM08.Tag <> cboTM08.Text Then
        For intQ = 0 To cboTM08.ListCount - 1
           If InStr(cboTM08.List(intQ), Trim(cboTM08.Text)) > 0 Then
              intX = intQ
              Exit For
           End If
        Next intQ
        If intX = -1 Then
            cboTM08.SetFocus
            cboTM08.Tag = cboTM08.Text
            Cancel = True
            Exit Sub
        Else
            cboTM08.ListIndex = intX
        End If
   End If
   textTM08 = Trim(Left(cboTM08.Text, 1))
   cboTM08.Tag = cboTM08.Text
End Sub
'end 2023/11/16

'Added by Lydia 2025/10/23
Private Sub cmdTFBaseNo_Click()
   
   Call frm020509.SetParent(Me, m_TM01 & m_TM02 & m_TM03 & m_TM04, "U")
   frm020509.Show vbModal
End Sub
