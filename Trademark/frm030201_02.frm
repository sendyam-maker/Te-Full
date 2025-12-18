VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm030201_02 
   BorderStyle     =   1  '單線固定
   Caption         =   "分案"
   ClientHeight    =   6780
   ClientLeft      =   144
   ClientTop       =   5604
   ClientWidth     =   9132
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6780
   ScaleWidth      =   9132
   Begin VB.CommandButton cmdFile 
      Caption         =   "檢視接洽單"
      CausesValidation=   0   'False
      Height          =   350
      Left            =   2040
      TabIndex        =   241
      Top             =   0
      Width           =   1065
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6024
      Left            =   60
      TabIndex        =   112
      Top             =   672
      Width           =   9048
      _ExtentX        =   15960
      _ExtentY        =   10626
      _Version        =   393216
      Tabs            =   5
      TabsPerRow      =   5
      TabHeight       =   420
      TabCaption(0)   =   "第一頁"
      TabPicture(0)   =   "frm030201_02.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label36"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label37"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label35"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label34(0)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label33"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label40"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label14"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label13"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label16"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label15"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label23"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label9"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label8"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label7"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Label6"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Label1(1)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Label1(8)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Label5"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Label25"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Label4"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Label3"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "Label26"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "Label27"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "Label28"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "Label38"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "Label39"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "Label12"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "Label11"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "Label42"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "lblDivCase"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "Label43"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "Label44"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "Label70"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "textTM02"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "textCP05"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "textTM02_2"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "textTM03"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "textTM04"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "textCP57"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).Control(39)=   "textSP59_2"
      Tab(0).Control(39).Enabled=   0   'False
      Tab(0).Control(40)=   "textSP59"
      Tab(0).Control(40).Enabled=   0   'False
      Tab(0).Control(41)=   "textSP58_2"
      Tab(0).Control(41).Enabled=   0   'False
      Tab(0).Control(42)=   "textSP58"
      Tab(0).Control(42).Enabled=   0   'False
      Tab(0).Control(43)=   "textTM23"
      Tab(0).Control(43).Enabled=   0   'False
      Tab(0).Control(44)=   "textTM23_2"
      Tab(0).Control(44).Enabled=   0   'False
      Tab(0).Control(45)=   "textTM06"
      Tab(0).Control(45).Enabled=   0   'False
      Tab(0).Control(46)=   "textTM07"
      Tab(0).Control(46).Enabled=   0   'False
      Tab(0).Control(47)=   "textTM05"
      Tab(0).Control(47).Enabled=   0   'False
      Tab(0).Control(48)=   "textTM29"
      Tab(0).Control(48).Enabled=   0   'False
      Tab(0).Control(49)=   "textCP26"
      Tab(0).Control(49).Enabled=   0   'False
      Tab(0).Control(50)=   "textTM28"
      Tab(0).Control(50).Enabled=   0   'False
      Tab(0).Control(51)=   "textTM10_2"
      Tab(0).Control(51).Enabled=   0   'False
      Tab(0).Control(52)=   "textTM08_2"
      Tab(0).Control(52).Enabled=   0   'False
      Tab(0).Control(53)=   "textTM08"
      Tab(0).Control(53).Enabled=   0   'False
      Tab(0).Control(54)=   "textCP13"
      Tab(0).Control(54).Enabled=   0   'False
      Tab(0).Control(55)=   "textCP13_2"
      Tab(0).Control(55).Enabled=   0   'False
      Tab(0).Control(56)=   "textTM10"
      Tab(0).Control(56).Enabled=   0   'False
      Tab(0).Control(57)=   "textCP06"
      Tab(0).Control(57).Enabled=   0   'False
      Tab(0).Control(58)=   "textCP07"
      Tab(0).Control(58).Enabled=   0   'False
      Tab(0).Control(59)=   "textCP10_2"
      Tab(0).Control(59).Enabled=   0   'False
      Tab(0).Control(60)=   "textCP10"
      Tab(0).Control(60).Enabled=   0   'False
      Tab(0).Control(61)=   "textCP14_2"
      Tab(0).Control(61).Enabled=   0   'False
      Tab(0).Control(62)=   "textCP14"
      Tab(0).Control(62).Enabled=   0   'False
      Tab(0).Control(63)=   "textCP48"
      Tab(0).Control(63).Enabled=   0   'False
      Tab(0).Control(64)=   "textCP16"
      Tab(0).Control(64).Enabled=   0   'False
      Tab(0).Control(65)=   "textCP17"
      Tab(0).Control(65).Enabled=   0   'False
      Tab(0).Control(66)=   "textCP18"
      Tab(0).Control(66).Enabled=   0   'False
      Tab(0).Control(67)=   "textTM23_3"
      Tab(0).Control(67).Enabled=   0   'False
      Tab(0).Control(68)=   "textTM05_1"
      Tab(0).Control(68).Enabled=   0   'False
      Tab(0).Control(69)=   "textTM01"
      Tab(0).Control(69).Enabled=   0   'False
      Tab(0).Control(70)=   "textTM80"
      Tab(0).Control(70).Enabled=   0   'False
      Tab(0).Control(71)=   "textTM80_2"
      Tab(0).Control(71).Enabled=   0   'False
      Tab(0).Control(72)=   "textTM81"
      Tab(0).Control(72).Enabled=   0   'False
      Tab(0).Control(73)=   "textTM81_2"
      Tab(0).Control(73).Enabled=   0   'False
      Tab(0).Control(74)=   "textTM72"
      Tab(0).Control(74).Enabled=   0   'False
      Tab(0).Control(75)=   "textTM72_2"
      Tab(0).Control(75).Enabled=   0   'False
      Tab(0).Control(76)=   "cboTM08"
      Tab(0).Control(76).Enabled=   0   'False
      Tab(0).Control(77)=   "cboTM72"
      Tab(0).Control(77).Enabled=   0   'False
      Tab(0).Control(78)=   "Label142"
      Tab(0).Control(78).Enabled=   0   'False
      Tab(0).Control(79)=   "txtDivCaseNo(2)"
      Tab(0).Control(79).Enabled=   0   'False
      Tab(0).Control(80)=   "txtDivCaseNo(4)"
      Tab(0).Control(80).Enabled=   0   'False
      Tab(0).Control(81)=   "txtDivCaseNo(3)"
      Tab(0).Control(81).Enabled=   0   'False
      Tab(0).Control(82)=   "txtDivCaseNo(1)"
      Tab(0).Control(82).Enabled=   0   'False
      Tab(0).Control(83)=   "txtDivCaseNo(0)"
      Tab(0).Control(83).Enabled=   0   'False
      Tab(0).Control(84)=   "chkWebApp"
      Tab(0).Control(84).Enabled=   0   'False
      Tab(0).Control(85)=   "FraLOS"
      Tab(0).Control(85).Enabled=   0   'False
      Tab(0).Control(86)=   "Check11"
      Tab(0).Control(86).Enabled=   0   'False
      Tab(0).Control(87)=   "Frame4"
      Tab(0).Control(87).Enabled=   0   'False
      Tab(0).Control(88)=   "Frame5"
      Tab(0).Control(88).Enabled=   0   'False
      Tab(0).Control(89)=   "grdList"
      Tab(0).Control(89).Enabled=   0   'False
      Tab(0).ControlCount=   90
      TabCaption(1)   =   "第二頁"
      TabPicture(1)   =   "frm030201_02.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label32"
      Tab(1).Control(1)=   "Label31"
      Tab(1).Control(2)=   "Label30(0)"
      Tab(1).Control(3)=   "Label29"
      Tab(1).Control(4)=   "Label20"
      Tab(1).Control(5)=   "Label19"
      Tab(1).Control(6)=   "Label18"
      Tab(1).Control(7)=   "Label17"
      Tab(1).Control(8)=   "Label10"
      Tab(1).Control(9)=   "Label55"
      Tab(1).Control(10)=   "Label22"
      Tab(1).Control(11)=   "Label21"
      Tab(1).Control(12)=   "Line1"
      Tab(1).Control(13)=   "Label1(115)"
      Tab(1).Control(14)=   "lblTM130"
      Tab(1).Control(15)=   "textCP64"
      Tab(1).Control(16)=   "textTM58"
      Tab(1).Control(17)=   "textTM09"
      Tab(1).Control(18)=   "textTM34"
      Tab(1).Control(19)=   "textTM25"
      Tab(1).Control(20)=   "textTM26"
      Tab(1).Control(21)=   "textTM24"
      Tab(1).Control(22)=   "textCP43"
      Tab(1).Control(23)=   "textTM44_2"
      Tab(1).Control(24)=   "textTM44"
      Tab(1).Control(25)=   "textTM45"
      Tab(1).Control(26)=   "textCP09_S"
      Tab(1).Control(27)=   "textCP09_S1"
      Tab(1).Control(28)=   "textCP09_S2"
      Tab(1).Control(29)=   "textCP09_S3"
      Tab(1).Control(30)=   "textTM130"
      Tab(1).Control(31)=   "cmdPriority"
      Tab(1).Control(32)=   "txtF0301"
      Tab(1).ControlCount=   33
      TabCaption(2)   =   "第三頁"
      TabPicture(2)   =   "frm030201_02.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label45"
      Tab(2).Control(1)=   "Label46"
      Tab(2).Control(2)=   "Label47"
      Tab(2).Control(3)=   "Label48"
      Tab(2).Control(4)=   "Label49"
      Tab(2).Control(5)=   "Label50"
      Tab(2).Control(6)=   "Label51"
      Tab(2).Control(7)=   "Label52"
      Tab(2).Control(8)=   "Label53"
      Tab(2).Control(9)=   "Label54"
      Tab(2).Control(10)=   "Label56"
      Tab(2).Control(11)=   "Label57"
      Tab(2).Control(12)=   "Label58"
      Tab(2).Control(13)=   "Label59"
      Tab(2).Control(14)=   "Label60"
      Tab(2).Control(15)=   "textTM82"
      Tab(2).Control(16)=   "textTM90"
      Tab(2).Control(17)=   "textTM86"
      Tab(2).Control(18)=   "textTM83"
      Tab(2).Control(19)=   "textTM91"
      Tab(2).Control(20)=   "textTM87"
      Tab(2).Control(21)=   "textTM84"
      Tab(2).Control(22)=   "textTM92"
      Tab(2).Control(23)=   "textTM88"
      Tab(2).Control(24)=   "textTM85"
      Tab(2).Control(25)=   "textTM93"
      Tab(2).Control(26)=   "textTM89"
      Tab(2).Control(27)=   "textCP50"
      Tab(2).Control(28)=   "textCP51"
      Tab(2).Control(29)=   "textCP52"
      Tab(2).ControlCount=   30
      TabCaption(3)   =   "第四頁"
      TabPicture(3)   =   "frm030201_02.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame2"
      Tab(3).Control(1)=   "Frame3"
      Tab(3).ControlCount=   2
      TabCaption(4)   =   "第五頁"
      TabPicture(4)   =   "frm030201_02.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Frame1"
      Tab(4).ControlCount=   1
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdList 
         Height          =   1020
         Left            =   950
         TabIndex        =   38
         Top             =   4650
         Width           =   7750
         _ExtentX        =   13674
         _ExtentY        =   1799
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
      Begin VB.Frame Frame5 
         Height          =   405
         Left            =   950
         TabIndex        =   247
         Top             =   5580
         Width           =   6860
         Begin VB.TextBox textCP142 
            Height          =   264
            Left            =   3230
            MaxLength       =   7
            TabIndex        =   255
            Top             =   120
            Width           =   945
         End
         Begin VB.OptionButton OptSendType 
            Caption         =   "指定日期"
            Height          =   180
            Index           =   3
            Left            =   2160
            TabIndex        =   254
            Top             =   170
            Width           =   1065
         End
         Begin VB.OptionButton OptSendType 
            Caption         =   "收款後"
            Height          =   180
            Index           =   2
            Left            =   1290
            TabIndex        =   253
            Top             =   170
            Width           =   870
         End
         Begin VB.OptionButton OptSendType 
            Caption         =   "立即"
            Height          =   180
            Index           =   1
            Left            =   30
            TabIndex        =   252
            Top             =   170
            Width           =   1260
         End
         Begin VB.Frame Frame6 
            Height          =   340
            Left            =   4200
            TabIndex        =   248
            Top             =   60
            Width           =   2170
            Begin VB.OptionButton Option1 
               Caption         =   "之前"
               Height          =   195
               Index           =   1
               Left            =   690
               TabIndex        =   251
               Top             =   95
               Width           =   705
            End
            Begin VB.OptionButton Option1 
               Caption         =   "當天"
               Height          =   195
               Index           =   0
               Left            =   0
               TabIndex        =   250
               Top             =   95
               Width           =   705
            End
            Begin VB.OptionButton Option1 
               Caption         =   "之後"
               Height          =   195
               Index           =   2
               Left            =   1380
               TabIndex        =   249
               Top             =   95
               Width           =   705
            End
         End
      End
      Begin VB.Frame Frame4 
         BorderStyle     =   0  '沒有框線
         Height          =   285
         Left            =   5790
         TabIndex        =   244
         Top             =   2310
         Visible         =   0   'False
         Width           =   2985
         Begin VB.TextBox textTM136 
            Height          =   264
            Left            =   1140
            MaxLength       =   1
            TabIndex        =   245
            Top             =   0
            Width           =   345
         End
         Begin VB.Label Label73 
            AutoSize        =   -1  'True
            Caption         =   "註冊證形式:             1:電子 2:紙本"
            Height          =   180
            Left            =   60
            TabIndex        =   246
            Top             =   40
            Width           =   2880
            WordWrap        =   -1  'True
         End
      End
      Begin VB.CheckBox Check11 
         Caption         =   "急件"
         ForeColor       =   &H00000000&
         Height          =   200
         Left            =   3510
         TabIndex        =   243
         Top             =   330
         Width           =   765
      End
      Begin VB.TextBox txtF0301 
         BorderStyle     =   0  '沒有框線
         Height          =   255
         Left            =   -74910
         Locked          =   -1  'True
         TabIndex        =   242
         Text            =   "txtF0301"
         Top             =   5400
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Frame FraLOS 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   0  '沒有框線
         Caption         =   "Frame2"
         Height          =   255
         Left            =   5016
         TabIndex        =   234
         Top             =   1680
         Width           =   3375
         Begin MSForms.TextBox txtLOSagree 
            Height          =   285
            Left            =   1890
            TabIndex        =   18
            Top             =   -8
            Width           =   405
            VariousPropertyBits=   671105051
            MaxLength       =   1
            Size            =   "714;494"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin VB.Label LBL6 
            Caption         =   "是否需要法律所配合：　　　(Y: 配合) "
            Height          =   195
            Left            =   30
            TabIndex        =   235
            Top             =   30
            Width           =   3135
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "案件基本資料"
         ForeColor       =   &H000000C0&
         Height          =   4665
         Left            =   -74880
         TabIndex        =   194
         Top             =   360
         Width           =   8805
         Begin MSForms.TextBox textTM121 
            Height          =   285
            Left            =   4410
            TabIndex        =   95
            Top             =   880
            Width           =   300
            VariousPropertyBits=   671105051
            MaxLength       =   1
            Size            =   "529;503"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "以EMail通知:        (Y:是   D:僅D/N）"
            Height          =   180
            Index           =   76
            Left            =   3390
            TabIndex        =   240
            Top             =   930
            Width           =   2715
         End
         Begin MSForms.TextBox textTM46 
            Height          =   300
            Left            =   1890
            TabIndex        =   94
            Top             =   880
            Width           =   372
            VariousPropertyBits=   671105051
            MaxLength       =   1
            Size            =   "656;529"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox textTM127 
            Height          =   300
            Left            =   1890
            TabIndex        =   96
            Top             =   1185
            Width           =   2700
            VariousPropertyBits=   671105051
            MaxLength       =   20
            Size            =   "4762;529"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox textTM35 
            Height          =   300
            Left            =   1620
            TabIndex        =   226
            Top             =   4110
            Width           =   2895
            VariousPropertyBits=   671105051
            MaxLength       =   50
            Size            =   "5106;529"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox textTM131 
            Height          =   765
            Left            =   1620
            TabIndex        =   103
            Top             =   3320
            Width           =   7155
            VariousPropertyBits=   -1466941413
            MaxLength       =   140
            ScrollBars      =   2
            Size            =   "12621;1349"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox textTM43 
            Height          =   300
            Left            =   1620
            TabIndex        =   102
            Top             =   3015
            Width           =   7155
            VariousPropertyBits=   671105051
            Size            =   "12621;529"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox textTM42 
            Height          =   300
            Left            =   1620
            TabIndex        =   101
            Top             =   2710
            Width           =   3825
            VariousPropertyBits=   671105051
            Size            =   "6747;529"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox textTM41 
            Height          =   300
            Left            =   1620
            TabIndex        =   100
            Top             =   2405
            Width           =   3000
            VariousPropertyBits=   671105051
            Size            =   "5292;529"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox textTM40 
            Height          =   300
            Left            =   1620
            TabIndex        =   99
            Top             =   2100
            Width           =   7155
            VariousPropertyBits=   671105051
            Size            =   "12621;529"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox textTM39 
            Height          =   300
            Left            =   1620
            TabIndex        =   98
            Top             =   1795
            Width           =   3825
            VariousPropertyBits=   671105051
            Size            =   "6747;529"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox textTM38 
            Height          =   300
            Left            =   1620
            TabIndex        =   97
            Top             =   1490
            Width           =   3000
            VariousPropertyBits=   671105051
            MaxLength       =   30
            Size            =   "5292;529"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox textTM56_1 
            Height          =   300
            Left            =   1620
            TabIndex        =   93
            Top             =   575
            Width           =   1212
            VariousPropertyBits=   671105051
            MaxLength       =   8
            Size            =   "2138;529"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox textTM56_2 
            Height          =   280
            Left            =   2850
            TabIndex        =   214
            TabStop         =   0   'False
            Top             =   575
            Width           =   5892
            VariousPropertyBits=   671105051
            BackColor       =   -2147483633
            Size            =   "10393;494"
            Value           =   "textTM56_2"
            SpecialEffect   =   0
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox textTM69_1 
            Height          =   300
            Left            =   1620
            TabIndex        =   92
            Top             =   270
            Width           =   1212
            VariousPropertyBits=   671105051
            MaxLength       =   8
            Size            =   "2138;529"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox textTM69_2 
            Height          =   280
            Left            =   2850
            TabIndex        =   213
            TabStop         =   0   'False
            Top             =   270
            Width           =   5892
            VariousPropertyBits=   671105051
            BackColor       =   -2147483633
            Size            =   "10393;494"
            Value           =   "textTM69_2"
            SpecialEffect   =   0
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "D/N是否列印申請人 :             (  Y:印 )"
            Height          =   180
            Index           =   32
            Left            =   90
            TabIndex        =   233
            Top             =   930
            Width           =   2820
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "CLIENT_MATTER_ID:"
            Height          =   180
            Index           =   169
            Left            =   90
            TabIndex        =   232
            Top             =   1235
            Width           =   1725
         End
         Begin VB.Label Label24 
            Caption         =   "客戶案件案號 :"
            Height          =   255
            Left            =   360
            TabIndex        =   227
            Top             =   4125
            Width           =   1215
         End
         Begin VB.Label Label69 
            AutoSize        =   -1  'True
            Caption         =   "定稿商標名稱 :"
            Height          =   180
            Left            =   255
            TabIndex        =   223
            Top             =   3320
            Width           =   1320
         End
         Begin VB.Label Label1 
            Caption         =   "聯絡人2(日) :"
            Height          =   255
            Index           =   74
            Left            =   420
            TabIndex        =   222
            Top             =   3028
            Width           =   1155
         End
         Begin VB.Label Label1 
            Caption         =   "聯絡人2(英) :"
            Height          =   255
            Index           =   73
            Left            =   420
            TabIndex        =   221
            Top             =   2723
            Width           =   1155
         End
         Begin VB.Label Label1 
            Caption         =   "聯絡人2(中) :"
            Height          =   255
            Index           =   72
            Left            =   420
            TabIndex        =   220
            Top             =   2418
            Width           =   1155
         End
         Begin VB.Label Label1 
            Caption         =   "聯絡人1(日) :"
            Height          =   255
            Index           =   71
            Left            =   420
            TabIndex        =   219
            Top             =   2113
            Width           =   1155
         End
         Begin VB.Label Label1 
            Caption         =   "聯絡人1(英) :"
            Height          =   255
            Index           =   70
            Left            =   420
            TabIndex        =   218
            Top             =   1808
            Width           =   1155
         End
         Begin VB.Label Label1 
            Caption         =   "聯絡人1(中) :"
            Height          =   255
            Index           =   69
            Left            =   420
            TabIndex        =   217
            Top             =   1503
            Width           =   1155
         End
         Begin VB.Label Label1 
            Caption         =   "固定請款對象 :"
            Height          =   255
            Index           =   34
            Left            =   90
            TabIndex        =   216
            Top             =   588
            Width           =   1245
         End
         Begin VB.Label Label1 
            Caption         =   "D/N固定列印對象 :"
            Height          =   270
            Index           =   33
            Left            =   90
            TabIndex        =   215
            Top             =   280
            Width           =   1520
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "代理人"
         ForeColor       =   &H000000C0&
         Height          =   2685
         Left            =   -74880
         TabIndex        =   193
         Top             =   3000
         Width           =   8775
         Begin VB.TextBox txtFA 
            Height          =   285
            Index           =   91
            Left            =   4980
            MaxLength       =   1
            TabIndex        =   84
            Top             =   780
            Width           =   330
         End
         Begin VB.Label Label72 
            AutoSize        =   -1  'True
            Caption         =   "商標以 EMail 通知：       （Y：是   D：僅D/N）"
            Height          =   180
            Left            =   3360
            TabIndex        =   239
            Top             =   840
            Width           =   3645
         End
         Begin MSForms.TextBox textFA109 
            Height          =   300
            Left            =   2100
            TabIndex        =   83
            Top             =   780
            Width           =   330
            VariousPropertyBits=   671105051
            MaxLength       =   1
            Size            =   "582;529"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox textFA07 
            Height          =   300
            Left            =   2100
            TabIndex        =   86
            Top             =   1380
            Width           =   2052
            VariousPropertyBits=   671105051
            MaxLength       =   10
            Size            =   "3619;529"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox textFA52 
            Height          =   300
            Left            =   2100
            TabIndex        =   89
            Top             =   1980
            Width           =   2052
            VariousPropertyBits=   671105051
            MaxLength       =   10
            Size            =   "3619;529"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox textFA08 
            Height          =   300
            Left            =   5610
            TabIndex        =   87
            Top             =   1380
            Width           =   3135
            VariousPropertyBits=   671105051
            MaxLength       =   35
            Size            =   "5530;529"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox textFA09 
            Height          =   300
            Left            =   2100
            TabIndex        =   88
            Top             =   1680
            Width           =   6645
            VariousPropertyBits=   671105051
            MaxLength       =   60
            Size            =   "11721;529"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox textFA53 
            Height          =   300
            Left            =   5610
            TabIndex        =   90
            Top             =   1980
            Width           =   3135
            VariousPropertyBits=   671105051
            MaxLength       =   35
            Size            =   "5530;529"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox textFA54 
            Height          =   300
            Left            =   2100
            TabIndex        =   91
            Top             =   2280
            Width           =   6645
            VariousPropertyBits=   671105051
            MaxLength       =   60
            Size            =   "11721;529"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox textFA106 
            Height          =   300
            Left            =   2100
            TabIndex        =   85
            Top             =   1080
            Width           =   2772
            VariousPropertyBits=   671105051
            MaxLength       =   30
            Size            =   "4890;529"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox textFA111 
            Height          =   300
            Left            =   2100
            TabIndex        =   81
            Top             =   180
            Width           =   1095
            VariousPropertyBits=   671105051
            MaxLength       =   8
            Size            =   "1931;529"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox textFA111_2 
            Height          =   285
            Left            =   3210
            TabIndex        =   205
            TabStop         =   0   'False
            Top             =   188
            Width           =   5505
            VariousPropertyBits=   671105051
            BackColor       =   -2147483633
            Size            =   "9710;503"
            Value           =   "textFA111_2"
            SpecialEffect   =   0
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox textFA107 
            Height          =   300
            Left            =   2100
            TabIndex        =   82
            Top             =   480
            Width           =   1095
            VariousPropertyBits=   671105051
            MaxLength       =   8
            Size            =   "1931;529"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox textFA107_2 
            Height          =   285
            Left            =   3210
            TabIndex        =   204
            TabStop         =   0   'False
            Top             =   488
            Width           =   5505
            VariousPropertyBits=   671105051
            BackColor       =   -2147483633
            Size            =   "9710;503"
            Value           =   "textFA107_2"
            SpecialEffect   =   0
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin VB.Label Label66 
            Alignment       =   1  '靠右對齊
            Caption         =   "聯絡人２(中)："
            Height          =   195
            Left            =   210
            TabIndex        =   231
            Top             =   2033
            Width           =   1905
         End
         Begin VB.Label Label67 
            Alignment       =   1  '靠右對齊
            Caption         =   "聯絡人２(英)："
            Height          =   195
            Left            =   4380
            TabIndex        =   230
            Top             =   2033
            Width           =   1245
         End
         Begin VB.Label Label71 
            AutoSize        =   -1  'True
            Caption         =   "商標D/N是否印申請人：          (Y:印)"
            Height          =   195
            Left            =   150
            TabIndex        =   229
            Top             =   833
            Width           =   2820
         End
         Begin VB.Label Label68 
            Alignment       =   1  '靠右對齊
            Caption         =   "聯絡人１(日)："
            Height          =   195
            Left            =   210
            TabIndex        =   212
            Top             =   1733
            Width           =   1905
         End
         Begin VB.Label Label41 
            Alignment       =   1  '靠右對齊
            Caption         =   "聯絡人２(日)："
            Height          =   195
            Index           =   0
            Left            =   210
            TabIndex        =   211
            Top             =   2333
            Width           =   1905
         End
         Begin VB.Label Label65 
            Alignment       =   1  '靠右對齊
            Caption         =   "聯絡人１(英)："
            Height          =   255
            Left            =   4380
            TabIndex        =   210
            Top             =   1403
            Width           =   1245
         End
         Begin VB.Label Label64 
            Alignment       =   1  '靠右對齊
            Caption         =   "聯絡人１(中)："
            Height          =   195
            Left            =   210
            TabIndex        =   209
            Top             =   1433
            Width           =   1905
         End
         Begin VB.Label Label62 
            Caption         =   "代理人商標財務編號："
            Height          =   195
            Left            =   150
            TabIndex        =   208
            Top             =   1133
            Width           =   1905
         End
         Begin VB.Label Label34 
            Caption         =   "商標D/N固定列印對象："
            Height          =   195
            Index           =   5
            Left            =   150
            TabIndex        =   207
            Top             =   233
            Width           =   1905
         End
         Begin VB.Label Label61 
            Caption         =   "商標固定請款對象："
            Height          =   195
            Left            =   150
            TabIndex        =   206
            Top             =   533
            Width           =   1905
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "申請人"
         ForeColor       =   &H000000C0&
         Height          =   2685
         Left            =   -74880
         TabIndex        =   192
         Top             =   330
         Width           =   8835
         Begin VB.TextBox txtCU 
            Height          =   264
            Index           =   126
            Left            =   4980
            MaxLength       =   1
            TabIndex        =   73
            Top             =   780
            Width           =   255
         End
         Begin VB.Label Label80 
            AutoSize        =   -1  'True
            Caption         =   "商標以 EMail 通知：     （Y：是   D：僅D/N）"
            Height          =   180
            Index           =   19
            Left            =   3360
            TabIndex        =   238
            Top             =   840
            Width           =   3555
         End
         Begin MSForms.TextBox textCU149 
            Height          =   300
            Left            =   2190
            TabIndex        =   72
            Top             =   780
            Width           =   315
            VariousPropertyBits=   671105051
            MaxLength       =   1
            Size            =   "556;529"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox textCU58 
            Height          =   300
            Left            =   2190
            TabIndex        =   75
            Top             =   1380
            Width           =   2052
            VariousPropertyBits=   671105051
            MaxLength       =   10
            Size            =   "3619;529"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox textCU59 
            Height          =   300
            Left            =   5550
            TabIndex        =   76
            Top             =   1380
            Width           =   3195
            VariousPropertyBits=   671105051
            MaxLength       =   35
            Size            =   "5636;529"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox textCU60 
            Height          =   300
            Left            =   2190
            TabIndex        =   77
            Top             =   1680
            Width           =   6555
            VariousPropertyBits=   671105051
            MaxLength       =   60
            Size            =   "11562;529"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox textCU61 
            Height          =   300
            Left            =   2190
            TabIndex        =   78
            Top             =   1980
            Width           =   2052
            VariousPropertyBits=   671105051
            MaxLength       =   10
            Size            =   "3619;529"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox textCU62 
            Height          =   300
            Left            =   5565
            TabIndex        =   79
            Top             =   1980
            Width           =   3165
            VariousPropertyBits=   671105051
            MaxLength       =   35
            Size            =   "5583;529"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox textCU63 
            Height          =   300
            Left            =   2190
            TabIndex        =   80
            Top             =   2280
            Width           =   6555
            VariousPropertyBits=   671105051
            MaxLength       =   60
            Size            =   "11562;529"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox textCU151 
            Height          =   300
            Left            =   2190
            TabIndex        =   70
            Top             =   180
            Width           =   1095
            VariousPropertyBits=   671105051
            MaxLength       =   8
            Size            =   "1931;529"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox textCU147 
            Height          =   300
            Left            =   2190
            TabIndex        =   71
            Top             =   480
            Width           =   1095
            VariousPropertyBits=   671105051
            MaxLength       =   8
            Size            =   "1931;529"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox textCU146 
            Height          =   300
            Left            =   2190
            TabIndex        =   74
            Top             =   1080
            Width           =   2772
            VariousPropertyBits=   671105051
            MaxLength       =   30
            Size            =   "4890;529"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.Label lblCU147 
            Height          =   285
            Left            =   3330
            TabIndex        =   237
            Top             =   488
            Width           =   5310
            VariousPropertyBits=   27
            Caption         =   "lblCU147"
            Size            =   "9366;503"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.Label lblCU151 
            Height          =   285
            Left            =   3330
            TabIndex        =   236
            Top             =   188
            Width           =   5310
            VariousPropertyBits=   27
            Caption         =   "lblCU151"
            Size            =   "9366;503"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin VB.Label Label80 
            AutoSize        =   -1  'True
            Caption         =   "商標D/N是否列印申請人：       (Y：印)"
            Height          =   180
            Index           =   23
            Left            =   120
            TabIndex        =   228
            Top             =   840
            Width           =   3000
         End
         Begin VB.Label Label63 
            AutoSize        =   -1  'True
            Caption         =   "聯絡人１(中)："
            Height          =   180
            Index           =   0
            Left            =   870
            TabIndex        =   203
            Top             =   1440
            Width           =   1260
         End
         Begin VB.Label Label63 
            Alignment       =   1  '靠右對齊
            AutoSize        =   -1  'True
            Caption         =   "聯絡人１(英)："
            Height          =   180
            Index           =   1
            Left            =   4350
            TabIndex        =   202
            Top             =   1440
            Width           =   1200
         End
         Begin VB.Label Label63 
            AutoSize        =   -1  'True
            Caption         =   "聯絡人１(日)："
            Height          =   180
            Index           =   2
            Left            =   870
            TabIndex        =   201
            Top             =   1740
            Width           =   1260
         End
         Begin VB.Label Label63 
            AutoSize        =   -1  'True
            Caption         =   "聯絡人２(中)："
            Height          =   180
            Index           =   3
            Left            =   870
            TabIndex        =   200
            Top             =   2040
            Width           =   1260
         End
         Begin VB.Label Label63 
            Alignment       =   1  '靠右對齊
            AutoSize        =   -1  'True
            Caption         =   "聯絡人２(英)："
            Height          =   180
            Index           =   4
            Left            =   4350
            TabIndex        =   199
            Top             =   2040
            Width           =   1200
         End
         Begin VB.Label Label63 
            AutoSize        =   -1  'True
            Caption         =   "聯絡人２(日)："
            Height          =   180
            Index           =   5
            Left            =   870
            TabIndex        =   198
            Top             =   2340
            Width           =   1260
         End
         Begin VB.Label Label80 
            AutoSize        =   -1  'True
            Caption         =   "商標D/N固定列印對象："
            Height          =   180
            Index           =   25
            Left            =   120
            TabIndex        =   197
            Top             =   240
            Width           =   2040
         End
         Begin VB.Label Label80 
            AutoSize        =   -1  'True
            Caption         =   "商標固定請款對象："
            Height          =   180
            Index           =   28
            Left            =   120
            TabIndex        =   196
            Top             =   540
            Width           =   2040
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "客戶彼所商標財務編號："
            Height          =   180
            Index           =   25
            Left            =   120
            TabIndex        =   195
            Top             =   1140
            Width           =   2040
         End
      End
      Begin VB.CheckBox chkWebApp 
         Caption         =   "電子送件"
         Height          =   255
         Left            =   3060
         TabIndex        =   20
         Top             =   1990
         Width           =   1050
      End
      Begin VB.TextBox txtDivCaseNo 
         Enabled         =   0   'False
         Height          =   300
         Index           =   0
         Left            =   5985
         MaxLength       =   3
         TabIndex        =   25
         Top             =   2580
         Visible         =   0   'False
         Width           =   390
      End
      Begin VB.TextBox txtDivCaseNo 
         Enabled         =   0   'False
         Height          =   300
         Index           =   1
         Left            =   6375
         MaxLength       =   6
         TabIndex        =   26
         Top             =   2580
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox txtDivCaseNo 
         Enabled         =   0   'False
         Height          =   300
         Index           =   3
         Left            =   7350
         MaxLength       =   1
         TabIndex        =   27
         Top             =   2580
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.TextBox txtDivCaseNo 
         Enabled         =   0   'False
         Height          =   300
         Index           =   4
         Left            =   7650
         MaxLength       =   2
         TabIndex        =   28
         Top             =   2580
         Visible         =   0   'False
         Width           =   390
      End
      Begin VB.TextBox txtDivCaseNo 
         Enabled         =   0   'False
         Height          =   270
         Index           =   2
         Left            =   7080
         MaxLength       =   1
         TabIndex        =   169
         Top             =   2610
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.CommandButton cmdPriority 
         Caption         =   "輸入(&V)"
         Height          =   280
         Left            =   -73440
         TabIndex        =   50
         Top             =   2642
         Width           =   1332
      End
      Begin VB.Label Label142 
         Caption         =   "送件方式 :"
         Height          =   200
         Left            =   90
         TabIndex        =   256
         Top             =   5730
         Width           =   920
      End
      Begin MSForms.ComboBox cboTM72 
         Height          =   300
         Left            =   3430
         TabIndex        =   7
         Top             =   1150
         Width           =   1550
         VariousPropertyBits=   679495707
         DisplayStyle    =   3
         Size            =   "2730;529"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "新細明體"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.ComboBox cboTM08 
         Height          =   290
         Left            =   1130
         TabIndex        =   6
         Top             =   1150
         Width           =   1380
         VariousPropertyBits=   679495707
         DisplayStyle    =   3
         Size            =   "2434;508"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "新細明體"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM72_2 
         Height          =   290
         Left            =   3820
         TabIndex        =   224
         TabStop         =   0   'False
         Top             =   860
         Visible         =   0   'False
         Width           =   910
         VariousPropertyBits=   671105051
         BackColor       =   16777152
         MaxLength       =   20
         Size            =   "1614;503"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM72 
         Height          =   300
         Left            =   3530
         TabIndex        =   9
         Top             =   850
         Visible         =   0   'False
         Width           =   290
         VariousPropertyBits=   671105051
         BackColor       =   16777152
         MaxLength       =   1
         Size            =   "503;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCP52 
         Height          =   300
         Left            =   -73260
         TabIndex        =   69
         Top             =   4890
         Width           =   7155
         VariousPropertyBits=   671105051
         MaxLength       =   60
         Size            =   "12621;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCP51 
         Height          =   300
         Left            =   -73260
         TabIndex        =   68
         Top             =   4559
         Width           =   7152
         VariousPropertyBits=   671105051
         MaxLength       =   60
         Size            =   "12615;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCP50 
         Height          =   300
         Left            =   -73260
         TabIndex        =   67
         Top             =   4236
         Width           =   7152
         VariousPropertyBits=   671105051
         MaxLength       =   60
         Size            =   "12615;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM130 
         Height          =   285
         Left            =   -70200
         TabIndex        =   51
         Top             =   2642
         Width           =   255
         VariousPropertyBits=   671105051
         MaxLength       =   1
         Size            =   "450;494"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM89 
         Height          =   300
         Left            =   -73260
         TabIndex        =   65
         Top             =   3590
         Width           =   7152
         VariousPropertyBits=   671105051
         MaxLength       =   185
         Size            =   "12615;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM93 
         Height          =   300
         Left            =   -73260
         TabIndex        =   66
         Top             =   3913
         Width           =   7152
         VariousPropertyBits=   671105051
         MaxLength       =   80
         Size            =   "12615;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM85 
         Height          =   300
         Left            =   -73260
         TabIndex        =   64
         Top             =   3267
         Width           =   7152
         VariousPropertyBits=   671105051
         MaxLength       =   80
         Size            =   "12615;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM88 
         Height          =   300
         Left            =   -73260
         TabIndex        =   62
         Top             =   2621
         Width           =   7152
         VariousPropertyBits=   671105051
         MaxLength       =   185
         Size            =   "12615;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM92 
         Height          =   300
         Left            =   -73260
         TabIndex        =   63
         Top             =   2944
         Width           =   7152
         VariousPropertyBits=   671105051
         MaxLength       =   80
         Size            =   "12615;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM84 
         Height          =   300
         Left            =   -73260
         TabIndex        =   61
         Top             =   2298
         Width           =   7152
         VariousPropertyBits=   671105051
         MaxLength       =   80
         Size            =   "12615;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM87 
         Height          =   300
         Left            =   -73260
         TabIndex        =   59
         Top             =   1652
         Width           =   7152
         VariousPropertyBits=   671105051
         MaxLength       =   185
         Size            =   "12615;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM91 
         Height          =   300
         Left            =   -73260
         TabIndex        =   60
         Top             =   1975
         Width           =   7152
         VariousPropertyBits=   671105051
         MaxLength       =   80
         Size            =   "12615;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM83 
         Height          =   300
         Left            =   -73260
         TabIndex        =   58
         Top             =   1329
         Width           =   7152
         VariousPropertyBits=   671105051
         MaxLength       =   80
         Size            =   "12615;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM86 
         Height          =   300
         Left            =   -73260
         TabIndex        =   56
         Top             =   683
         Width           =   7152
         VariousPropertyBits=   671105051
         MaxLength       =   185
         Size            =   "12615;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM90 
         Height          =   300
         Left            =   -73260
         TabIndex        =   57
         Top             =   1006
         Width           =   7152
         VariousPropertyBits=   671105051
         MaxLength       =   80
         Size            =   "12615;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM82 
         Height          =   300
         Left            =   -73260
         TabIndex        =   55
         Top             =   360
         Width           =   7152
         VariousPropertyBits=   671105051
         BackColor       =   16777215
         MaxLength       =   80
         Size            =   "12615;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM81_2 
         Height          =   290
         Left            =   6480
         TabIndex        =   172
         TabStop         =   0   'False
         Top             =   4370
         Width           =   2180
         VariousPropertyBits=   671105055
         ForeColor       =   -2147483641
         Size            =   "3836;503"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM81 
         Height          =   300
         Left            =   5480
         TabIndex        =   37
         Top             =   4350
         Width           =   980
         VariousPropertyBits=   671105051
         MaxLength       =   9
         Size            =   "1720;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM80_2 
         Height          =   290
         Left            =   1980
         TabIndex        =   171
         TabStop         =   0   'False
         Top             =   4370
         Width           =   2180
         VariousPropertyBits=   671105055
         ForeColor       =   -2147483641
         Size            =   "3836;503"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM80 
         Height          =   300
         Left            =   960
         TabIndex        =   36
         Top             =   4350
         Width           =   980
         VariousPropertyBits=   671105051
         MaxLength       =   9
         Size            =   "1720;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM01 
         Height          =   300
         Left            =   1140
         TabIndex        =   11
         Top             =   1420
         Width           =   610
         VariousPropertyBits=   671105051
         MaxLength       =   3
         Size            =   "1080;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM05_1 
         Height          =   880
         Left            =   1440
         TabIndex        =   29
         Top             =   2870
         Width           =   7280
         VariousPropertyBits=   -1467989989
         MaxLength       =   140
         ScrollBars      =   2
         Size            =   "12832;1552"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCP09_S3 
         Height          =   300
         Left            =   -71370
         TabIndex        =   48
         Top             =   1990
         Width           =   465
         VariousPropertyBits=   671105051
         MaxLength       =   2
         Size            =   "820;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCP09_S2 
         Height          =   300
         Left            =   -71820
         TabIndex        =   47
         Top             =   1990
         Width           =   345
         VariousPropertyBits=   671105051
         MaxLength       =   1
         Size            =   "609;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCP09_S1 
         Height          =   300
         Left            =   -72900
         TabIndex        =   46
         Top             =   1990
         Width           =   975
         VariousPropertyBits=   671105051
         MaxLength       =   6
         Size            =   "1720;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM23_3 
         Height          =   290
         Left            =   5480
         TabIndex        =   166
         TabStop         =   0   'False
         Top             =   3740
         Width           =   3200
         VariousPropertyBits=   671105055
         MaxLength       =   20
         Size            =   "5636;503"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCP09_S 
         Height          =   300
         Left            =   -73440
         TabIndex        =   45
         Top             =   1990
         Width           =   465
         VariousPropertyBits=   671105051
         MaxLength       =   1
         Size            =   "820;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCP18 
         Height          =   300
         Left            =   1170
         TabIndex        =   24
         Top             =   2580
         Width           =   1100
         VariousPropertyBits=   671105051
         MaxLength       =   8
         Size            =   "1931;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCP17 
         Height          =   300
         Left            =   3990
         TabIndex        =   23
         Top             =   2280
         Width           =   1460
         VariousPropertyBits=   671105051
         MaxLength       =   8
         Size            =   "2566;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCP16 
         Height          =   300
         Left            =   1170
         TabIndex        =   22
         Top             =   2300
         Width           =   1460
         VariousPropertyBits=   671105051
         MaxLength       =   8
         Size            =   "2566;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCP48 
         Height          =   300
         Left            =   5976
         TabIndex        =   1
         Top             =   276
         Width           =   1212
         VariousPropertyBits=   671105051
         MaxLength       =   8
         Size            =   "2143;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM45 
         Height          =   300
         Left            =   -73440
         TabIndex        =   40
         Top             =   686
         Width           =   2655
         VariousPropertyBits=   671105051
         MaxLength       =   50
         Size            =   "4683;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM44 
         Height          =   300
         Left            =   -69660
         TabIndex        =   41
         Top             =   686
         Width           =   1125
         VariousPropertyBits=   671105051
         MaxLength       =   9
         Size            =   "1984;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM44_2 
         Height          =   285
         Left            =   -68520
         TabIndex        =   156
         TabStop         =   0   'False
         Top             =   686
         Width           =   2232
         VariousPropertyBits=   671105055
         Size            =   "3937;494"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCP43 
         Height          =   300
         Left            =   -73440
         TabIndex        =   39
         Top             =   360
         Width           =   2655
         VariousPropertyBits=   671105051
         MaxLength       =   9
         Size            =   "4683;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCP14 
         Height          =   300
         Left            =   1140
         TabIndex        =   0
         Top             =   270
         Width           =   732
         VariousPropertyBits=   671105051
         MaxLength       =   6
         Size            =   "1291;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCP14_2 
         Height          =   285
         Left            =   1890
         TabIndex        =   121
         TabStop         =   0   'False
         Top             =   278
         Width           =   1572
         VariousPropertyBits=   671105055
         MaxLength       =   20
         Size            =   "2773;503"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCP10 
         Height          =   300
         Left            =   1140
         TabIndex        =   2
         Top             =   570
         Width           =   730
         VariousPropertyBits=   671105051
         MaxLength       =   6
         Size            =   "1291;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCP10_2 
         Height          =   290
         Left            =   1890
         TabIndex        =   120
         TabStop         =   0   'False
         Top             =   570
         Width           =   1570
         VariousPropertyBits=   671105055
         MaxLength       =   20
         Size            =   "2773;503"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCP07 
         Height          =   300
         Left            =   5980
         TabIndex        =   5
         Top             =   850
         Width           =   1210
         VariousPropertyBits=   671105051
         MaxLength       =   7
         Size            =   "2143;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCP06 
         Height          =   300
         Left            =   1140
         TabIndex        =   4
         Top             =   850
         Width           =   1220
         VariousPropertyBits=   671105051
         MaxLength       =   7
         Size            =   "2143;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM10 
         Height          =   300
         Left            =   5980
         TabIndex        =   3
         Top             =   580
         Width           =   610
         VariousPropertyBits=   671105051
         MaxLength       =   20
         Size            =   "1080;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCP13_2 
         Height          =   288
         Left            =   6864
         TabIndex        =   119
         TabStop         =   0   'False
         Top             =   1152
         Width           =   1236
         VariousPropertyBits=   671105055
         MaxLength       =   20
         Size            =   "2180;508"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCP13 
         Height          =   300
         Left            =   5980
         TabIndex        =   10
         Top             =   1150
         Width           =   850
         VariousPropertyBits=   671105051
         MaxLength       =   6
         Size            =   "1503;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM08 
         Height          =   300
         Left            =   3530
         TabIndex        =   8
         Top             =   540
         Visible         =   0   'False
         Width           =   290
         VariousPropertyBits=   671105051
         BackColor       =   16777152
         MaxLength       =   1
         Size            =   "512;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM08_2 
         Height          =   290
         Left            =   3820
         TabIndex        =   118
         TabStop         =   0   'False
         Top             =   550
         Visible         =   0   'False
         Width           =   910
         VariousPropertyBits=   671105051
         BackColor       =   16777152
         MaxLength       =   20
         Size            =   "1605;512"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM10_2 
         Height          =   290
         Left            =   6640
         TabIndex        =   117
         TabStop         =   0   'False
         Top             =   580
         Width           =   1570
         VariousPropertyBits=   671105055
         MaxLength       =   20
         Size            =   "2773;503"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM28 
         Height          =   300
         Left            =   5980
         TabIndex        =   16
         Top             =   1430
         Width           =   370
         VariousPropertyBits=   671105051
         MaxLength       =   20
         Size            =   "656;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCP26 
         Height          =   300
         Left            =   1440
         TabIndex        =   19
         Top             =   2000
         Width           =   370
         VariousPropertyBits=   671105051
         MaxLength       =   20
         Size            =   "656;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM29 
         Height          =   300
         Left            =   6220
         TabIndex        =   21
         Top             =   1940
         Width           =   370
         VariousPropertyBits=   671105051
         MaxLength       =   1
         Size            =   "656;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM05 
         Height          =   300
         Left            =   1440
         TabIndex        =   30
         Top             =   2870
         Width           =   7280
         VariousPropertyBits=   671105051
         MaxLength       =   40
         Size            =   "12832;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM07 
         Height          =   300
         Left            =   1440
         TabIndex        =   32
         Top             =   3160
         Width           =   7280
         VariousPropertyBits=   671105051
         MaxLength       =   40
         Size            =   "12832;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM06 
         Height          =   300
         Left            =   1440
         TabIndex        =   31
         Top             =   3450
         Width           =   7280
         VariousPropertyBits=   671105051
         MaxLength       =   60
         Size            =   "12832;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM23_2 
         Height          =   290
         Left            =   1980
         TabIndex        =   116
         TabStop         =   0   'False
         Top             =   3740
         Width           =   2180
         VariousPropertyBits=   671105055
         ForeColor       =   -2147483641
         Size            =   "3836;503"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM23 
         Height          =   300
         Left            =   960
         TabIndex        =   33
         Top             =   3720
         Width           =   980
         VariousPropertyBits=   671105051
         MaxLength       =   9
         Size            =   "1720;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM24 
         Height          =   300
         Left            =   -73440
         TabIndex        =   42
         Top             =   1012
         Width           =   7152
         VariousPropertyBits=   671105051
         MaxLength       =   80
         Size            =   "12615;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM26 
         Height          =   300
         Left            =   -73440
         TabIndex        =   44
         Top             =   1664
         Width           =   7152
         VariousPropertyBits=   671105051
         MaxLength       =   80
         Size            =   "12615;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM25 
         Height          =   300
         Left            =   -73440
         TabIndex        =   43
         Top             =   1338
         Width           =   7152
         VariousPropertyBits=   671105051
         MaxLength       =   185
         Size            =   "12615;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM34 
         Height          =   300
         Left            =   -73440
         TabIndex        =   49
         Top             =   2316
         Width           =   2835
         VariousPropertyBits=   671105051
         MaxLength       =   50
         Size            =   "5001;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM09 
         Height          =   300
         Left            =   -73470
         TabIndex        =   52
         Top             =   2970
         Width           =   7155
         VariousPropertyBits=   671105051
         MaxLength       =   395
         Size            =   "12621;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM58 
         Height          =   650
         Left            =   -73470
         TabIndex        =   53
         Top             =   3300
         Width           =   7155
         VariousPropertyBits=   -1466941413
         MaxLength       =   2000
         ScrollBars      =   2
         Size            =   "12621;1147"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCP64 
         Height          =   650
         Left            =   -73470
         TabIndex        =   54
         Top             =   4260
         Width           =   7155
         VariousPropertyBits=   -1466941413
         MaxLength       =   2000
         ScrollBars      =   2
         Size            =   "12621;1147"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textSP58 
         Height          =   300
         Left            =   960
         TabIndex        =   34
         Top             =   4040
         Width           =   980
         VariousPropertyBits=   671105051
         MaxLength       =   9
         Size            =   "1720;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textSP58_2 
         Height          =   290
         Left            =   1980
         TabIndex        =   115
         TabStop         =   0   'False
         Top             =   4050
         Width           =   2180
         VariousPropertyBits=   671105055
         ForeColor       =   -2147483641
         Size            =   "3836;503"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textSP59 
         Height          =   300
         Left            =   5480
         TabIndex        =   35
         Top             =   4040
         Width           =   980
         VariousPropertyBits=   671105051
         MaxLength       =   9
         Size            =   "1720;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textSP59_2 
         Height          =   290
         Left            =   6480
         TabIndex        =   114
         TabStop         =   0   'False
         Top             =   4050
         Width           =   2180
         VariousPropertyBits=   671105055
         ForeColor       =   -2147483641
         Size            =   "3836;503"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCP57 
         Height          =   290
         Left            =   3510
         TabIndex        =   113
         TabStop         =   0   'False
         Top             =   1700
         Width           =   1040
         VariousPropertyBits=   671105051
         ForeColor       =   -2147483641
         Size            =   "1826;503"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM04 
         Height          =   300
         Left            =   3060
         TabIndex        =   15
         Top             =   1410
         Width           =   490
         VariousPropertyBits=   671105051
         MaxLength       =   2
         Size            =   "868;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM03 
         Height          =   300
         Left            =   2820
         TabIndex        =   14
         Top             =   1420
         Width           =   260
         VariousPropertyBits=   671105051
         MaxLength       =   1
         Size            =   "450;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM02_2 
         Height          =   300
         Left            =   2580
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   1420
         Visible         =   0   'False
         Width           =   260
         VariousPropertyBits=   671105051
         MaxLength       =   1
         Size            =   "450;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCP05 
         Height          =   300
         Left            =   1140
         TabIndex        =   17
         Top             =   1700
         Width           =   1220
         VariousPropertyBits=   671105051
         MaxLength       =   7
         Size            =   "2143;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM02 
         Height          =   300
         Left            =   1740
         TabIndex        =   12
         Top             =   1420
         Width           =   1100
         VariousPropertyBits=   671105051
         MaxLength       =   6
         Size            =   "1931;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label70 
         Caption         =   "特殊商標 :"
         Height          =   250
         Left            =   2540
         TabIndex        =   225
         Top             =   1190
         Width           =   850
      End
      Begin VB.Label Label60 
         Caption         =   "徵求同意書對象(日) :"
         Height          =   255
         Left            =   -74940
         TabIndex        =   191
         Top             =   4903
         Width           =   1695
      End
      Begin VB.Label Label59 
         Caption         =   "徵求同意書對象(英) :"
         Height          =   255
         Left            =   -74940
         TabIndex        =   190
         Top             =   4572
         Width           =   1695
      End
      Begin VB.Label Label58 
         Caption         =   "徵求同意書對象(中) :"
         Height          =   255
         Left            =   -74940
         TabIndex        =   189
         Top             =   4248
         Width           =   1695
      End
      Begin VB.Label lblTM130 
         AutoSize        =   -1  'True
         Caption         =   "特殊出名公司:          (J:智權公司 空白:系統預設)"
         Height          =   180
         Left            =   -71430
         TabIndex        =   188
         Top             =   2692
         Width           =   3690
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "案件備註欄：不可銷卷案請加註 ""不銷卷"" 字樣！  與他案合併計算結餘請註明""與某案號合併計算結餘""！"
         ForeColor       =   &H000000FF&
         Height          =   180
         Index           =   115
         Left            =   -74730
         TabIndex        =   187
         Top             =   5160
         Width           =   8220
      End
      Begin VB.Label Label57 
         Caption         =   "申請地址5(英) :"
         Height          =   255
         Left            =   -74940
         TabIndex        =   186
         Top             =   3600
         Width           =   1215
      End
      Begin VB.Label Label56 
         Caption         =   "申請地址5(日) :"
         Height          =   255
         Left            =   -74940
         TabIndex        =   185
         Top             =   3924
         Width           =   1335
      End
      Begin VB.Label Label54 
         Caption         =   "申請地址5(中) :"
         Height          =   255
         Left            =   -74940
         TabIndex        =   184
         Top             =   3276
         Width           =   1335
      End
      Begin VB.Label Label53 
         Caption         =   "申請地址4(英) :"
         Height          =   255
         Left            =   -74940
         TabIndex        =   183
         Top             =   2628
         Width           =   1215
      End
      Begin VB.Label Label52 
         Caption         =   "申請地址4(日) :"
         Height          =   255
         Left            =   -74940
         TabIndex        =   182
         Top             =   2952
         Width           =   1335
      End
      Begin VB.Label Label51 
         Caption         =   "申請地址4(中) :"
         Height          =   255
         Left            =   -74940
         TabIndex        =   181
         Top             =   2304
         Width           =   1335
      End
      Begin VB.Label Label50 
         Caption         =   "申請地址3(英) :"
         Height          =   255
         Left            =   -74940
         TabIndex        =   180
         Top             =   1656
         Width           =   1215
      End
      Begin VB.Label Label49 
         Caption         =   "申請地址3(日) :"
         Height          =   255
         Left            =   -74940
         TabIndex        =   179
         Top             =   1980
         Width           =   1335
      End
      Begin VB.Label Label48 
         Caption         =   "申請地址3(中) :"
         Height          =   255
         Left            =   -74940
         TabIndex        =   178
         Top             =   1332
         Width           =   1335
      End
      Begin VB.Label Label47 
         Caption         =   "申請地址2(英) :"
         Height          =   255
         Left            =   -74940
         TabIndex        =   177
         Top             =   684
         Width           =   1215
      End
      Begin VB.Label Label46 
         Caption         =   "申請地址2(日) :"
         Height          =   255
         Left            =   -74940
         TabIndex        =   176
         Top             =   1008
         Width           =   1335
      End
      Begin VB.Label Label45 
         Caption         =   "申請地址2(中) :"
         Height          =   255
         Left            =   -74940
         TabIndex        =   175
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label44 
         AutoSize        =   -1  'True
         Caption         =   "申請人5 :"
         Height          =   180
         Left            =   4470
         TabIndex        =   174
         Top             =   4440
         Width           =   720
      End
      Begin VB.Label Label43 
         AutoSize        =   -1  'True
         Caption         =   "申請人4 :"
         Height          =   180
         Left            =   120
         TabIndex        =   173
         Top             =   4410
         Width           =   720
      End
      Begin VB.Label lblDivCase 
         AutoSize        =   -1  'True
         Caption         =   "分割母案本所案號:"
         Height          =   180
         Left            =   4440
         TabIndex        =   170
         Top             =   2640
         Visible         =   0   'False
         Width           =   1610
      End
      Begin VB.Label Label42 
         AutoSize        =   -1  'True
         Caption         =   "案件名稱 :"
         Height          =   180
         Left            =   120
         TabIndex        =   168
         Top             =   2890
         Width           =   810
      End
      Begin VB.Label Label11 
         Caption         =   "案件日文名稱 :"
         Height          =   260
         Left            =   120
         TabIndex        =   138
         Top             =   3440
         Width           =   1220
      End
      Begin VB.Label Label12 
         Caption         =   "案件英文名稱 :"
         Height          =   260
         Left            =   120
         TabIndex        =   137
         Top             =   3150
         Width           =   1220
      End
      Begin VB.Line Line1 
         X1              =   -73170
         X2              =   -71100
         Y1              =   2130
         Y2              =   2130
      End
      Begin VB.Label Label39 
         AutoSize        =   -1  'True
         Caption         =   "申請人國籍 :"
         Height          =   180
         Left            =   4470
         TabIndex        =   165
         Top             =   3780
         Width           =   990
      End
      Begin VB.Label Label38 
         Caption         =   "點數 :"
         Height          =   260
         Left            =   120
         TabIndex        =   163
         Top             =   2600
         Width           =   620
      End
      Begin VB.Label Label28 
         Caption         =   "規費 :"
         Height          =   260
         Left            =   3030
         TabIndex        =   162
         Top             =   2330
         Width           =   620
      End
      Begin VB.Label Label27 
         Caption         =   "費用 :"
         Height          =   260
         Left            =   120
         TabIndex        =   161
         Top             =   2320
         Width           =   620
      End
      Begin VB.Label Label26 
         Caption         =   "承辦期限 :"
         Height          =   252
         Left            =   5016
         TabIndex        =   160
         Top             =   288
         Width           =   852
      End
      Begin VB.Label Label21 
         Caption         =   "查名本所案號 :"
         Height          =   255
         Left            =   -74880
         TabIndex        =   159
         Top             =   2003
         Width           =   1335
      End
      Begin VB.Label Label22 
         Caption         =   "彼所案號："
         Height          =   252
         Left            =   -74880
         TabIndex        =   158
         Top             =   700
         Width           =   972
      End
      Begin VB.Label Label55 
         Caption         =   "FC代理人 :"
         Height          =   252
         Left            =   -70560
         TabIndex        =   157
         Top             =   700
         Width           =   852
      End
      Begin VB.Label Label10 
         Caption         =   "相關總收文號 :"
         Height          =   255
         Left            =   -74880
         TabIndex        =   155
         Top             =   373
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "承辦人 :"
         Height          =   255
         Left            =   120
         TabIndex        =   152
         Top             =   293
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "案件性質 :"
         Height          =   260
         Left            =   120
         TabIndex        =   151
         Top             =   600
         Width           =   980
      End
      Begin VB.Label Label25 
         Caption         =   "法定期限 :"
         Height          =   250
         Left            =   5020
         TabIndex        =   150
         Top             =   870
         Width           =   970
      End
      Begin VB.Label Label5 
         Caption         =   "本所期限 :"
         Height          =   260
         Left            =   120
         TabIndex        =   149
         Top             =   870
         Width           =   980
      End
      Begin VB.Label Label1 
         Caption         =   "申請國家 :"
         Height          =   250
         Index           =   8
         Left            =   5020
         TabIndex        =   148
         Top             =   600
         Width           =   970
      End
      Begin VB.Label Label1 
         Caption         =   "智權人員 :"
         Height          =   250
         Index           =   1
         Left            =   5020
         TabIndex        =   147
         Top             =   1170
         Width           =   970
      End
      Begin VB.Label Label6 
         Caption         =   "商標種類 :"
         Height          =   260
         Left            =   120
         TabIndex        =   146
         Top             =   1170
         Width           =   980
      End
      Begin VB.Label Label7 
         Caption         =   "卷宗性質 :"
         Height          =   252
         Left            =   5016
         TabIndex        =   145
         Top             =   1440
         Width           =   972
      End
      Begin VB.Label Label8 
         Caption         =   "是否算案件數 :"
         Height          =   260
         Left            =   120
         TabIndex        =   144
         Top             =   2020
         Width           =   1220
      End
      Begin VB.Label Label9 
         Caption         =   "(1:申請 2:異議 3:評定 4:廢止)"
         Height          =   250
         Left            =   6460
         TabIndex        =   143
         Top             =   1440
         Width           =   2290
      End
      Begin VB.Label Label23 
         Caption         =   "(N:不算)"
         Height          =   260
         Left            =   1920
         TabIndex        =   142
         Top             =   2020
         Width           =   740
      End
      Begin VB.Label Label15 
         Caption         =   "是否取消閉卷 :"
         Height          =   250
         Left            =   4900
         TabIndex        =   141
         Top             =   1980
         Width           =   1210
      End
      Begin VB.Label Label16 
         Caption         =   "(Y:取消)"
         Height          =   250
         Left            =   6810
         TabIndex        =   140
         Top             =   1990
         Width           =   1090
      End
      Begin VB.Label Label13 
         Caption         =   "案件中文名稱 :"
         Height          =   260
         Left            =   120
         TabIndex        =   139
         Top             =   2890
         Width           =   1220
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "申請人1:"
         Height          =   180
         Left            =   120
         TabIndex        =   136
         Top             =   3780
         Width           =   680
      End
      Begin VB.Label Label40 
         Caption         =   "本案期限："
         Height          =   260
         Left            =   60
         TabIndex        =   135
         Top             =   4680
         Width           =   980
      End
      Begin VB.Label Label17 
         Caption         =   "申請地址1(中) :"
         Height          =   252
         Left            =   -74880
         TabIndex        =   134
         Top             =   1026
         Width           =   1332
      End
      Begin VB.Label Label18 
         Caption         =   "申請地址1(日) :"
         Height          =   252
         Left            =   -74880
         TabIndex        =   133
         Top             =   1678
         Width           =   1332
      End
      Begin VB.Label Label19 
         Caption         =   "申請地址1(英) :"
         Height          =   252
         Left            =   -74880
         TabIndex        =   132
         Top             =   1352
         Width           =   1212
      End
      Begin VB.Label Label20 
         Caption         =   "分所案號 :"
         Height          =   252
         Left            =   -74880
         TabIndex        =   131
         Top             =   2330
         Width           =   1092
      End
      Begin VB.Label Label29 
         Caption         =   "商品類別 :"
         Height          =   252
         Left            =   -74880
         TabIndex        =   130
         Top             =   2986
         Width           =   1212
      End
      Begin VB.Label Label30 
         Caption         =   "優先權資料 :"
         Height          =   252
         Index           =   0
         Left            =   -74880
         TabIndex        =   129
         Top             =   2656
         Width           =   1332
      End
      Begin VB.Label Label31 
         Caption         =   "案件備註 :"
         Height          =   255
         Left            =   -74880
         TabIndex        =   128
         Top             =   3300
         Width           =   975
      End
      Begin VB.Label Label32 
         Caption         =   "進度備註 :"
         Height          =   255
         Left            =   -74880
         TabIndex        =   127
         Top             =   4260
         Width           =   975
      End
      Begin VB.Label Label33 
         AutoSize        =   -1  'True
         Caption         =   "申請人2 :"
         Height          =   180
         Left            =   120
         TabIndex        =   126
         Top             =   4100
         Width           =   720
      End
      Begin VB.Label Label34 
         AutoSize        =   -1  'True
         Caption         =   "申請人3 :"
         Height          =   180
         Index           =   0
         Left            =   4470
         TabIndex        =   125
         Top             =   4100
         Width           =   720
      End
      Begin VB.Label Label35 
         Caption         =   "取消收文日 :"
         Height          =   260
         Left            =   2430
         TabIndex        =   124
         Top             =   1720
         Width           =   1100
      End
      Begin VB.Label Label37 
         Caption         =   "收文日 :"
         Height          =   260
         Left            =   120
         TabIndex        =   122
         Top             =   1720
         Width           =   860
      End
      Begin VB.Label Label36 
         Caption         =   "轉本所案號 :"
         Height          =   260
         Left            =   120
         TabIndex        =   123
         Top             =   1450
         Width           =   1100
      End
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "下一筆(&N)"
      Height          =   350
      Left            =   3105
      TabIndex        =   104
      Top             =   0
      Width           =   955
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   350
      Left            =   8325
      TabIndex        =   109
      Top             =   0
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   350
      Left            =   6360
      TabIndex        =   107
      Top             =   0
      Width           =   850
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "回前畫面(&U)"
      Height          =   350
      Left            =   7215
      TabIndex        =   108
      Top             =   0
      Width           =   1100
   End
   Begin VB.CommandButton cmdRelate 
      Caption         =   "相關卷號(&F)"
      Height          =   350
      Left            =   5280
      TabIndex        =   106
      Top             =   0
      Width           =   1100
   End
   Begin VB.CommandButton cmdCaseProgress 
      Caption         =   "案件進度(&C)"
      Height          =   350
      Left            =   4080
      TabIndex        =   105
      Top             =   0
      Width           =   1200
   End
   Begin MSForms.TextBox textTM29_2 
      Height          =   285
      Left            =   6780
      TabIndex        =   164
      TabStop         =   0   'False
      Top             =   345
      Width           =   2295
      VariousPropertyBits=   671105051
      ForeColor       =   255
      MaxLength       =   20
      Size            =   "4043;494"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textCP09 
      Height          =   285
      Left            =   1020
      TabIndex        =   111
      TabStop         =   0   'False
      Top             =   345
      Width           =   2535
      VariousPropertyBits=   671105051
      MaxLength       =   20
      Size            =   "4466;494"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textTMKey 
      Height          =   285
      Left            =   5100
      TabIndex        =   110
      TabStop         =   0   'False
      Top             =   345
      Width           =   1575
      VariousPropertyBits=   671105051
      Size            =   "2773;494"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label41 
      AutoSize        =   -1  'True
      Caption         =   "S商品類別輸在                   ""案件備註""欄!!!"
      ForeColor       =   &H000000FF&
      Height          =   360
      Index           =   1
      Left            =   120
      TabIndex        =   167
      Top             =   0
      Visible         =   0   'False
      Width           =   2025
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "收文號 :"
      Height          =   180
      Index           =   0
      Left            =   270
      TabIndex        =   154
      Top             =   405
      Width           =   630
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "本所案號 :"
      Height          =   180
      Left            =   4140
      TabIndex        =   153
      Top             =   405
      Width           =   810
   End
End
Attribute VB_Name = "frm030201_02"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2023/10/17 Form2.0已修改 ; grdList從MSFlexGrid改為MSHFlexGrid
'Memo by Lydia 2021/08/30 改成Form2.0 ; 所有TextBox(除了txtDivCaseNo)、lblCU151、lblCU147; grdList改字型=新細明體-ExtB
'Memo By Sindy 2012/12/4 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/29 員工編號欄已修改
'Memo By Sindy 2010/8/11 日期欄已修改
Option Explicit

Dim m_TM01 As String
Dim m_TM02 As String
Dim m_TM03 As String
Dim m_TM04 As String

Dim m_CPKeyList() As String
Dim m_CPKeyCount As Integer
' 收文號
Dim m_CP09 As String
' 案件性質
Dim m_CP10 As String
' 國家代碼
Dim m_TM10 As String
' 卷宗性質
Dim m_TM28 As String
' 是否閉卷
Dim m_TM29 As String
'91.12.22 ADD BY SONIA
' 是否新案件
Dim m_CP31 As String
'910626 Sieg 601
' 收據編號
Dim m_CP60 As String
' 申請人
Dim m_TM23 As String
'add by nickc 2007/01/15
Dim m_TM78 As String
Dim m_TM79 As String
Dim m_TM80 As String
Dim m_TM81 As String
Dim m_TM44 As String 'Added by Lydia 2024/06/13
Dim m_CP65 As String 'Add By Sindy 2010/8/6
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
'
Dim m_CurrSel As Integer
' 優先權畫面所使用的變數
Dim m_Pa(1 To 4) As String '本所案號
Dim m_Priority(1 To 6) As String
'Add By Cheng 2002/06/12
Dim m_strCP06 As String '原本所期限
Dim m_strCP07 As String '原法定期限
Dim m_TM22 As String '專用期止日
'Add By Cheng 2002/08/22
'Mark by Lydia 2024/06/13
'Dim m_strCust1 As String '申請人1
'Dim m_strCust2 As String '申請人2
'Dim m_strCust3 As String '申請人3
''add by nickc 2007/01/15
'Dim m_strCust4 As String
'Dim m_strCust5 As String
'end --- 'Mark by Lydia 2024/06/13

'add by nickc 2005/03/17 加乘註記
Dim m_CP98 As String
Dim m_CP101 As String
Dim m_CP104 As String
Dim m_CP30 As String 'Add by Morgan 2011/4/22
Dim m_CP27 As String 'Add by Sindy 2012/6/1
Dim m_CP16 As String 'add by sonia 2013/10/31
Dim m_CP31isYGetCP05 As String 'Add By Sindy 2014/1/29
Dim m_textTM44_FA03 As String 'Add By Sindy 2014/2/14
Dim m_CP14 As String 'Added by Lydia 2020/05/20
'Added by Lydia 2020/05/20 法律所案源收文
Dim m_LOS15 As String '案源單號
Dim m_LOS01 As String '案源總收文號
Dim m_LOS02 As String 'Added by Lydia 2020/06/09 案源案件類型
Dim m_LOS07 As String '放棄日期
Public m_CP143 As String, m_CP36 As String, m_CP21 As String 'Add By Sindy 2020/10/20
Dim m_CP149 As String 'Added by Lydia 2022/03/09 分案日
Dim m_Txt As Object 'Add by Sindy 2022/3/16
Dim strCP122 As String 'Add by Amy 2022/11/17
Dim strPTM As String, strSPT As String 'Added by Lydia 2023/11/16 暫存商標種類及特殊商標的Combo.ItemData
Dim m_CP141 As String 'Add By Sindy 2024/1/22
Dim m_strRefText As String 'Modify By Sindy 2024/5/3
Dim strMsgCloseCancel As String 'Added by Lydia 2025/06/30 取消閉卷時，若下一程序有未過期且已上N之延展102、使用宣誓105期限，都自動取消NP06、NP11、NP12，並彈訊息提醒已還原XXX期限。

Private Sub cmdCancel_Click()
   Unload Me
   frm030201_01.Show
End Sub

Private Sub cmdCaseProgress_Click()
   frm030201_03.SetData 0, m_TM01, True
   frm030201_03.SetData 1, m_TM02, False
   frm030201_03.SetData 2, m_TM03, False
   frm030201_03.SetData 3, m_TM04, False
   frm030201_03.SetData 4, m_CP09, False
   frm030201_03.SetParent Me
   Me.Hide
   frm030201_03.Show
   frm030201_03.QueryData
End Sub

Private Sub cmdExit_Click()
   Unload Me
   Unload frm030201_01
End Sub

'Add by Amy 2022/11/10 檢視接洽單
Private Sub cmdFile_Click()
    frm090801_Q.SetParent Me
    frm090801_Q.m_blnCallPrint = True
    frm090801_Q.Text5 = txtF0301
    Call frm090801_Q.cmdok_Click(4)
    frm090801_Q.Show 'Add by Amy 2022/11/17
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
      frm030201_01.Show
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
   Dim bolHadPoMsg As Boolean
      
    'Added by Lydia 2021/09/02 檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
    If PUB_ChkUniText(Me, , True, "TextBox") = False Then
         Exit Sub
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
         strExc(9) = m_TM23
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
      
      'Add By Cheng 2002/08/23
      If Me.textTM01.Text <> "" And Me.textTM02.Text <> "" Then
         MsgBox "本次只更新轉本所案號的資料，不更新其他欄位的修改!!!", vbCritical
      'Added by Lydia 2023/12/14
      Else
        '檢查智財協作在分案時若未建立相關案號(caserelation1)時則跳提醒程序人員，但可選擇輸或不輸 !
         'Modified by Lydia 2023/12/15 PS及CPS之智財協作967，TT及S之智財協作737，L之智財協作7601，(也可用案件性質中文判斷)在分案時若未建立相關案號且為ACS且為TIPS的案件時，提醒文字：「案件性質為智財協作，請先依接洽單輸入相關卷號資料」。
'         If m_TM01 = "S" And textCP10 = "737" Then
'            If PUB_IfCaseRelation1Exists(m_TM01, m_TM02, m_TM03, m_TM04) = False Then
'               If MsgBox("案件性質為" & textCP10_2 & "，請確認接洽單是否有相關案號，是否補輸入？", vbExclamation + vbYesNo + vbDefaultButton1) = vbYes Then
'                  Exit Sub
'               End If
'            End If
'         End If
         If m_TM01 = "S" And InStr(textCP10_2, "智財協作") > 0 Then
            If PUB_ChkACSforTIPS(m_TM01 & m_TM02 & m_TM03 & m_TM04, , True) = False Then
               MsgBox "案件性質為" & textCP10_2 & "，請先依接洽單輸入相關卷號資料", vbExclamation
               Exit Sub
            End If
         End If
         'end 2023/12/15
      'end 2023/12/14
      End If
      
      'Add By Sindy 2014/2/12 若財務處已開立收據,且收據的公司別與案件的特殊出名公司不符時,
      '顯示訊息,讓使用者可選擇是否修改,預設在"是"
      strSql = "select cp60,a0k01,a0k11 from caseprogress,acc0k0" & _
               " where cp09='" & m_CP09 & "' and cp60 is not null and cp60=a0k01(+)"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strSql)
      If intI = 1 Then
         If "" & RsTemp.Fields("a0k01") <> "" Then
            If (textTM130 = "J" Or "" & RsTemp.Fields("a0k11") = "J") And _
               textTM130 <> "" & RsTemp.Fields("a0k11") And _
               textTM130.Visible = True Then
               'MODIFY BY SONIA 2014/5/30 阿蓮要改訊息
               'If MsgBox("財務處開立的收據公司與分案之特殊出名公司不符,是否修改特殊出名公司？", vbExclamation + vbYesNo + vbDefaultButton1, "重要訊息！") = vbYes Then
               'MODIFY BY SONIA 2014/7/9 阿蓮要改訊息 CFT-16783
               'If MsgBox("財務處開立的收據公司與分案之特殊出名公司不符,請確認是否為特殊出名公司？", vbExclamation + vbYesNo + vbDefaultButton1, "重要訊息！") = vbNo Then
               If MsgBox("財務處開立的收據公司與分案之特殊出名公司不符,是否要修改？", vbExclamation + vbYesNo + vbDefaultButton1, "重要訊息！") = vbYes Then
                  'add by sonia 2014/7/9
                  SSTab1.Tab = 1
                  textTM130.SetFocus
                  'end 2014/7/9
                  Exit Sub
               End If
            End If
         End If
      End If
      '2014/2/12 END
      
      'Added by Lydia 2020/01/21 所有系統別分案，若該收文號存在於下一程序(NP24)，若修改案件性質則彈訊息
      If textCP10 <> m_CP10 Then
          If Pub_CheckNP24Exists(textCP09.Text) = True Then
          End If
      End If
      'end 2020/01/21
      
      'Add By Sindy 2020/10/20
      If textCP10 = "210" Then '陳述意見書
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
     If strSrvDate(1) >= 法律所案源收文啟用日 And m_TM01 = "FCT" Then
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
      'Modify By Sindy 2021/9/8 查名併案,不檢查接洽單PDF檔
      'Modify By Sindy 2021/9/13 查名併案,要多增加用收文日判斷要不要檢查接洽單PDF檔:倘併入之查名本所案號收文日為109年前，即不管制接洽單電子檔
      'Modify By Sindy 2022/3/33 增加判斷沒有輸入承辦人時,先不檢查接洽單PDF檔
      If (textTM01 <> "" And textTM02 <> "") Or _
         ((textCP09_S <> "" And textCP09_S1 <> "") And DBDATE(textCP05) <= 20201231) Or _
         Trim(textCP14) = "" Then
         '不需檢查接洽單PDF檔
      Else
      '2021/3/9 END
         'Add By Sindy 2020/12/17
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
         '2020/12/17
      End If
      
      OnUpdateField
      'edit by nick 2004/11/03
      'OnSaveData
      If OnSaveData = False Then MsgBox "存檔失敗，請洽系統管理員 !", vbCritical: Screen.MousePointer = vbDefault: Exit Sub
      
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
      If (m_TM01 = "T" Or m_TM01 = "TF" Or m_TM01 = "FCT" Or m_TM01 = "CFT") And InStr("3,4", m_TM28) > 0 And textTM28 = "1" Then
         'Modify By Sindy 2024/5/3 + m_strRefText
         ShowMaintainForm m_CP09, "N", "分案", , m_strRefText
         MsgBox "請輸入專用期限！", vbInformation
      End If
      'end 2023/01/18
      'Added by Lydia 2025/06/30 取消閉卷時，若下一程序有未過期且已上N之延展102、使用宣誓105期限，都自動取消NP06、NP11、NP12，並彈訊息提醒已還原XXX期限。
      If strMsgCloseCancel <> "" Then
         MsgBox "已還原「" & strMsgCloseCancel & "」期限", vbInformation, "取消閉卷"
      End If
      
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
         frm030201_01.Show
      End If
   End If
End Sub

Private Sub cmdPriority_Click()
Dim strPCase(1 To 4) As String 'Added by Lydia 2023/02/09

   ' 修改優先權資料
   'Modify by Amy 2014/04/17 +, m_Priority(4), m_Priority(5)
   'Modify by Sindy 2017/10/12 +, m_Priority(6)
   'Modify by Sindy 2019/1/23 + m_TM01 & m_TM02 & m_TM03 & m_TM04
   'Added by Lydia 2023/02/09 FCT分割子案(案件性質為分案且CP31='Y')
   If m_TM01 = "FCT" And m_CP10 = "308" And m_CP31 = "Y" And m_Priority(1) & m_Priority(2) & m_Priority(3) = "" And txtDivCaseNo(0).Text <> "" And txtDivCaseNo(1).Text & txtDivCaseNo(2).Text <> "" Then
      '1.若該分割案已有優先權資料則直接帶出；
      '2.若該分割案無優先權資料，則以分案畫面之分割母案本所案號抓其優先權資料
      strPCase(1) = txtDivCaseNo(0).Text
      strPCase(2) = txtDivCaseNo(1).Text & txtDivCaseNo(2).Text
      strPCase(3) = txtDivCaseNo(3).Text
      strPCase(4) = txtDivCaseNo(4).Text
      If strPCase(3) = "" Then strPCase(3) = "0"
      If strPCase(4) = "" Then strPCase(4) = "00"
      ClsPDReadPriority strPCase, m_Priority(1), m_Priority(2), m_Priority(3), m_Priority(4), m_Priority(5), m_Priority(6), textTM09.Text
   End If
   ModifyPriority m_Priority(1), m_Priority(2), m_Priority(3), , , m_TM01 & m_TM02 & m_TM03 & m_TM04, , , m_Priority(4), m_Priority(5), m_Priority(6)
End Sub

Private Sub cmdRelate_Click()
   Where1103ComeFrom Me, m_TM01, m_TM02, m_TM03, m_TM04
End Sub

Private Sub Form_Load()
   ' 設定控制項的背景顏色
   textTMKey.BackColor = &H8000000F
   textTM08_2.BackColor = &H8000000F
   textTM72_2.BackColor = &H8000000F 'Add By Sindy 2019/1/4
   textTM10_2.BackColor = &H8000000F
   textTM23_2.BackColor = &H8000000F
   'add by nickc 2007/02/01
   textTM80_2.BackColor = &H8000000F
   textTM81_2.BackColor = &H8000000F
   
   textTM23_3.BackColor = &H8000000F
   textTM29_2.BackColor = &H8000000F
   textTM44_2.BackColor = &H8000000F
   textSP58_2.BackColor = &H8000000F
   textSP59_2.BackColor = &H8000000F
   
   textCP09.BackColor = &H8000000F
   textCP10_2.BackColor = &H8000000F
   textCP13_2.BackColor = &H8000000F
   textCP14_2.BackColor = &H8000000F
   textCP57.BackColor = &H8000000F
   
   SSTab1.Tab = 0
   MoveFormToCenter Me
   SSTab1.TabVisible(3) = True
   SSTab1.TabVisible(4) = True
   
   'Added by Lydia 2020/05/20 法律所案源收文
   FraLOS.Visible = False
   FraLOS.BackColor = &H8000000F
   txtLOSagree.Text = ""
   'end 2020/05/20
   'Add By Sindy 2024/1/22
   Frame5.BorderStyle = 0
   Frame6.BorderStyle = 0
   '2024/1/22 END
   
   'Add By Sindy 2023/12/11
   If strSrvDate(1) < 指定日期啟用日 Then
      Label142.Visible = False
      Frame5.Visible = False
      Frame6.Visible = False
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
   textTM131 = Empty 'Add By Sindy 2015/7/14
   textTM06 = Empty
   textTM07 = Empty
   textTM08 = Empty
   textTM08_2 = Empty
   'Add By Sindy 2019/1/4
   textTM72 = Empty
   textTM72_2 = Empty
   '2019/1/4 END
   textTM09 = Empty
   textTM10 = Empty
   textTM10_2 = Empty
   textTM23 = Empty
   textTM23.Tag = "" 'Add By Sindy 2014/4/9
   textTM23_2 = Empty
   'Add By Cheng 2002/11/12
   textTM23_3 = Empty
   textTM24 = Empty
   textTM25 = Empty
   textTM26 = Empty
   textTM28 = Empty
   textTM29 = Empty
   textTM34 = Empty
   textTM35 = Empty 'Add By Sindy 2014/5/26
   textTM38 = Empty 'Add By Sindy 2014/2/14
   textTM39 = Empty 'Add By Sindy 2014/2/14
   textTM40 = Empty 'Add By Sindy 2014/2/14
   textTM41 = Empty 'Add By Sindy 2014/2/14
   textTM42 = Empty 'Add By Sindy 2014/2/14
   textTM43 = Empty 'Add By Sindy 2014/2/14
   textTM44 = Empty
   textTM44.Tag = "" 'Add By Sindy 2014/4/9
   textTM44_2 = Empty
   textTM45 = Empty
   textTM56_1 = Empty: textTM56_2 = Empty 'Add By Sindy 2014/2/14
   textTM58 = Empty
   textTM69_1 = Empty: textTM69_2 = Empty 'Add By Sindy 2014/2/14
   textTM46 = Empty: textTM127 = Empty 'Add By Sindy 2020/3/4
   textTM121 = Empty 'Add By Sindy 2022/3/16
   'add by nickc 2007/01/15
   textTM80 = Empty
   textTM80_2 = Empty
   textTM81 = Empty
   textTM81_2 = Empty
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
   textTM130 = Empty 'Add By Sindy 2013/12/16
   
   textSP58 = Empty
   textSP58_2 = Empty
   textSP59 = Empty
   textSP59_2 = Empty
   
   textCP05 = Empty
   textCP06 = Empty
   textCP07 = Empty
   textCP09 = Empty
   textCP09_S = Empty
   'Add By Cheng 2002/09/18
   textCP09_S1 = Empty
   textCP09_S2 = Empty
   textCP09_S3 = Empty
   textCP10 = Empty
   textCP10_2 = Empty
   textCP13 = Empty
   textCP13_2 = Empty
   textCP14 = Empty
   textCP14_2 = Empty
   'add by nick 2004/10/05
   textCP16 = Empty
   textCP17 = Empty
   textCP18 = Empty
   textCP26 = Empty
   textCP43 = Empty
   textCP48 = Empty
   textCP48.Tag = Empty 'Add By Sindy 2014/12/16
   'Add By Sindy 2013/12/16
   textCP50 = Empty
   textCP51 = Empty
   textCP52 = Empty
   '2013/12/16 END
   textCP57 = Empty
   textCP64 = Empty
   
   'Add By Sindy 2014/2/14
   'Modified by Lydia 2021/08/30  改成Form2.0 ;; Label30(16)=> lblCU147、Label30(15)=> lblCU151
   textCU147 = Empty
   textCU151 = Empty
   lblCU147.Caption = Empty
   lblCU151.Caption = Empty
   'end 2021/08/30
   textCU149 = Empty 'Add By Sindy 2020/3/4
   textCU146 = Empty
   textCU58 = Empty
   textCU59 = Empty
   textCU60 = Empty
   textCU61 = Empty
   textCU62 = Empty
   textCU63 = Empty
   textFA107 = Empty: textFA107_2 = Empty
   textFA111 = Empty: textFA111_2 = Empty
   textFA109 = Empty 'Add By Sindy 2020/3/4
   textFA106 = Empty
   textFA07 = Empty
   textFA08 = Empty
   textFA09 = Empty
   textFA52 = Empty
   textFA53 = Empty
   textFA54 = Empty
   '2014/2/14 END
   
   'Add By Sindy 2022/3/16
   For Each m_Txt In txtFA
      m_Txt = Empty
   Next
   For Each m_Txt In txtCU
      m_Txt = Empty
   Next
   '2022/3/16 END
   
   m_strCountry = Empty
   txtF0301 = Empty 'Add by Amy 2022/11/10
   
   'Add By Sindy 2023/4/27
   textCP06.Enabled = True
   textCP07.Enabled = True
   '2023/4/27 END
   
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
'      If IsNull(rsTmp.Fields("TM05")) = False Then
'         textTM05 = rsTmp.Fields("TM05")
'      End If
'      SetTMSPFieldOldData "TM05", textTM05, 0
      ' 案件名稱
      If IsNull(rsTmp.Fields("TM05")) = False Then
         textTM05_1 = rsTmp.Fields("TM05")
      End If
      SetTMSPFieldOldData "TM05", textTM05_1, 0
      
      'Add By Sindy 2015/7/14
      '定稿商標名稱
      If IsNull(rsTmp.Fields("TM131")) = False Then
         textTM131 = rsTmp.Fields("TM131")
      End If
      SetTMSPFieldOldData "TM131", textTM131, 0
      '2015/7/14 END
      
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
      'Add By Sindy 2019/1/4
      ' 特殊商標
      If IsNull(rsTmp.Fields("TM72")) = False Then
         textTM72 = rsTmp.Fields("TM72")
         textTM72_Validate False
      End If
      SetTMSPFieldOldData "TM72", textTM72, 0
      '2019/1/4 END
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
      End If
      SetTMSPFieldOldData "TM10", m_TM10, 0
      
      'Added by Lydia 2023/11/16 內外商之分案及商標基本資料維護之商標種類、特殊商標欄位增加下拉功能
      Pub_SetTMcombo "1", cboTM08, textTM08, IIf(m_TM10 <> "000", True, False), strPTM '商標種類
      Pub_SetTMcombo "2", cboTM72, textTM72, IIf(m_TM10 <> "000", True, False), strSPT '特殊商標種類
      'end 2023/11/16
      
      ' 申請人
      If IsNull(rsTmp.Fields("TM23")) = False Then
         m_TM23 = rsTmp.Fields("TM23")
         textTM23 = rsTmp.Fields("TM23")
         'Modify By Sindy 2014/2/14
         'textTM23_Validate False
         'textTM23_2 = GetCustomerName(textTM23, 0)
         'Add By Cheng 2002/11/12
         'textTM23_3 = GetNationName(GetCustomerNation(textTM23), 0)
         '2014/2/14 END
      End If
      SetTMSPFieldOldData "TM23", textTM23, 0
      'Add By Cheng 2002/08/22
      'm_strCust1 = "" & Me.textTM23.Text 'Mark by Lydia 2024/06/13
      'add by nickc 2007/01/15
      'Modify By Sindy 2014/2/14
      textTM23_Validate False
      '2014/2/14 END
      If IsNull(rsTmp.Fields("TM78")) = False Then
         m_TM78 = rsTmp.Fields("TM78")
         textSP58 = rsTmp.Fields("TM78")
         textSP58_2 = GetCustomerName(textSP58, 0)
      End If
      SetTMSPFieldOldData "TM78", textSP58, 0
      
      If IsNull(rsTmp.Fields("TM79")) = False Then
         m_TM79 = rsTmp.Fields("TM79")
         textSP59 = rsTmp.Fields("TM79")
         textSP59_2 = GetCustomerName(textSP59, 0)
      End If
      SetTMSPFieldOldData "TM79", textSP59, 0
      
      If IsNull(rsTmp.Fields("TM80")) = False Then
         m_TM80 = rsTmp.Fields("TM80")
         textTM80 = rsTmp.Fields("TM80")
         textTM80_2 = GetCustomerName(textTM80, 0)
      End If
      SetTMSPFieldOldData "TM80", textTM80, 0
      
      If IsNull(rsTmp.Fields("TM81")) = False Then
         m_TM81 = rsTmp.Fields("TM81") 'Modify by Amy 2024/07/15 bug-原:m_TM80
         textTM81 = rsTmp.Fields("TM81")
         textTM81_2 = GetCustomerName(textTM81, 0)
      End If
      SetTMSPFieldOldData "TM81", textTM81, 0
      'Mark by Lydia 2024/06/13
      'm_strCust2 = "" & Me.textSP58.Text
      'm_strCust3 = "" & Me.textSP59.Text
      'm_strCust4 = "" & Me.textTM80.Text
      'm_strCust5 = "" & Me.textTM81.Text
      'end 2024/06/13
      
      ' 申請地址
      If IsNull(rsTmp.Fields("TM82")) = False Then
         textTM82 = rsTmp.Fields("TM82")
      End If
      SetTMSPFieldOldData "TM82", textTM82, 0
      ' 申請地址
      If IsNull(rsTmp.Fields("TM86")) = False Then
         textTM86 = rsTmp.Fields("TM86")
      End If
      SetTMSPFieldOldData "TM86", textTM86, 0
      ' 申請地址
      If IsNull(rsTmp.Fields("TM90")) = False Then
         textTM90 = rsTmp.Fields("TM90")
      End If
      SetTMSPFieldOldData "TM90", textTM90, 0
      ' 申請地址
      If IsNull(rsTmp.Fields("TM83")) = False Then
         textTM83 = rsTmp.Fields("TM83")
      End If
      SetTMSPFieldOldData "TM83", textTM83, 0
      ' 申請地址
      If IsNull(rsTmp.Fields("TM87")) = False Then
         textTM87 = rsTmp.Fields("TM87")
      End If
      SetTMSPFieldOldData "TM87", textTM87, 0
      ' 申請地址
      If IsNull(rsTmp.Fields("TM91")) = False Then
         textTM91 = rsTmp.Fields("TM91")
      End If
      SetTMSPFieldOldData "TM91", textTM91, 0
      ' 申請地址
      If IsNull(rsTmp.Fields("TM84")) = False Then
         textTM84 = rsTmp.Fields("TM84")
      End If
      SetTMSPFieldOldData "TM84", textTM84, 0
      ' 申請地址
      If IsNull(rsTmp.Fields("TM88")) = False Then
         textTM88 = rsTmp.Fields("TM88")
      End If
      SetTMSPFieldOldData "TM88", textTM88, 0
      ' 申請地址
      If IsNull(rsTmp.Fields("TM92")) = False Then
         textTM92 = rsTmp.Fields("TM92")
      End If
      SetTMSPFieldOldData "TM92", textTM92, 0
      ' 申請地址
      If IsNull(rsTmp.Fields("TM85")) = False Then
         textTM85 = rsTmp.Fields("TM85")
      End If
      SetTMSPFieldOldData "TM85", textTM85, 0
      ' 申請地址
      If IsNull(rsTmp.Fields("TM89")) = False Then
         textTM89 = rsTmp.Fields("TM89")
      End If
      SetTMSPFieldOldData "TM89", textTM89, 0
      ' 申請地址
      If IsNull(rsTmp.Fields("TM93")) = False Then
         textTM93 = rsTmp.Fields("TM93")
      End If
      SetTMSPFieldOldData "TM93", textTM93, 0
      
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
         textTM44 = rsTmp.Fields("TM44")
         textTM44_Validate False
      'add by sonia 2017/12/1 以防更代沒記錄
         textTM44.Enabled = False
      Else
         textTM44.Enabled = True
      'end 2017/12/1
      End If
      SetTMSPFieldOldData "TM44", textTM44, 0
      m_TM44 = "" & rsTmp.Fields("TM44")  'Added by Lydia 2024/06/13
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
      
      textTM130 = "" 'Add by Amy 2016/11/07 因先跑基本檔後會再執行此Function,導致textTM130.Tag又被設一次
      'Add By Sindy 2013/12/16
      If IsNull(rsTmp.Fields("TM130")) = False Then
         textTM130 = rsTmp.Fields("TM130")
      End If
      textTM130.Tag = textTM130
      SetTMSPFieldOldData "TM130", textTM130, 0
      '2013/12/16 END
      'Modify by Amy 2016/08/29
      'Modify by Amy 2017/03/21 +新案才帶 (CFT-013586 舊案不應該帶)
      'Mark by Amy 2017/11/24 個案為空白會預設申請人出名公司,若個案改為空仍會一直預帶(CFP-029915)-秀玲:拿掉
'      If textTM130 = MsgText(601) And m_CP31 = "Y" Then
'        'Add by Amy 2016/08/12 +客戶檔收據公司別
'        textTM130 = GetReceiptCmp(Left(textTM23, 8), Mid(textTM23, 9, 1), m_TM01, textTM10)
'      End If
      'Add By Sindy 2014/2/14 第五頁
      If IsNull(rsTmp.Fields("TM56")) = False Then
         textTM56_1 = rsTmp.Fields("TM56")
      End If
      textTM56_1.Tag = textTM56_1
      textTM56_1_Validate False
      If IsNull(rsTmp.Fields("TM69")) = False Then
         textTM69_1 = rsTmp.Fields("TM69")
      End If
      textTM69_1.Tag = textTM69_1
      textTM69_1_Validate False
      
      'Add By Sindy 2020/3/4
      If IsNull(rsTmp.Fields("TM46")) = False Then
         textTM46 = rsTmp.Fields("TM46")
      End If
      textTM46.Tag = textTM46
      If IsNull(rsTmp.Fields("TM127")) = False Then
         textTM127 = rsTmp.Fields("TM127")
      End If
      textTM127.Tag = textTM127
      '2020/3/4 END
      'Add By Sindy 2022/3/16
      If IsNull(rsTmp.Fields("TM121")) = False Then
         textTM121 = rsTmp.Fields("TM121")
      End If
      textTM121.Tag = textTM121
      '2022/3/16 END
      
      If IsNull(rsTmp.Fields("TM38")) = False Then
         textTM38 = rsTmp.Fields("TM38")
      End If
      textTM38.Tag = textTM38
      If IsNull(rsTmp.Fields("TM39")) = False Then
         textTM39 = rsTmp.Fields("TM39")
      End If
      textTM39.Tag = textTM39
      If IsNull(rsTmp.Fields("TM40")) = False Then
         textTM40 = rsTmp.Fields("TM40")
      End If
      textTM40.Tag = textTM40
      If IsNull(rsTmp.Fields("TM41")) = False Then
         textTM41 = rsTmp.Fields("TM41")
      End If
      textTM41.Tag = textTM41
      If IsNull(rsTmp.Fields("TM42")) = False Then
         textTM42 = rsTmp.Fields("TM42")
      End If
      textTM42.Tag = textTM42
      If IsNull(rsTmp.Fields("TM43")) = False Then
         textTM43 = rsTmp.Fields("TM43")
      End If
      textTM43.Tag = textTM43
      '2014/2/14 END
      
      'Added by Morgan 2022/12/23
      textTM136 = "" & rsTmp.Fields("TM136")
      textTM136.Tag = textTM136
      'end 2022/12/23
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
        Select Case m_TM01
        Case "S"
            ' 案件名稱
            If IsNull(rsTmp.Fields("SP05")) = False Then
               textTM05_1 = rsTmp.Fields("SP05")
            End If
            SetTMSPFieldOldData "SP05", textTM05_1, 0
        Case Else
            ' 案件中文名稱
            If IsNull(rsTmp.Fields("SP05")) = False Then
               textTM05 = rsTmp.Fields("SP05")
            End If
            SetTMSPFieldOldData "SP05", textTM05, 0
        End Select
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
         m_TM23 = rsTmp.Fields("SP08") 'Add by Amy 2024/07/15 bug:S-008305 會一直彈"非新案或新案已發文者，不可修改申請人!!!"
         textTM23 = rsTmp.Fields("SP08")
         'Modify By Sindy 2014/2/14
         'textTM23_Validate False
         'textTM23_2 = GetCustomerName(textTM23, 0)
         'Add By Cheng 2002/11/12
         'textTM23_3 = GetNationName(GetCustomerNation(textTM23), 0)
         '2014/2/14 END
      End If
      SetTMSPFieldOldData "SP08", textTM23, 0
      'Add By Cheng 2002/08/22
      'm_strCust1 = "" & Me.textTM23.Text 'Mark by Lydia 2024/06/13
      'Modify By Sindy 2014/2/14
      textTM23_Validate False
      '2014/2/14 END
      ' 第二申請人及第三申請人
'edit by nickc 2007/01/15
'      If m_TM01 = "CFC" Then
         If IsNull(rsTmp.Fields("SP58")) = False Then
            m_TM78 = rsTmp.Fields("SP58") 'Add by Amy 2024/07/15
            textSP58 = rsTmp.Fields("SP58")
            'Modify By Cheng 2002/09/23
'            textSP58_Validate False
            textSP58_2 = GetCustomerName(textSP58, 0)
         End If
         If IsNull(rsTmp.Fields("SP59")) = False Then
            m_TM79 = rsTmp.Fields("SP59") 'Add by Amy 2024/07/15
            textSP59 = rsTmp.Fields("SP59")
            'Modify By Cheng 2002/09/23
'            textSP59_Validate False
            textSP59_2 = GetCustomerName(textSP59, 0)
         End If
         'Add By Cheng 2002/08/22
         'Mark by Lydia 2024/06/13
         'm_strCust2 = "" & Me.textSP58.Text
         'm_strCust3 = "" & Me.textSP59.Text
         'end 2024/06/13
'      End If
'add by nickc 2007/01/15
         SetTMSPFieldOldData "SP58", textSP58, 0
         SetTMSPFieldOldData "SP59", textSP59, 0
         If IsNull(rsTmp.Fields("SP65")) = False Then
            m_TM80 = rsTmp.Fields("SP65") 'Add by Amy 2024/07/15
            textTM80 = rsTmp.Fields("SP65")
            textTM80_2 = GetCustomerName(textTM80, 0)
         End If
         If IsNull(rsTmp.Fields("SP66")) = False Then
            m_TM81 = rsTmp.Fields("SP66") 'Add by Amy 2024/07/15
            textTM81 = rsTmp.Fields("SP66")
            textTM81_2 = GetCustomerName(textTM81, 0)
         End If
         'Mark by Lydia 2024/06/13
         'm_strCust4 = "" & Me.textTM80.Text
         'm_strCust5 = "" & Me.textTM81.Text
         'end 2024/06/13
         SetTMSPFieldOldData "SP65", textTM80, 0
         SetTMSPFieldOldData "SP66", textTM81, 0
      '商品類別
      If IsNull(rsTmp.Fields("SP73")) = False Then
         textTM09 = rsTmp.Fields("SP73")
      End If
      SetTMSPFieldOldData "SP73", textTM09, 0
      
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
         textTM44 = rsTmp.Fields("SP26")
         textTM44_Validate False
      End If
      SetTMSPFieldOldData "SP26", textTM44, 0
      m_TM44 = "" & rsTmp.Fields("SP26")  'Added by Lydia 2024/06/13
      ' 彼所案號
      If IsNull(rsTmp.Fields("SP27")) = False Then
         textTM45 = rsTmp.Fields("SP27")
      End If
      SetTMSPFieldOldData "SP27", textTM45, 0
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
      'Mark by Amy 2017/11/24 個案為空白會預設申請人出名公司,若個案改為空仍會一直預帶(CFP-029915)-秀玲:拿掉
'      If textTM130 = MsgText(601) Then
'        '客戶檔收據公司別
'        textTM130 = GetReceiptCmp(Left(textTM23, 8), Mid(textTM23, 9, 1), m_TM01, textTM10)
'      End If
      'end 2016/11/07
      
      'Add By Sindy 2014/2/14 第五頁
      'Add by Sindy 2021/5/5
      ' 客戶案件案號
      If IsNull(rsTmp.Fields("SP29")) = False Then
         textTM35 = rsTmp.Fields("SP29")
      End If
      SetTMSPFieldOldData "SP29", textTM35, 0
      '2021/5/5 END
      If IsNull(rsTmp.Fields("SP37")) = False Then
         textTM56_1 = rsTmp.Fields("SP37")
      End If
      textTM56_1.Tag = textTM56_1
      textTM56_1_Validate False
      If IsNull(rsTmp.Fields("SP67")) = False Then
         textTM69_1 = rsTmp.Fields("SP67")
      End If
      textTM69_1.Tag = textTM69_1
      textTM69_1_Validate False
      
      'Add By Sindy 2020/3/4
      If IsNull(rsTmp.Fields("SP33")) = False Then
         textTM46 = rsTmp.Fields("SP33")
      End If
      textTM46.Tag = textTM46
      If IsNull(rsTmp.Fields("SP84")) = False Then
         textTM127 = rsTmp.Fields("SP84")
      End If
      textTM127.Tag = textTM127
      '2020/3/4 END
      'Add By Sindy 2022/3/16
      If IsNull(rsTmp.Fields("SP80")) = False Then
         textTM121 = rsTmp.Fields("SP80")
      End If
      textTM121.Tag = textTM121
      '2022/3/16 END
      
      'Add By Sindy 2014/5/16
      If IsNull(rsTmp.Fields("SP30")) = False Then
         textTM38.MaxLength = 60
         textTM38 = rsTmp.Fields("SP30")
      End If
      textTM38.Tag = textTM38
      '2014/5/16 END
      'Add By Sindy 2021/5/5 聯絡人2
      If IsNull(rsTmp.Fields("SP75")) = False Then
         textTM41.MaxLength = 60
         textTM41 = rsTmp.Fields("SP75")
      End If
      textTM41.Tag = textTM41
      '2021/5/5 END
      '2014/2/14 END
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
   
   strSql = "SELECT * FROM CaseProgress " & _
            "WHERE CP09 = '" & m_CP09 & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
     
      '費用 add by sonia 2013/10/31
      m_CP16 = ""
      If IsNull(rsTmp.Fields("CP16")) = False Then
         m_CP16 = rsTmp.Fields("CP16")
      End If
      '2013/10/31 end
     
     ' 案件性質
      If IsNull(rsTmp.Fields("CP10")) = False Then
         m_CP10 = rsTmp.Fields("CP10")
         textCP10 = rsTmp.Fields("CP10")
         textCP10_Validate False
      End If
      
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
      End If
      SetCPFieldOldData "CP06", textCP06, 1
      'Add By Cheng 2002/06/12
      m_strCP06 = "" & rsTmp.Fields("CP06")
      
      ' 法定期限
      If IsNull(rsTmp.Fields("CP07")) = False Then
         textCP07 = TAIWANDATE(rsTmp.Fields("CP07"))
      End If
      SetCPFieldOldData "CP07", textCP07, 1
      'Add By Cheng 2002/06/12
      m_strCP07 = "" & rsTmp.Fields("CP07")
      ' 業務區
      'Modify By Cheng 2002/11/05
      If "" & rsTmp.Fields("CP12") <> "" Then
          SetCPFieldOldData "CP12", rsTmp.Fields("CP12"), 0
      End If
      
      ' 智權人員
      If IsNull(rsTmp.Fields("CP13")) = False Then
         textCP13 = rsTmp.Fields("CP13")
         textCP13_Validate False
      End If
      SetCPFieldOldData "CP13", textCP13, 0

      ' 承辦人員
      If IsNull(rsTmp.Fields("CP14")) = False Then
         textCP14 = rsTmp.Fields("CP14")
         textCP14_Validate False
      End If
      SetCPFieldOldData "CP14", textCP14, 0
      textCP14.Tag = textCP14.Text 'Add By Sindy 2022/12/26
      m_CP14 = textCP14 'Added by Lydia 2020/05/20
      'Added by Lydia 2021/06/18 CFT案件承辦人若空白時，預設為國家檔之CFT承辦人。
      'Modified by Lydia 2022/03/18 有關S案（申請國為「台灣」除外）之分案，請比照CFT案，當承辦人空白時，預設為國家檔之CFT承辦人
      'If m_TM01 = "CFT" And m_CP14 = "" Then
      'Modified by Lydia 2022/09/21 增加CFC案
      If (m_TM01 = "CFT" Or m_TM01 = "CFC" Or (m_TM01 = "S" And m_TM10 <> "000")) And m_CP14 = "" Then
          'modify by sonia 2023/12/22 加傳本所案號 m_TM01~m_TM04，否則CFC案件會錯誤CFC-000810
          Call GetNA69("", m_TM10, textCP13.Text, strTemp, m_TM01, m_TM02, m_TM03, m_TM04)
          textCP14 = strTemp
          If strTemp <> "" Then textCP14_Validate False
      End If
      'end 2021/06/18
      'Add By Sindy 2022/8/24 先使分案時,一併更新北所分案日
      SetCPFieldOldData "CP157", "" & rsTmp.Fields("CP157"), 1
      '2022/8/24 END
      
      ' 費用
      If IsNull(rsTmp.Fields("CP16")) = False Then
         textCP16 = rsTmp.Fields("CP16")
      End If
      textCP16.Tag = textCP16.Text 'Add By Sindy 2023/9/22
      SetCPFieldOldData "CP16", textCP16, 1
      ' 規費
      If IsNull(rsTmp.Fields("CP17")) = False Then
         textCP17 = rsTmp.Fields("CP17")
      End If
      textCP17.Tag = textCP17.Text 'Add By Sindy 2023/9/22
      SetCPFieldOldData "CP17", textCP17, 1
      ' 點數
      If IsNull(rsTmp.Fields("CP18")) = False Then
         textCP18 = rsTmp.Fields("CP18")
      End If
      SetCPFieldOldData "CP18", textCP18, 1
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
      
      'Add by Sindy 2012/6/1
      m_CP27 = Empty
      If IsNull(rsTmp.Fields("CP27")) = False Then
         m_CP27 = rsTmp.Fields("CP27")
      End If
      '2012/6/1 End
      
      'Add by Morgan 2011/4/22
      m_CP30 = "" & rsTmp.Fields("cp30")
      SetCPFieldOldData "CP30", m_CP30, 0
      'end 2011/4/22
      
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
      ' 承辦期限
      strTemp = Empty
      If IsNull(rsTmp.Fields("CP48")) = False Then
         textCP48 = TAIWANDATE(rsTmp.Fields("CP48"))
         strTemp = rsTmp.Fields("CP48")
      End If
      textCP48.Tag = textCP48.Text 'Add By Sindy 2014/12/16
      SetCPFieldOldData "CP48", strTemp, 1
      
      'Add By Sindy 2013/12/31
      ' 被授權人(中)
      textCP50 = Empty
      If IsNull(rsTmp.Fields("CP50")) = False Then: textCP50 = rsTmp.Fields("CP50")
      SetCPFieldOldData "CP50", textCP50, 0
      ' 被授權人(英)
      textCP51 = Empty
      If IsNull(rsTmp.Fields("CP51")) = False Then: textCP51 = rsTmp.Fields("CP51")
      SetCPFieldOldData "CP51", textCP51, 0
      ' 被授權人(日)
      textCP52 = Empty
      If IsNull(rsTmp.Fields("CP52")) = False Then: textCP52 = rsTmp.Fields("CP52")
      SetCPFieldOldData "CP52", textCP52, 0
      '2013/12/31 END
      
      ' 取消收文日期
      If IsNull(rsTmp.Fields("CP57")) = False Then
         textCP57 = TAIWANDATE(rsTmp.Fields("CP57"))
      End If
      ' 是否新案件 91.12.22 ADD BY SONIA
      m_CP31 = ""
      If IsNull(rsTmp.Fields("CP31")) = False Then
         m_CP31 = rsTmp.Fields("CP31")
      End If
      ' CreateID Add By Sindy 2010/8/6
      m_CP65 = ""
      If IsNull(rsTmp.Fields("CP65")) = False Then
         m_CP65 = rsTmp.Fields("CP65")
      End If
      
      'Add By Sindy 2013/8/26
      chkWebApp.Value = 0
      If IsNull(rsTmp.Fields("CP118")) = False Then
         If rsTmp.Fields("CP118") = "Y" Then
            chkWebApp.Value = 1
         End If
      End If
      SetCPFieldOldData "CP118", IIf(chkWebApp.Value = 1, "Y", ""), 0
      '2013/8/26 END
      
      'Add By Sindy 2012/9/10
      '收據/CF帳單編號有值時,費用、規費、點數欄位鎖住
      If "" & rsTmp.Fields("CP60") > "" Or "" & rsTmp.Fields("CP61") > "" Or _
         "" & rsTmp.Fields("CP62") > "" Or "" & rsTmp.Fields("CP63") > "" Or _
         "" & rsTmp.Fields("CP87") > "" Or "" & rsTmp.Fields("CP88") > "" Then
         textCP16.Enabled = False
         textCP17.Enabled = False
         textCP18.Enabled = False
      End If
      '發文日跟發文規費有值時,規費欄位鎖住
      If Val("" & rsTmp.Fields("CP27")) > 0 Or Val("" & rsTmp.Fields("CP84")) > 0 Then
         textCP17.Enabled = False
      End If
      '2012/9/10 End
      
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
      
      '910626 Sieg 601
      '收據編號
      If IsNull(rsTmp.Fields("CP60")) = False Then
         m_CP60 = rsTmp.Fields("CP60")
      Else
         m_CP60 = ""
      End If
        'Add By Cheng 2004/01/30
        '若有收據/請款編號資料, 不可修改費用, 規費, 點數欄位
        If "" & rsTmp("CP60").Value <> "" Then
            Me.textCP16.Enabled = False
            Me.textCP17.Enabled = False
            Me.textCP18.Enabled = False
        End If
        'End
      ' 進度備註
      If IsNull(rsTmp.Fields("CP64")) = False Then
         textCP64 = rsTmp.Fields("CP64")
      End If
      SetCPFieldOldData "CP64", textCP64, 0
      
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
      m_CP149 = "" & rsTmp.Fields("CP149") 'Added by Lydia 2022/03/09 分案日
      
      'Modify by Sindy 2024/1/22 送件方式
      m_CP141 = "" & rsTmp.Fields("CP141")
      SetCPFieldOldData "CP141", m_CP141, 0
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
      If Frame6.Visible = True Then
         If "" & rsTmp.Fields("CP164") = "1" Then
            Option1(0).Value = True
         ElseIf "" & rsTmp.Fields("CP164") = "2" Then
            Option1(1).Value = True
         ElseIf "" & rsTmp.Fields("CP164") = "3" Then
            Option1(2).Value = True
         End If
      End If
      '2024/1/22 END
      
      ' 卷宗性質不為1時, 案件中英日文名稱從案件進度檔中帶入
      If IsEmptyText(m_CP10) = False Then
         If m_TM28 <> "1" Then
            'textTM05 = Empty
            'textTM06 = Empty
            'textTM07 = Empty
            Set rsSubTmp = New ADODB.Recordset
            strSubSQL = "SELECT * FROM CaseProgress " & _
                        "WHERE CP01 = '" & m_TM01 & "' AND " & _
                              "CP02 = '" & m_TM02 & "' AND " & _
                              "CP03 = '" & m_TM03 & "' AND " & _
                              "CP04 = '" & m_TM04 & "' AND " & _
                              "CP31 = 'Y' "
            rsSubTmp.CursorLocation = adUseClient
            rsSubTmp.Open strSubSQL, cnnConnection, adOpenStatic, adLockReadOnly
            If rsSubTmp.RecordCount > 0 Then
               rsSubTmp.MoveFirst
                Select Case m_TM01
                Case "T", "FCT", "CFT", "TF"
                    ' 對造案件名稱
                    If IsNull(rsSubTmp.Fields("CP37")) = False Then
                       If IsEmptyText(rsSubTmp.Fields("CP37")) = False Then
                          textTM05_1 = rsSubTmp.Fields("CP37")
                       End If
                    End If
                    SetCPFieldOldData "CP37", textTM05_1, 0
                Case Else
                    Select Case m_TM01
                    Case "S"
                        ' 對造案件中文名稱
                        If IsNull(rsSubTmp.Fields("CP37")) = False Then
                           If IsEmptyText(rsSubTmp.Fields("CP37")) = False Then
                              textTM05_1 = rsSubTmp.Fields("CP37")
                           End If
                        End If
                        SetCPFieldOldData "CP37", textTM05_1, 0
                    Case Else
                        ' 對造案件中文名稱
                        If IsNull(rsSubTmp.Fields("CP37")) = False Then
                           If IsEmptyText(rsSubTmp.Fields("CP37")) = False Then
                              textTM05 = rsSubTmp.Fields("CP37")
                           End If
                        End If
                        SetCPFieldOldData "CP37", textTM05, 0
                        ' 對造案件英文名稱
                        If IsNull(rsSubTmp.Fields("CP38")) = False Then
                           If IsEmptyText(rsSubTmp.Fields("CP38")) = False Then
                              textTM06 = rsSubTmp.Fields("CP38")
                           End If
                        End If
                        SetCPFieldOldData "CP38", textTM06, 0
                        ' 對造案件日文名稱
                        If IsNull(rsSubTmp.Fields("CP39")) = False Then
                           If IsEmptyText(rsSubTmp.Fields("CP39")) = False Then
                              textTM07 = rsSubTmp.Fields("CP39")
                           End If
                        End If
                        SetCPFieldOldData "CP39", textTM07, 0
                    End Select
                End Select
            End If
            rsSubTmp.Close
            Set rsSubTmp = Nothing
         End If
      End If
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
   SetFrame4 'Added by Morgan 2023/1/14
End Sub

' 讀取資料庫
Public Sub QueryData()
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   ' 顯示畫面為第一頁
   SSTab1.Tab = 0
   
   ' 已閉卷
   m_TM29 = Empty
   textTM29_2 = Empty
   strCP122 = "" 'Add by Amy 2022/11/17
   
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
      strCP122 = "" & rsTmp.Fields("CP122") 'Add by Amy 2022/11/17
      If IsNull(rsTmp.Fields("CP140")) = False Then: txtF0301 = rsTmp.Fields("CP140") 'Add by Amy 2022/11/10
   End If
   rsTmp.Close
    'Add By Cheng 2003/11/11
    Select Case m_TM01
    Case "T", "FCT", "CFT", "TF", "S"
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
        Me.textTM07.Enabled = True
        Me.Label42.Visible = False
        Me.textTM05_1.Visible = False
        Me.textTM05_1.Enabled = False
    End Select
   ' 本所案號
   'modify by sonia CFT-009179-2-00
   'textTMKey = m_TM01 & m_TM02 & m_TM03 & m_TM04
   textTMKey = m_TM01 & "-" & m_TM02 & "-" & m_TM03 & "-" & m_TM04
   ' 收文號
   textCP09 = m_CP09
      
   '2007/8/13 ADD BY SONIA銷卷提醒
   CheckCaseDestroy m_TM01, m_TM02, m_TM03, m_TM04
   '2007/8/13 END
   
'   Select Case m_TM01
'      ' 系統類別為CFT的為讀取商標基本檔
'      Case "T", "TF", "CFT", "FCT":
'         QueryTradeMark
'         '92.10.31 ADD BY SONIA
''         textTM05.MaxLength = 40
''         textTM07.MaxLength = 40
'         '92.10.31 END
'      Case Else:
'         QueryServicePractice
'         '92.10.31 ADD BY SONIA
'         TextTM05.MaxLength = 60
'         textTM07.MaxLength = 60
'         '92.10.31 END
'   End Select
   'Modify By Sindy 2012/6/1 把查詢基本檔的程式寫到此函數裡,共用之
   Call QueryMainFile
   
   ' 取得案件進度檔的欄位
   QueryCaseProgress
      
   'Modify By Sindy 2014/1/29
   m_CP31isYGetCP05 = GetCP31isY_CP05(m_TM01, m_TM02, m_TM03, m_TM04) '取得本所案號新案件的收文日
   'Add By Sindy 2013/12/16
   textTM130.Visible = False
   lblTM130.Visible = False
   'If strSrvDate(1) >= InvoiceStartDate Then
   If Val(m_CP31isYGetCP05) >= Val(InvoiceStartDate) Then
      'Modify By Sindy 2014/2/10 改非台灣新案都可以收J智權公司
      'If m_TM01 = "CFT" And m_CP31 = "Y" Then
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
   
   'Add By Sindy 2013/8/26
   'modify by sonia 2017/6/16 所有FCT案件性質都顯示
   'If m_TM01 = "FCT" And textTM10 = "000" And textCP10 = "101" Then
   If m_TM01 = "FCT" And textTM10 = "000" Then
   'end 2017/6/16
      chkWebApp.Visible = True
   Else
      chkWebApp.Visible = False
   End If
   '2013/8/26 END
   
   ' 是否閉卷
   If m_TM29 = "Y" Then
      EnableTextBox textTM29, True
      textTM29_2 = "已閉卷"
   Else
      EnableTextBox textTM29, False
      textTM29_2 = Empty
   End If
   
   ' 計算承辦期限
   '2011/5/16 modify by sonia 陳金蓮說CFT案件不預設承辦期限,多為收款後送件,若有需要自行輸入
   'modify by sonia 2023/2/24 還原預設但改規則
   If IsEmptyText(textCP48) = True Then
   'If IsEmptyText(textCP48) = True And m_TM01 <> "CFT" Then
      ReCaculateCP48
   End If
   
   ' 系統類別為CFC時可輸入三個申請人
'edit by nickc 2007/02/13
'   If m_TM01 = "CFC" Then
      EnableTextBox textSP58, True
      EnableTextBox textSP59, True
'   Else
'      EnableTextBox textSP58, False
'      EnableTextBox textSP59, False
'   End If
   
   ' 依讀取的是商標基本檔還是服務業務基本檔來更新控制項的狀態
   Select Case m_TM01
      Case "T", "TF", "CFT", "FCT":
         EnableTextBox textTM09, True
      Case Else:
         'Modify By Sindy 2018/11/21 ex:S-005720要輸商品類別
         'EnableTextBox textTM09, False
         EnableTextBox textTM09, True
   End Select
   
   ' 讀取優先權資料
   m_Pa(1) = m_TM01
   m_Pa(2) = m_TM02
   m_Pa(3) = m_TM03
   m_Pa(4) = m_TM04
   'edit by nickc 2007/02/06 不用 dll 了
   'objPublicData.ReadPriority m_Pa, m_Priority(1), m_Priority(2), m_Priority(3)
   'Modify by Amy 2014/04/17 +, m_Priority(4), m_Priority(5)
   'Modify by Sindy 2017/10/12 +, m_Priority(6)
   ClsPDReadPriority m_Pa, m_Priority(1), m_Priority(2), m_Priority(3), m_Priority(4), m_Priority(5), m_Priority(6)
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
   'Modify By Sindy 2010/10/15
   'ShowMaintainForm m_CP09
   'Modify By Sindy 2012/6/1 +me
   'Modify By Sindy 2012/8/7 FCT案碰上未輸申請案號之案件，在分案及發文時會彈出基本資料供補輸，煩請控制案件性質為”回代理人函”及”告代理人函”時，不要彈，謝謝,蓮
   If m_CP10 <> "720" And m_CP10 <> "719" Then
   '2012/8/7 End
      'Modify By Sindy 2024/5/3 + m_strRefText
      ShowMaintainForm m_CP09, "", "分案", Me, m_strRefText
   End If
   '2010/10/15 End
   'Add by Morgan 2003/12/07
   If (m_TM01 = "FCT") Then
      Call PUB_CheckSales(m_TM01, m_TM02, m_TM03, m_TM04, textCP05, textCP13, textCP13_2)
   End If
   'End 2003/12/07
   'Add By Cheng 2004/05/13
   '若非C類來函, Enable轉本所案號欄位
   'Modify By Sindy 2012/6/1 C類來函或已發文案件須鎖住轉本所案號欄位, 若為併號請以聯絡單通知電腦中心處理
   'If m_CP09 < "C" Then
   '2012/7/4 modify by sonia 開放719告知代理人,720回覆代理人
   'If m_CP09 < "C" And Val(m_CP27) = 0 Then
   If (m_CP09 < "C" And Val(m_CP27) = 0) Or (m_CP10 = "719" Or m_CP10 = "720") Then
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

   '2013/10/31 add by sonia 非台灣新申請案收費0,第一次分案時要提醒二案案件備註加註同時合併計算結餘" T-189182(T-188512)
   'Modified by Lydia 2022/03/09 改判斷分案日;  ex.2021/06/18 CFT案件承辦人若空白時，預設為國家檔之CFT承辦人
   'If textTM10 <> "000" And textCP10 = "101" And textCP14 = "" And Val(m_CP16) = 0 Then
   If textTM10 <> "000" And textCP10 = "101" And Val(m_CP149) = 0 And Val(m_CP16) = 0 Then
      MsgBox "此新申請案未收費, 若有前案則請至第二頁頁籤之案件備註欄加註與前案號合併計算結餘(前案之案件備註也要加註)!!!", vbExclamation + vbOKOnly
      SSTab1.Tab = 1
      textTM58_GotFocus
      textTM58.SetFocus
   End If
   '201/10/31 end
    
   'Added by Lydia 2020/05/20 法律所案源收文
   Call ReadLOS
   Call SetLOSagree
   'Modify by Amy 2022/11/10 +接洽單電子收文才顯示「檢視接洽單」鈕
   cmdFile.Visible = False
   Check11.Visible = False '急件 'Add by Amy 2022/11/17
   Check11.Value = 0 'Add By Sindy 2023/1/10 要先清欄位值,再後續判斷是否急件
   'Modify by Amy 2023/01/03 8碼(結案單)不可開接洽單會錯: + And Len(txtF0301) = 10
   If strSrvDate(1) >= 接洽單電子收文啟用日 And txtF0301 <> MsgText(601) And Len(txtF0301) = 10 Then
      cmdFile.Visible = True
      'Add by Amy 2022/11/17 +急件
      Check11.Visible = True
      If strCP122 = "Y" Then Check11.Value = 1
      'end 2022/11/17
      'Add by Amy 2023/01/07 直接開啟接洽單-桂英
      frm090801_Q.SetParent Me
      frm090801_Q.m_blnCallPrint = True
      frm090801_Q.Text5 = txtF0301
      Call frm090801_Q.cmdok_Click(4)
      frm090801_Q.Show
      'end 2023/01/07
   End If
   
   'Add By Sindy 2024/1/30 各部門分案時，若本所期限與法定期限與接洽單的本所期限與法定期限不同時，要提醒
   Call PUB_ChkCRLdtCP06CP07(m_CP09)
End Sub

'Add By Sindy 2012/6/1 為防使用者在前基本檔維護作業有修改資料, 因此基本檔資料再重新讀取一次
Public Sub QueryMainFile()
   Select Case m_TM01
      ' 系統類別為CFT的為讀取商標基本檔
      Case "T", "TF", "CFT", "FCT":
         QueryTradeMark
         '92.10.31 ADD BY SONIA
'         textTM05.MaxLength = 40
'         textTM07.MaxLength = 40
         '92.10.31 END
         'Add By Sindy 2014/2/14 案件聯絡人
         textTM38.Enabled = True
         textTM39.Enabled = True
         textTM40.Enabled = True
         textTM41.Enabled = True
         textTM42.Enabled = True
         textTM43.Enabled = True
         '2014/2/14 END
         textTM131.Enabled = True 'Add By Sindy 2015/7/14
      Case Else:
         QueryServicePractice
         '92.10.31 ADD BY SONIA
         textTM05.MaxLength = 60
         textTM07.MaxLength = 60
         '92.10.31 END
         'Add By Sindy 2014/2/14 無案件聯絡人
         textTM38.Enabled = True
         textTM39.Enabled = False
         textTM40.Enabled = False
         textTM41.Enabled = True 'False
         textTM42.Enabled = False
         textTM43.Enabled = False
         '2014/2/14 END
         textTM131.Enabled = False 'Add By Sindy 2015/7/14
   End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
   'Add by Amy 2022/11/17
   If PUB_CheckFormExist("frm090801_Q") = True Then
        Unload frm090801_Q
   End If
   PUB_SendMailCache 'Add by Sindy 2010/6/18
   m_CP09 = Empty
   'Add By Cheng 2002/07/19
   Set frm030201_02 = Nothing
End Sub

'Add By Sindy 2024/1/22
Private Sub OptSendType_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
   Dim oOpt As OptionButton
   If OptSendType(Index).Tag = "1" Then
      OptSendType(Index).Value = False
      OptSendType(Index).Tag = "0"
      If Index = 3 Then
         textCP142.Text = ""
         textCP142.Enabled = False
         If Frame6.Visible = True Then
            Option1(0).Value = False
            Option1(1).Value = False
            Option1(2).Value = False
            Option1(0).Enabled = False
            Option1(1).Enabled = False
            Option1(2).Enabled = False
         End If
      End If
      
   Else
      For Each oOpt In OptSendType
         If oOpt.Index = Index Then
            oOpt.Tag = "1"
         Else
            oOpt.Tag = "0"
         End If
      Next
      'If Index = 3 Then
      If Index = 3 And OptSendType(Index).Value Then
         textCP142.Enabled = True
         textCP142.SetFocus
         If Frame6.Visible Then
            Option1(0).Enabled = True
            Option1(1).Enabled = True
            Option1(2).Enabled = True
         End If
      Else
         textCP142.Text = ""
         textCP142.Enabled = False
         If Frame6.Visible Then
            Option1(0).Value = False
            Option1(1).Value = False
            Option1(2).Value = False
            Option1(0).Enabled = False
            Option1(1).Enabled = False
            Option1(2).Enabled = False
         End If
      End If
   End If
End Sub

Private Sub textCP05_LostFocus()
    'Add By Cheng 2003/10/14
    If Me.textCP05.Text <> "" Then
        ReCaculateCP48
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
         textCP05_GotFocus
      Else
            'Modify By Cheng 2003/10/14
'         If IsEmptyText(textCP48) = True Then
'            ReCaculateCP48
'         End If
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
      ' 檢查是否為民國日期
      If CheckIsTaiwanDate(textCP06, False) = False Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "本所期限的日期格式不正確"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      'Added by Lydia 2020/07/09 本所期限檢查：若本所期限非工作天則直接調整至最近的工作天
      Else
          textCP06.Text = TransDate(PUB_GetWorkDay1(textCP06, True), 1)
      'end 2020/07/09
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

'Modified by Lydia 2021/08/30 改成Form2.0 ; Integer=> MSForms.ReturnInteger
Private Sub textCP09_S_KeyPress(KeyAscii As MSForms.ReturnInteger)
   'Add By Cheng 2002/09/17
   KeyAscii = UpperCase(KeyAscii)
End Sub

'查名本所案號系統類別
'' 查名收文號
Private Sub textCP09_S_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   
   Cancel = False
   'Modify By Cheng 2002/09/17
'   If IsEmptyText(textCP09_S) = False Then
'      strSQL = "SELECT * FROM CaseProgress " & _
'               "WHERE CP09 = '" & textCP09_S & "' "
'      rsTmp.CursorLocation = adUseClient
'      rsTmp.Open strSQL, cnnConnection, adOpenStatic, adLockReadOnly
'      If rsTmp.RecordCount <= 0 Then
'         Cancel = True
'         strTit = "檢核資料"
'         strMsg = "查名收文號不存在"
'         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
'         textCP09_S_GotFocus
'      End If
'      rsTmp.Close
'   End If
'   Set rsTmp = Nothing
   If Me.textCP09_S.Text <> "" Then
      If Me.textCP09_S.Text <> "S" Then
         MsgBox "查名本所案號的系統類別類輸入錯誤!!!", vbExclamation + vbOKOnly
         Cancel = True
         Me.textCP09_S.SetFocus
         textCP09_S_GotFocus
      End If
   End If
End Sub

Private Sub textCP09_S1_GotFocus()
   InverseTextBox textCP09_S1
End Sub

Private Sub textCP09_S2_GotFocus()
   InverseTextBox textCP09_S2
End Sub

Private Sub textCP09_S3_GotFocus()
   InverseTextBox textCP09_S3
End Sub

Private Sub textCP09_S3_LostFocus()
   If textCP09_S2 = "" Then textCP09_S2 = "0"
   If textCP09_S3 = "" Then textCP09_S3 = "00"
   Call ChkSPDataErr
'   'Add By Cheng 2002/09/17
'   Dim strTit As String
'   Dim strMsg As String
'   Dim nResponse
'   Dim rsTmp As New ADODB.Recordset
'   Dim strSql As String
'
'   'Add By Cheng 2002/09/17
'   If textCP09_S = "S" And IsEmptyText(textCP09_S1) = False Then
'      strSql = "SELECT CP09 FROM CaseProgress " & _
'               "WHERE " & ChgCaseprogress(Me.textCP09_S.Text & Me.textCP09_S1.Text & Me.textCP09_S2.Text & Me.textCP09_S3.Text)
'      rsTmp.CursorLocation = adUseClient
'      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
'      If rsTmp.RecordCount <= 0 Then
'         strTit = "檢核資料"
'         strMsg = "查名本所案號不存在"
'         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
'         Me.textCP09_S.SetFocus
'         textCP09_S_GotFocus
'      End If
'      rsTmp.Close
'   End If
'   Set rsTmp = Nothing
End Sub
'Modify By Sindy 2015/6/22
Private Function ChkSPDataErr() As Boolean
   'Add By Cheng 2002/09/17
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   
   ChkSPDataErr = False
   'Add By Cheng 2002/09/17
   If textCP09_S = "S" And IsEmptyText(textCP09_S1) = False Then
      strSql = "SELECT CP09 FROM CaseProgress " & _
               "WHERE " & ChgCaseprogress(Me.textCP09_S.Text & Me.textCP09_S1.Text & Me.textCP09_S2.Text & Me.textCP09_S3.Text)
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount <= 0 Then
         strTit = "檢核資料"
         strMsg = "查名本所案號不存在"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         Me.textCP09_S1.SetFocus
         textCP09_S1_GotFocus
         ChkSPDataErr = True
         rsTmp.Close
         Set rsTmp = Nothing
         Exit Function
      End If
      rsTmp.Close
      'Add By Sindy 2015/6/22 檢查輸入之S案號,其案件名稱有無'未成卷'的字樣
      strSql = "SELECT SP01,SP02,SP03,SP04,SP05 FROM ServicePractice " & _
               "WHERE SP01='" & textCP09_S & "' and SP02='" & textCP09_S1 & "' and SP03='" & textCP09_S2 & "' and SP04='" & textCP09_S3 & "' " & _
               "and instr(SP05,'未成卷')>0"
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         strTit = "檢核資料"
         strMsg = "查名本所案號欄不可輸入未成卷的查名案號！"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         Me.textCP09_S1.SetFocus
         textCP09_S1_GotFocus
         ChkSPDataErr = True
         rsTmp.Close
         Set rsTmp = Nothing
         Exit Function
      End If
      rsTmp.Close
      '2015/6/22 END
   End If
   Set rsTmp = Nothing
End Function

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
      'end 2020/05/20
      End If
      SetFrame4 'Added by Morgan 2022/12/23
  End If
End Sub

'Add By Sindy 2010/11/29
'Modified by Lydia 2021/08/30 改成Form2.0 ; Integer=> MSForms.ReturnInteger
Private Sub textCP13_KeyPress(KeyAscii As MSForms.ReturnInteger)
KeyAscii = UpperCase(KeyAscii)
End Sub

' 智權人員
Private Sub textCP13_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Dim rsTmp As New ADODB.Recordset
   'Added by Lydia 2019/02/14
   Dim m_SalesST15 As String '畫面上智權人員的收文部門
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

'Add By Sindy 2010/11/29
'Modified by Lydia 2021/08/30 改成Form2.0 ; Integer=> MSForms.ReturnInteger
Private Sub textCP14_KeyPress(KeyAscii As MSForms.ReturnInteger)
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

'Add By Sindy 2023/12/11
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
      ElseIf Not ChkWorkDay(DBDATE(textCP142)) Then
         MsgBox "指定送件日期必須是工作天 !", vbExclamation, "輸入指定日期錯誤"
         Cancel = True
         textCP142.SetFocus
         textCP142_GotFocus
      ElseIf Val(textCP142) < Val(strSrvDate(2)) Then
         MsgBox "指定日期不可小於系統日！", vbExclamation, "輸入指定日期錯誤"
         Cancel = True
         textCP142.SetFocus
         textCP142_GotFocus
      ElseIf textCP142 <> "" And textCP07 <> "" And Val(textCP142) > Val(textCP07) Then
         MsgBox "指定日期不可大於法定期限！", vbExclamation, "輸入指定日期錯誤"
         Cancel = True
         textCP142.SetFocus
         textCP142_GotFocus
      End If
   End If
End Sub
'2023/12/11 END

' 費用
Private Sub textCP16_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textCP16) = False Then
      If IsNumeric(textCP16) = False Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "費用為數值資料"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP16_GotFocus
      End If
   End If
End Sub

' 規費
Private Sub textCP17_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textCP17) = False Then
      If IsNumeric(textCP17) = False Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "規費為數值資料"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP17_GotFocus
      End If
   End If
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
         strMsg = "點數為數值資料"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP18_GotFocus
      End If
   End If
End Sub

'Modified by Lydia 2021/08/30 改成Form2.0 ; Integer=> MSForms.ReturnInteger
Private Sub textCP26_KeyPress(KeyAscii As MSForms.ReturnInteger)
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
         strMsg = "相關總收文號不可為本身之收文號"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         GoTo EXITSUB
      End If
      
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
      rsTmp.Close
   End If
EXITSUB:
   Set rsTmp = Nothing
End Sub

' 承辦期限
Private Sub textCP48_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textCP48) = False Then
      If CheckIsTaiwanDate(textCP48, False) = False Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "承辦期限的日期格式不正確"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP48_GotFocus
      'add by sonia 2025/3/14
      ElseIf Not ChkWorkDay(DBDATE(textCP48)) Then
         MsgBox "承辦期限必須是工作天 ! 將自動更新為前一個工作天", vbExclamation, "輸入承辦期限錯誤"
         textCP48.Text = TransDate(PUB_GetWorkDay1(textCP48, True), 1)
      'end 2025/3/14
      End If
   End If
End Sub

Private Sub textCP50_GotFocus()
   InverseTextBox textCP50
End Sub

Private Sub textCP51_GotFocus()
   InverseTextBox textCP51
End Sub

Private Sub textCP52_GotFocus()
   InverseTextBox textCP52
End Sub

' 徵求同意書對象中文名稱
Private Sub textCP50_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textCP50) = False Then
      If CheckLengthIsOK(textCP50, 60, False) = False Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "徵求同意書對象中文名稱內容太長"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP50_GotFocus
      End If
   End If
End Sub

' 徵求同意書對象英文名稱
Private Sub textCP51_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textCP51) = False Then
      If CheckLengthIsOK(textCP51, 60, False) = False Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "徵求同意書對象英文名稱內容太長"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP51_GotFocus
      End If
   End If
End Sub

' 徵求同意書對象日文名稱
Private Sub textCP52_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textCP52) = False Then
      If CheckLengthIsOK(textCP52, 60, False) = False Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "徵求同意書對象日文名稱內容太長"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP52_GotFocus
      End If
   End If
End Sub

' 進度備註
Private Sub textCP64_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If CheckLengthIsOK(textCP64, textCP64.MaxLength, False) = False Then
      Cancel = True
      strTit = "檢核資料"
      strMsg = "進度備註內容太長"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textCP64_GotFocus
   End If
End Sub

'Add By Sindy 2014/2/13
Private Sub textCU146_GotFocus()
   CloseIme
   TextInverse textCU146
End Sub

'Add By Sindy 2014/2/13
Private Sub textCU147_GotFocus()
   CloseIme
   TextInverse textCU147
End Sub

'Modified by Lydia 2021/08/30 改成Form2.0 ; Integer=> MSForms.ReturnInteger
Private Sub textCU147_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textCU147_LostFocus()
   textCU147_Validate False
End Sub

Private Sub textCU147_Validate(Cancel As Boolean)
   'Modified by Lydia 2021/08/30 改成Form2.0 ; Label30(16)=> lblCU147、Label30(15)=> lblCU151
   lblCU147.Caption = ""
   If textCU147.Text <> "" Then
      textCU147 = textCU147 & String(9 - Len(textCU147), "0")
      lblCU147.Caption = ChgType(4, textCU147.Text)
      If lblCU147.Caption = "" Then Cancel = True: Exit Sub
   End If
   'end 2021/08/30
End Sub

Private Sub textCU149_GotFocus()
   CloseIme
   TextInverse textCU149
End Sub

'Modified by Lydia 2021/08/30 改成Form2.0 ; Integer=> MSForms.ReturnInteger
Private Sub textCU149_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = UpperCase(KeyAscii)
End Sub

'商標D/N是否列印申請人
Private Sub textCU149_Validate(Cancel As Boolean)
   If textCU149.Text = "" Then Exit Sub
   If textCU149.Text <> "Y" Then
      ShowMsg "輸入錯誤 !"
      Cancel = True
   End If
End Sub

'Add By Sindy 2014/2/13
Private Sub textCU151_GotFocus()
   CloseIme
   TextInverse textCU151
End Sub

'Modified by Lydia 2021/08/30 改成Form2.0 ; Integer=> MSForms.ReturnInteger
Private Sub textCU151_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textCU151_LostFocus()
   textCU151_Validate False
End Sub

Private Sub textCU151_Validate(Cancel As Boolean)
   'Modified by Lydia 2021/08/30 改成Form2.0 ; Label30(16)=> lblCU147、Label30(15)=> lblCU151
   lblCU151.Caption = ""
   If textCU151 <> "" Then
      textCU151 = textCU151 & String(9 - Len(textCU151), "0")
      lblCU151.Caption = ChgType(4, textCU151.Text)
      If lblCU151.Caption = "" Then Cancel = True: Exit Sub
   End If
   'end 2021/08/30
End Sub

'Add By Sindy 2014/2/13
Private Sub textCU58_GotFocus()
   OpenIme
   TextInverse textCU58
End Sub
Private Sub textCU58_Validate(Cancel As Boolean)
   If textCU58.Text <> "" Then
      If Not CheckLengthIsOK(textCU58, textCU58.MaxLength, False) Then
         Cancel = True
         MsgBox "聯絡人１(中)內容太長", vbOKOnly, "檢核資料"
         textCU58_GotFocus
      End If
   End If
End Sub

'Add By Sindy 2014/2/13
Private Sub textCU59_GotFocus()
   CloseIme
   TextInverse textCU59
End Sub

'Add By Sindy 2014/2/13
Private Sub textCU60_GotFocus()
   OpenIme
   TextInverse textCU60
End Sub
Private Sub textCU60_Validate(Cancel As Boolean)
   If textCU60.Text = "" Then Exit Sub
    '中日文欄位尾碼加空白，最大可輸長度減一(因可能會有造字無法存入問題)
   If Not CheckLengthIsOK(textCU60, textCU60.MaxLength - 1, False) Then
      Cancel = True
      MsgBox "聯絡人１(日)內容太長", vbOKOnly, "檢核資料"
      textCU60_GotFocus
   End If
End Sub

'Add By Sindy 2014/2/13
Private Sub textCU61_GotFocus()
   OpenIme
   TextInverse textCU61
End Sub
Private Sub textCU61_Validate(Cancel As Boolean)
   If textCU61.Text = "" Then Exit Sub
   If Not CheckLengthIsOK(textCU61, textCU61.MaxLength, False) Then
      Cancel = True
      MsgBox "聯絡人２(中)內容太長", vbOKOnly, "檢核資料"
      textCU61_GotFocus
   End If
End Sub

'Add By Sindy 2014/2/13
Private Sub textCU62_GotFocus()
   CloseIme
   TextInverse textCU62
End Sub

'Add By Sindy 2014/2/13
Private Sub textCU63_GotFocus()
   OpenIme
   TextInverse textCU63
End Sub

Private Sub textCU63_Validate(Cancel As Boolean)
   If textCU63.Text = "" Then Exit Sub
    '中日文欄位尾碼加空白，最大可輸長度減一(因可能會有造字無法存入問題)
   If Not CheckLengthIsOK(textCU63, textCU63.MaxLength - 1, False) Then
      Cancel = True
      MsgBox "聯絡人２(日)內容太長", vbOKOnly, "檢核資料"
      textCU63_GotFocus
   End If
End Sub

'Add By Sindy 2014/2/13
'代理人聯絡人1(中)
Private Sub textFA07_GotFocus()
   InverseTextBox textFA07
   OpenIme
End Sub
Private Sub textFA07_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textFA07) = False Then
      If StrLength(textFA07) > 10 Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "代理人聯絡人1(中)內容太長"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textFA07_GotFocus
      End If
   End If
   If Cancel = False Then CloseIme
End Sub

'Add By Sindy 2014/2/13
Private Sub textFA08_GotFocus()
   InverseTextBox textFA08
End Sub

'Add By Sindy 2014/2/13
'代理人聯絡人1(日)
Private Sub textFA09_GotFocus()
   InverseTextBox textFA09
   OpenIme
End Sub
Private Sub textFA09_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textFA09) = False Then
      If StrLength(textFA09) > textFA09.MaxLength Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "代理人聯絡人1(日)內容太長"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textFA09_GotFocus
      End If
   End If
   If Cancel = False Then CloseIme
End Sub

'Add By Sindy 2014/2/13
Private Sub textFA106_GotFocus()
   InverseTextBox textFA106
End Sub

'Add By Sindy 2014/2/13
'代理人商標固定請款對象
Private Sub textFA107_GotFocus()
   InverseTextBox textFA107
End Sub

'Modified by Lydia 2021/08/30 改成Form2.0 ; Integer=> MSForms.ReturnInteger
Private Sub textFA107_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textFA107_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   Cancel = False
   textFA107_2 = Empty
   If IsEmptyText(textFA107) = False Then
      If (textFA107 & String(9 - Len(textFA107), "0")) = textTM44 Or _
         Mid(textFA107 & String(9 - Len(textFA107), "0"), 1, 8) = m_textTM44_FA03 Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "代理人商標固定請款對象不可為該筆資料的代理人"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textFA107_GotFocus
      End If
      Select Case Mid(textFA107, 1, 1)
         Case "X":
            textFA107_2 = GetCustomerName(textFA107, 0)
         Case "Y":
            textFA107_2 = GetFAgentName(textFA107)
         Case Else:
            textFA107_2 = Empty
            Cancel = True
            strTit = "檢核資料"
            strMsg = "代理人商標固定請款對項代號不正確"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textFA107_GotFocus
            GoTo EXITSUB
      End Select
      If IsEmptyText(textFA107_2) = True Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "代理人商標固定請款對項代號不存在"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textFA107_GotFocus
      End If
   End If
EXITSUB:
End Sub

Private Sub textFA109_GotFocus()
   InverseTextBox textFA109
End Sub

'Modified by Lydia 2021/08/30 改成Form2.0 ; Integer=> MSForms.ReturnInteger
Private Sub textFA109_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = UpperCase(KeyAscii)
End Sub

' 商標D/N是否列印申請人
Private Sub textFA109_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   Select Case textFA109
      Case "Y", "":
      Case Else:
'         Select Case m_EditMode
'            Case 1, 2:
               Cancel = True
               strTit = "檢核資料"
               strMsg = "商標D/N是否列印申請人只可輸入Y"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textFA109_GotFocus
'         End Select
   End Select
End Sub
'2011/3/4 End

'Add By Sindy 2014/2/13
Private Sub textFA111_GotFocus()
   InverseTextBox textFA111
End Sub
'Modified by Lydia 2021/08/30 改成Form2.0 ; Integer=> MSForms.ReturnInteger
Private Sub textFA111_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textFA111_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   textFA111_2 = Empty
   If IsEmptyText(textFA111) = False Then
      Select Case Mid(textFA111, 1, 1)
         Case "X":
            textFA111_2 = GetCustomerName(textFA111, 0)
         Case "Y":
            textFA111_2 = GetFAgentName(textFA111)
         Case Else:
            textFA111_2 = Empty
            Cancel = True
            strTit = "檢核資料"
            strMsg = "代理人商標D/N固定列印對象代號不正確"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textFA111_GotFocus
            GoTo EXITSUB
      End Select
      If IsEmptyText(textFA111_2) = True Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "代理人商標D/N固定列印對象代號不存在"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textFA111_GotFocus
      End If
   End If
EXITSUB:
End Sub

'Add By Sindy 2014/2/13
'代理人聯絡人2(中)
Private Sub textFA52_GotFocus()
   InverseTextBox textFA52
   OpenIme
End Sub
Private Sub textFA52_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textFA52) = False Then
      If StrLength(textFA52) > 10 Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "代理人聯絡人2(中)內容太長"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textFA52_GotFocus
      End If
   End If
   If Cancel = False Then CloseIme
End Sub

'Add By Sindy 2014/2/13
Private Sub textFA53_GotFocus()
   InverseTextBox textFA53
End Sub

'Add By Sindy 2014/2/13
'代理人聯絡人2(日)
Private Sub textFA54_GotFocus()
   InverseTextBox textFA54
   OpenIme
End Sub
Private Sub textFA54_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textFA54) = False Then
      If StrLength(textFA54) > textFA54.MaxLength Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "代理人聯絡人2(日)內容太長"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textFA54_GotFocus
      End If
   End If
   If Cancel = False Then CloseIme
End Sub

'add by nickc 2007/01/15
'Modified by Lydia 2021/08/30 改成Form2.0 ; Integer=> MSForms.ReturnInteger
Private Sub textSP58_KeyPress(KeyAscii As MSForms.ReturnInteger)
KeyAscii = UpperCase(KeyAscii)
End Sub

'Modified by Lydia 2021/08/30 改成Form2.0 ; Integer=> MSForms.ReturnInteger
Private Sub textSP59_KeyPress(KeyAscii As MSForms.ReturnInteger)
KeyAscii = UpperCase(KeyAscii)
End Sub

'Modified by Lydia 2021/08/30 改成Form2.0 ; Integer=> MSForms.ReturnInteger
Private Sub textTM01_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = UpperCase(KeyAscii)
End Sub

' 本所案號的系統別
Private Sub textTM01_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   Cancel = False
   textTM02_2.Visible = False
   textTM02_2.Locked = True
   textTM02_2.TabStop = False
   textTM02.MaxLength = 6
   If IsEmptyText(textTM01) = False Then
      If textTM01 <> m_TM01 Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "轉本所案號中的系統別不正確"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
        textTM01_GotFocus
         GoTo EXITSUB
      End If
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
      'add by sonia 2017/6/16 CFT案加可選下一程序緩衝期限CFT-017235
      If strTM01 = "CFT" Then
         strSql = strSql & " UNION SELECT NP01,NP07,NP08,NP09,NP10,NP11,NP12,NP13,NP14,NP15,NP22 FROM NextProgress " & _
               "WHERE NP02 = '" & strTM01 & "' AND " & _
                     "NP03 = '" & strTM02 & "' AND " & _
                     "NP04 = '" & strTM03 & "' AND " & _
                     "NP05 = '" & strTM04 & "' AND " & _
                     "(NP06 IS NULL OR NP06 <> 'Y') AND NP07='312'"
      End If
      'end 2017/6/16
      strSql = strSql & " UNION SELECT CP09,CP10,CP06,CP07,CP13,CP57,CP58,CP08,CP40,CP64,0" & _
         " FROM CASEPROGRESS WHERE CP01 = '" & strTM01 & "' AND CP02 = '" & strTM02 & "'" & _
         " AND CP03 = '" & strTM03 & "' AND CP04 = '" & strTM04 & "'" & _
         " AND CP09<'C' and cp10<>'303' and cp07>0 AND CP27 IS NULL AND CP57 IS NULL"
   End If
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
            grdList.TextMatrix(nIndex, 1) = GetCaseTypeName(m_TM01, rsTmp.Fields("NP07"))
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
      'Added by Lydia 2023/10/17
      If grdList.Rows >= 2 Then
         grdList.FixedRows = 1
      End If
      'end 2023/10/17
   End If
   rsTmp.Close
   
   'add by sonia 2021/4/23 若下一程序有相同案件性質未續辦則提醒
   strSql = "Select * From CaseProgress,nextprogress Where CP09='" & m_CP09 & "' AND CP01=NP02(+) AND CP02=NP03(+) AND CP03=NP04(+) AND CP04=NP05(+) AND CP10=NP07(+) AND NP06 IS NULL AND NP01 IS NOT NULL"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      MsgBox "此案件性質於下一程序檔仍有未續辦期限，請注意是否消本案期限!!!", vbExclamation + vbOKOnly
   End If
   rsTmp.Close
   'end 2021/4/23
   
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
      Case "T", "TF", "FCT", "CFT":
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

'Modified by Lydia 2021/08/30 改成Form2.0 ; Integer=> MSForms.ReturnInteger
Private Sub textTM03_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textTM04_LostFocus()
   Dim strTM01 As String
   Dim strTM02 As String
   Dim strTM03 As String
   Dim strTM04 As String

'Add By Cheng 2002/12/31
If (Me.textTM01.Text = "" And Me.textTM02.Text <> "") Or (Me.textTM01.Text <> "" And Me.textTM02.Text = "") Then
    MsgBox "轉本所案號輸入不完整!!!", vbExclamation + vbOKOnly
    Me.textTM01.SetFocus
    textTM01_GotFocus
    Exit Sub
End If

If textTM01 <> "" And textTM02 <> "" Then
   strTM01 = textTM01
   strTM02 = textTM02
   If strTM02 = "TF" Then: strTM02 = strTM02 & textTM02_2
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
   chkNewTMNo = True
   'edit by nickc 2007/02/06 不用 dll 了
   'If objPublicData.GetMaxNumber(strNo(1), strExc(0)) Then
   If ClsPDGetMaxNumber(strNo(1), strExc(0)) Then
      '2006/7/5 MODIFY BY SONIA 只判斷前二欄位
      'If strNo(1) & strNo(2) & strNo(3) & strNo(4) > strNo(1) & String(6 - Len(strExc(0)), "0") & strExc(0) Then
      If strNo(1) & strNo(2) > strNo(1) & String(6 - Len(strExc(0)), "0") & strExc(0) Then
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
            '若系統日小於等於法定期限
            '91.12.22 MODIFY BY SONIA 改判斷收文日
            'If ServerDate <= Val(m_strCP07) Then
            If DBDATE(Val(textCP05)) <= Val(m_strCP07) Then
            '91.12.22 END
               'Modify By Sindy 2023/10/19 mark:不用檢查,因人員進來時欄位是鎖住的,不能修改,
               '                           檢查了,反而擋住系統要把本所期限改為系統日
'               '本所期限及法定期限不可修改
'               SetCPFieldNewData "CP06", DBDATE(m_strCP06)
'               SetCPFieldNewData "CP07", DBDATE(m_strCP07)
'               'add by sonia 2017/8/22
'               If (Val(DBDATE(Val(textCP06))) <> Val(m_strCP06)) Or (Val(DBDATE(Val(textCP07))) <> Val(m_strCP07)) Then
'                  MsgBox "注意!!延展案收文日小於等於法定期限，不可在分案修改本所期限或法定期限，其他欄位仍會更新!!!", vbCritical
'               End If
'               'end 2017/8/22
            '若系統日大於法定期限
            Else
               '91.12.22 ADD BY SONIA
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
                     If GetCF28(m_TM01, m_TM10, Me.textCP10.Text) <> 0 Then
                        'modify by sonia 2023/4/14 改用textCP07計算,以免m_TM22與textCP07不同而未注意到(T-131010)
                        'm_strCP07 = DBDATE(CompDate(1, (GetCF28(m_TM01, m_TM10, Me.textCP10.Text)), Format(m_TM22)))
                        m_strCP07 = DBDATE(CompDate(1, (GetCF28(m_TM01, m_TM10, Me.textCP10.Text)), Format(textCP07)))
                     End If
                  End If
                  If m_strCP07 <> "" Then
                     SetCPFieldNewData "CP07", DBDATE(m_strCP07)
                     '本所期限 = 法定期限 - 2天
                    'Modify By Cheng 2003/09/01
'                     m_strCP06 = DBDATE(Format(DateSerial(Val(DBYEAR(m_strCP07)), Val(DBMONTH(m_strCP07)), Val(DBDAY(m_strCP07)) - 2)))
                     'Modify By Sindy 2014/10/6 台灣案之本所期限設定
                     If m_TM10 = "000" And Val(strSrvDate(1)) >= 台灣案所限新規則啟用日 Then
                        m_strCP06 = PUB_GetOurDeadline(DBDATE(m_strCP07))
                     Else
                     '2014/10/6 END
                        m_strCP06 = DBDATE(DateAdd("d", -2, ChangeWStringToWDateString(DBDATE(m_strCP07))))
                     End If
                     m_strCP06 = PUB_GetWorkDay1(m_strCP06, True) 'Added by Lydia 2020/07/09 本所期限檢查：若本所期限非工作天則直接調整至最近的工作天
                     SetCPFieldNewData "CP06", DBDATE(m_strCP06)
                  End If
                  '91.11.3 END
               '91.12.22 ADD BY SONIA
               Else
                  SetCPFieldNewData "CP06", DBDATE(textCP06)
                  SetCPFieldNewData "CP07", DBDATE(textCP07)
               End If
               '91.12.22 END
            End If
            '91.11.3 MODIFY BY SONIA 規費為 CP17 非 CP07
            '規費
            'SetCPFieldNewData "CP07", (Val(GetCF08(m_TM01, m_TM10, Me.textCP10.Text)) * 2)
            '91.12.22 CANCEL BY SONIA
            'If m_TM01 = "FCT" Then
            '   textCP17 = Val(GetCF08(m_TM01, m_TM10, Me.textCP10.Text)) * 2
            '   textCP18 = (textCP16 - textCP17) / 1000
            'End If
            '91.12.22 END
            '91.11.3 END
         
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
   'Add By Sindy 2022/8/24
   If textCP14 <> "" Then
      SetCPFieldNewData "CP157", strSrvDate(1)
   Else
      SetCPFieldNewData "CP157", Empty
   End If
   '2022/8/24 END
   ' 費用
   SetCPFieldNewData "CP16", textCP16
   ' 規費
   SetCPFieldNewData "CP17", textCP17
   ' 點數
   SetCPFieldNewData "CP18", textCP18
   ' 相關總收文號
   SetCPFieldNewData "CP43", textCP43
   ' 是否算案件數
   SetCPFieldNewData "CP26", textCP26
   
   'Add By Sindy 2024/1/22
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
   '2024/1/22 END
   
   'Add By Sindy 2013/8/26
   ' 是否電子送件
   If chkWebApp.Visible = True Then
      If chkWebApp.Value = 1 Then
         SetCPFieldNewData "CP118", "Y"
      Else
         SetCPFieldNewData "CP118", Empty
      End If
   Else
      SetCPFieldNewData "CP118", Empty
   End If
   '2013/8/26 END
   
   'Add by Morgan 2011/4/22 延期要紀錄NP22
   If textCP10 = "303" Then
      SetCPFieldNewData "CP30", m_CP30
   End If
   
   'Add By Sindy 2013/12/31
   ' 被授權人(中)
   SetCPFieldNewData "CP50", textCP50
   ' 被授權人(英)
   SetCPFieldNewData "CP51", textCP51
   ' 被授權人(日)
   SetCPFieldNewData "CP52", textCP52
   '2013/12/31 END
   
   ' 承辦期限
    'Modify By Cheng 2002/11/12
'   SetCPFieldNewData "CP48", textCP48
   SetCPFieldNewData "CP48", DBDATE(textCP48)
   ' 進度備註
    strCP64 = Me.textCP64.Text
    'Modify By Cheng 2003/09/05
    '取消
    'Begin
'    'Add By Cheng 2003/06/16
'    '若有輸入查名本所案號
'    If Me.textCP09_S.Text <> "" And Me.textCP09_S1.Text <> "" Then
'        strCP64 = strCP64 & IIf(strCP64 <> "", ",", "") & "原查名本所案號：" & Me.textCP09_S.Text & "-" & Me.textCP09_S1.Text & "-" & Left(Me.textCP09_S2.Text & "0", 1) & "-" & Left(Me.textCP09_S3.Text & "00", 2)
'    End If
    'End
   SetCPFieldNewData "CP64", strCP64
   
   'Add By Sindy 2020/10/20
   If textCP10 = "210" Then '陳述意見書
      SetCPFieldNewData "CP143", IIf(m_CP143 <> "", DBDATE(m_CP143), "")
      SetCPFieldNewData "CP36", m_CP36
      SetCPFieldNewData "CP21", m_CP21
   End If
   '2020/10/20 END
   
   ' 卷宗性質為非申請時, 更新案件進度檔的對造案件名稱
   If textTM28 <> "1" Then
        Select Case m_TM01
        Case "T", "FCT", "CFT", "TF", "S"
            ' 對造案件名稱
            SetCPFieldNewData "CP37", textTM05_1
        Case Else
            ' 對造案件名稱(中)
            SetCPFieldNewData "CP37", textTM05
            ' 對造案件名稱(英)
            SetCPFieldNewData "CP38", textTM06
            ' 對造案件名稱(日)
            SetCPFieldNewData "CP39", textTM07
        End Select
   End If
   
   Select Case m_TM01
      ' 系統類別為CFT的為更新商標基本檔
      Case "T", "TF", "CFT", "FCT":
        'Modify By Cheng 2003/02/24
        '取消卷宗性質為申請的限制
'         ' 卷宗性質為非申請時, 不更新基本檔
'         If textTM28 = "1" Then
'            ' 案件中文名稱
'            SetTMSPFieldNewData "TM05", textTM05
            ' 案件名稱
            SetTMSPFieldNewData "TM05", textTM05_1
'            ' 案件英文名稱
'            SetTMSPFieldNewData "TM06", textTM06
'            ' 案件日文名稱
'            SetTMSPFieldNewData "TM07", textTM07
'         End If
         'Add By Sindy 2015/7/14
         ' 定稿商標名稱
         SetTMSPFieldNewData "TM131", textTM131
         '2015/7/14 END
         ' 商標種類
         SetTMSPFieldNewData "TM08", textTM08
         'Add By Sindy 2019/1/4
         ' 特殊商標
         SetTMSPFieldNewData "TM72", textTM72
         '2019/1/4 END
         ' 商品類別
         SetTMSPFieldNewData "TM09", textTM09
         ' 申請國家
         SetTMSPFieldNewData "TM10", textTM10
         ' 申請人
         SetTMSPFieldNewData "TM23", textTM23
         ' 申請地址(中)
         SetTMSPFieldNewData "TM24", textTM24
         ' 申請地址(英)
         SetTMSPFieldNewData "TM25", textTM25
         ' 申請地址(日)
         SetTMSPFieldNewData "TM26", textTM26
         'add by nickc 2007/01/15
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
         'Add By Sindy 2013/12/16
         ' 特殊出名公司
         SetTMSPFieldNewData "TM130", textTM130
         '2013/12/16 END
      Case Else:
        'Modify By Cheng 2003/02/27
        '取消卷宗性質為申請的限制
'         ' 卷宗性質為非申請時, 不更新基本檔
'         If textTM28 = "1" Then
            Select Case m_TM01
            Case "S"
                ' 案件中文名稱
                SetTMSPFieldNewData "SP05", textTM05_1
            Case Else
                ' 案件中文名稱
                SetTMSPFieldNewData "SP05", textTM05
            End Select
            ' 案件英文名稱
            SetTMSPFieldNewData "SP06", textTM06
            ' 案件日文名稱
            SetTMSPFieldNewData "SP07", textTM07
'         End If
         ' 申請人
         SetTMSPFieldNewData "SP08", textTM23
'edit by nickc 2007/01/15
'         If m_TM01 = "CFC" Then
            ' 申請人2
            SetTMSPFieldNewData "SP58", textSP58
            ' 申請人3
            SetTMSPFieldNewData "SP59", textSP59
'         End If
         'add by nickc 2007/01/15
         SetTMSPFieldNewData "SP65", textTM80
         SetTMSPFieldNewData "SP66", textTM81
         SetTMSPFieldNewData "SP73", textTM09
         
         ' 申請國家
'         SetTMSPFieldNewData "SP09", m_TM10
         SetTMSPFieldNewData "SP09", Me.textTM10.Text
         ' FC代理人
         If IsEmptyText(textTM44) = False Then
            SetTMSPFieldNewData "SP26", textTM44 & String(9 - Len(textTM44), "0")
         Else
            SetTMSPFieldNewData "SP26", textTM44
         End If
         ' 彼所案號
         SetTMSPFieldNewData "SP27", textTM45
         ' 案件備註
         SetTMSPFieldNewData "SP18", textTM58
         ' 客戶案件案號
         SetTMSPFieldNewData "SP29", textTM35 'Add By Sindy 2021/5/5
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
Private Sub OnUpdateTradeMark()
   Dim strTmp As String
   Dim strSql As String
   Dim nIndex As Integer
   Dim bFirst As Boolean
   Dim bDifference As Boolean
   Dim rsTmp As New ADODB.Recordset 'Add By Sindy 2017/3/28
   
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
               strExc(6) = textTM23
               strExc(7) = strExc(0)
               '911118 nick 新增申請人
               strExc(8) = m_TM23
               'edit by nickc 2007/02/06 不用 dll 了
               'If Not objLawDll.UpdAcc0k0(strExc(), True) Then
               If Not ClsLawUpdAcc0k0(strExc(), True) Then
                  textTM23.SetFocus
                  Exit Sub
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
   'add by nickc 2007/01/16
   If textSP58 <> "" Then
      'edit by nickc 2007/02/06 不用 dll 了
      'If objPublicData.GetCustomerNameAndAddress(textSP58, strExc(0), strTmp1(1), strTmp1(2), strTmp1(3)) Then
      If ClsPDGetCustomerNameAndAddress(textSP58, strExc(0), strTmp1(1), strTmp1(2), strTmp1(3)) Then
         '修改申請人時
         If InStr(ChangeCustomerL(m_TM78), ChangeCustomerL(textSP58)) = 0 Then
            If m_CP60 <> "" Then
               strExc(1) = m_TM01
               strExc(2) = m_TM02
               strExc(3) = m_TM03
               strExc(4) = m_TM04
               strExc(5) = m_CP60
               strExc(6) = textSP58
               strExc(7) = strExc(0)
               strExc(8) = m_TM78
            End If
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
            If m_CP60 <> "" Then
               strExc(1) = m_TM01
               strExc(2) = m_TM02
               strExc(3) = m_TM03
               strExc(4) = m_TM04
               strExc(5) = m_CP60
               strExc(6) = textSP59
               strExc(7) = strExc(0)
               strExc(8) = m_TM79
            End If
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
            If m_CP60 <> "" Then
               strExc(1) = m_TM01
               strExc(2) = m_TM02
               strExc(3) = m_TM03
               strExc(4) = m_TM04
               strExc(5) = m_CP60
               strExc(6) = textTM80
               strExc(7) = strExc(0)
               strExc(8) = m_TM80
            End If
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
            If m_CP60 <> "" Then
               strExc(1) = m_TM01
               strExc(2) = m_TM02
               strExc(3) = m_TM03
               strExc(4) = m_TM04
               strExc(5) = m_CP60
               strExc(6) = textTM81
               strExc(7) = strExc(0)
               strExc(8) = m_TM81
            End If
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
   
   ' 更新案件進度檔
   strSql = "UPDATE TradeMark SET "
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
                  "WHERE TM01 = '" & m_TM01 & "' AND " & _
                        "TM02 = '" & m_TM02 & "' AND " & _
                        "TM03 = '" & m_TM03 & "' AND " & _
                        "TM04 = '" & m_TM04 & "' "
   ' 執行SQL指令
   If bDifference = True Then
        Pub_SeekTbLog strSql 'Added by Lydia 2022/01/06 增加Log ; ex.1/5發現FCT-48403申請人日文地址夾到符號，因為當天只有通知申請案號，所以增加分案修改基本檔的log
        cnnConnection.Execute strSql
   End If
   
   'Add By Sindy 2017/3/28
   '外商分案時, 案件有 FC代理人, 且代理人國籍為日本時,
   '不論是商標案件或服務業務案件, 存檔時定稿語文欄若為空值時,
   '一律更新為3.日文
   If textTM44 <> "" Then
      strSql = "SELECT fa01,fa02,fa10 FROM fagent" & _
               " WHERE fa01=" & CNULL(Left(Me.textTM44.Text, 8)) & _
               " and fa02=" & CNULL(Mid(Me.textTM44.Text, 9, 1))
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         If Left("" & rsTmp.Fields("fa10"), 3) = "011" Then '日本
            strSql = "UPDATE TradeMark SET TM53='3'" & _
                     " WHERE TM01 = '" & m_TM01 & "' AND " & _
                        "TM02 = '" & m_TM02 & "' AND " & _
                        "TM03 = '" & m_TM03 & "' AND " & _
                        "TM04 = '" & m_TM04 & "' AND " & _
                        "TM53 is null"
            cnnConnection.Execute strSql
         End If
      End If
      rsTmp.Close
      Set rsTmp = Nothing
   End If
   '2017/3/28 END
End Sub

' 更新服務業務基本檔的相關欄位
Private Sub OnUpdateServicePractice()
   Dim strTmp As String
   Dim strSql As String
   Dim nIndex As Integer
   Dim bFirst As Boolean
   Dim bDifference As Boolean
   Dim rsTmp As New ADODB.Recordset 'Add By Sindy 2017/3/28
   
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
   If bDifference = True Then
      Pub_SeekTbLog strSql 'Added by Lydia 2022/01/06 增加Log ; ex.1/5發現FCT-48403申請人日文地址夾到符號，因為當天只有通知申請案號，所以增加分案修改基本檔的log
      cnnConnection.Execute strSql
   End If
   
   'Add By Sindy 2017/3/28
   '外商分案時, 案件有 FC代理人, 且代理人國籍為日本時,
   '不論是商標案件或服務業務案件, 存檔時定稿語文欄若為空值時,
   '一律更新為3.日文
   If textTM44 <> "" Then
      strSql = "SELECT fa01,fa02,fa10 FROM fagent" & _
               " WHERE fa01=" & CNULL(Left(Me.textTM44.Text, 8)) & _
               " and fa02=" & CNULL(Mid(Me.textTM44.Text, 9, 1))
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         If Left("" & rsTmp.Fields("fa10"), 3) = "011" Then '日本
            strSql = "UPDATE ServicePractice SET sp34='3'" & _
                     " WHERE SP01 = '" & m_TM01 & "' AND " & _
                        "SP02 = '" & m_TM02 & "' AND " & _
                        "SP03 = '" & m_TM03 & "' AND " & _
                        "SP04 = '" & m_TM04 & "' AND " & _
                        "sp34 is null"
            cnnConnection.Execute strSql
         End If
      End If
      rsTmp.Close
      Set rsTmp = Nothing
   End If
   '2017/3/28 END
End Sub

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
'edit by nick 2004/11/03
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
   Dim strCP09B  As String 'Add by Amy 2020/10/20
   Dim strCP122_Now As String 'Add by Amy 2022/11/17 急件
   Dim douStPrice As Double, douLowPrice As Double
   Dim strEP05 As String 'Add By Sindy 2024/12/9
   
   OnSaveData = True
   
'911106 nick transation
On Error GoTo CheckingErr
cnnConnection.BeginTrans

   'Modify By Cheng 2002/08/22
   '若有輸入轉本所案號
   If Me.textTM01.Text <> "" And Me.textTM02.Text <> "" Then
      'Add By Cheng 2002/09/09
      '判斷是否新增商標或服務業務基本案
      Select Case m_TM01
         Case "T", "TF", "FCT", "CFT":
            StrSQLa = "SELECT * FROM TRADEMARK WHERE " & ChgTradeMark(Me.textTM01.Text & Me.textTM02.Text & IIf(Me.textTM02_2.Visible, Me.textTM02_2.Text, "") & Me.textTM03.Text & Me.textTM04.Text)
         Case Else:
            StrSQLa = "SELECT * FROM SERVICEPRACTICE WHERE " & ChgService(Me.textTM01.Text & Me.textTM02.Text & IIf(Me.textTM02_2.Visible, Me.textTM02_2.Text, "") & Me.textTM03.Text & Me.textTM04.Text)
      End Select
      rsA.CursorLocation = adUseClient
      rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
      If rsA.RecordCount <= 0 Then
         Select Case m_TM01
            Case "T", "TF", "FCT", "CFT":
               If PUB_ReadTradeMarkData(tm(), m_TM01, m_TM02, m_TM03, m_TM04) Then
                  tm(1) = Me.textTM01.Text
                  tm(2) = Me.textTM02.Text & IIf(Me.textTM02_2.Visible, Me.textTM02_2.Text, "")
'                  tm(3) = Me.textTM03.Text
                  tm(3) = Left(Me.textTM03.Text & "0", 1)
                  tm(4) = Left(Me.textTM04.Text & "00", 2)
                  If PUB_AddNewTradeMark(tm()) Then
                    'Add By Cheng 2002/12/03
                    Else
                        GoTo CheckingErr
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
                    'Add By Cheng 2002/12/03
                    Else
                        GoTo CheckingErr
                  End If
               End If
         End Select
      'Add By Cheng 2002/12/06
      '若基本檔有資料, 若是否新案欄為'Y'更新為Null
      Else
            'Modify by Morgan 2007/5/30 要用收文號更新
            'strSQL = " Update CaseProgress Set CP31=DECODE(CP31,'Y',NULL,CP31) WHERE " & ChgCaseprogress(Me.textTM01.Text & Me.textTM02.Text & IIf(Me.textTM02_2.Visible, Me.textTM02_2.Text, "") & Me.textTM03.Text & Me.textTM04.Text)
            strSql = " Update CaseProgress Set CP31=NULL WHERE CP09='" & textCP09 & "'"
            'END 2007/5/30
            cnnConnection.Execute strSql, intI
      End If
      If rsA.State <> adStateClosed Then rsA.Close
      Set rsA = Nothing
      
      'Modify By Sindy 2013/8/26
      'strSql = "Update CASEPROGRESS SET CP01='" & Me.textTM01 & "',CP02='" & Left(Me.textTM02.Text & Me.textTM02_2.Text & "000000", 6) & "',CP03='" & Left(Me.textTM03.Text & "0", 1) & "',CP04='" & Left(Me.textTM04.Text & "00", 2) & "',CP43='' WHERE CP09='" & m_CP09 & "'"
      strSql = "Update CASEPROGRESS SET CP01='" & Me.textTM01 & "',CP02='" & Left(Me.textTM02.Text & Me.textTM02_2.Text & "000000", 6) & "',CP03='" & Left(Me.textTM03.Text & "0", 1) & "',CP04='" & Left(Me.textTM04.Text & "00", 2) & "'," & _
               "CP43=''" & IIf(chkWebApp.Visible, ",cp118='" & IIf(chkWebApp.Value = 1, "Y", "") & "'", "") & " WHERE CP09='" & m_CP09 & "'"
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
      If Frame6.Visible = True Then
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
         strTM03 = textTM03
         If IsEmptyText(strTM03) = True Then: strTM03 = "0"
         strTM04 = textTM04
         If IsEmptyText(strTM04) = True Then: strTM04 = "00"
         
         ' 檢查原始檔是否存在
         If IsDataRecordExist(strTM01, strTM02, strTM03, strTM04) = False Then
            Set objCopyTM = New ClsCopyTM
            objCopyTM.SetSrc m_TM01, m_TM02, m_TM03, m_TM04
            objCopyTM.SetDes strTM01, strTM02, strTM03, strTM04
            objCopyTM.CopyTradeMark
            Set objCopyTM = Nothing
            
            m_TM01 = strTM01
            m_TM02 = strTM02
            m_TM03 = strTM03
            m_TM04 = strTM04
            
            Select Case m_TM01
               ' 更新商標基本檔
               Case "T", "TF", "CFT", "FCT":
                  OnUpdateTradeMark
               ' 更新服務業務基本檔
               Case Else:
                  OnUpdateServicePractice
            End Select
            
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            ' 儲存優先權資料
            m_Pa(1) = m_TM01
            m_Pa(2) = m_TM02
            m_Pa(3) = m_TM03
            m_Pa(4) = m_TM04
            'edit by nickc 2007/02/06 不用 dll 了
            'objPublicData.SavePriority m_Pa, m_Priority(1), m_Priority(2), m_Priority(3)
            'Modify by Amy 2014/04/17 +, m_Priority(4), m_Priority(5)
            'Modify by Sindy 2017/10/12 +, m_Priority(6)
            ClsPDSavePriority m_Pa, m_Priority(1), m_Priority(2), m_Priority(3), m_Priority(4), m_Priority(5), m_Priority(6)
         Else
            m_TM01 = strTM01
            m_TM02 = strTM02
            m_TM03 = strTM03
            m_TM04 = strTM04
         End If
      Else
         Select Case m_TM01
            ' 更新商標基本檔
            Case "T", "TF", "CFT", "FCT":
               OnUpdateTradeMark
            ' 更新服務業務基本檔
            Case Else:
               OnUpdateServicePractice
         End Select
         
         ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
         ' 儲存優先權資料
         'edit by nickc 2007/02/06 不用 dll 了
         'objPublicData.SavePriority m_Pa, m_Priority(1), m_Priority(2), m_Priority(3)
         'Modify by Amy 2014/04/17 +, m_Priority(4), m_Priority(5)
         'Modify by Sindy 2017/10/12 +, m_Priority(6)
         ClsPDSavePriority m_Pa, m_Priority(1), m_Priority(2), m_Priority(3), m_Priority(4), m_Priority(5), m_Priority(6)
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
      'Add By Sindy 2023/9/22 寫成共用函數,有修改資料同時需記錄在接洽單上
      Dim bolModifyCRL As Boolean
      Dim strModCP10 As String, strModCP16 As String, strModCP17 As String
      bolModifyCRL = False
      If m_CP10 <> textCP10 Then
         bolModifyCRL = True
         strModCP10 = textCP10.Text
      End If
      'Add By Sindy 2022/12/5 修改費用,規費時; 發Mail通知正本財務處,副本智權人員
      If Val(Me.textCP16.Text) <> Val(Me.textCP16.Tag) Or Val(Me.textCP17.Text) <> Val(Me.textCP17.Tag) Then
         bolModifyCRL = True
         strModCP16 = Me.textCP16.Text
         strModCP17 = Me.textCP17.Text
      End If
      If bolModifyCRL = True Then
         If PUB_ModCrLCRCData(m_CP09, txtF0301, strModCP10, m_CP10 _
            , textTM10, textCP64, strModCP16, strModCP17, Me.textCP16.Tag, Me.textCP17.Tag) = False Then
            GoTo CheckingErr
         End If
      End If
      '2023/9/22 END
      ' 若有修改案件性質時,
      'modify by sonia 2024/11/6 +申請國家CFT-024746
      If textCP10 <> m_CP10 Or textTM10 <> m_TM10 Then
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
         
         'Modify By Sindy 2022/12/16
         Call ClsPDGetCaseLowPrice(m_TM01, textTM10, textCP10, douStPrice, douLowPrice, textTM08, "", txtF0301)
         ' 更新案件進度檔的標準價及底價欄位
         strSql = "UPDATE CaseProgress SET CP33 = " & douStPrice & ", " & _
                                          "CP34 = " & douLowPrice & " " & _
                  "WHERE CP09 = '" & m_CP09 & "' "
         cnnConnection.Execute strSql
         
'         'Add By Sindy 2022/12/15 有修改案件性質
'         If PUB_ModCrLCRCData(m_CP09, txtF0301, textCP10, textTM10, textCP64) = False Then
'            GoTo CheckingErr
'         End If
'         '2022/12/15 END
      End If
      
      ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
      ' 案件性質為查名, 補收款, 後金時更新收文日為系統日
      Select Case textCP10
         Case "001", "705", "909":
            'Modify By Sindy 2012/12/22 陳金蓮提
            If textCP10 = "001" And m_TM10 <> "000" Then
               '不考慮系統類別,只要案件性質是"001"查名者,申請國家非"000"者不上發文日
            Else
            '2012/12/22 End
               strSql = "UPDATE CaseProgress SET CP27 = " & DBDATE(SystemDate()) & " " & _
                        "WHERE CP09 = '" & m_CP09 & "' "
               cnnConnection.Execute strSql
            End If
         'Add By Sindy 2020/10/20 陳述意見書
         Case "210":
            strExc(0) = "select cp09,cp10 from caseprogress where cp01='" & m_TM01 & "' and cp02='" & m_TM02 & "'" & _
                        " and cp03='" & m_TM03 & "' and cp04='" & m_TM04 & "' and cp158=0 and cp159=0" & _
                        " and cp10='214'"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 0 Then
               strCP09B = AutoNo("B", 6)
               strSql = "insert into caseprogress(cp01,cp02,cp03,cp04,cp05,cp06,cp09,cp10,cp12,cp13,cp14,cp20,cp26,cp32,cp43) " & _
                              "values (" & CNULL(m_TM01) & "," & CNULL(m_TM02) & "," & CNULL(m_TM03) & _
                              "," & CNULL(m_TM04) & "," & CNULL(strSrvDate(1)) & "," & CompWorkDay(4, DBDATE(textCP05)) & "," & CNULL(strCP09B) & ",'214'," & _
                              CNULL(GetSalesArea(IIf(m_TM01 = "FCT", PUB_GetFCTSalesNo(m_TM01, m_TM02, m_TM03, m_TM04), PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04)))) & "," & _
                              CNULL(IIf(m_TM01 = "FCT", PUB_GetFCTSalesNo(m_TM01, m_TM02, m_TM03, m_TM04), PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04))) & "," & _
                              CNULL(textCP14) & ",'N','N','N'," & CNULL(m_CP09) & ")"
               cnnConnection.Execute strSql
            End If
      End Select
      
      '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
      ' 計算承辦期限
      'strDay = GetWorkDays(m_TM01, m_TM10, textCP10)
      'If IsEmptyText(strDay) = False Then
      '   strDate = DBDATE(DateSerial(Val(DBYEAR(textCP05)), Val(DBMONTH(textCP05)), Val(DBDAY(textCP05)) + Val(strDay)))
      '   strSQL = "UPDATE CaseProgress SET CP48 = " & strDate & " " & _
      '            "WHERE CP09 = '" & m_CP09 & "' "
      '   cnnConnection.Execute strSQL
      'End If
      
      '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
      'Modify By Cheng 2002/09/18
      ' 若有輸入查名本所案號時, 更新此該查名本所案號的案件進度資料的本所案號為本案的本所案號
'      ' 若有輸入查名收文號時, 更新此該查名收文號的案件進度資料的本所案號為本案的本所案號
'      If IsEmptyText(textCP09_S) = False Then
'         strSQL = "UPDATE CaseProgress SET CP01 = '" & m_TM01 & "', " & _
'                                          "CP02 = '" & m_TM02 & "', " & _
'                                          "CP03 = '" & m_TM03 & "', " & _
'                                          "CP04 = '" & m_TM04 & "' " & _
'                  "WHERE CP09 = '" & textCP09_S & "' "
'         cnnConnection.Execute strSQL
'      End If
      If textCP09_S.Text = "S" And IsEmptyText(textCP09_S1) = False Then
         'add by nickc 2005/10/28 清未結餘的可結餘日期
         strSql = "UPDATE CaseProgress SET cp109=null " & _
                  "WHERE " & ChgCaseprogress(Me.textCP09_S.Text & Me.textCP09_S1.Text & Me.textCP09_S2.Text & Me.textCP09_S3.Text) & " and cp59 is null "
         cnnConnection.Execute strSql
         'edit by nickc 2006/07/18 加入 cp31=null
         strSql = "UPDATE CaseProgress SET cp31=null " & _
                  "WHERE " & ChgCaseprogress(Me.textCP09_S.Text & Me.textCP09_S1.Text & Me.textCP09_S2.Text & Me.textCP09_S3.Text) & " "
         cnnConnection.Execute strSql
         
         strSql = "UPDATE CaseProgress SET CP01 = '" & m_TM01 & "', CP02 = '" & m_TM02 & "', " & _
                                          "CP03 = '" & m_TM03 & "', CP04 = '" & m_TM04 & "', " & _
                                          "CP64=CP64||Decode(CP64,Null,'','，')||'" & "原查名本所案號：" & Me.textCP09_S.Text & "-" & Me.textCP09_S1.Text & "-" & Left(Me.textCP09_S2.Text & "0", 1) & "-" & Left(Me.textCP09_S3.Text & "00", 2) & "' " & _
                  "WHERE " & ChgCaseprogress(Me.textCP09_S.Text & Me.textCP09_S1.Text & Me.textCP09_S2.Text & Me.textCP09_S3.Text)
         cnnConnection.Execute strSql
        'Add By Cheng 2003/06/16
        strSql = "Update ServicePractice Set SP18=SP18||Decode(SP18,Null,'','，')||'轉入商標：" & m_TM01 & "-" & m_TM02 & "-" & m_TM03 & "-" & m_TM04 & "' Where " & ChgService(Me.textCP09_S.Text & Me.textCP09_S1.Text & Left(Me.textCP09_S2.Text & "0", 1) & Left(Me.textCP09_S3.Text & "00", 2))
        cnnConnection.Execute strSql
        '2005/4/18 ADD BY SONIA 1~4欄原查名本所案號,5~8欄新商標本所案號
        If PUB_UpdOther(Me.textCP09_S.Text, Me.textCP09_S1.Text, Left(Me.textCP09_S2.Text & "0", 1), Left(Me.textCP09_S3.Text & "00", 2), m_TM01, m_TM02, m_TM03, m_TM04) = False Then
           GoTo CheckingErr
        End If
        '2005/4/18 END
      End If
      
      '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
      ' 若案件性質為救濟程序時或爭議程序更新基本檔的欄位
      Select Case Mid(textCP10, 1, 1)
         ' 救濟程序
         Case "4":
            Select Case m_TM01:
               Case "T", "TF", "FCT", "CFT":
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
               Case "T", "TF", "FCT", "CFT":
                  strSql = "UPDATE TradeMark SET TM19 = 'Y' " & _
                           "WHERE TM01 = '" & m_TM01 & "' AND " & _
                                 "TM02 = '" & m_TM02 & "' AND " & _
                                 "TM03 = '" & m_TM03 & "' AND " & _
                                 "TM04 = '" & m_TM04 & "' "
                  cnnConnection.Execute strSql
               Case Else:
            End Select
      End Select
   
      '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
      ' 更新使用者所選取的本案期限資料
      For nIndex = 1 To grdList.Rows - 1
         ' 判斷該列是否有被選取
         If grdList.TextMatrix(nIndex, 0) = "V" Then
            strNP07 = grdList.TextMatrix(nIndex, 9)
            strNP22 = grdList.TextMatrix(nIndex, 10)
            'Modified by Lydia 2021/08/31 +更新NP24
            strSql = "UPDATE NextProgress SET NP06 = 'Y', NP24='" & m_CP09 & "'  " & _
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
      '92.3.13 CANCEL BY SONIA 與外商再討論過後, 決定刪除此功能
      '''''''''''''''''''''''''''''''''''''''''''''''''''''''
      ' 若此案為母案且申請國家為歐盟, 美國或新加坡時
      'If m_TM03 = "0" And m_TM04 = "00" And (IsEmptyText(textTM01) = True Or IsEmptyText(textTM02) = True) And (textTM10 = "239" Or textTM10 = "101" Or textTM10 = "014") Then
      '   If textTM10 = "239" And textCP10 = "101" And IsEmptyText(m_strCountry) = False Then
      '      If IsEmptyText(textTM09) = False Then
      '         For nIndex = 1 To GetSubStringCount(textTM09)
      '            strProduct = GetSubString(textTM09, nIndex)
      '            For nSubIndex = 1 To GetSubStringCount(m_strCountry)
      '               strCountry = GetSubString(m_strCountry, nSubIndex)
      '               Set objCopyTM = New ClsCopyTM
      '               objCopyTM.SetSrc m_TM01, m_TM02, m_TM03, m_TM04
      '               objCopyTM.SetDes m_TM01, m_TM02, CStr(Val(m_TM03 + nIndex)), Format(CStr(Val(m_TM04) + nSubIndex), "00")
      '               objCopyTM.SetExtraField "TM09", strProduct
      '               objCopyTM.SetExtraField "TM10", strCountry
      '               objCopyTM.CopyTradeMark
      '               Set objCopyTM = Nothing
      '            Next nSubIndex
      '         Next nIndex
      '      Else
      '         For nSubIndex = 1 To GetSubStringCount(m_strCountry)
      '            strCountry = GetSubString(m_strCountry, nSubIndex)
      '            Set objCopyTM = New ClsCopyTM
      '            objCopyTM.SetSrc m_TM01, m_TM02, m_TM03, m_TM04
      '            objCopyTM.SetDes m_TM01, m_TM02, m_TM03, Format(CStr(Val(m_TM04) + nSubIndex), "00")
      '            objCopyTM.SetExtraField "TM10", strCountry
      '            objCopyTM.CopyTradeMark
      '            Set objCopyTM = Nothing
      '         Next nSubIndex
      '      End If
      '   Else
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
      'End If
      '92.3.13 END
        'Add By Cheng 2004/04/14
        '更新分割案件關係資料
        If m_CP10 = "308" Then
            If PUB_UpdateDivisionCase(m_TM01, m_TM02, m_TM03, m_TM04, Me.txtDivCaseNo(0).Text, Me.txtDivCaseNo(1).Text & Me.txtDivCaseNo(2).Text, Me.txtDivCaseNo(3).Text, Me.txtDivCaseNo(4).Text) = False Then
                GoTo CheckingErr
            End If
        End If
        'End
        'add by nick 2004/09/27 若沒費用，則不請款不開收據
        If Val(textCP16.Text) = 0 Then
            strSql = "UPDATE CaseProgress SET cp20='N',cp32='N' " & _
                     "WHERE CP09 = '" & m_CP09 & "' "
            cnnConnection.Execute strSql
        Else
            strSql = "UPDATE CaseProgress SET cp20=null,cp32=null " & _
                     "WHERE CP09 = '" & m_CP09 & "' "
            cnnConnection.Execute strSql
        End If
        
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
   End If
   
   ' 通知前畫面該筆收文資料已存檔
   frm030201_01.SetDataComplete m_CP09
   
   'add by nickc 2005/03/17 加入加乘註記及寄件值
   m_CP98 = "": m_CP101 = "": m_CP104 = ""
   If PUB_GetFlagValue(m_CP09, m_CP98, m_CP101, m_CP104) = True Then
      strSql = "update caseprogress set cp98=" & m_CP98 & ",cp101=" & m_CP101 & ",cp104=" & m_CP104 & " WHERE CP09 = '" & m_CP09 & "' "
      cnnConnection.Execute strSql
   End If
   'PUB_UpdateCaseValue m_CP09 'Remove by Morgan 2005/4/13 改由 trigger 更新
   
   'Add By Sindy 2014/2/14 更新第四頁及第五頁欄位資料
   '申請人
   If Trim(textTM23.Text) <> "" And Frame2.Enabled = True Then
      If textCU58.Tag <> textCU58 Or textCU59.Tag <> textCU59 Or textCU60.Tag <> textCU60 Or _
         textCU61.Tag <> textCU61 Or textCU62.Tag <> textCU62 Or textCU63.Tag <> textCU63 Or _
         textCU146.Tag <> textCU146 Or textCU147.Tag <> textCU147 Or textCU151.Tag <> textCU151 Or _
         textCU149.Tag <> textCU149 Or txtCU(126).Tag <> txtCU(126).Text Then
         'Add By Sindy 2014/3/13
         If Trim(textCU147.Tag) <> "" And Trim(textCU147) = "" Then
            If MsgBox("確定要清空『商標固定請款對象』資料？", vbExclamation + vbYesNo + vbDefaultButton2, "重要訊息！") = vbNo Then
               SSTab1.Tab = 3
               textCU147.SetFocus
               cnnConnection.RollbackTrans
               OnSaveData = False
               Exit Function
            End If
         End If
         '2014/3/13 END
         strTemp = ""
         If textCU58.Tag <> textCU58 Then strTemp = strTemp & ",cu58=" & CNULL(ChgSQL(textCU58))
         If textCU59.Tag <> textCU59 Then strTemp = strTemp & ",cu59=" & CNULL(ChgSQL(textCU59))
         If textCU60.Tag <> textCU60 Then strTemp = strTemp & ",cu60=" & CNULL(ChgSQL(textCU60))
         If textCU61.Tag <> textCU61 Then strTemp = strTemp & ",cu61=" & CNULL(ChgSQL(textCU61))
         If textCU62.Tag <> textCU62 Then strTemp = strTemp & ",cu62=" & CNULL(ChgSQL(textCU62))
         If textCU63.Tag <> textCU63 Then strTemp = strTemp & ",cu63=" & CNULL(ChgSQL(textCU63))
         If textCU146.Tag <> textCU146 Then strTemp = strTemp & ",cu146=" & CNULL(ChgSQL(textCU146))
         If textCU149.Tag <> textCU149 Then strTemp = strTemp & ",cu149=" & CNULL(ChgSQL(textCU149)) 'Add By Sindy 2020/3/4
         If txtCU(126).Tag <> txtCU(126) Then strTemp = strTemp & ",cu126=" & CNULL(ChgSQL(txtCU(126))) 'Add By Sindy 2022/3/16
         If textCU147.Tag <> textCU147 Then
            If Trim(textCU147) <> "" Then
               strTemp = strTemp & ",cu147=" & CNULL(textCU147 & String(9 - Len(textCU147), "0"))
            Else
               strTemp = strTemp & ",cu147=null"
            End If
         End If
         If textCU151.Tag <> textCU151 Then
            If Trim(textCU151) <> "" Then
               strTemp = strTemp & ",cu151=" & CNULL(textCU151 & String(9 - Len(textCU151), "0"))
            Else
               strTemp = strTemp & ",cu151=null"
            End If
         End If
         strSql = "update customer set " & Mid(strTemp, 2) & _
                  " where cu01='" & Left(textTM23, 8) & "' and cu02='" & Right(textTM23, 1) & "'"
         Pub_SeekTbLog strSql
         cnnConnection.Execute strSql
      End If
   End If
   '代理人
   If Trim(textTM44.Text) <> "" And Frame3.Enabled = True Then
      If textFA07.Tag <> textFA07 Or textFA08.Tag <> textFA08 Or textFA09.Tag <> textFA09 Or _
         textFA52.Tag <> textFA52 Or textFA53.Tag <> textFA53 Or textFA54.Tag <> textFA54 Or _
         textFA106.Tag <> textFA106 Or textFA107.Tag <> textFA107 Or textFA111.Tag <> textFA111 Or _
         textFA109.Tag <> textFA109 Or txtFA(91).Tag <> txtFA(91).Text Then
         strTemp = ""
         If textFA07.Tag <> textFA07 Then strTemp = strTemp & ",FA07=" & CNULL(ChgSQL(textFA07))
         If textFA08.Tag <> textFA08 Then strTemp = strTemp & ",FA08=" & CNULL(ChgSQL(textFA08))
         If textFA09.Tag <> textFA09 Then strTemp = strTemp & ",FA09=" & CNULL(ChgSQL(textFA09))
         If textFA52.Tag <> textFA52 Then strTemp = strTemp & ",FA52=" & CNULL(ChgSQL(textFA52))
         If textFA53.Tag <> textFA53 Then strTemp = strTemp & ",FA53=" & CNULL(ChgSQL(textFA53))
         If textFA54.Tag <> textFA54 Then strTemp = strTemp & ",FA54=" & CNULL(ChgSQL(textFA54))
         If textFA106.Tag <> textFA106 Then strTemp = strTemp & ",FA106=" & CNULL(ChgSQL(textFA106))
         If textFA109.Tag <> textFA109 Then strTemp = strTemp & ",FA109=" & CNULL(ChgSQL(textFA109)) 'Add By Sindy 2020/3/4
         If txtFA(91).Tag <> txtFA(91) Then strTemp = strTemp & ",FA91=" & CNULL(ChgSQL(txtFA(91))) 'Add By Sindy 2022/3/16
         If textFA107.Tag <> textFA107 Then
            If Trim(textFA107) <> "" Then
               strTemp = strTemp & ",FA107=" & CNULL(textFA107 & String(9 - Len(textFA107), "0"))
            Else
               strTemp = strTemp & ",FA107=null"
            End If
         End If
         If textFA111.Tag <> textFA111 Then
            If Trim(textFA111) <> "" Then
               strTemp = strTemp & ",FA111=" & CNULL(textFA111 & String(9 - Len(textFA111), "0"))
            Else
               strTemp = strTemp & ",FA111=null"
            End If
         End If
         strSql = "update fagent set " & Mid(strTemp, 2) & _
                  " where FA01='" & Left(textTM44, 8) & "' and FA02='" & Right(textTM44, 1) & "'"
         Pub_SeekTbLog strSql
         cnnConnection.Execute strSql
      End If
   End If
   '案件基本檔
   Select Case m_TM01
      Case "T", "TF", "CFT", "FCT":
         If textTM56_1.Tag <> textTM56_1 Or textTM69_1.Tag <> textTM69_1 Or _
            textTM38.Tag <> textTM38 Or textTM39.Tag <> textTM39 Or textTM40.Tag <> textTM40 Or _
            textTM41.Tag <> textTM41 Or textTM42.Tag <> textTM42 Or textTM43.Tag <> textTM43 Or _
            textTM46.Tag <> textTM46 Or textTM127.Tag <> textTM127 Or textTM121.Tag <> textTM121 Then
            strTemp = ""
            If textTM56_1.Tag <> textTM56_1 Then
               If Trim(textTM56_1) <> "" Then
                  strTemp = strTemp & ",TM56=" & CNULL(textTM56_1 & String(9 - Len(textTM56_1), "0"))
               Else
                  strTemp = strTemp & ",TM56=null"
               End If
            End If
            If textTM69_1.Tag <> textTM69_1 Then
               If Trim(textTM69_1) <> "" Then
                  strTemp = strTemp & ",TM69=" & CNULL(textTM69_1 & String(9 - Len(textTM69_1), "0"))
               Else
                  strTemp = strTemp & ",TM69=null"
               End If
            End If
            If textTM46.Tag <> textTM46 Then strTemp = strTemp & ",TM46=" & CNULL(ChgSQL(textTM46)) 'Add By Sindy 2020/3/4
            If textTM121.Tag <> textTM121 Then strTemp = strTemp & ",TM121=" & CNULL(ChgSQL(textTM121)) 'Add By Sindy 2022/3/16
            If textTM127.Tag <> textTM127 Then strTemp = strTemp & ",TM127=" & CNULL(ChgSQL(textTM127)) 'Add By Sindy 2020/3/4
            If textTM38.Tag <> textTM38 Then strTemp = strTemp & ",TM38=" & CNULL(ChgSQL(textTM38))
            If textTM39.Tag <> textTM39 Then strTemp = strTemp & ",TM39=" & CNULL(ChgSQL(textTM39))
            If textTM40.Tag <> textTM40 Then strTemp = strTemp & ",TM40=" & CNULL(ChgSQL(textTM40))
            If textTM41.Tag <> textTM41 Then strTemp = strTemp & ",TM41=" & CNULL(ChgSQL(textTM41))
            If textTM42.Tag <> textTM42 Then strTemp = strTemp & ",TM42=" & CNULL(ChgSQL(textTM42))
            If textTM43.Tag <> textTM43 Then strTemp = strTemp & ",TM43=" & CNULL(ChgSQL(textTM43))
            strSql = "update trademark set " & Mid(strTemp, 2) & _
                     " where tm01='" & m_TM01 & "' and tm02='" & m_TM02 & "' and tm03='" & m_TM03 & "' and tm04='" & m_TM04 & "'"
            Pub_SeekTbLog strSql
            cnnConnection.Execute "begin user_data.user_enabled:=1; " & strSql & "; end;"
         End If
      'Add By Sindy 2014/5/16
      Case Else:
         If textTM56_1.Tag <> textTM56_1 Or textTM69_1.Tag <> textTM69_1 Or _
            textTM38.Tag <> textTM38 Or textTM41.Tag <> textTM41 Or textTM46.Tag <> textTM46 Or _
            textTM127.Tag <> textTM127 Or textTM121.Tag <> textTM121 Then
            strTemp = ""
            If textTM56_1.Tag <> textTM56_1 Then
               If Trim(textTM56_1) <> "" Then
                  strTemp = strTemp & ",SP37=" & CNULL(textTM56_1 & String(9 - Len(textTM56_1), "0"))
               Else
                  strTemp = strTemp & ",SP37=null"
               End If
            End If
            If textTM69_1.Tag <> textTM69_1 Then
               If Trim(textTM69_1) <> "" Then
                  strTemp = strTemp & ",SP67=" & CNULL(textTM69_1 & String(9 - Len(textTM69_1), "0"))
               Else
                  strTemp = strTemp & ",SP67=null"
               End If
            End If
            If textTM46.Tag <> textTM46 Then strTemp = strTemp & ",SP33=" & CNULL(ChgSQL(textTM46)) 'Add By Sindy 2020/3/4
            If textTM121.Tag <> textTM121 Then strTemp = strTemp & ",SP80=" & CNULL(ChgSQL(textTM121)) 'Add By Sindy 2022/3/16
            If textTM127.Tag <> textTM127 Then strTemp = strTemp & ",SP84=" & CNULL(ChgSQL(textTM127)) 'Add By Sindy 2020/3/4
            If textTM38.Tag <> textTM38 Then strTemp = strTemp & ",SP30=" & CNULL(ChgSQL(textTM38))
            If textTM41.Tag <> textTM41 Then strTemp = strTemp & ",SP75=" & CNULL(ChgSQL(textTM41)) 'Add by Sindy 2021/5/5
            strSql = "update servicePractice set " & Mid(strTemp, 2) & _
                     " where sp01='" & m_TM01 & "' and sp02='" & m_TM02 & "' and sp03='" & m_TM03 & "' and sp04='" & m_TM04 & "'"
            Pub_SeekTbLog strSql
            cnnConnection.Execute "begin user_data.user_enabled:=1; " & strSql & "; end;"
         End If
      '2014/5/16 END
   End Select
   '2014/2/14 END
   
    'Added by Morgan 2022/12/23
    '註冊證形式
    If textTM136.Visible And textTM136.Tag <> textTM136 Then
      strSql = "Update trademark Set tm136='" & textTM136 & "' " & _
                  "WHERE tm01='" & m_TM01 & "' and tm02='" & m_TM02 & "'" & _
                   " and tm03='" & m_TM03 & "' and tm04='" & m_TM04 & "'"
      cnnConnection.Execute strSql
    End If
    'end 2022/12/23
    
    'Added by Lydia 2020/05/20 法律所案源收文：如果案件性質或申請國家有變化,則需要對應分案
    If strSrvDate(1) >= 法律所案源收文啟用日 And m_TM01 = "FCT" And m_LOS07 = "" Then  '排除已放棄的案源
        'Modified by Lydia 2020/07/23 重新整理: 因為案源收文已設定不可變更案件性質和申請國家,所以只要判斷有案源
        'If textTM10.Text <> m_TM10 Or m_CP10 <> textCP10.Text Or (m_LOS15 = "" And txtLOSagree = "Y") Then
        '    Call PUB_UpdateCP10toPT(m_TM01, m_TM02, m_TM03, m_TM04, m_CP09, m_CP10, m_TM10, textCP10.Text, textTM10.Text, textCP06.Text, textCP13, textTM23, IIf(m_LOS15 = "" And txtLOSagree = "Y", True, False))
        'End If
        '
        'If m_CP14 = "" And textCP14.Text <> "" Then
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
        If m_LOS15 <> "" And m_CP14 = "" And textCP14.Text <> "" Then
            Call PUB_UpdateLOS01(m_TM01, m_TM02, m_TM03, m_TM04, m_CP09, textTM23 & "," & textSP58 & "," & textSP59 & "," & textTM80 & "," & textTM81, txtLOSagree)
        End If
        'end 2020/07/23
    End If
    'end 2020/05/20
    
    'Add by Amy 2022/11/17 +CP122 急件
    If Check11.Visible = True Then
        strCP122_Now = "N"
        If Check11.Value = 1 Then strCP122_Now = "Y"
        'Memo DB資料若為null,回存N,避免與內商混淆
        If strCP122 <> strCP122_Now Then
            strSql = "Update CaseProgress Set CP122=" & CNULL(strCP122_Now) & " Where cp09='" & m_CP09 & "' "
            cnnConnection.Execute strSql
        End If
    End If
    'end 2022/11/17
    
    'Add By Sindy 2022/12/26 分案後要發通知信
    If strSrvDate(1) >= 接洽單電子收文啟用日 Then
      'Modify By Sindy 2023/1/3 + And Left(PUB_GetStaffST15(textCP13, 1), 2) <> "F1"
      'Modify By Sindy 2023/11/1 系統別為CFT、S(台灣案除外)之所有收文案件,分案後都要發mail通知CFT承辦組人員
      'modify by sonia 2023/12/22改申請國家非台灣者都發,否則CFC沒發到
      'If textCP14.Tag <> textCP14.Text And Trim(textCP14.Text) <> "" And _
         (m_TM01 = "CFT" Or (m_TM01 = "S" And m_TM10 <> "000")) Then
      'Modify By Sindy 2024/10/28 strSrvDate(1) >= 外商承辦歷程啟用日
      If textCP14.Tag <> textCP14.Text And Trim(textCP14.Text) <> "" And (m_TM10 <> "000" Or strSrvDate(1) >= 外商承辦歷程啟用日) Then
         'Modify By Sindy 2024/12/9 FC案承辦人=程序人員時,不用發mail
         strEP05 = ""
         strExc(0) = "select * from ENGINEERPROGRESS where ep02='" & textCP09 & "'"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            strEP05 = "" & RsTemp.Fields("EP05")
         End If
         If strEP05 = "" Then strEP05 = textCP14.Text
         If Not ((m_TM01 = "FCT" Or (m_TM01 = "S" And m_TM10 = "000")) _
                 And PUB_GetST03(strEP05) = "F12") Then
         '2024/12/9 END
            strExc(1) = "分案通知 " & textTM05_1 & "（本所案號：" & m_TM01 & "-" & m_TM02 & IIf(m_TM03 & m_TM04 = "000", "", "-" & m_TM03 & "-" & m_TM04) & "）"
            'Modify By Sindy 2023/1/31 案號後面增加顯示申請國家
            'Modify By Sindy 2023/2/2 只需發給承辦人員,智權人員不需要通知
            'Modify By Sindy 2023/3/8 當CFT收文之案件為「新案」，但案件性質非「申請」，請在分案通知加上「＊請補案件基本資料」
            'Modify By Sindy 2023/10/19 若法定期限≦系統日提醒承辦人,在給承辦人的EMAIL中法定期限欄後面加註：已過期
            'Modify By Sindy 2023/12/11 +指定送件日期
            'Modify By Sindy 2024/5/3 + m_strRefText
            'Modify By Sindy 2024/12/9 textCP14 + " " + textCP14_2 => strEP05 + " " + GetPrjSalesNM(strEP05)
            strExc(10) = "本所案號：" + m_TM01 & "-" & m_TM02 & "-" & m_TM03 & "-" & m_TM04 + vbCrLf + _
                         "申請國家：" + textTM10 + " " + textTM10_2 + vbCrLf + _
                         "案件名稱：" + textTM05_1 + vbCrLf + _
                         "案件性質：" + textCP10_2 + vbCrLf + _
                         "收文日　：" + ChangeWStringToTDateString(DBDATE(textCP05)) + vbCrLf + _
                         "智權人員：" + textCP13 + " " + textCP13_2 + vbCrLf + _
                         "承辦人　：" + strEP05 + " " + GetPrjSalesNM(strEP05) + vbCrLf + _
                         "承辦期限：" + ChangeWStringToTDateString(DBDATE(textCP48)) + vbCrLf + _
                         IIf(Trim(textCP142) <> "", "指定送件日期：" + ChangeWStringToTDateString(DBDATE(textCP142)) + IIf(Option1(0).Value = True, "當天", IIf(Option1(1).Value = True, "之後", "之後")) + vbCrLf, "") + _
                         "本所期限：" + ChangeWStringToTDateString(DBDATE(textCP06)) + vbCrLf + _
                         "法定期限：" + ChangeWStringToTDateString(DBDATE(textCP07)) + IIf(Val(textCP07) <= Val(strSrvDate(2)) And Val(textCP07) > 0, " (已過期)", "") + vbCrLf + _
                         "是否急件：" + IIf(Check11.Value = 1, "是", "否") & vbCrLf & _
                         IIf(m_CP31 = "Y" And textCP10 <> "101", "★請補案件基本資料" & vbCrLf, "") & _
                         IIf(m_strRefText <> "", "★" & m_strRefText & "★" & vbCrLf, "")
            'Modify By Sindy 2024/1/30
            'CFT案件性質為「申請」，在分案後，除通知承辦人員外，請同時副本給程序人員
            strExc(9) = ""
            If m_TM01 = "CFT" And Trim(textCP10) = "101" Then
               'Modify By Sindy 2024/2/2
               If PUB_GetST06(textCP13) = "1" Then '北所
                  strExc(9) = Pub_GetSpecMan("CFT程序人員-北所")
               Else
                  strExc(9) = Pub_GetSpecMan("CFT程序人員-分所")
               End If
               '2024/2/2 END
            End If
            '2024/1/30 END
            'Modify By Sindy 2024/1/30 +mc09 :副本
            'Modify By Sindy 2024/4/22 +程序人員改至收受者;因人員休假時才會轉發職代
            'Modify By Sindy 2024/12/9 textCP14.Text => strEP05
            strSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08,mc09)" & _
               " values ('" & strUserNum & "','" & strEP05 & ";" & strExc(9) & "',to_char(sysdate,'yyyymmdd'),to_char(sysdate,'hh24miss')" & _
               ",'" & ChgSQL(strExc(1)) & "','" & ChgSQL(strExc(10)) & "','')"
            cnnConnection.Execute strSql, intI
         End If
      End If
    End If
    '2022/12/26 END
    
'911106 nick transation
    cnnConnection.CommitTrans
    Exit Function
    
CheckingErr:
    cnnConnection.RollbackTrans
    MsgBox (Err.Description)
    OnSaveData = False
End Function

Private Function CheckDataValid() As Boolean
Dim strTit As String
Dim strMsg As String
Dim nResponse
Dim nVal As Currency
'Add By Cheng 2002/09/18
'Dim rsTmp As New ADODB.Recordset
'Dim strSql As String
Dim ii As Integer
Dim strCode(0 To 7) As String
   
   CheckDataValid = False
    'Modify By Cheng 2002/11/22
    '若非執行轉本所案號功能
   If Me.textTM01.Text = "" Or Me.textTM02.Text = "" Then
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
        Select Case m_TM01
        Case "T", "FCT", "CFT", "TF", "S"
            ' 案件名稱不可空白
            If IsEmptyText(textTM05_1) = True Then
               strTit = "檢核資料"
               strMsg = "案件名稱不可空白"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textTM05_1.SetFocus
               GoTo EXITSUB
            End If
        Case Else
            ' 案件名稱不可同時為空白
            If IsEmptyText(textTM05) = True And IsEmptyText(textTM06) = True And IsEmptyText(textTM07) = True Then
               strTit = "檢核資料"
               strMsg = "案件名稱(中)(英)(日)不可全為空白"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textTM05.SetFocus
               GoTo EXITSUB
            End If
        End Select
         
         'Add By Sindy 2014/12/16 分案的承辦期限更改時,檢查下列條件
         If textCP48.Tag <> textCP48.Text And IsEmptyText(textCP48) = False Then
            '承辦期限若小於系統日
            If Val(textCP48) < Val(strSrvDate(2)) Then
               If MsgBox("承辦期限小於系統日，是否正確？", vbExclamation + vbYesNo + vbDefaultButton1, "重要訊息！") = vbNo Then
                  textCP48.SetFocus
                  GoTo EXITSUB
               End If
            End If
            '承辦期限若大於本所期限
            'Modify By Sindy 2023/10/19 增加案件性質為延展102且本所期限欄被鎖住時,不必檢查
            If Not (textCP10 = "102" And textCP06.Enabled = False And m_TM01 = "CFT") Then
            '2023/10/19 END
               If IsEmptyText(textCP06) = False Then
                  If Val(textCP48) > Val(textCP06) Then
                     If MsgBox("承辦期限大於本所期限，是否正確？", vbExclamation + vbYesNo + vbDefaultButton1, "重要訊息！") = vbNo Then
                        textCP48.SetFocus
                        GoTo EXITSUB
                     End If
                  End If
               End If
            End If
         End If
'         ' 承辦期限不可超過本所期限
'         If IsEmptyText(textCP06) = False And IsEmptyText(textCP48) = False Then
'            If Val(textCP48) > Val(textCP06) Then
'               strTit = "檢核資料"
'               strMsg = "承辦期限不可超過本所期限"
'               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
'               textCP06.SetFocus
'               GoTo EXITSUB
'            End If
'         End If
         '2014/12/16 END

       'Add By Sindy 2022/11/22
       If strSrvDate(1) >= 接洽單電子收文啟用日 Then
          'Modify By Sindy 2023/4/12 + , , , , textCP09
          'Modify By Sindy 2023/10/19 bolOnlyCountCP06參數原為False 改+IIf(textCP06.Enabled = True, False, True)
          strExc(10) = "" 'Add By Sindy 2023/10/24
          If PUB_CRLUseCP07CheckCP06(m_CP31, textTM10, m_TM01, textCP10, textCP06, textCP07, strExc(10), , _
                                 IIf(textCP06.Enabled = True, False, True), textCP09) = False Then
             If textCP06.Enabled = True Then textCP06.SetFocus
             GoTo EXITSUB
          'Add By Sindy 2023/10/19
          ElseIf textCP10 = "102" Then
            '有計算出本所期限,填入日期鎖住欄位
            If Val(strExc(10)) > 0 Then
               textCP06.Text = strExc(10)
            End If
            '2023/10/19 END
          End If
       End If
       '2022/11/22 END
       
       ' 案件性質為延展或延期時本所期限及法定期限不可為空白
       '2009/3/30 MODIFY BY SONIA 加105使用宣誓
       '2013/10/7 MODIFY BY SONIA 加729復權T-183459
       If textCP10 = "102" Or textCP10 = "303" Or textCP10 = "105" Or textCP10 = "729" Then
          If IsEmptyText(textCP06) = True Then
             strTit = "檢核資料"
            'Modify By Cheng 2002/10/29
    '         strMsg = "案件性質為延展, 本所期限不可為空白"
             strMsg = "案件性質為延展, 使用宣誓, 復權或延期時, 本所期限不可為空白"
             nResponse = MsgBox(strMsg, vbOKOnly, strTit)
             'textCP06.SetFocus
             If textCP06.Visible = True And textCP06.Enabled = True Then textCP06.SetFocus
             GoTo EXITSUB
          End If
          If IsEmptyText(textCP07) = True Then
             strTit = "檢核資料"
            'Modify By Cheng 2002/10/29
    '         strMsg = "案件性質為延展, 法定期限不可為空白"
             strMsg = "案件性質為延展, 使用宣誓, 復權或延期時, 法定期限不可為空白"
             nResponse = MsgBox(strMsg, vbOKOnly, strTit)
             If textCP07.Visible = True And textCP07.Enabled = True Then textCP07.SetFocus
             GoTo EXITSUB
          End If
       End If
       If IsEmptyText(textCP06) = False And IsEmptyText(textCP07) = False Then
          If Val(textCP06) > Val(textCP07) Then
             strTit = "檢核資料"
             strMsg = "本所期限的日期不可超過法定期限的日期"
             nResponse = MsgBox(strMsg, vbOKOnly, strTit)
             If textCP06.Visible = True And textCP06.Enabled = True Then textCP06.SetFocus
             GoTo EXITSUB
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
       
       'Add By Sindy 2011/01/06
       '外商(S)申請人1或FC代理人至少要輸入一個
       '其他的一定要輸入申請人1
       If m_TM01 = "S" Then
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
       
       ' 商標種類不可空白
       'If IsEmptyText(textTM08) = True Then
       '   strTit = "檢核資料"
       '   strMsg = "商標種類不可空白"
       '   nResponse = MsgBox(strMsg, vbOKOnly, strTit)
       '   GoTo ExitSub
       'End If
       ' 卷宗性質不可空白
       'If IsEmptyText(textTM28) = True Then
       '   strTit = "檢核資料"
       '   strMsg = "卷宗性質不可空白"
       '   nResponse = MsgBox(strMsg, vbOKOnly, strTit)
       '   GoTo ExitSub
       'End If
        ' 點數=(費用-規費) / 1000
        If IsEmptyText(textCP16) = False Then
'           If Val(textCP18) <> Format(((Val(textCP16) - Val(textCP17)) / 1000), "0.0") Then
'              strTit = "檢核資料"
'              strMsg = "點數應為 " & CStr(Format(((Val(textCP16) - Val(textCP17)) / 1000), "0.0"))
'              nResponse = MsgBox(strMsg, vbOKOnly, strTit)
'              textCP16.SetFocus
'              GoTo EXITSUB
'           End If
            'Add By Sindy 2012/9/10
            If Format(((Val(textCP16) - Val(textCP17)) / 1000), "0.0") <> Format(Val(textCP18), "0.0") Then
               strTit = "檢核資料"
               strMsg = "(費用 - 規費) / 1000 <> 點數 !!"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textCP16.SetFocus
               GoTo EXITSUB
            End If
            '2012/9/10 End
        Else
           nVal = 0
           If IsEmptyText(textCP18) = False Then
              If textCP18 <> "0" Then
                 strTit = "檢核資料"
                 strMsg = "點數應為空白或0"
                 nResponse = MsgBox(strMsg, vbOKOnly, strTit)
                 textCP18.SetFocus
                 GoTo EXITSUB
              End If
           End If
        End If
        ' 申請人
'edit by nickc 2007/01/15
        If m_TM01 = "CFC" Then
'           If IsEmptyText(textTM23) = True And IsEmptyText(textSP58) = True And IsEmptyText(textSP59) = True Then
           If IsEmptyText(textTM23) = True And IsEmptyText(textSP58) = True And IsEmptyText(textSP59) = True And IsEmptyText(textTM80) = True And IsEmptyText(textTM81) = True Then
              strTit = "檢核資料"
              strMsg = "申請人不可全為空白"
              nResponse = MsgBox(strMsg, vbOKOnly, strTit)
              textTM23.SetFocus
              GoTo EXITSUB
           End If
        End If
        
        'Add By Cheng 2002/07/11
        '若案件性質為"自請撤回"(306)或"自請撤銷"(307)時, 第二頁的"相關總收文號"欄不可空白
        '2009/10/14 MODIFY BY SONIA 加退費(725)
        '2012/7/2 MODIFY BY SONIA 加暫緩審理(310)
        'modify By Sindy 2010/12/27 加(延期303)
        '2013/9/25 modify by sonia 加補正(201),申請意見書(202),更正(302),催審(305)
        If Me.textCP10.Text = "306" Or Me.textCP10.Text = "307" Or Me.textCP10.Text = "725" Or Me.textCP10 = "310" Or textCP10 = "303" Or Me.textCP10.Text = "201" Or Me.textCP10.Text = "202" Or Me.textCP10 = "302" Or textCP10 = "305" Then
           '相關總收文號不可為空白
           If IsEmptyText(Me.textCP43.Text) = True Then
              strTit = "檢核資料"
              strMsg = textCP10_2 & "案件, 請輸入相關總收文號!!! 可按 案件進度 按鈕 點選 !!"
              nResponse = MsgBox(strMsg, vbOKOnly, strTit)
              Me.SSTab1.Tab = 1
              Me.textCP43.SetFocus
              GoTo EXITSUB
           End If
           'Add By Sindy 2016/5/10 不可以輸C類來函 : "自請撤回"(306)或"自請撤銷"(307)
           If Me.textCP10.Text = "306" Or Me.textCP10.Text = "307" Then
              If Left(Trim(Me.textCP43.Text), 1) = "C" Then
                 strTit = "檢核資料"
                 strMsg = "相關總收文號不可以輸入C類來函!!!"
                 nResponse = MsgBox(strMsg, vbOKOnly, strTit)
                 Me.SSTab1.Tab = 1
                 Me.textCP43.SetFocus
                 GoTo EXITSUB
              End If
           End If
           '2016/5/10 END
        End If
        
        'Add By Cheng 2002/09/17
'        If textCP09_S = "S" And IsEmptyText(textCP09_S1) = False Then
'           strSql = "SELECT * FROM CaseProgress " & _
'                    "WHERE " & ChgCaseprogress(Me.textCP09_S.Text & Me.textCP09_S1.Text & Me.textCP09_S2.Text & Me.textCP09_S3.Text)
'           rsTmp.CursorLocation = adUseClient
'           rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
'           If rsTmp.RecordCount <= 0 Then
'              strTit = "檢核資料"
'              strMsg = "查名本所案號不存在"
'              nResponse = MsgBox(strMsg, vbOKOnly, strTit)
'              Me.SSTab1.Tab = 1
'              Me.textCP09_S.SetFocus
'              textCP09_S_GotFocus
'              If rsTmp.State <> adStateClosed Then rsTmp.Close
'              Set rsTmp = Nothing
'              GoTo EXITSUB
'           End If
'           rsTmp.Close
'        End If
'       Set rsTmp = Nothing
        
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
            '若未輸分割案母案資料
            Else
                 'edit by nickc 2006/07/20
'                If MsgBox("是否輸入分割案母案資料???", vbExclamation + vbYesNo) = vbYes Then
'                    Me.txtDivCaseNo(0).SetFocus
'                    GoTo EXITSUB
'                End If
                If m_CP31 = "Y" Then
                    MsgBox "分割母案案號一定要輸！", vbExclamation
                    Me.txtDivCaseNo(0).SetFocus
                    GoTo EXITSUB
                End If
            End If
        End If
        'End
        
         'Add By Sindy 2013/12/31 724徵求同意書
         If textCP10 = "724" Then
            If IsEmptyText(textCP50) = True And IsEmptyText(textCP51) = True And IsEmptyText(textCP52) = True Then
               MsgBox "請於第三頁頁籤輸入徵求同意書對象(中/英/日)！", vbExclamation
               SSTab1.Tab = 2
               textCP50.SetFocus
               GoTo EXITSUB
            End If
         End If
         '2013/12/31 END
         
         'Modify By Sindy 2017/3/28 S案在櫃台收文時會控管「類別」欄必須輸入，
         '但有些案件無法指定類別, 故請取消控管, 並在收文及分案時改以提醒方式
         If textTM09.Visible = True And textTM09.Enabled = True And textCP10 = "001" Then
             If Trim(textTM09) = "" Then
               If MsgBox("查名，是否有商品類別要輸入？", vbExclamation + vbYesNo + vbDefaultButton1, "注意！") = vbYes Then
                  SSTab1.Tab = 1
                  textTM09.SetFocus
                  GoTo EXITSUB
               End If
             End If
         End If
         '2017/3/28 END
         
    '若執行轉本所案號功能
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
   'Add By Cheng 2002/09/09
'   'Add By Cheng 2002/08/23
'   If Me.textTM01.Text <> "" And Me.textTM02.Text <> "" Then
'      MsgBox "不可同時做" & """轉本所案號""" & "及" & """修改其他欄位""" & "，否則資料可能有不一致的情形產生!!!", vbExclamation + vbOKOnly
'   End If
End Sub

Private Sub textTM05_1_GotFocus()
    TextInverse Me.textTM05_1
End Sub

Private Sub textTM05_1_Validate(Cancel As Boolean)
   Cancel = False
   If CheckLengthIsOK(textTM05_1, textTM05_1.MaxLength, False) = False Then
      Cancel = True
      MsgBox "案件名稱內容太長", vbOKOnly, "檢核資料"
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
      If CheckLengthIsOK(textTM05, 140, False) = False Then
         Cancel = True
         MsgBox "案件中文名稱內容太長", vbOKOnly, "檢核資料"
         textTM05_GotFocus
      End If
   Else '服務業務
      If CheckLengthIsOK(textTM05, 140, False) = False Then
         Cancel = True
         MsgBox "案件中文名稱內容太長", vbOKOnly, "檢核資料"
         textTM05_GotFocus
      End If
   End If
   '92.10.31 END
   'edit by nickc 2007/06/06 切換輸入法改用API
   'If Cancel = False Then: textTM05.IMEMode = 2
   If Cancel = False Then CloseIme
End Sub

' 案件英文名稱
Private Sub textTM06_Validate(Cancel As Boolean)
   Cancel = False
   If CheckLengthIsOK(textTM06, textTM06.MaxLength, False) = False Then
      Cancel = True
      MsgBox "案件英文名稱內容太長", vbOKOnly, "檢核資料"
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
      If CheckLengthIsOK(textTM07, 40, False) = False Then
         Cancel = True
         MsgBox "案件日文名稱內容太長", vbOKOnly, "檢核資料"
         textTM07_GotFocus
      End If
   Else '服務業務
      If CheckLengthIsOK(textTM07, 60, False) = False Then
         Cancel = True
         MsgBox "案件日文名稱內容太長", vbOKOnly, "檢核資料"
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
      If textTM08 = "7" And m_TM01 = "FCT" Then
         textTM09 = "證"
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
         If textTM08 <> "7" Then 'Add By Sindy 2015/6/30 +if
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
         GoTo EXITSUB
      End If
      '91.11.10 add by sonia
      If m_TM01 = "FCT" And textTM10 <> 台灣國家代號 Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "申請國家代號不正確"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textTM10_GotFocus
         GoTo EXITSUB
      End If
      '91.11.10 END
      
      If m_TM10 <> textTM10 Then SetLOSagree  'Added by Lydia 2020/05/20 法律所案源收文
   End If
   SetFrame4 'Added by Morgan 2022/12/23
EXITSUB:
End Sub

Private Sub textTM121_GotFocus()
   CloseIme
   TextInverse textTM121
End Sub

Private Sub textTM121_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = UpperCase(KeyAscii)
   'Modified by Morgan 2014/6/4 +可輸D
   If KeyAscii <> 8 And KeyAscii <> Asc("Y") And KeyAscii <> Asc("D") Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub textTM127_GotFocus()
   TextInverse textTM127
   CloseIme
End Sub

'Add By Sindy 2013/12/16
Private Sub textTM130_GotFocus()
   TextInverse textTM130
   CloseIme
End Sub

'Modified by Lydia 2021/08/30 改成Form2.0 ; Integer=> MSForms.ReturnInteger
Private Sub textTM130_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 8 And KeyAscii <> Asc("J") Then
      KeyAscii = 0
      Beep
   End If
End Sub

'Add By Sindy 2015/7/14
Private Sub textTM131_GotFocus()
   InverseTextBox textTM131
   '切換輸入法改用API
   OpenIme
End Sub
' 定稿商標名稱
Private Sub textTM131_Validate(Cancel As Boolean)
Dim strTit As String
Dim strMsg As String
Dim nResponse
   
   Cancel = False
   If CheckLengthIsOK(textTM131, textTM131.MaxLength, False) = False Then
      Cancel = True
      MsgBox "定稿商標名稱內容太長", vbOKOnly, "檢核資料"
      textTM131_GotFocus
   End If
   '切換輸入法改用API
   If Cancel = False Then CloseIme
End Sub
'2015/7/14 END

Private Sub textTM136_GotFocus()
   TextInverse textTM136
End Sub

Private Sub textTM136_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 8 And KeyAscii <> Asc("1") And KeyAscii <> Asc("2") Then
      KeyAscii = 0
   End If
End Sub

'Modified by Lydia 2021/08/30 改成Form2.0 ; Integer=> MSForms.ReturnInteger
Private Sub textTM23_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = UpperCase(KeyAscii)
End Sub

'Add By Sindy 2014/2/14
Private Sub textTM23_LostFocus()
   textTM23_Validate False
End Sub

' 申請人
Private Sub textTM23_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Dim strTemp As String
   'Add By Sindy 2014/2/14
   Dim strCU58 As String, strCU59 As String, strCU60 As String, strCU61 As String, strCU62 As String
   Dim strCU63 As String, strCU146 As String, strCU147 As String, strCU151 As String, strCU149 As String
   '2014/2/14 END
   Dim strCU126 As String 'Add By Sindy 2022/3/16
   
   Cancel = False
   textTM23_2 = Empty
   textTM23_3 = Empty
   'Add By Sindy 2014/2/14
   Frame2.Enabled = False
   If textTM23.Tag <> textTM23.Text Then
      textCU58 = ""
      textCU59 = ""
      textCU60 = ""
      textCU61 = ""
      textCU62 = ""
      textCU63 = ""
      textCU146 = ""
      'Modified by Lydia 2021/08/30 改成Form2.0 ; Label30(16)=> lblCU147、Label30(15)=> lblCU151
      textCU147 = "": lblCU147 = ""
      textCU151 = "": lblCU151 = ""
      'end 2021/08/30
      textCU149 = "": textCU149.Tag = textCU149 'Add By Sindy 2020/3/4
      txtCU(126) = "": txtCU(126).Tag = txtCU(126) 'Add By Sindy 2022/3/16
      textCU58.Tag = textCU58
      textCU59.Tag = textCU59
      textCU60.Tag = textCU60
      textCU61.Tag = textCU61
      textCU62.Tag = textCU62
      textCU63.Tag = textCU63
      textCU146.Tag = textCU146
      textCU147.Tag = textCU147
      textCU151.Tag = textCU151
   End If
   '2014/2/14 END
   If IsEmptyText(textTM23) = False Then
      Me.textTM23.Text = ChangeCustomerL(Me.textTM23.Text)
      'Modify By Sindy 2014/2/14
      'textTM23_2 = GetCustomerName(textTM23, 0)
      Frame2.Enabled = True
      textTM23_2 = GetCustomerName(textTM23, 0, strCU58, strCU59, strCU60, strCU61, strCU62, strCU63, strCU146, strCU147, strCU151, strCU149, strCU126)
      If textTM23.Tag <> textTM23.Text Then
         textCU58 = strCU58
         textCU59 = strCU59
         textCU60 = strCU60
         textCU61 = strCU61
         textCU62 = strCU62
         textCU63 = strCU63
         textCU146 = strCU146
         textCU147 = strCU147
         textCU147_Validate False
         textCU151 = strCU151
         textCU151_Validate False
         textCU149 = strCU149: textCU149.Tag = textCU149 'Add By Sindy 2020/3/4
         txtCU(126) = strCU126: txtCU(126).Tag = txtCU(126) 'Add By Sindy 2022/3/16
         textCU58.Tag = textCU58
         textCU59.Tag = textCU59
         textCU60.Tag = textCU60
         textCU61.Tag = textCU61
         textCU62.Tag = textCU62
         textCU63.Tag = textCU63
         textCU146.Tag = textCU146
         textCU147.Tag = textCU147
         textCU151.Tag = textCU151
      End If
      '2014/2/14 END
      
      If IsEmptyText(textTM23_2) = True Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "申請人代碼<" & textTM23 & ">不正確"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textTM23_GotFocus
      Else
        'Add By Cheng 2002/08/22
        'Mark by Lydia 2024/06/13
        'If Me.textTM23.Text <> m_strCust1 Then
        If ChangeCustomerL(Me.textTM23.Text) <> m_TM23 Then
           If Not PUB_EditCustOk(m_CP09, m_TM01, m_TM02, m_TM03, m_TM04) Then Cancel = True
        End If
            '在OnUpdateTrademark下執行
'         If Cancel = False Then
'            '910701 Sieg 601
'            If m_CP60 <> "" And InStr(ChangeCustomerL(m_TM23), ChangeCustomerL(textTM23)) = 0 Then
'               strExc(1) = m_TM01
'               strExc(2) = m_TM02
'               strExc(3) = m_TM03
'               strExc(4) = m_TM04
'               strExc(5) = m_CP60
'               strExc(6) = textTM23
'               strExc(7) = textTM23_2
'               '911118 nick 新增申請人
'               strExc(8) = m_TM23
'               If Not objLawDll.UpdAcc0k0(strExc()) Then
'                  textTM23_2 = ""
'                  textTM24 = ""
'                  textTM25 = ""
'                  textTM26 = ""
'                  Cancel = True
'                  Exit Sub
'               End If
'            End If
            strTemp = GetCustomerNation(textTM23)
            If IsEmptyText(strTemp) = False Then
               textTM23_3 = GetNationName(strTemp, 0)
            End If
            ' 91.01.22 modify by louis (更新申請人地址)
            '2005/11/18 MODIFY BY SONIA 修改申請人時才更新地址
            'UpdateCustomerAddress
            If InStr(ChangeCustomerL(m_TM23), ChangeCustomerL(textTM23)) = 0 Then
               UpdateCustomerAddress
            End If
            '2005/11/18 END
'        End If
      End If
   End If
   textTM23.Tag = textTM23 'add By Sindy 2014/2/14
   If Cancel = True Then textTM23_GotFocus
   
EXITSUB:
End Sub

' 申請人
Private Sub textSP58_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   Cancel = False
   textSP58_2 = Empty
   If IsEmptyText(textSP58) = False Then
        Me.textSP58.Text = ChangeCustomerL(Me.textSP58.Text)
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
      'Mark by Lydia 2024/06/13
      'If Me.textSP58.Text <> m_strCust2 Then
      If ChangeCustomerL(Me.textSP58.Text) <> m_TM78 Then
         If Not PUB_EditCustOk(m_CP09, m_TM01, m_TM02, m_TM03, m_TM04) Then Cancel = True
      End If
      'add by nickc 2007/01/15
      If InStr(ChangeCustomerL(m_TM78), ChangeCustomerL(textSP58)) = 0 Then
         UpdateCustomerAddress2
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
        Me.textSP59.Text = ChangeCustomerL(Me.textSP59.Text)
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
      'Mark by Lydia 2024/06/13
      'If Me.textSP59.Text <> m_strCust3 Then
      If ChangeCustomerL(Me.textSP59.Text) <> m_TM79 Then
         If Not PUB_EditCustOk(m_CP09, m_TM01, m_TM02, m_TM03, m_TM04) Then Cancel = True
      End If
      'add by nickc 2007/01/15
      If InStr(ChangeCustomerL(m_TM79), ChangeCustomerL(textSP59)) = 0 Then
        UpdateCustomerAddress3
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
      'edit by nickc 2007/10/17 修正
      'If CheckLengthIsOK(textTM24, 70) = False Then
      If CheckLengthIsOK(textTM24, textTM24.MaxLength, False) = False Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "申請地址1(中)內容太長"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textTM24_GotFocus
      End If
   End If
End Sub

' 申請地址(英)
Private Sub textTM25_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textTM25) = False Then
      If CheckLengthIsOK(textTM25, textTM25.MaxLength, False) = False Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "申請地址1(英)內容太長"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
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
      If CheckLengthIsOK(textTM26, textTM26.MaxLength, False) = False Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "申請地址1(日)內容太長"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
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
            '2012/12/19 ADD BY SONIA FCT-032571
            'modify by sonia 2019/10/4 +202申請意見書
            Case "501", "202"
               If textTM28 <> "1" Then
                  Cancel = True
                  strTit = "檢核資料"
                  strMsg = "卷宗性質不正確"
                  nResponse = MsgBox(strMsg, vbOKOnly, strTit)
                  textTM28_GotFocus
               End If
            '2012/12/19 END
            'add by sonia 2019/10/4
            Case "210"
               If textTM28 = "1" Then
                  Cancel = True
                  strTit = "檢核資料"
                  strMsg = "卷宗性質不正確"
                  nResponse = MsgBox(strMsg, vbOKOnly, strTit)
                  textTM28_GotFocus
               End If
            'end 2019/10/4
            Case Else:
               '91.12.26 CANCEL BY SONIA
               'If textTM28 <> "1" Then
               '   Cancel = True
               '   strTit = "檢核資料"
               '   strMsg = "卷宗性質不正確"
               '   nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               '   textTM28_GotFocus
               'End If
               '91.12.26 END
         End Select
      End If
   End If
End Sub

'Modified by Lydia 2021/08/30 改成Form2.0 ; Integer=> MSForms.ReturnInteger
Private Sub textTM29_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = UpperCase(KeyAscii)
End Sub

' 是否取消閉卷
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
      If m_TM29 = "Y" Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "此案已閉卷, 應該要取消閉卷 ?"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textTM29_GotFocus
      End If
   '2006/5/12 END
   End If
End Sub

Private Sub textTM34_GotFocus()
   InverseTextBox textTM34
End Sub

'Modified by Lydia 2021/08/30 改成Form2.0 ; Integer=> MSForms.ReturnInteger
Private Sub textTM34_KeyPress(KeyAscii As MSForms.ReturnInteger)
KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textTM34_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If CheckLengthIsOK(textTM34, textTM34.MaxLength, False) = False Then
      Cancel = True
      strTit = "檢核資料"
      strMsg = "分所號內容太長"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textTM34_GotFocus
   End If
End Sub

'Add By Sindy 2014/2/13
'案件聯絡人1(中)
Private Sub textTM38_GotFocus()
   OpenIme
   InverseTextBox textTM38
End Sub

Private Sub textTM38_LostFocus()
   textTM38_Validate False
End Sub

Private Sub textTM38_Validate(Cancel As Boolean)
Dim strTit As String
Dim strMsg As String
Dim nResponse
   
   Cancel = False
    'Memo by Lydia 2017/06/14 聯絡人(中)改為30字;若為服務業務會改為60字
    If CheckLengthIsOK(textTM38, textTM38.MaxLength, False) = False Then
      Cancel = True
      MsgBox "聯絡人1(中)內容太長", vbOKOnly, "檢核資料"
      textTM38_GotFocus
   End If
   If Cancel = False Then CloseIme
End Sub

'Add By Sindy 2014/2/13
Private Sub textTM39_GotFocus()
   InverseTextBox textTM39
End Sub


'Add By Sindy 2014/2/13
'案件聯絡人1(日)
Private Sub textTM40_GotFocus()
   OpenIme
   InverseTextBox textTM40
End Sub

Private Sub textTM40_LostFocus()
   textTM40_Validate False
End Sub

Private Sub textTM40_Validate(Cancel As Boolean)
Dim strTit As String
Dim strMsg As String
Dim nResponse
   
   Cancel = False
   'Modified by Lydia 2017/06/14
   'If CheckLengthIsOK(textTM40, textTM40.MaxLength) = False Then
   If CheckLengthIsOK(textTM40, 60, False) = False Then
      Cancel = True
      MsgBox "聯絡人1(日)內容太長", vbOKOnly, "檢核資料"
      textTM40_GotFocus
   End If
   If Cancel = False Then CloseIme
End Sub

'Add By Sindy 2014/2/13
'案件聯絡人2(中)
Private Sub textTM41_GotFocus()
   OpenIme
   InverseTextBox textTM41
End Sub

Private Sub textTM41_LostFocus()
   textTM41_Validate False
End Sub

Private Sub textTM41_Validate(Cancel As Boolean)
Dim strTit As String
Dim strMsg As String
Dim nResponse
   
   Cancel = False
   'Modified by Lydia 2017/06/14 聯絡人(中)改為30字
   'If CheckLengthIsOK(textTM41, textTM41.MaxLength) = False Then
   If CheckLengthIsOK(textTM41, 60, False) = False Then
      Cancel = True
      MsgBox "聯絡人2(中)內容太長", vbOKOnly, "檢核資料"
      textTM41_GotFocus
   End If
   If Cancel = False Then CloseIme
End Sub

'Add By Sindy 2014/2/13
Private Sub textTM42_GotFocus()
   InverseTextBox textTM42
End Sub

'Add By Sindy 2014/2/13
'案件聯絡人2(日)
Private Sub textTM43_GotFocus()
   OpenIme
   InverseTextBox textTM43
End Sub

Private Sub textTM43_LostFocus()
   textTM43_Validate False
End Sub

Private Sub textTM43_Validate(Cancel As Boolean)
Dim strTit As String
Dim strMsg As String
Dim nResponse
   
   Cancel = False
   'Modified by Lydia 2017/06/14
   'If CheckLengthIsOK(textTM43, textTM43.MaxLength) = False Then
   If CheckLengthIsOK(textTM43, 60, False) = False Then
      Cancel = True
      MsgBox "聯絡人2(日)內容太長", vbOKOnly, "檢核資料"
      textTM43_GotFocus
   End If
   If Cancel = False Then CloseIme
End Sub

'Modified by Lydia 2021/08/30 改成Form2.0 ; Integer=> MSForms.ReturnInteger
Private Sub textTM44_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = UpperCase(KeyAscii)
End Sub

'Add By Sindy 2014/2/14
Private Sub textTM44_LostFocus()
   textTM44_Validate False
End Sub

' FC代理人
Private Sub textTM44_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   'Add By Sindy 2014/2/14
   Dim strFA07 As String, strFA08 As String, strFA09 As String, strFA52 As String, strFA53 As String
   Dim strFA54 As String, strFA106 As String, strFA107 As String, strFA111 As String, strFA109 As String
   '2014/2/14 END
   Dim strFA91 As String 'Add By Sindy 2022/3/16
   
   Cancel = False
   textTM44_2 = Empty
   'Add By Sindy 2014/2/14
   m_textTM44_FA03 = ""
   Frame3.Enabled = False
   If textTM44.Tag <> textTM44.Text Then
      textFA07 = ""
      textFA08 = ""
      textFA09 = ""
      textFA52 = ""
      textFA53 = ""
      textFA54 = ""
      textFA106 = ""
      textFA107 = "": textFA107_2 = ""
      textFA111 = "": textFA111_2 = ""
      textFA109 = "": textFA109.Tag = textFA109 'Add By Sindy 2020/3/4
      txtFA(91) = "": txtFA(91).Tag = txtFA(91) 'Add By Sindy 2022/3/16
      textFA07.Tag = textFA07
      textFA08.Tag = textFA08
      textFA09.Tag = textFA09
      textFA52.Tag = textFA52
      textFA53.Tag = textFA53
      textFA54.Tag = textFA54
      textFA106.Tag = textFA106
      textFA107.Tag = textFA107
      textFA111.Tag = textFA111
   End If
   '2014/2/14 END
   If IsEmptyText(textTM44) = False Then
      Me.textTM44.Text = ChangeCustomerL(Me.textTM44.Text)
      'Modify By Sindy 2014/2/14
      'textTM44_2 = GetFAgentName(textTM44)
      Frame3.Enabled = True
      textTM44_2 = GetFAgentName(textTM44, m_textTM44_FA03, strFA07, strFA08, strFA09, strFA52, strFA53, strFA54, strFA106, strFA107, strFA111, strFA109, strFA91)
      If textTM44.Tag <> textTM44.Text Then
         textFA07 = strFA07
         textFA08 = strFA08
         textFA09 = strFA09
         textFA52 = strFA52
         textFA53 = strFA53
         textFA54 = strFA54
         textFA106 = strFA106
         textFA107 = strFA107
         textFA107_Validate False
         textFA111 = strFA111
         textFA111_Validate False
         textFA109 = strFA109: textFA109.Tag = textFA109 'Add By Sindy 2020/3/4
         txtFA(91) = strFA91: txtFA(91).Tag = txtFA(91) 'Add By Sindy 2022/3/16
         textFA07.Tag = textFA07
         textFA08.Tag = textFA08
         textFA09.Tag = textFA09
         textFA52.Tag = textFA52
         textFA53.Tag = textFA53
         textFA54.Tag = textFA54
         textFA106.Tag = textFA106
         textFA107.Tag = textFA107
         textFA111.Tag = textFA111
      End If
      '2014/2/14 END
      If IsEmptyText(textTM44_2) = True Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "FC代理人<" & textTM44 & ">不正確"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textTM44_GotFocus
      End If
   'add by sonia 2020/11/11
   Else
      If textTM10 = "000" And textTM44 = "" Then
         If MsgBox("台灣案但沒有FC代理人，是否要補輸入？", vbExclamation + vbYesNo + vbDefaultButton1, "重要訊息！") = vbYes Then
            Cancel = True
            SSTab1.Tab = 1
            textTM44.SetFocus
            Exit Sub
         End If
      End If
   'end 2020/11/11
   End If
   textTM44.Tag = textTM44 'add By Sindy 2014/2/14
End Sub

Private Sub textTM46_GotFocus()
   InverseTextBox textTM46
End Sub
'Modified by Lydia 2021/08/30 改成Form2.0 ; Integer=> MSForms.ReturnInteger
Private Sub textTM46_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = UpperCase(KeyAscii)
End Sub
' D/N是否列印申請人
Private Sub textTM46_Validate(Cancel As Boolean)
Dim strTit As String
Dim strMsg As String
Dim nResponse
   
   Cancel = False
   If IsYesOrSpace(textTM46) = False Then
      Cancel = True
      strTit = "檢核資料"
      strMsg = "請輸入Y或空白"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textTM46_GotFocus
   End If
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
   End Select
End Function

'Add By Sindy 2014/2/13
'案件固定請款對象
Private Sub textTM56_1_GotFocus()
   InverseTextBox textTM56_1
End Sub

'Modified by Lydia 2021/08/30 改成Form2.0 ; Integer=> MSForms.ReturnInteger
Private Sub textTM56_1_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textTM56_1_LostFocus()
   textTM56_1_Validate False
End Sub

Private Sub textTM56_1_Validate(Cancel As Boolean)
Dim strTemp As String
Dim strTit As String
Dim strMsg As String
Dim nResponse
Dim strTempName As String
   
   Cancel = False
   textTM56_2 = ""
   If IsEmptyText(textTM56_1) = False Then
      strTemp = textTM56_1
      If Len(strTemp) < 8 Then: strTemp = strTemp & String(8 - Len(strTemp), "0")
      If Left(Me.textTM56_1.Text, 1) = "X" Then
         textTM56_2 = GetAgentOrCustName(strTemp)
      Else
         If PUB_GetAgentName(Me.textTM01.Text, strTemp, strTempName) Then
            textTM56_2 = strTempName
         End If
      End If
      
      If IsEmptyText(textTM56_2) = True Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "案件固定請款對象代號不存在"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textTM56_1_GotFocus
      End If
   End If
End Sub

' 案件備註
Private Sub textTM58_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If CheckLengthIsOK(textTM58, textTM58.MaxLength, False) = False Then
      Cancel = True
      strTit = "檢核資料"
      strMsg = "案件備註內容太長"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textTM58_GotFocus
   End If
End Sub

Private Sub ReCaculateCP48()
   Dim strDay As String
   Dim strDate As String
   Dim strTemp As String
   
   ' 檢查案件性質
   If IsEmptyText(textCP10) = True Then
      GoTo EXITSUB
   End If
   ' 檢查收文日
   If IsEmptyText(textCP05) = True Then
      GoTo EXITSUB
   End If
   ' 檢查申請國家
   If IsEmptyText(textTM10) = True Then
      GoTo EXITSUB
   End If
   
   ' 承辦期限的日期應為來函收文日加上工作天數
   ' 工作天數由系統別+國家代碼+案件性質搜尋案件收費表的工作天數
''''edit by nickc 2007/10/12 改抓有時效性的
''''   strDay = GetWorkDays(m_TM01, textTM10, textCP10)
''''   If IsEmptyText(strDay) = False Then
''''      strDate = DBDATE(textCP05)
''''      ' 90.07.03 modify by louis (承辦期限以實際工作天數計算)
''''      'strTemp = DBDATE(Format(DateSerial(Val(DBYEAR(strDate)), Val(DBMONTH(strDate)), Val(DBDAY(strDate)) + Val(strDay))))
''''      strTemp = DBDATE(CompWorkDay(Val(strDay), DBDATE(strDate), 0))
''''      textCP48 = TAIWANDATE(strTemp)
''''   End If
   'modify by sonia 2023/2/24 CFT,CFC,S非台灣案一律改為收文日+3工作天
   'textCP48 = TAIWANDATE(Pub_GetHandleDay(m_TM01, m_TM10, textCP10, DBDATE(textCP05), DBDATE(textCP06), textCP09))
   If textTM10 <> "000" Then
      'add by sonia 2023/3/10 剔除部分案件性質
      If textCP10 = "107" Or textCP10 = "705" Or textCP10 = "706" Or textCP10 = "711" Or textCP10 = "714" Or textCP10 = "721" Then
         GoTo EXITSUB
      End If
         textCP48 = CompWorkDay(3, DBDATE(textCP05), 0)
         If textCP06 <> "" Then
             If textCP48 > Val(DBDATE(textCP06)) Then
                 textCP48 = textCP06
             End If
         End If
         '本所期限小於系統日時，承辦期限設定為系統日
         If textCP48 <> "" And textCP06 <> "" And Val(DBDATE(textCP06)) < Val(strSrvDate(1)) Then
            textCP48 = strSrvDate(2)
         End If
         '承辦期限小於系統日時，承辦期限設定為系統日
         If textCP48 <> "" And Val(DBDATE(textCP48)) < Val(strSrvDate(1)) Then
            textCP48 = strSrvDate(2)
         End If
         textCP48 = TAIWANDATE(textCP48)
   Else
      textCP48 = TAIWANDATE(Pub_GetHandleDay(m_TM01, m_TM10, textCP10, DBDATE(textCP05), DBDATE(textCP06), textCP09))
   End If
   'end 2023/2/24
EXITSUB:
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

Private Sub textCP09_S_GotFocus()
   InverseTextBox textCP09_S
End Sub

Private Sub textCP10_GotFocus()
   InverseTextBox textCP10
End Sub

Private Sub textCP13_GotFocus()
   InverseTextBox textCP13
End Sub

Private Sub textCP14_GotFocus()
   InverseTextBox textCP14
End Sub

Private Sub textCP16_GotFocus()
   InverseTextBox textCP16
End Sub

Private Sub textCP17_GotFocus()
   InverseTextBox textCP17
End Sub

Private Sub textCP18_GotFocus()
   InverseTextBox textCP18
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

Private Sub textCP48_GotFocus()
   InverseTextBox textCP48
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

Private Sub textTM35_GotFocus()
   InverseTextBox textTM35
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

'Add By Cheng 2002/05/23
Private Function TxtValidate() As Boolean
Dim objTxt As Object
Dim ii As Integer
Dim Cancel As Boolean
Dim strTit As String
Dim strMsg As String
Dim nResponse
Dim strCU01 As String, strCU02 As String 'Add By Sindy 2013/12/16
Dim bolColor As Boolean 'Added by Morgan 2023/7/27

TxtValidate = False

'Modify By Cheng 2002/11/22
'若執行轉本所案號功能
If Me.textTM01.Text <> "" And Me.textTM02.Text <> "" Then
    '910722 Sieg
    If textTM01 <> "" And textTM02 <> "" Then
       Dim strTM01 As String
       Dim strTM02 As String
       Dim strTM03 As String
       Dim strTM04 As String
       strTM01 = textTM01
       strTM02 = textTM02
       If strTM02 = "TF" Then: strTM02 = strTM02 & textTM02_2
       strTM03 = textTM03
       If IsEmptyText(strTM03) = True Then: strTM03 = "0"
       strTM04 = textTM04
       If IsEmptyText(strTM04) = True Then: strTM04 = "00"
    
       strExc(1) = strTM01
       strExc(2) = strTM02
       strExc(3) = strTM03
       strExc(4) = strTM04
       If IsDataRecordExist(strTM01, strTM02, strTM03, strTM04) Then
          
       Else
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
    
'若非執行轉本所案號功能
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
    
    If Me.textCP09_S.Enabled = True Then
       Cancel = False
       textCP09_S_Validate Cancel
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
    
    If Me.textCP16.Enabled = True Then
       Cancel = False
       textCP16_Validate Cancel
       If Cancel = True Then
          Exit Function
       End If
    End If
    
    If Me.textCP17.Enabled = True Then
       Cancel = False
       textCP17_Validate Cancel
       If Cancel = True Then
          Exit Function
       End If
    End If
    
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
    
    If Me.textCP43.Enabled = True Then
       Cancel = False
       textCP43_Validate Cancel
       If Cancel = True Then
          Exit Function
       End If
    End If
    
    If Me.textCP48.Enabled = True Then
       Cancel = False
       textCP48_Validate Cancel
       If Cancel = True Then
          Exit Function
       End If
    End If
   
   'Add By Sindy 2013/12/31
   If Me.textCP50.Enabled = True Then
      Cancel = False
      textCP50_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If Me.textCP51.Enabled = True Then
      Cancel = False
      textCP51_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If Me.textCP52.Enabled = True Then
      Cancel = False
      textCP52_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   '2013/12/31 END
   
   If Me.textCP64.Enabled = True Then
      Cancel = False
      textCP64_Validate Cancel
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
   
   If Me.textTM23.Enabled = True Then
      Cancel = False
      textTM23_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   'Add By Sindy 2015/7/14
   If Me.textTM131.Enabled = True Then
      Cancel = False
      textTM131_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   '2015/7/14 END
   
   'Add By Sindy 2020/3/4
   If Me.textTM46.Enabled = True Then
      Cancel = False
      textTM46_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If Me.textCU149.Enabled = True Then
      Cancel = False
      textCU149_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If Me.textFA109.Enabled = True Then
      Cancel = False
      textFA109_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   ' 2020/3/4 END
   
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
   
   'Add By Sindy 2019/1/4
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
   
   'Add By Sindy 2014/2/14
   If Me.textCU147.Enabled = True Then
      Cancel = False
      textCU147_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If Me.textCU151.Enabled = True Then
      Cancel = False
      textCU151_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If Me.textCU58.Enabled = True Then
      Cancel = False
      textCU58_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If Me.textCU60.Enabled = True Then
      Cancel = False
      textCU60_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If Me.textCU61.Enabled = True Then
      Cancel = False
      textCU61_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If Me.textCU63.Enabled = True Then
      Cancel = False
      textCU63_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If Me.textFA07.Enabled = True Then
      Cancel = False
      textFA07_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If Me.textFA09.Enabled = True Then
      Cancel = False
      textFA09_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If Me.textFA107.Enabled = True Then
      Cancel = False
      textFA107_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If Me.textFA111.Enabled = True Then
      Cancel = False
      textFA111_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If Me.textFA52.Enabled = True Then
      Cancel = False
      textFA52_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If Me.textFA54.Enabled = True Then
      Cancel = False
      textFA54_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If Me.textTM38.Enabled = True Then
      Cancel = False
      textTM38_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If Me.textTM40.Enabled = True Then
      Cancel = False
      textTM40_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If Me.textTM41.Enabled = True Then
      Cancel = False
      textTM41_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If Me.textTM43.Enabled = True Then
      Cancel = False
      textTM43_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If Me.textTM56_1.Enabled = True Then
      Cancel = False
      textTM56_1_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If Me.textTM69_1.Enabled = True Then
      Cancel = False
      textTM69_1_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   '2014/2/14 END
   
   If Me.textTM58.Enabled = True Then
      Cancel = False
      textTM58_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   'Add By Sindy 2015/6/22
   If Me.textCP09_S3.Enabled = True Then
      If ChkSPDataErr = True Then
         'textCP09_S1.SetFocus
         'textCP09_S1_GotFocus
         Exit Function
      End If
   End If
   '2015/6/22 END

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
               Me.SSTab1.Tab = 1
               textTM130.SetFocus
               Exit Function
            End If
         End If
      Next ii
   End If
   '2013/12/16 END
   
   '92.3.13 ADD BY SONIA
   If textCP10 = "108" And m_Priority(3) = "" Then
      strTit = "檢核資料"
      strMsg = "案件性質為 主張優先權 時, 請輸入優先權資料!!"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      Cancel = True
      Exit Function
   End If
   '92.3.13 END
   
   'Added by Morgan 2022/12/15
   'Modified by Morgan 2023/1/13 排除申請
   If textTM136.Visible And textCP10 <> "101" Then
      If strSrvDate(1) > "20230000" Then
         If textTM136 = "" Then
            MsgBox "請輸入註冊證形式！", vbExclamation
            textTM136.SetFocus
            Exit Function
         End If
      End If
   End If
   'end 2022/12/15
   
   'Added by Morgan 2023/7/26
   '若接洽單有商標圖而基本檔無代表圖時詢問是否存入
   If m_TM01 = "CFT" And textCP10 = "101" And txtF0301 <> "" Then
      strExc(0) = "select crif02,crif05 from consultrecimagef where crif01='" & txtF0301 & "'" & _
         " and not exists(select * from imgbytefile where ibf01='" & m_TM01 & "' and ibf02='" & m_TM02 & "' and ibf03='" & m_TM03 & "' and ibf04='" & m_TM04 & "')"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         If MsgBox("是否將接洽單之附圖放入案件代表圖？", vbYesNo + vbDefaultButton2 + vbQuestion) = vbYes Then
            strTit = App.path & "\$$NowPic.jpg"
            Cancel = False
            bolColor = IIf(RsTemp.Fields("crif02") = "2", True, False)
            If PUB_GetFtpFile(RsTemp.Fields("crif05"), strTit, UCase("consultrecimagef")) = True Then
               frmPic001.oCP01 = m_TM01
               frmPic001.oCP02 = m_TM02
               frmPic001.oCP03 = m_TM03
               frmPic001.oCP04 = m_TM04
               'frmPic001.StrMenu
               Cancel = Not frmPic001.UploadFile(strTit, bolColor)
               Unload frmPic001
            End If
            If Dir(strTit) <> "" Then Kill strTit
            If Cancel Then Exit Function
         End If
      End If
   End If
   'end 2023/7/26
   
   'Add by Sindy 2023/12/11
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
         '檢查指定送件日相關欄位
         If Frame6.Visible = True Then
            If Option1(0).Value = False And Option1(1).Value = False And Option1(2).Value = False Then
               MsgBox "有輸入指定送件日，當天或之前或之後請擇一。", vbExclamation
               Exit Function
            End If
         End If
      End If
   Else
      Option1(0).Value = False
      Option1(1).Value = False
      Option1(2).Value = False
   End If
   '2023/12/11 END
End If

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
         If GetAgentAndState(strExc(1), strExc(3), , , , textTM01, strExc(2), False) = False Then
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
'取得案件收費表的下次期限
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

'Add By Sindy 2014/2/13
Private Sub textTM69_1_GotFocus()
   TextInverse Me.textTM69_1
End Sub

'Modified by Lydia 2021/08/30 改成Form2.0 ; Integer=> MSForms.ReturnInteger
Private Sub textTM69_1_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textTM69_1_LostFocus()
   textTM69_1_Validate False
End Sub

Private Sub textTM69_1_Validate(Cancel As Boolean)
Dim strTemp As String
Dim strTit As String
Dim strMsg As String
Dim nResponse
Dim strTempName As String
   
   Cancel = False
   textTM69_2 = ""
   If IsEmptyText(textTM69_1) = False Then
      strTemp = textTM69_1
      ' 不滿八碼補0
      If Len(strTemp) < 8 Then: strTemp = strTemp & String(8 - Len(strTemp), "0")
      If Left(Me.textTM69_1.Text, 1) = "X" Then
         textTM69_2 = GetAgentOrCustName(strTemp)
      Else
         If PUB_GetAgentName(Me.textTM01.Text, strTemp, strTempName) Then
            textTM69_2 = strTempName
         End If
      End If
      If IsEmptyText(textTM69_2) = True Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "案件D/N固定列印對象代號不存在"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textTM69_1_GotFocus
      End If
   End If
End Sub

Private Sub textTM72_GotFocus()
    TextInverse Me.textTM72
End Sub

'Added by Lydia 2023/11/14
Private Sub textTM72_KeyPress(KeyAscii As MSForms.ReturnInteger)
    KeyAscii = UpperCase(KeyAscii)
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

'add by nickc 2007/01/15
Private Sub textTM80_GotFocus()
InverseTextBox textTM80
End Sub

'Modified by Lydia 2021/08/30 改成Form2.0 ; Integer=> MSForms.ReturnInteger
Private Sub textTM80_KeyPress(KeyAscii As MSForms.ReturnInteger)
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
         textTM80_GotFocus
      End If
   End If
   If Cancel = False Then
      'Mark by Lydia 2024/06/13
      'If Me.textTM80.Text <> m_strCust4 Then
      If ChangeCustomerL(Me.textTM80.Text) <> m_TM80 Then
         If Not PUB_EditCustOk(m_CP09, m_TM01, m_TM02, m_TM03, m_TM04) Then Cancel = True
      End If
      If InStr(ChangeCustomerL(m_TM80), ChangeCustomerL(textTM80)) = 0 Then
        UpdateCustomerAddress4
      End If
   End If
   If Cancel = True Then textTM80_GotFocus
End Sub

Private Sub textTM81_GotFocus()
InverseTextBox textTM81
End Sub

'Modified by Lydia 2021/08/30 改成Form2.0 ; Integer=> MSForms.ReturnInteger
Private Sub textTM81_KeyPress(KeyAscii As MSForms.ReturnInteger)
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
      If textTM80_2 = Empty Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "申請人代碼<" & textTM81 & ">不正確"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textTM81_GotFocus
      End If
   End If
   If Cancel = False Then
      'Mark by Lydia 2024/06/13
      'If Me.textTM81.Text <> m_strCust5 Then
      If ChangeCustomerL(Me.textTM81.Text) <> m_TM81 Then
         If Not PUB_EditCustOk(m_CP09, m_TM01, m_TM02, m_TM03, m_TM04) Then Cancel = True
      End If
      If InStr(ChangeCustomerL(m_TM81), ChangeCustomerL(textTM81)) = 0 Then
        UpdateCustomerAddress5
      End If
   End If
   If Cancel = True Then textTM81_GotFocus
End Sub

'add by nickc 2007/01/15
Private Sub textTM82_GotFocus()
InverseTextBox textTM82
End Sub
Private Sub textTM82_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textTM82) = False Then
      If CheckLengthIsOK(textTM82, textTM82.MaxLength, False) = False Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "申請地址2(中)內容太長"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textTM82_GotFocus
      End If
   End If
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
      If CheckLengthIsOK(textTM83, textTM83.MaxLength, False) = False Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "申請地址3(中)內容太長"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textTM83_GotFocus
      End If
   End If
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
      If CheckLengthIsOK(textTM84, textTM84.MaxLength, False) = False Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "申請地址4(中)內容太長"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textTM84_GotFocus
      End If
   End If
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
      If CheckLengthIsOK(textTM85, textTM85.MaxLength, False) = False Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "申請地址5(中)內容太長"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textTM85_GotFocus
      End If
   End If
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
      If CheckLengthIsOK(textTM86, textTM86.MaxLength, False) = False Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "申請地址2(英)內容太長"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
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
      If CheckLengthIsOK(textTM87, textTM87.MaxLength, False) = False Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "申請地址3(英)內容太長"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
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
      If CheckLengthIsOK(textTM88, textTM88.MaxLength, False) = False Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "申請地址4(英)內容太長"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
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
      If CheckLengthIsOK(textTM89, textTM89.MaxLength, False) = False Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "申請地址5(英)內容太長"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
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
      If CheckLengthIsOK(textTM90, textTM90.MaxLength, False) = False Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "申請地址2(日)內容太長"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
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
      If CheckLengthIsOK(textTM91, textTM91.MaxLength, False) = False Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "申請地址3(日)內容太長"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
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
      If CheckLengthIsOK(textTM92, textTM92.MaxLength, False) = False Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "申請地址4(日)內容太長"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
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
      If CheckLengthIsOK(textTM93, textTM93.MaxLength, False) = False Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "申請地址5(日)內容太長"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textTM93_GotFocus
      End If
   End If
End Sub

Private Sub txtCU_GotFocus(Index As Integer)
   If txtCU(Index).Enabled = True And txtCU(Index).Locked = False Then
      InverseTextBox txtCU(Index)
      CloseIme
   End If
End Sub

Private Sub txtCU_KeyPress(Index As Integer, KeyAscii As Integer)
   If txtCU(Index).Enabled = True And txtCU(Index).Locked = False Then
      Select Case Index
         Case 126
            KeyAscii = UpperCase(KeyAscii)
            If KeyAscii <> 89 And KeyAscii <> 68 And KeyAscii <> 8 Then
               KeyAscii = 0
               Beep
            End If
      End Select
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
    Me.txtDivCaseNo(0).Text = ""
    Me.txtDivCaseNo(1).Text = ""
    Me.txtDivCaseNo(2).Text = ""
    Me.txtDivCaseNo(3).Text = ""
    Me.txtDivCaseNo(4).Text = ""
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
    StrSQLa = StrSQLa & " Union Select TM10, TM23, '', '', '', '', '' From Trademark Where TM01='" & strCode(0) & "' And TM02='" & strCode(1) & "' And TM03='" & strCode(2) & "' And TM04='" & strCode(3) & "' "
    StrSQLa = StrSQLa & " Union Select LC15, LC11,'', '', '', '', '' From Lawcase Where LC01='" & strCode(0) & "' And LC02='" & strCode(1) & "' And LC03='" & strCode(2) & "' And LC04='" & strCode(3) & "' "
    StrSQLa = StrSQLa & " Union Select '000', HC05, '', '', '', '', '' From Hirecase Where HC01='" & strCode(0) & "' And HC02='" & strCode(1) & "' And HC03='" & strCode(2) & "' And HC04='" & strCode(3) & "' "
    StrSQLa = StrSQLa & " Union Select SP09, SP08, '', '', '', '', '' From Servicepractice Where SP01='" & strCode(0) & "' And SP02='" & strCode(1) & "' And SP03='" & strCode(2) & "' And SP04='" & strCode(3) & "' "
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
    StrSQLa = StrSQLa & " Union Select TM10, TM23, '', '', '', '', '' From Trademark Where TM01='" & strCode(4) & "' And TM02='" & strCode(5) & "' And TM03='" & strCode(6) & "' And TM04='" & strCode(7) & "' "
    StrSQLa = StrSQLa & " Union Select LC15, LC11, '', '', '', '', '' From Lawcase Where LC01='" & strCode(4) & "' And LC02='" & strCode(5) & "' And LC03='" & strCode(6) & "' And LC04='" & strCode(7) & "' "
    StrSQLa = StrSQLa & " Union Select '000', HC05, '', '', '', '', '' From Hirecase Where HC01='" & strCode(4) & "' And HC02='" & strCode(5) & "' And HC03='" & strCode(6) & "' And HC04='" & strCode(7) & "' "
    StrSQLa = StrSQLa & " Union Select SP09, SP08, '', '', '' ,'', '' From Servicepractice Where SP01='" & strCode(4) & "' And SP02='" & strCode(5) & "' And SP03='" & strCode(6) & "' And SP04='" & strCode(7) & "' "
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

'add by nickc 2007/01/15
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
            strTM90 = rsTmp.Fields("CU29")
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
            strTM91 = rsTmp.Fields("CU29")
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
            strTM92 = rsTmp.Fields("CU29")
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
            strTM93 = rsTmp.Fields("CU29")
         End If
      End If
   End If
   rsTmp.Close
   Set rsTmp = Nothing
   
   textTM85 = strTM85
   textTM89 = strTM89
   textTM93 = strTM93
End Sub

'Add By Sindy 2014/2/13
Private Function ChgType(ByVal Sty As Integer, ByVal txt As String) As String
Dim strTmp As String, strTmp1 As String
   
   Select Case Sty
      Case 0
         If ClsPDGetCaseProperty("P", txt, strTmp) = True Then
            ChgType = strTmp
         Else
            ChgType = ""
         End If
      Case 1, 5
         '只檢查智權人員代號是否存在, 不管是否仍在職
         If PUB_GetStaffNameDept(txt, strTmp, strTmp1, False) = True Then
            ChgType = strTmp & "," & strTmp1
         Else
            ChgType = ""
         End If
      Case 2
         If ClsPDGetNation(txt, strTmp) = True Then
            ChgType = strTmp
         Else
            ChgType = ""
         End If
      Case 3
         If ClsPDGetCaseSource(txt, strTmp) = True Then
            ChgType = strTmp
         Else
            ChgType = ""
         End If
      Case 4
         If ClsLawLawGetName(txt, strTmp) = True Then
            ChgType = strTmp
         Else
            ChgType = ""
         End If
   End Select
End Function

'Add By Sindy 2014/2/13
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
         rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
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
         rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
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

'Added by Lydia 2017/06/14
Private Sub textTM39_Validate(Cancel As Boolean)
   Cancel = False
    If CheckLengthIsOK(textTM39, 35, False) = False Then
      Cancel = True
      MsgBox "聯絡人1(英)內容太長", vbOKOnly, "檢核資料"
      textTM39_GotFocus
   End If
End Sub
Private Sub textTM42_Validate(Cancel As Boolean)
   Cancel = False
    If CheckLengthIsOK(textTM42, 35, False) = False Then
      Cancel = True
      MsgBox "聯絡人2(英)內容太長", vbOKOnly, "檢核資料"
      textTM42_GotFocus
   End If
End Sub

'Added by Lydia 2020/05/20 法律所案源收文：案件性質=>案源案件類型
Private Sub SetLOSagree()
Dim m_LOSkind As String

    If strSrvDate(1) >= 法律所案源收文啟用日 And m_TM01 = "FCT" Then
        'Modified by Lydia 2020/06/29 直接用案源檔的類型
        'm_LOSkind = PUB_GetLOSkind(m_TM01, textCP10, textTM10)
        m_LOSkind = m_LOS02
        txtLOSagree.Text = ""
        FraLOS.Visible = False
        If textTM10 = "000" Then
            If Left(m_LOSkind, 1) = "C" And m_LOS01 = "" Then 'C類-未分案通知
                 FraLOS.Visible = True
                 txtLOSagree.Text = "Y"
            End If
        End If
    End If
    
End Sub

Private Sub txtFA_GotFocus(Index As Integer)
   If txtFA(Index).Enabled = True And txtFA(Index).Locked = False Then
      InverseTextBox txtFA(Index)
      CloseIme
   End If
End Sub
Private Sub txtFA_KeyPress(Index As Integer, KeyAscii As Integer)
   If txtFA(Index).Enabled = True And txtFA(Index).Locked = False Then
      Select Case Index
         Case 91
            KeyAscii = UpperCase(KeyAscii)
            If KeyAscii <> 89 And KeyAscii <> 68 And KeyAscii <> 8 Then
               KeyAscii = 0
               Beep
            End If
      End Select
   End If
End Sub

'Modified by Lydia 2021/08/30 改成Form2.0 ; Integer=> MSForms.ReturnInteger
Private Sub txtLOSagree_KeyPress(KeyAscii As MSForms.ReturnInteger)
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
   If strSrvDate(1) >= 法律所案源收文啟用日 And m_TM01 = "FCT" Then
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
Private Sub SetFrame4()
   Frame4.Visible = False
   'Modified by Morgan 2023/2/6 不必限制未發文--陳金蓮
   'If textTM10 = "000" And Val(m_CP27) = 0 Then
   If textTM10 = "000" Then
      'Added by Morgan 2023/1/13 申請也可設定--陳金蓮
      If textCP10 = "101" Then
         Frame4.Visible = True
      'end 2023/1/13
      Else
         If PUB_TWCertPty(m_TM01, textCP10, m_TM02, m_TM03, m_TM04) = True Then
            Frame4.Visible = True
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

