VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm010012_07 
   BorderStyle     =   1  '單線固定
   Caption         =   "內部收文"
   ClientHeight    =   5916
   ClientLeft      =   1068
   ClientTop       =   2688
   ClientWidth     =   8952
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5916
   ScaleWidth      =   8952
   Begin VB.CommandButton cmdExit 
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   8100
      TabIndex        =   24
      Top             =   30
      Width           =   800
   End
   Begin VB.TextBox textLC08_2 
      BorderStyle     =   0  '沒有框線
      ForeColor       =   &H000000FF&
      Height          =   264
      Left            =   2805
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   480
      Width           =   1935
   End
   Begin VB.TextBox textLCKey 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   1020
      Locked          =   -1  'True
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   480
      Width           =   1755
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "回前畫面(&U)"
      Height          =   400
      Left            =   6840
      TabIndex        =   23
      Top             =   30
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Left            =   6015
      TabIndex        =   22
      Top             =   30
      Width           =   800
   End
   Begin VB.CommandButton cmdCaseProgress 
      Caption         =   "案件進度(&C)"
      Height          =   400
      Left            =   4785
      TabIndex        =   21
      Top             =   30
      Width           =   1200
   End
   Begin VB.CommandButton cmdRelate 
      Caption         =   "相關案號(&F)"
      Height          =   400
      Left            =   3540
      TabIndex        =   20
      Top             =   30
      Width           =   1200
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5124
      Left            =   60
      TabIndex        =   36
      Top             =   768
      Width           =   8832
      _ExtentX        =   15579
      _ExtentY        =   9038
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabHeight       =   529
      TabCaption(0)   =   "基本資料"
      TabPicture(0)   =   "frm010012_07.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label22"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label21"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label40"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label19"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label18"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Line1"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label9"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label10"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label23"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label8"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label7"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label25"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label6"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label1(1)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Label5"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Label4"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Label3"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Label20"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Label16"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Label2(4)"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Label12"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "Label11"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "Label13"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "Label37"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "Label14"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "textLC06"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "textLC07"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "textLC05"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "textCP13_2"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "textCP14_2"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "textCP64"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "textCP29_2"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "grdList"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "textLC08"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "textCP21"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "textCP54"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "textCP53"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "textCP43"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "textCP26"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).Control(39)=   "textCP06"
      Tab(0).Control(39).Enabled=   0   'False
      Tab(0).Control(40)=   "textCP07"
      Tab(0).Control(40).Enabled=   0   'False
      Tab(0).Control(41)=   "textCP49"
      Tab(0).Control(41).Enabled=   0   'False
      Tab(0).Control(42)=   "textCP13"
      Tab(0).Control(42).Enabled=   0   'False
      Tab(0).Control(43)=   "textCP10_2"
      Tab(0).Control(43).Enabled=   0   'False
      Tab(0).Control(44)=   "textCP10"
      Tab(0).Control(44).Enabled=   0   'False
      Tab(0).Control(45)=   "textCP29"
      Tab(0).Control(45).Enabled=   0   'False
      Tab(0).Control(46)=   "textCP14"
      Tab(0).Control(46).Enabled=   0   'False
      Tab(0).Control(47)=   "textLC16"
      Tab(0).Control(47).Enabled=   0   'False
      Tab(0).Control(48)=   "textLC13"
      Tab(0).Control(48).Enabled=   0   'False
      Tab(0).Control(49)=   "textCP05"
      Tab(0).Control(49).Enabled=   0   'False
      Tab(0).ControlCount=   50
      TabCaption(1)   =   "當事人／對造資料"
      TabPicture(1)   =   "frm010012_07.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label2(3)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label2(2)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label2(1)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label2(0)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Label36"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Label27"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Label26"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Label24"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "Label15"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "Label17"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "Label28"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "Label29"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "textLC11_2"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "textCP40"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "textCP41"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "textCP42"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "textLC43_2"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "textLC44_2"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).Control(18)=   "textLC45_2"
      Tab(1).Control(18).Enabled=   0   'False
      Tab(1).Control(19)=   "textLC46_2"
      Tab(1).Control(19).Enabled=   0   'False
      Tab(1).Control(20)=   "textLC11"
      Tab(1).Control(20).Enabled=   0   'False
      Tab(1).Control(21)=   "textLC43"
      Tab(1).Control(21).Enabled=   0   'False
      Tab(1).Control(22)=   "textLC44"
      Tab(1).Control(22).Enabled=   0   'False
      Tab(1).Control(23)=   "textLC45"
      Tab(1).Control(23).Enabled=   0   'False
      Tab(1).Control(24)=   "textLC46"
      Tab(1).Control(24).Enabled=   0   'False
      Tab(1).ControlCount=   25
      Begin VB.TextBox textLC46 
         Height          =   264
         Left            =   -73830
         MaxLength       =   9
         TabIndex        =   31
         Top             =   1710
         Width           =   1095
      End
      Begin VB.TextBox textLC45 
         Height          =   264
         Left            =   -73830
         MaxLength       =   9
         TabIndex        =   30
         Top             =   1410
         Width           =   1095
      End
      Begin VB.TextBox textLC44 
         Height          =   264
         Left            =   -73830
         MaxLength       =   9
         TabIndex        =   29
         Top             =   1110
         Width           =   1095
      End
      Begin VB.TextBox textLC43 
         Height          =   264
         Left            =   -73830
         MaxLength       =   9
         TabIndex        =   28
         Top             =   810
         Width           =   1095
      End
      Begin VB.TextBox textLC11 
         Height          =   264
         Left            =   -73830
         MaxLength       =   9
         TabIndex        =   27
         Top             =   510
         Width           =   1095
      End
      Begin VB.TextBox textCP05 
         Height          =   264
         Left            =   1100
         MaxLength       =   7
         TabIndex        =   0
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox textLC13 
         Height          =   264
         Left            =   1920
         MaxLength       =   1
         TabIndex        =   4
         Top             =   1500
         Width           =   372
      End
      Begin VB.TextBox textLC16 
         Height          =   264
         Left            =   5385
         MaxLength       =   50
         TabIndex        =   5
         Top             =   1500
         Width           =   3330
      End
      Begin VB.TextBox textCP14 
         Height          =   264
         Left            =   1100
         MaxLength       =   6
         TabIndex        =   6
         Top             =   1800
         Width           =   732
      End
      Begin VB.TextBox textCP29 
         Height          =   264
         Left            =   5400
         MaxLength       =   6
         TabIndex        =   7
         Top             =   1800
         Width           =   732
      End
      Begin VB.TextBox textCP10 
         Height          =   264
         Left            =   1100
         MaxLength       =   6
         TabIndex        =   8
         Top             =   2100
         Width           =   732
      End
      Begin VB.TextBox textCP10_2 
         BorderStyle     =   0  '沒有框線
         Height          =   264
         Left            =   1900
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   2100
         Width           =   2175
      End
      Begin VB.TextBox textCP13 
         Height          =   264
         Left            =   5400
         MaxLength       =   6
         TabIndex        =   9
         Top             =   2100
         Width           =   735
      End
      Begin VB.TextBox textCP49 
         Height          =   264
         Left            =   1100
         MaxLength       =   300
         TabIndex        =   10
         Top             =   2400
         Width           =   2292
      End
      Begin VB.TextBox textCP07 
         Height          =   264
         Left            =   5400
         MaxLength       =   7
         TabIndex        =   14
         Top             =   2700
         Width           =   1095
      End
      Begin VB.TextBox textCP06 
         Height          =   264
         Left            =   1100
         MaxLength       =   7
         TabIndex        =   13
         Top             =   2700
         Width           =   1215
      End
      Begin VB.TextBox textCP26 
         Height          =   264
         Left            =   1440
         MaxLength       =   20
         TabIndex        =   15
         Top             =   3000
         Width           =   372
      End
      Begin VB.TextBox textCP43 
         Height          =   264
         Left            =   5820
         MaxLength       =   9
         TabIndex        =   16
         Top             =   3000
         Width           =   2172
      End
      Begin VB.TextBox textCP53 
         Height          =   264
         Left            =   5400
         MaxLength       =   7
         TabIndex        =   11
         Top             =   2376
         Width           =   1095
      End
      Begin VB.TextBox textCP54 
         Height          =   264
         Left            =   6804
         MaxLength       =   7
         TabIndex        =   12
         Top             =   2400
         Width           =   1095
      End
      Begin VB.TextBox textCP21 
         Height          =   264
         Left            =   1100
         MaxLength       =   1
         TabIndex        =   17
         Top             =   3315
         Width           =   372
      End
      Begin VB.TextBox textLC08 
         Height          =   264
         Left            =   5820
         MaxLength       =   1
         TabIndex        =   18
         Top             =   3315
         Width           =   372
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdList 
         Height          =   960
         Left            =   1080
         TabIndex        =   82
         Top             =   4080
         Width           =   7632
         _ExtentX        =   13462
         _ExtentY        =   1693
         _Version        =   393216
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
      Begin MSForms.TextBox textCP29_2 
         Height          =   264
         Left            =   6180
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   1800
         Width           =   2295
         VariousPropertyBits=   671107103
         Size            =   "3413;476"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textLC46_2 
         Height          =   270
         Left            =   -72720
         TabIndex        =   80
         TabStop         =   0   'False
         Top             =   1710
         Width           =   4215
         VariousPropertyBits=   671107103
         Size            =   "7435;476"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textLC45_2 
         Height          =   270
         Left            =   -72720
         TabIndex        =   78
         TabStop         =   0   'False
         Top             =   1410
         Width           =   4215
         VariousPropertyBits=   671107103
         Size            =   "7435;476"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textLC44_2 
         Height          =   270
         Left            =   -72720
         TabIndex        =   76
         TabStop         =   0   'False
         Top             =   1110
         Width           =   4215
         VariousPropertyBits=   671107103
         Size            =   "7435;476"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textLC43_2 
         Height          =   270
         Left            =   -72720
         TabIndex        =   74
         TabStop         =   0   'False
         Top             =   810
         Width           =   4215
         VariousPropertyBits=   671107103
         Size            =   "7435;476"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCP42 
         Height          =   300
         Left            =   -73440
         TabIndex        =   34
         Top             =   2670
         Width           =   6915
         VariousPropertyBits=   671107099
         MaxLength       =   600
         Size            =   "12197;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCP41 
         Height          =   300
         Left            =   -73440
         TabIndex        =   33
         Top             =   2385
         Width           =   6915
         VariousPropertyBits=   671107099
         MaxLength       =   600
         Size            =   "12197;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCP40 
         Height          =   300
         Left            =   -73440
         TabIndex        =   32
         Top             =   2100
         Width           =   6915
         VariousPropertyBits=   671107099
         MaxLength       =   600
         Size            =   "12197;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textLC11_2 
         Height          =   270
         Left            =   -72720
         TabIndex        =   69
         TabStop         =   0   'False
         Top             =   510
         Width           =   4215
         VariousPropertyBits=   671107103
         Size            =   "7435;476"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCP64 
         Height          =   390
         Left            =   1095
         TabIndex        =   19
         Top             =   3615
         Width           =   7575
         VariousPropertyBits=   -1467987941
         MaxLength       =   2000
         ScrollBars      =   2
         Size            =   "13361;688"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCP14_2 
         Height          =   264
         Left            =   1900
         TabIndex        =   44
         TabStop         =   0   'False
         Top             =   1800
         Width           =   2175
         VariousPropertyBits=   671107103
         Size            =   "3413;476"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCP13_2 
         Height          =   264
         Left            =   6180
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   2100
         Width           =   2295
         VariousPropertyBits=   671107103
         Size            =   "3413;476"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textLC05 
         Height          =   300
         Left            =   1440
         TabIndex        =   1
         Top             =   630
         Width           =   6975
         VariousPropertyBits=   671107099
         MaxLength       =   160
         Size            =   "12303;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textLC07 
         Height          =   300
         Left            =   1440
         TabIndex        =   3
         Top             =   1200
         Width           =   6975
         VariousPropertyBits=   671107099
         MaxLength       =   160
         Size            =   "12303;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textLC06 
         Height          =   300
         Left            =   1440
         TabIndex        =   2
         Top             =   930
         Width           =   6975
         VariousPropertyBits=   671107099
         MaxLength       =   160
         Size            =   "12303;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label29 
         Caption         =   "當事人5 :"
         Height          =   195
         Left            =   -74730
         TabIndex        =   81
         Top             =   1710
         Width           =   855
      End
      Begin VB.Label Label28 
         Caption         =   "當事人4 :"
         Height          =   195
         Left            =   -74730
         TabIndex        =   79
         Top             =   1410
         Width           =   855
      End
      Begin VB.Label Label17 
         Caption         =   "當事人3 :"
         Height          =   195
         Left            =   -74730
         TabIndex        =   77
         Top             =   1110
         Width           =   855
      End
      Begin VB.Label Label15 
         Caption         =   "當事人2 :"
         Height          =   195
         Left            =   -74730
         TabIndex        =   75
         Top             =   810
         Width           =   855
      End
      Begin VB.Label Label24 
         Caption         =   "對造日文名稱 :"
         Height          =   255
         Left            =   -74730
         TabIndex        =   73
         Top             =   2700
         Width           =   1275
      End
      Begin VB.Label Label26 
         Caption         =   "對造英文名稱 :"
         Height          =   255
         Left            =   -74730
         TabIndex        =   72
         Top             =   2400
         Width           =   1275
      End
      Begin VB.Label Label27 
         Caption         =   "對造中文名稱 :"
         Height          =   255
         Left            =   -74730
         TabIndex        =   71
         Top             =   2100
         Width           =   1275
      End
      Begin VB.Label Label36 
         Caption         =   "當事人1 :"
         Height          =   195
         Left            =   -74730
         TabIndex        =   70
         Top             =   510
         Width           =   855
      End
      Begin VB.Label Label14 
         Caption         =   "進度備註 :"
         Height          =   255
         Left            =   120
         TabIndex        =   68
         Top             =   3650
         Width           =   975
      End
      Begin VB.Label Label37 
         Caption         =   "收文日 :"
         Height          =   255
         Left            =   120
         TabIndex        =   67
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label13 
         Caption         =   "案件中文名稱 :"
         Height          =   195
         Left            =   120
         TabIndex        =   66
         Top             =   660
         Width           =   1215
      End
      Begin VB.Label Label11 
         Caption         =   "案件日文名稱 :"
         Height          =   195
         Left            =   120
         TabIndex        =   65
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label Label12 
         Caption         =   "案件英文名稱 :"
         Height          =   195
         Left            =   120
         TabIndex        =   64
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "是否為智慧財產權案 :"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   63
         Top             =   1500
         Width           =   1755
      End
      Begin VB.Label Label16 
         Caption         =   "(Y:是)"
         Height          =   255
         Left            =   2520
         TabIndex        =   62
         Top             =   1500
         Width           =   1095
      End
      Begin VB.Label Label20 
         Caption         =   "分所案號 :"
         Height          =   255
         Left            =   4500
         TabIndex        =   61
         Top             =   1500
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "承辦人 :"
         Height          =   255
         Left            =   120
         TabIndex        =   60
         Top             =   1800
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "協辦人員 :"
         Height          =   255
         Left            =   4500
         TabIndex        =   59
         Top             =   1800
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "案件性質 :"
         Height          =   255
         Left            =   120
         TabIndex        =   58
         Top             =   2115
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "智權人員 :"
         Height          =   255
         Index           =   1
         Left            =   4500
         TabIndex        =   57
         Top             =   2100
         Width           =   855
      End
      Begin VB.Label Label6 
         Caption         =   "當事人稱謂 :"
         Height          =   255
         Left            =   120
         TabIndex        =   56
         Top             =   2400
         Width           =   1155
      End
      Begin VB.Label Label25 
         Caption         =   "法定期限 :"
         Height          =   255
         Left            =   4500
         TabIndex        =   55
         Top             =   2730
         Width           =   975
      End
      Begin VB.Label Label7 
         Caption         =   "本所期限 :"
         Height          =   255
         Left            =   120
         TabIndex        =   54
         Top             =   2700
         Width           =   975
      End
      Begin VB.Label Label8 
         Caption         =   "是否算案件數 :"
         Height          =   255
         Left            =   120
         TabIndex        =   53
         Top             =   3000
         Width           =   1215
      End
      Begin VB.Label Label23 
         Caption         =   "(N:不算)"
         Height          =   255
         Left            =   2040
         TabIndex        =   52
         Top             =   3030
         Width           =   1575
      End
      Begin VB.Label Label10 
         Caption         =   "相關總收文號 :"
         Height          =   255
         Left            =   4500
         TabIndex        =   51
         Top             =   3030
         Width           =   1335
      End
      Begin VB.Label Label9 
         Caption         =   "聘任期間 :"
         Height          =   255
         Left            =   4500
         TabIndex        =   50
         Top             =   2400
         Width           =   975
      End
      Begin VB.Line Line1 
         X1              =   6540
         X2              =   6720
         Y1              =   2520
         Y2              =   2520
      End
      Begin VB.Label Label18 
         Caption         =   "是否取締案 :"
         Height          =   255
         Left            =   120
         TabIndex        =   49
         Top             =   3330
         Width           =   1215
      End
      Begin VB.Label Label19 
         Caption         =   "(Y:取締)"
         Height          =   255
         Left            =   1980
         TabIndex        =   48
         Top             =   3315
         Width           =   1095
      End
      Begin VB.Label Label40 
         Caption         =   "本案期限："
         Height          =   255
         Left            =   105
         TabIndex        =   47
         Top             =   4080
         Width           =   975
      End
      Begin VB.Label Label21 
         Caption         =   "是否取消閉卷 :"
         Height          =   255
         Left            =   4500
         TabIndex        =   46
         Top             =   3315
         Width           =   1215
      End
      Begin VB.Label Label22 
         Caption         =   "(Y:取消)"
         Height          =   255
         Left            =   6420
         TabIndex        =   45
         Top             =   3315
         Width           =   1095
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Height          =   180
         Index           =   0
         Left            =   -72120
         TabIndex        =   40
         Top             =   504
         Width           =   45
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Height          =   180
         Index           =   1
         Left            =   -72120
         TabIndex        =   39
         Top             =   1530
         Width           =   45
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Height          =   180
         Index           =   2
         Left            =   -72120
         TabIndex        =   38
         Top             =   2490
         Width           =   45
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Height          =   180
         Index           =   3
         Left            =   -72000
         TabIndex        =   37
         Top             =   3624
         Width           =   45
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "本所案號："
      Height          =   180
      Index           =   6
      Left            =   96
      TabIndex        =   35
      Top             =   504
      Width           =   900
   End
End
Attribute VB_Name = "frm010012_07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2023/10/16 Form2.0已修改 ; grdList從MSFlexGrid改為MSHFlexGrid
'Memo by Morgan 2021/5/13 改成Form2.0 (textLC05,grdList...)
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/25 員工編號欄已修改
'Modify By Sindy 2010/7/22 日期欄已修改
Option Explicit

Dim m_LC01 As String
Dim m_LC02 As String
Dim m_LC03 As String
Dim m_LC04 As String

Dim m_CPKeyList() As String
Dim m_CPKeyCount As Integer
' 收文日
Dim m_CP05 As String
' 收文號
Dim m_CP09 As String
' 案件性質
Dim m_CP10 As String
' 收據編號
Dim m_CP60 As String
' 國家代碼
Dim m_LC15 As String
' 是否閉卷
Dim m_LC08 As String
' 相關總收文號
Dim m_CP43 As String

' 宣告欄位內容結構
Private Type FIELDITEM
   fiName As String
   fiOldData As String
   fiNewData As String
   fiType As Integer
End Type
' 儲存商標基本檔或服務業務基本檔檔案欄位的串列
Dim m_LCList() As FIELDITEM
Dim m_LCCount As Integer
' 儲存案件進度檔檔案欄位的串列
Dim m_CPList() As FIELDITEM
Dim m_CPCount As Integer
'
Dim m_CurrSel As Integer
Dim m_strCP06 As String '原本所期限
Dim m_strCP07 As String '原法定期限
Dim m_bolSetCP27 As Boolean '是否上發文日
Dim m_LOS04 As String, m_LOS04_1 As String 'Added by Lydia 2023/02/13 法務案之案源單號, 介紹人, 介紹人第一人

Public Sub SetData(ByVal strData As String, ByVal nType As Integer, ByVal bClear As Boolean)
   If bClear Then
      m_LC01 = Empty
      m_LC02 = Empty
      m_LC03 = Empty
      m_LC04 = Empty
      m_CP10 = Empty
      '92.03.27 nick
      m_CP09 = Empty
   End If
   
   Select Case nType
      Case 0: m_LC01 = strData
      Case 1: m_LC02 = strData
      Case 2: m_LC03 = strData & String(1 - Len(strData), "0")
      Case 3: m_LC04 = strData & String(2 - Len(strData), "0")
      Case 4: m_CP10 = strData
      Case 6:
         m_CP43 = strData
         textCP43 = m_CP43
      Case 7:
         m_CP09 = strData
   End Select
End Sub

Private Sub ClearAll()
   textLCKey = Empty
   textLC08_2 = Empty
   textCP05 = Empty
   textCP06 = Empty
   textCP07 = Empty
   textCP10 = Empty
   textCP10_2 = Empty
   textCP13 = Empty
   textCP13_2 = Empty
   textCP14 = Empty
   textCP14_2 = Empty
   '911114 nick 邱小姐說刪除
   'textCP16 = Empty
   'textCP17 = Empty
   'textCP18 = Empty
   'textCP19 = Empty
   textCP21 = Empty
   textCP26 = Empty
   textCP29 = Empty
   textCP29_2 = Empty
   textCP43 = Empty
   textCP49 = Empty
   textLC05 = Empty
   textLC06 = Empty
   textLC07 = Empty
   textLC08 = Empty
   textLC08_2 = Empty
   textCP64 = Empty
   textLC11 = Empty
   textLC11_2 = Empty
   'Add By Sindy 2011/1/19
   textLC43 = Empty
   textLC43_2 = Empty
   textLC44 = Empty
   textLC44_2 = Empty
   textLC45 = Empty
   textLC45_2 = Empty
   textLC46 = Empty
   textLC46_2 = Empty
   '2011/1/19 End
   textLC13 = Empty
   textLC16 = Empty
   textCP53 = Empty
   textCP54 = Empty
End Sub

Private Sub cmdCancel_Click()
   Unload Me
   frm010001.Show
End Sub

Private Sub cmdCaseProgress_Click()
   frm010012_03.SetData 0, m_LC01, True
   frm010012_03.SetData 1, m_LC02, False
   frm010012_03.SetData 2, m_LC03, False
   frm010012_03.SetData 3, m_LC04, False
   frm010012_03.SetData 4, m_CP09, False
   'Modified by Lydia 2020/04/21 改為Form型態
   'frm010012_03.SetParent "frm010012_07"
   frm010012_03.SetParent Me
   Me.Hide
   frm010012_03.Show
   frm010012_03.QueryData
End Sub

Private Sub cmdExit_Click()
   Unload Me
   Unload frm010001
End Sub

Private Sub cmdOK_Click()
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   If CheckDataValid = True Then
      If ValidateInput() = False Then
         Exit Sub
      End If
      'Added by Lydia 2015/02/04 所有內部收文, 若有輸入本所期限或法定期限者, 檢查期限不可小於系統日
      'Modified by Lydia 2017/07/31 改為預設和檢查
      'If PUB_CheckCP0607(0, textCP06.Text, textCP07.Text) = False Then Exit Sub
      'Modified by Lyddia 2023/11/08 傳入必需欄位
      'If PUB_CheckCP0607(0, textCP06, textCP07) = False Then Exit Sub
      If PUB_CheckCP0607(0, textCP06, textCP07, "", m_LC15, m_LC01, textCP10) = False Then Exit Sub
      
      ' 設定滑鼠游標為等待狀態
      Screen.MousePointer = vbHourglass
      ' 儲存資料
      OnUpdateField
        'Modify By Cheng 2002/11/06
'      'OnSaveData
      If OnSaveData = False Then MsgBox "存檔失敗，請洽系統管理員 !", vbCritical: Screen.MousePointer = vbDefault: Exit Sub
      
      'Added by Lydia 2023/02/08 內部收文補收款，智權人員為SXX部門時，要發MAIL給杜協理及智權人員
      'Modified by Lydia 2023/02/13 增加判斷案源之介紹人
      'If (m_LC01 = "L" Or m_LC01 = "LA" Or m_LC01 = "ACS" Or m_LC01 = "CFL") And textCP10 <> "" And InStr(textCP10_2, "補收款") > 0 And Left(GetST15(textCP13), 1) = "S" Then
      If m_LOS04_1 <> "" Then
         strExc(5) = m_LOS04_1
      Else
         strExc(5) = textCP13
      End If
      If (m_LC01 = "L" Or m_LC01 = "LA" Or m_LC01 = "ACS" Or m_LC01 = "CFL") And textCP10 <> "" And InStr(textCP10_2, "補收款") > 0 And Left(GetST15(strExc(5)), 1) = "S" Then
      'end 2023/02/13
          strExc(0) = m_LC01 & "-" & m_LC02 & "-" & m_LC03 & "-" & m_LC04
          strExc(1) = "本所案號：" & strExc(0) & vbCrLf & _
                           "案件名稱：" & textLC05 & vbCrLf & _
                           "當事人1：" & textLC11 & " " & textLC11_2 & vbCrLf & _
                           "相關國家：" & GetPrjNationName(m_LC15) & vbCrLf & _
                           "補收款費用：0" & vbCrLf & _
                           "補收款備註：" & Trim(textCP64)
          strExc(2) = Pub_GetSpecMan("全所智權部主管")
          'Modified by Lydia 2023/02/13
          'If InStr(strExc(2), textCP13) = 0 Then
          '    strExc(2) = strExc(2) & ";" & textCP13
          If InStr(strExc(2), strExc(5)) = 0 Then
              strExc(2) = strExc(2) & ";" & strExc(5)
          'end 2023/02/13
          End If
          PUB_SendMail strUserNum, strExc(2), "", strExc(0) & "內部收文補收款通知!", strExc(1)
      End If
      'end 2023/02/08
      
      ' 設定滑鼠游標為預設
      Screen.MousePointer = vbDefault

      ' 回到收文的畫面
      frm010001.SetData m_CP09, 0, True
      frm010001.SetData m_LC01, 1, False
      frm010001.SetData m_LC02, 2, False
      frm010001.SetData m_LC03, 3, False
      frm010001.SetData m_LC04, 4, False
      frm010001.Show
      ClearAll
      Unload Me
   End If
End Sub

Private Sub cmdRelate_Click()
   Where1103ComeFrom Me, m_LC01, m_LC02, m_LC03, m_LC04
End Sub

' 畫面被載入時
Private Sub Form_Load()
   ' 設定控制項的背景顏色
   textLCKey.BackColor = &H8000000F
   textLC11_2.BackColor = &H8000000F
   'Add By Sindy 2011/1/19
   textLC43_2.BackColor = &H8000000F
   textLC44_2.BackColor = &H8000000F
   textLC45_2.BackColor = &H8000000F
   textLC46_2.BackColor = &H8000000F
   '2011/1/19 End
   textCP10_2.BackColor = &H8000000F
   textCP13_2.BackColor = &H8000000F
   textCP14_2.BackColor = &H8000000F
   textCP29_2.BackColor = &H8000000F
   textLC08_2.BackColor = &H8000000F
   
   'add by sonia 2019/8/7
   Me.SSTab1.Tab = 0
   If m_LC01 = "ACS" Then
      SSTab1.TabCaption(1) = "當事人"
      textLC13.Visible = False
      textLC13.Enabled = False
      textCP29.Visible = False
      textCP29.Enabled = False
      textCP49.Visible = False
      textCP49.Enabled = False
      textCP53.Visible = False
      textCP53.Enabled = False
      textCP54.Visible = False
      textCP54.Enabled = False
      textCP21.Visible = False
      textCP21.Enabled = False
      Label2(4).Visible = False
      Label16.Visible = False
      Label4.Visible = False
      Label6.Visible = False
      Label9.Visible = False
      Label18.Visible = False
      Label19.Visible = False
      textCP40.Visible = False
      textCP40.Enabled = False
      textCP41.Visible = False
      textCP41.Enabled = False
      textCP42.Visible = False
      textCP42.Enabled = False
      Label24.Visible = False
      Label26.Visible = False
      Label27.Visible = False
   End If
   'end 2019/8/7

   MoveFormToCenter Me
End Sub

' 初始化 GridList
Private Sub InitialGrdList()
   grdList.Clear
   grdList.Rows = 1
   '911111 nick
   'grdList.Cols = 11
   grdList.Cols = 12
   grdList.ColWidth(0) = 300
   grdList.row = 0
   grdList.col = 1
   grdList.Text = "下一程序"
   grdList.ColWidth(1) = 1200
   grdList.col = 2
   grdList.Text = "本所期限"
   grdList.ColWidth(2) = 1000
   grdList.col = 3
   grdList.Text = "法定期限"
   grdList.ColWidth(3) = 1000
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
   grdList.ColWidth(7) = 1200
   grdList.col = 8
   grdList.Text = "收文號"
   grdList.ColWidth(8) = 0
   grdList.col = 9
   grdList.Text = "下一程序代號"
   grdList.ColWidth(9) = 0
   grdList.col = 10
   grdList.Text = "序號"
   grdList.ColWidth(10) = 0
   '911111 nick add
   grdList.col = 11
   grdList.Text = "序號"
   grdList.ColWidth(11) = 0
End Sub

Private Sub grdList_KeyDown(KeyCode As Integer, Shift As Integer)
   ' 案件性質必須為延期的才可以選取
   If KeyCode = vbKeySpace Then
      If grdList.row > 0 Then
         grdList.col = 0
         If grdList.Text = "V" Then
            grdList.Text = Empty
         Else
            'Added by Lydia 2021/08/31 各系統之分案作業和內部收文作業：勾選下一程序的期限，且該收文的案件性質與下一程序的案件性質不同，請SHOW訊息提醒
            If Pub_CheckNpTheSameShow(m_LC01, textCP10, Trim("" & grdList.TextMatrix(grdList.row, 9))) = False Then
                Exit Sub
            End If
            'end 2021/08/31
            grdList.Text = "V"
            '911018 nick 當有勾選第一筆時，將本所期限，法定期限，備註，相關總收文號更新
            '911111 nick 邱小姐說改成若本所期限沒值時，以勾的該筆代 本所期限，法定期限，備註，相關總收文號 到上方
            '            智權人員沒值時，以勾的該筆代智權人員
            'If grdList.Row = 1 Then
             If textCP06.Text = "" Then
                grdList.col = 2
                textCP06 = grdList.Text
                grdList.col = 3
                textCP07 = grdList.Text
                grdList.col = 8
                textCP43 = grdList.Text
                grdList.col = 6
                Dim nIndex As Integer
                For nIndex = 0 To m_LCCount - 1
                     If m_LCList(nIndex).fiName = "CP64" Then
                         m_LCList(nIndex).fiNewData = m_LCList(nIndex).fiNewData & grdList.Text
                     End If
                Next nIndex
             End If
             If textCP13.Text = "" Then
                grdList.col = 11
                textCP13 = grdList.Text
                '911115 nick
                textCP13_2 = GetStaffName(textCP13)
             End If
            'End If
         End If
      End If
   End If
EXITSUB:
End Sub

Private Sub grdList_SelChange()
   If grdList.row > 0 Then
      grdList.col = 0
      If grdList.Text = "V" Then
         grdList.Text = Empty
      Else
         grdList.Text = "V"
            'Added by Lydia 2021/08/31 各系統之分案作業和內部收文作業：勾選下一程序的期限，且該收文的案件性質與下一程序的案件性質不同，請SHOW訊息提醒
            If Pub_CheckNpTheSameShow(m_LC01, textCP10, Trim("" & grdList.TextMatrix(grdList.row, 9))) = False Then
                Exit Sub
            End If
            'end 2021/08/31
            '911018 nick 當有勾選第一筆時，將本所期限，法定期限，備註，相關總收文號更新
            '911111 nick 邱小姐說改成若本所期限沒值時，以勾的該筆代 本所期限，法定期限，備註，相關總收文號 到上方
            '            智權人員沒值時，以勾的該筆代智權人員
            'If grdList.Row = 1 Then
             If textCP06.Text = "" Then
                grdList.col = 2
                textCP06 = grdList.Text
                grdList.col = 3
                textCP07 = grdList.Text
                grdList.col = 8
                textCP43 = grdList.Text
                grdList.col = 6
                Dim nIndex As Integer
                For nIndex = 0 To m_LCCount - 1
                     If m_LCList(nIndex).fiName = "CP64" Then
                         m_LCList(nIndex).fiNewData = m_LCList(nIndex).fiNewData & grdList.Text
                     End If
                Next nIndex
             End If
             If textCP13.Text = "" Then
                grdList.col = 11
                textCP13 = grdList.Text
                '911115 nick
                textCP13_2 = GetStaffName(textCP13)
             End If
            'End If
      End If
   End If
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

' 清除商標基本檔檔案欄位串列
Private Sub ClearLCFieldList()
   If m_LCCount > 0 Then
      Erase m_LCList
   End If
   m_LCCount = 0
End Sub

' 設定商標基本檔或服務業務基本檔欄位串列中的欄位內容
Private Sub SetLCFieldOldData(ByVal strFieldName As String, ByVal strFieldData As String, ByVal nFieldType As Integer)
   Dim nPos As Integer
   Dim bFind As Boolean
   bFind = False
   For nPos = 0 To m_LCCount - 1
      If m_LCList(nPos).fiName = strFieldName Then
         bFind = True
         m_LCList(nPos).fiOldData = strFieldData
         m_LCList(nPos).fiNewData = strFieldData
         m_LCList(nPos).fiType = nFieldType
         Exit For
      End If
   Next nPos
   If bFind = False Then
      ReDim Preserve m_LCList(m_LCCount + 1)
      m_LCList(m_LCCount).fiName = strFieldName
      m_LCList(m_LCCount).fiOldData = strFieldData
      m_LCList(m_LCCount).fiNewData = strFieldData
      m_LCList(m_LCCount).fiType = nFieldType
      m_LCCount = m_LCCount + 1
   End If
End Sub

' 設定商標基本檔或服務業務基本檔欄位串列中的欄位內容
Private Sub SetLCFieldNewData(ByVal strFieldName As String, ByVal strFieldData As String)
   Dim nPos As Integer
   Dim bFind As Boolean
   bFind = False
   For nPos = 0 To m_LCCount - 1
      If m_LCList(nPos).fiName = strFieldName Then
         bFind = True
         m_LCList(nPos).fiNewData = strFieldData
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

' 更新欄位的內容
Private Sub OnUpdateField()
   SetCPFieldNewData "CP01", m_LC01
   SetCPFieldNewData "CP02", m_LC02
   SetCPFieldNewData "CP03", m_LC03
   SetCPFieldNewData "CP04", m_LC04
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
   ' 收文號
   'Modify by Morgan 2004/2/18
   '新增才要重抓收文號
    If frm010001.intModifyKind = 0 Then
        m_CP09 = AutoNo("B", 6)
    End If
   SetCPFieldNewData "CP09", m_CP09
   ' 案件性質
   SetCPFieldNewData "CP10", textCP10
   ' 業務區
   SetCPFieldNewData "CP12", GetSalesArea(textCP13)
   ' 智權人員
   SetCPFieldNewData "CP13", textCP13
   ' 承辦人員
   SetCPFieldNewData "CP14", textCP14
   '911114 nick 邱小姐說刪除
   ' 費用
   'SetCPFieldNewData "CP16", textCP16
   ' 規費
   'SetCPFieldNewData "CP17", textCP17
   ' 點數
   'SetCPFieldNewData "CP18", textCP18
   ' 後金
   'SetCPFieldNewData "CP19", textCP19
   
   '911114 nick
   SetCPFieldNewData "CP11", "90"
   'SetCPFieldNewData "CP08", ""
   'SetCPFieldNewData "CP48", ""
   SetCPFieldNewData "CP20", "N"
   SetCPFieldNewData "CP32", "N"
   
   ' 是否取締案
   SetCPFieldNewData "CP21", textCP21
   ' 是否算案件數
   SetCPFieldNewData "CP26", textCP26
   ' 協辦人員
   If Not IsEmptyText(textCP29) Then
      '911113 nick 協辦人員應該為 5-6 碼
      'SetCPFieldNewData "CP29", textCP29 & String(9 - Len(textCP29), "0")
      SetCPFieldNewData "CP29", textCP29
   Else
      SetCPFieldNewData "CP29", Empty
   End If
   ' 對造名稱(中)
   SetCPFieldNewData "CP40", textCP40
   ' 對造名稱(英)
   SetCPFieldNewData "CP41", textCP41
   ' 對造名稱(日)
   SetCPFieldNewData "CP42", textCP42
   ' 相關總收文號
   SetCPFieldNewData "CP43", textCP43
   ' 當事人稱謂
   SetCPFieldNewData "CP49", textCP49
   ' 進度備註
   SetCPFieldNewData "CP64", textCP64
   ' 聘任期間起
   If Not IsEmptyText(textCP53) Then
      SetCPFieldNewData "CP53", DBDATE(textCP53)
   Else
      SetCPFieldNewData "CP53", Empty
   End If
   ' 聘任期間迄
   If Not IsEmptyText(textCP54) Then
      SetCPFieldNewData "CP54", DBDATE(textCP54)
   Else
      SetCPFieldNewData "CP54", Empty
   End If
   
   Select Case m_LC01
      ' 系統類別為L的為法務
      Case "L":
         ' 案件名稱(中)
         SetLCFieldNewData "LC05", textLC05
         ' 案件名稱(英)
         SetLCFieldNewData "LC06", textLC06
         ' 案件名稱(日)
         SetLCFieldNewData "LC07", textLC07
         ' 當事人1
         If Not IsEmptyText(textLC11) Then
            SetLCFieldNewData "LC11", textLC11 & String(9 - Len(textLC11), "0")
         Else
            SetLCFieldNewData "LC11", Empty
         End If
         'Add By Sindy 2011/1/19
         ' 當事人2
         If Not IsEmptyText(textLC43) Then
            SetLCFieldNewData "LC43", textLC43 & String(9 - Len(textLC43), "0")
         Else
            SetLCFieldNewData "LC43", Empty
         End If
         ' 當事人3
         If Not IsEmptyText(textLC44) Then
            SetLCFieldNewData "LC44", textLC44 & String(9 - Len(textLC44), "0")
         Else
            SetLCFieldNewData "LC44", Empty
         End If
         ' 當事人4
         If Not IsEmptyText(textLC45) Then
            SetLCFieldNewData "LC45", textLC45 & String(9 - Len(textLC45), "0")
         Else
            SetLCFieldNewData "LC45", Empty
         End If
         ' 當事人5
         If Not IsEmptyText(textLC46) Then
            SetLCFieldNewData "LC46", textLC46 & String(9 - Len(textLC46), "0")
         Else
            SetLCFieldNewData "LC46", Empty
         End If
         '2011/1/19 End
         ' 是否為智慧財產權
         SetLCFieldNewData "LC13", textLC13
         ' 分所案號
         SetLCFieldNewData "LC16", textLC16
      Case "LA":
         ' 當事人1
         If Not IsEmptyText(textLC11) Then
            SetLCFieldNewData "HC05", textLC11 & String(9 - Len(textLC11), "0")
         Else
            SetLCFieldNewData "HC05", Empty
         End If
         'Add By Sindy 2011/1/19
         ' 當事人2
         If Not IsEmptyText(textLC43) Then
            SetLCFieldNewData "HC24", textLC43 & String(9 - Len(textLC43), "0")
         Else
            SetLCFieldNewData "HC24", Empty
         End If
         ' 當事人3
         If Not IsEmptyText(textLC44) Then
            SetLCFieldNewData "HC25", textLC44 & String(9 - Len(textLC44), "0")
         Else
            SetLCFieldNewData "HC25", Empty
         End If
         ' 當事人4
         If Not IsEmptyText(textLC45) Then
            SetLCFieldNewData "HC26", textLC45 & String(9 - Len(textLC45), "0")
         Else
            SetLCFieldNewData "HC26", Empty
         End If
         ' 當事人5
         If Not IsEmptyText(textLC46) Then
            SetLCFieldNewData "HC27", textLC46 & String(9 - Len(textLC46), "0")
         Else
            SetLCFieldNewData "HC27", Empty
         End If
         '2011/1/19 End
         ' 案件中文名稱
         SetLCFieldNewData "HC06", textLC05
         ' 分所案號
         SetLCFieldNewData "HC07", textLC16
   End Select
End Sub

'Modify By Cheng 2002/11/06
'Private Function OnSaveData()
Public Function OnSaveData() As Boolean
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   Dim strCF13 As String
   Dim strCF14 As String
   Dim strDay As String
   Dim strDate As String
   Dim strNP07 As String
   Dim strNP22 As String
   Dim strTemp As String
   Dim nIndex As Integer
   Dim nSubIndex As Integer
      
'Add By Cheng 2002/11/06
On Error GoTo ErrorHandler
OnSaveData = True
cnnConnection.BeginTrans

   ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   ' 更新案件進度檔
    'Modify By Cheng 2002/11/06
'   SaveNewCaseProgress
    'Modify By Cheng 2003/03/28
    '若為新增
    If frm010001.intModifyKind = 0 Then
        If SaveNewCaseProgress = False Then GoTo ErrorHandler
    '若為修改
    ElseIf frm010001.intModifyKind = 1 Then
        OnUpdateCaseProgress
    End If
   
   '92.2.28 ADD BY SONIA
   If textCP10 = "0" Then
      strSql = "UPDATE CASEPROGRESS SET CP27=" & GetTodayDate & _
               "WHERE CP09 = '" & m_CPList(4).fiNewData & "' "
      cnnConnection.Execute strSql
      
   'Added by Morgan 2022/8/18 +m_bolSetCP27
   ElseIf m_bolSetCP27 Then
      strSql = "UPDATE CASEPROGRESS SET CP27=" & DBDATE(textCP05) & _
               "WHERE CP09 = '" & m_CPList(4).fiNewData & "' "
      cnnConnection.Execute strSql
      
   End If
   '92.2.28 END
   
   Select Case m_LC01
      ' 更新商標基本檔
      '911114 nick 系統別錯誤
      'Case "LC":
      Case "LA":
        'Modify By Cheng 2002/11/06
'         OnUpdateLowCase
         If OnUpdateHireCase = False Then GoTo ErrorHandler
      ' 更新服務業務基本檔
      Case Else:
        'Modify By Cheng 2002/11/06
'         OnUpdateHireCase
         If OnUpdateLawCase = False Then GoTo ErrorHandler
   End Select

   ' 更新基本檔是否閉卷, 閉卷日期, 閉卷原因
   If textLC08 = "Y" Then
      Select Case m_LC01
         ' 更新基本檔
         ' 法務
         Case "L":
            '911114 ncik tablename 錯誤
            'strSQL = "UPDATE LOWCASE SET LC08=NULL, LC09=NULL, LC10=NULL " & _
                     "WHERE LC01 = '" & m_LC01 & "' AND " & _
                           "LC02 = '" & m_LC02 & "' AND " & _
                           "LC03 = '" & m_LC03 & "' AND " & _
                           "LC04 = '" & m_LC04 & "' "
            strSql = "UPDATE LaWCASE SET LC08=NULL, LC09=NULL, LC10=NULL " & _
                     "WHERE LC01 = '" & m_LC01 & "' AND " & _
                           "LC02 = '" & m_LC02 & "' AND " & _
                           "LC03 = '" & m_LC03 & "' AND " & _
                           "LC04 = '" & m_LC04 & "' "
         ' 顧問
         Case "LA":
            strSql = "UPDATE HIRECASE SET HC09=NULL, HC10=NULL, HC11=NULL " & _
                     "WHERE HC01 = '" & m_LC01 & "' AND " & _
                           "HC02 = '" & m_LC02 & "' AND " & _
                           "HC03 = '" & m_LC03 & "' AND " & _
                           "HC04 = '" & m_LC04 & "' "
      End Select
      cnnConnection.Execute strSql
   End If
    
   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   ' 更新使用者所選取的本案期限資料
   For nIndex = 1 To grdList.Rows - 1
      ' 判斷該列是否有被選取
      If grdList.TextMatrix(nIndex, 0) = "V" Then
         strNP07 = grdList.TextMatrix(nIndex, 9)
         strNP22 = grdList.TextMatrix(nIndex, 10)
         'Modify By Sindy 2016/11/2 + ,np24=" & CNULL(m_CP09) &
         strSql = "UPDATE NextProgress SET NP06 = 'Y',np24=" & CNULL(m_CP09) & _
                  " WHERE NP02 = '" & m_LC01 & "' AND " & _
                        "NP03 = '" & m_LC02 & "' AND " & _
                        "NP04 = '" & m_LC03 & "' AND " & _
                        "NP05 = '" & m_LC04 & "' AND " & _
                        "NP07 = " & strNP07 & " AND " & _
                        "NP22 = " & strNP22 & " "
         Pub_SeekTbLog strSql 'Added by Lydia 2021/08/31 各系統之分案作業和內部收文作業，若畫面勾選下一程序期限且存檔有上續辦Y的都寫Log以便事後能追蹤
         cnnConnection.Execute strSql
      End If
   Next nIndex
   '911018 nick 當有相關總收文號時，要將總收文號該筆更新成續辦，因為只會有一筆時才會讀出來秀畫面，所以不用np22
   '91.11.10 MODIFY BY SONIA
   'If textCP43 <> "" Then
   '     strSQL = "update nextprogress set np06='Y' where np01='" & textCP43 & "' "
   '     cnnConnection.Execute strSQL
   'End If
   '91.11.10 END
'Add By Cheng 2002/11/06
cnnConnection.CommitTrans
Exit Function
ErrorHandler:
    OnSaveData = False
    '911113 nick 程弘漏寫
    cnnConnection.RollbackTrans
End Function

' 新增案件進度檔
'Modify By Cheng 2002/11/06
'Private Sub SaveNewCaseProgress()
Private Function SaveNewCaseProgress() As Boolean
   Dim strTmp As String
   Dim strSql As String
   Dim nIndex As Integer
   
'Add By Cheng 2002/11/06
On Error GoTo ErrorHandler
SaveNewCaseProgress = True

   strSql = "INSERT INTO CaseProgress ("
   For nIndex = 0 To m_CPCount - 1
      If Not IsEmptyText(m_CPList(nIndex).fiNewData) Then
         If nIndex <> 0 Then strSql = strSql & ","
         strSql = strSql & m_CPList(nIndex).fiName
      End If
   Next nIndex
   strSql = strSql & ") "
   strSql = strSql & "VALUES ("
   For nIndex = 0 To m_CPCount - 1
      If Not IsEmptyText(m_CPList(nIndex).fiNewData) Then
         If nIndex <> 0 Then strSql = strSql & ","
         If m_CPList(nIndex).fiType = 0 Then
            '911028 nick 加 chgsql
            'strSQL = strSQL & "'" & m_CPList(nIndex).fiNewData & "'"
            strSql = strSql & "'" & ChgSQL(m_CPList(nIndex).fiNewData) & "'"
         Else
            strSql = strSql & m_CPList(nIndex).fiNewData
         End If
      End If
   Next nIndex
   strSql = strSql & ") "
   
   cnnConnection.Execute strSql
'Add By Cheng 2002/11/06
Exit Function
ErrorHandler:
    SaveNewCaseProgress = False
End Function

' 更新法務基本檔的相關欄位
'Modify By Cheng 2002/11/06
'Private Sub OnUpdateLowCase()
Private Function OnUpdateLawCase() As Boolean
   Dim strTmp As String
   Dim strSql As String
   Dim nIndex As Integer
   Dim bFirst As Boolean
   Dim bDifference As Boolean
 
 'Add By Cheng 2002/11/06
 On Error GoTo ErrorHandler
 OnUpdateLawCase = True
 
   ' 更新案件進度檔
   '911114 nick table name 錯誤
   'strSQL = "UPDATE LOWCASE SET "
   strSql = "UPDATE LaWCASE SET "
   bFirst = True
   bDifference = False
   For nIndex = 0 To m_LCCount - 1
      strTmp = Empty
      If m_LCList(nIndex).fiOldData <> m_LCList(nIndex).fiNewData Then
         bDifference = True
         If m_LCList(nIndex).fiType = 0 Then
            If m_LCList(nIndex).fiNewData = Empty Then
               strTmp = m_LCList(nIndex).fiName & " = " & "NULL"
            Else
               strTmp = m_LCList(nIndex).fiName & " = '" & ChgSQL(m_LCList(nIndex).fiNewData) & "'"
            End If
         Else
            If m_LCList(nIndex).fiNewData = Empty Then
               strTmp = m_LCList(nIndex).fiName & " = " & "NULL"
            Else
               strTmp = m_LCList(nIndex).fiName & " = " & m_LCList(nIndex).fiNewData
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
                  "WHERE LC01 = '" & m_LC01 & "' AND " & _
                        "LC02 = '" & m_LC02 & "' AND " & _
                        "LC03 = '" & m_LC03 & "' AND " & _
                        "LC04 = '" & m_LC04 & "' "
   ' 執行SQL指令
   If bDifference = True Then: cnnConnection.Execute strSql
'Add By Cheng 2002/11/06
Exit Function
ErrorHandler:
    OnUpdateLawCase = False
End Function

' 更新顧問基本檔的相關欄位
'Modify By Cheng 2002/11/06
'Private Sub OnUpdateHireCase()
Private Function OnUpdateHireCase() As Boolean
   Dim strTmp As String
   Dim strSql As String
   Dim nIndex As Integer
   Dim bFirst As Boolean
   Dim bDifference As Boolean
 
'Add By Cheng 2002/11/06
On Error GoTo ErrorHandler
OnUpdateHireCase = True
   
   ' 更新案件進度檔
   strSql = "UPDATE HIRECASE SET "
   bFirst = True
   bDifference = False
   For nIndex = 0 To m_LCCount - 1
      strTmp = Empty
      If m_LCList(nIndex).fiOldData <> m_LCList(nIndex).fiNewData Then
         bDifference = True
         If m_LCList(nIndex).fiType = 0 Then
            If m_LCList(nIndex).fiNewData = Empty Then
               strTmp = m_LCList(nIndex).fiName & " = " & "NULL"
            Else
               strTmp = m_LCList(nIndex).fiName & " = '" & ChgSQL(m_LCList(nIndex).fiNewData) & "'"
            End If
         Else
            If m_LCList(nIndex).fiNewData = Empty Then
               strTmp = m_LCList(nIndex).fiName & " = " & "NULL"
            Else
               strTmp = m_LCList(nIndex).fiName & " = " & m_LCList(nIndex).fiNewData
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
                  "WHERE HC01 = '" & m_LC01 & "' AND " & _
                        "HC02 = '" & m_LC02 & "' AND " & _
                        "HC03 = '" & m_LC03 & "' AND " & _
                        "HC04 = '" & m_LC04 & "' "
   ' 執行SQL指令
   If bDifference = True Then: cnnConnection.Execute strSql
'Add By Cheng 2002/11/06
Exit Function
ErrorHandler:
    OnUpdateHireCase = False
End Function

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
     
      ' 案件性質
      If IsNull(rsTmp.Fields("CP10")) = False Then
         m_CP10 = rsTmp.Fields("CP10")
         textCP10 = rsTmp.Fields("CP10")
         textCP10_Validate False
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
      SetCPFieldOldData "CP12", rsTmp.Fields("CP12"), 0
      
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
      
      'Added by Morgan 2021/5/13
      ' 協辦人員
      If IsNull(rsTmp.Fields("CP29")) = False Then
         textCP29 = rsTmp.Fields("CP29")
         textCP29_Validate False
      End If
      SetCPFieldOldData "CP29", textCP29, 0
      'end 2021/5/13
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
      ' 進度備註
      If IsNull(rsTmp.Fields("CP64")) = False Then
         textCP64 = rsTmp.Fields("CP64")
      End If
      SetCPFieldOldData "CP64", textCP64, 0
      '911018 nick 不要此欄位
      ' 取消收文日期
      'If IsNull(rsTmp.Fields("CP57")) = False Then
      '   textCP57 = TAIWANDATE(rsTmp.Fields("CP57"))
      'End If
      
      '910626 Sieg
      '收據編號
      If IsNull(rsTmp.Fields("CP60")) = False Then
         m_CP60 = rsTmp.Fields("CP60")
      Else
         m_CP60 = ""
      End If
      
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
               
      ' 卷宗性質不為1時, 案件中英日文名稱從案件進度檔中帶入
      
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Sub

' 讀取資料庫
Public Sub QueryData()
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   
   ' 顯示畫面為第一頁
   'SSTab1.Tab = 0
   
   textLC08_2 = Empty
   m_LC08 = Empty
   
   ' 先清除基本檔欄位串列
   ClearLCFieldList
   ' 先清除案件進度檔欄位串列
   ClearCPFieldList
   
   m_CP05 = TAIWANDATE(SystemDate())
   textCP05 = m_CP05
   textCP10 = m_CP10
   textCP10_Validate False
    
   'Added by Lydia 2020/10/26 法律所人員做內部收文,則智權人員掛本人(ex. LA-3346陳頌恩指定法務人員林律師，所以陳亮之印接洽單，現在又回中所辦。)
   If Left(Pub_StrUserSt03, 1) = "L" Then
        textCP13 = strUserNum
        textCP13_Validate False
   Else
   'end 2020/10/26
        '2008/10/31 add by sonia
        textCP13 = PUB_GetAKindSalesNo(m_LC01, m_LC02, m_LC03, m_LC04)
        textCP13_Validate False
        '2008/10/31 end
   End If 'Added by Lydia 2020/10/26
   
   ' 本所案號
   textLCKey = m_LC01 & m_LC02 & m_LC03 & m_LC04
      
   Select Case m_LC01
      ' 系統類別為L的為法務基本檔
      '911114 nick 系統別錯誤
      'Case "LC":
      Case "LA":
         QueryHireCase
      ' 顧問
      Case Else:
         QueryLawCase
   End Select
   
   ' 取得案件進度檔的欄位
   '92.03.27 nick 修正
   If frm010001.intModifyKind = 0 Then
        QueryCaseProgressWithNewCP
   Else
        QueryCaseProgress
   End If
   

   ' 是否閉卷
   If m_LC08 = "Y" Then
      textLC08_2 = "本案已閉卷"
   Else
      textLC08_2 = Empty
   End If
   
   ' 依讀取的是法務
   If m_LC01 = "L" Then
      EnableTextBox textLC13, True
   Else
      EnableTextBox textLC13, False
   End If
   
   ' 系統類別為LA, 案件性質為聘任時, 要顯示輸入聘任期間
   If m_LC01 = "LA" And m_CP10 = "0" Then
      EnableTextBox textCP53, True
      EnableTextBox textCP54, True
      Label9.Visible = True
      textCP53.Visible = True
      textCP54.Visible = True
      Line1.Visible = True
   Else
      EnableTextBox textCP53, False
      EnableTextBox textCP54, False
      Label9.Visible = False
      textCP53.Visible = False
      textCP54.Visible = False
      Line1.Visible = False
   End If

   ' 更新本案期限的資料
   UpdateGrdList m_LC01, m_LC02, m_LC03, m_LC04
   
   '911018 nick 新增時要待下一程序資料     本所期限，法定期限，收文號==>相關總收文號，備註==>進度備註    #只有一筆時，且本所案號和案件性質都要輸入且找的到
   If frm010001.intModifyKind = 0 Then
        If m_LC01 <> "" And m_LC02 <> "" And m_LC03 <> "" And m_LC04 <> "" And m_CP10 <> "" Then
            Dim nick911018rs As New ADODB.Recordset
            Dim nickstrsql As String
            Set nick911018rs = New ADODB.Recordset
            '911111 nick 邱小姐說要加入 np06 is null  np06<>'Y'(包含 null) 同意義
            'nickstrsql = "select * from nextprogress where np02='" & m_LC01 & "' and np03='" & m_LC02 & "' and np04='" & m_LC03 & "' and np05='" & m_LC04 & "' and np07=" & m_CP10 & " "
            '911203 nick 改成先抓 null 若 0 筆，則再抓 <>'Y'，但如果大於一筆，則都不代
            'nickstrsql = "select * from nextprogress where np02='" & m_LC01 & "' and np03='" & m_LC02 & "' and np04='" & m_LC03 & "' and np05='" & m_LC04 & "' and np07=" & m_CP10 & " and (np06 <>'Y' or np06 is null) "
            nickstrsql = "select * from nextprogress where np02='" & m_LC01 & "' and np03='" & m_LC02 & "' and np04='" & m_LC03 & "' and np05='" & m_LC04 & "' and np07=" & m_CP10 & " and np06 is null "
            nick911018rs.CursorLocation = adUseClient
            nick911018rs.Open nickstrsql, cnnConnection, adOpenStatic, adLockReadOnly
            If nick911018rs.RecordCount = 1 Then
                textCP06 = ChangeWStringToTString(CheckStr(nick911018rs.Fields("np08").Value))
                textCP07 = ChangeWStringToTString(CheckStr(nick911018rs.Fields("np09").Value))
                textCP43 = CheckStr(nick911018rs.Fields("np01").Value)
                If Not IsNull(nick911018rs.Fields("np15").Value) Then
                   SetCPFieldNewData "CP64", CheckStr(nick911018rs.Fields("np15").Value)
                Else
                   SetCPFieldNewData "CP64", Empty
                End If
                '91.11.10 ADD BY SONIA
                textCP13 = CheckStr(nick911018rs.Fields("np10").Value)
                textCP13_Validate False
                '91.11.10 END
                '911030 nick 自動上勾
                Dim nickI As Integer
                For nickI = 1 To grdList.Rows - 1
                    'edit by nick 2004/09/08
                    'If Trim(grdList.TextMatrix(nickI, 9)) = Trim(CheckStr(nick911018rs.Fields("np07").Value)) And grdList.TextMatrix(nickI, 2) = textCP06 And grdList.TextMatrix(nickI, 3) = textCP07 Then
                    If Trim(grdList.TextMatrix(nickI, 9)) = Trim(CheckStr(nick911018rs.Fields("np07").Value)) And Val(grdList.TextMatrix(nickI, 2)) = Val(textCP06) And Val(grdList.TextMatrix(nickI, 3)) = Val(textCP07) And textCP10.Text <> "411" Then
                        grdList.TextMatrix(nickI, 0) = "V"
                    End If
                Next nickI
            Else
                '911203 nick 改成先抓 null 若 0 筆，則再抓 <>'Y'，但如果大於一筆，則都不代
                If nick911018rs.RecordCount = 0 Then
                    Set nick911018rs = New ADODB.Recordset
                    nick911018rs.CursorLocation = adUseClient
                    nick911018rs.Open nickstrsql, cnnConnection, adOpenStatic, adLockReadOnly
                    nickstrsql = "select * from nextprogress where np02='" & m_LC01 & "' and np03='" & m_LC02 & "' and np04='" & m_LC03 & "' and np05='" & m_LC04 & "' and np07=" & m_CP10 & " and np06 <>'Y' "
                    If nick911018rs.RecordCount = 1 Then
                        textCP06 = ChangeWStringToTString(CheckStr(nick911018rs.Fields("np08").Value))
                        textCP07 = ChangeWStringToTString(CheckStr(nick911018rs.Fields("np09").Value))
                        textCP43 = CheckStr(nick911018rs.Fields("np01").Value)
                        If Not IsNull(nick911018rs.Fields("np15").Value) Then
                           SetCPFieldNewData "CP64", CheckStr(nick911018rs.Fields("np15").Value)
                        Else
                           SetCPFieldNewData "CP64", Empty
                        End If
                        textCP13 = CheckStr(nick911018rs.Fields("np10").Value)
                        textCP13_Validate False
                        For nickI = 1 To grdList.Rows - 1
                            'edit by nick 2004/09/08
                            'If Trim(grdList.TextMatrix(nickI, 9)) = Trim(CheckStr(nick911018rs.Fields("np07").Value)) And grdList.TextMatrix(nickI, 2) = textCP06 And grdList.TextMatrix(nickI, 3) = textCP07 Then
                            If Trim(grdList.TextMatrix(nickI, 9)) = Trim(CheckStr(nick911018rs.Fields("np07").Value)) And Val(grdList.TextMatrix(nickI, 2)) = Val(textCP06) And Val(grdList.TextMatrix(nickI, 3)) = Val(textCP07) And textCP10.Text <> "411" Then
                                grdList.TextMatrix(nickI, 0) = "V"
                            End If
                        Next nickI
                    End If
                End If
            End If
        End If
   End If
   
   'Added by Lydia 2023/02/13 取得案源資料
   strSql = "select cp05,cp09,los15,los04 from caseprogress,lawofficesource where cp01='" & m_LC01 & "' and cp02='" & m_LC02 & "' and cp03='" & m_LC03 & "' and cp04='" & m_LC04 & "' " & _
                "and substr(cp09,1,1)='A' and cp162=los15(+) and los04 is not null order by cp05 desc "
   intI = 1
   Set rsTmp = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
       m_LOS04 = "" & rsTmp.Fields("los04")
       strSql = PUB_GetNowStaff(m_LOS04, m_LOS04_1)
       m_LOS04 = strSql
   End If
   'end 2023/02/13
   
   ' 設定輸入的位置
   SetInputEntry
   '92.03.27 nick 當查詢時，將確定 disabled
   If frm010001.intModifyKind = 2 Then
        cmdOK.Enabled = False
   End If
End Sub

Public Sub QueryLawCase()
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   
   strSql = "SELECT * FROM LAWCASE " & _
            "WHERE LC01 = '" & m_LC01 & "' AND " & _
                  "LC02 = '" & m_LC02 & "' AND " & _
                  "LC03 = '" & m_LC03 & "' AND " & _
                  "LC04 = '" & m_LC04 & "' "
                  
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      ' 案件中文名稱
      If IsNull(rsTmp.Fields("LC05")) = False Then
         textLC05 = rsTmp.Fields("LC05")
      End If
      SetLCFieldOldData "LC05", textLC05, 0
      ' 案件英文名稱
      If IsNull(rsTmp.Fields("LC06")) = False Then
         textLC06 = rsTmp.Fields("LC06")
      End If
      '911114 nick  欄位寫錯了
      'SetLCFieldOldData "LC065", textLC06, 0
      SetLCFieldOldData "LC06", textLC06, 0
      ' 案件日文名稱
      If IsNull(rsTmp.Fields("LC07")) = False Then
         textLC07 = rsTmp.Fields("LC07")
      End If
      SetLCFieldOldData "LC07", textLC07, 0
      ' 當事人1
      If IsNull(rsTmp.Fields("LC11")) = False Then
         textLC11 = rsTmp.Fields("LC11")
         textLC11_Validate False
      End If
      SetLCFieldOldData "LC11", textLC11, 0
      'Add By Sindy 2011/1/19
      ' 當事人2
      If IsNull(rsTmp.Fields("LC43")) = False Then
         textLC43 = rsTmp.Fields("LC43")
         textLC43_Validate False
      End If
      SetLCFieldOldData "LC43", textLC43, 0
      ' 當事人3
      If IsNull(rsTmp.Fields("LC44")) = False Then
         textLC44 = rsTmp.Fields("LC44")
         textLC44_Validate False
      End If
      SetLCFieldOldData "LC44", textLC44, 0
      ' 當事人4
      If IsNull(rsTmp.Fields("LC45")) = False Then
         textLC45 = rsTmp.Fields("LC45")
         textLC45_Validate False
      End If
      SetLCFieldOldData "LC45", textLC45, 0
      ' 當事人5
      If IsNull(rsTmp.Fields("LC46")) = False Then
         textLC46 = rsTmp.Fields("LC46")
         textLC46_Validate False
      End If
      SetLCFieldOldData "LC46", textLC46, 0
      '2011/1/19 End
      ' 是否為智慧財產權
      If IsNull(rsTmp.Fields("LC13")) = False Then
         textLC13 = rsTmp.Fields("LC13")
      End If
      SetLCFieldOldData "LC13", textLC13, 0
      ' 分所案號
      If IsNull(rsTmp.Fields("LC16")) = False Then
         textLC16 = rsTmp.Fields("LC16")
      End If
      SetLCFieldOldData "LC16", textLC16, 0
      ' 相關國家
      If IsNull(rsTmp.Fields("LC15")) = False Then
         m_LC15 = rsTmp.Fields("LC15")
      End If
      ' 是否閉卷
      If IsNull(rsTmp.Fields("LC08")) = False Then
         m_LC08 = rsTmp.Fields("LC08")
      End If
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Sub

Public Sub QueryHireCase()
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   
   strSql = "SELECT * FROM HIRECASE " & _
            "WHERE HC01 = '" & m_LC01 & "' AND " & _
                  "HC02 = '" & m_LC02 & "' AND " & _
                  "HC03 = '" & m_LC03 & "' AND " & _
                  "HC04 = '" & m_LC04 & "' "
                  
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      ' 當事人1
      If IsNull(rsTmp.Fields("HC05")) = False Then
         textLC11 = rsTmp.Fields("HC05")
         textLC11_Validate False
      End If
      SetLCFieldOldData "HC05", textLC11, 0
      'Add By Sindy 2011/1/19
      ' 當事人2
      If IsNull(rsTmp.Fields("HC24")) = False Then
         textLC43 = rsTmp.Fields("HC24")
         textLC43_Validate False
      End If
      SetLCFieldOldData "HC24", textLC43, 0
      ' 當事人3
      If IsNull(rsTmp.Fields("HC25")) = False Then
         textLC44 = rsTmp.Fields("HC25")
         textLC44_Validate False
      End If
      SetLCFieldOldData "HC25", textLC44, 0
      ' 當事人4
      If IsNull(rsTmp.Fields("HC26")) = False Then
         textLC45 = rsTmp.Fields("HC26")
         textLC45_Validate False
      End If
      SetLCFieldOldData "HC26", textLC45, 0
      ' 當事人5
      If IsNull(rsTmp.Fields("HC27")) = False Then
         textLC46 = rsTmp.Fields("HC27")
         textLC46_Validate False
      End If
      SetLCFieldOldData "HC27", textLC46, 0
      '2011/1/19 End
      ' 案件名稱
      If IsNull(rsTmp.Fields("HC06")) = False Then
         textLC05 = rsTmp.Fields("HC06")
      End If
      SetLCFieldOldData "HC06", textLC05, 0
      ' 分所案號
      If IsNull(rsTmp.Fields("HC07")) = False Then
         textLC16 = rsTmp.Fields("HC07")
      End If
      SetLCFieldOldData "HC07", textLC16, 0
      ' 是否閉卷
      If IsNull(rsTmp.Fields("HC09")) = False Then
         m_LC08 = rsTmp.Fields("HC09")
      End If
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Sub

' 取得案件進度檔的欄位內容
Private Sub QueryCaseProgressWithNewCP()
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   Dim strTemp As String
   Dim strSubSQL As String
   Dim rsSubTmp As ADODB.Recordset
   
   SetCPFieldOldData "CP01", Empty, 0
   SetCPFieldOldData "CP02", Empty, 0
   SetCPFieldOldData "CP03", Empty, 0
   SetCPFieldOldData "CP04", Empty, 0
   SetCPFieldOldData "CP09", Empty, 0
   ' 案件性質
   SetCPFieldOldData "CP10", Empty, 0
   ' 收文日
   SetCPFieldOldData "CP05", Empty, 1
   ' 本所期限
   SetCPFieldOldData "CP06", Empty, 1
   ' 法定期限
   SetCPFieldOldData "CP07", Empty, 1
   ' 業務區
   SetCPFieldOldData "CP12", Empty, 0
   ' 智權人員
   SetCPFieldOldData "CP13", Empty, 0
   ' 承辦人員
   SetCPFieldOldData "CP14", Empty, 0
   ' 費用
   SetCPFieldOldData "CP16", Empty, 1
   ' 規費
   SetCPFieldOldData "CP17", Empty, 1
   ' 點數
   SetCPFieldOldData "CP18", Empty, 1
   ' 後金
   SetCPFieldOldData "CP19", Empty, 1
   ' 是否取締案
   SetCPFieldOldData "CP21", Empty, 0
   ' 是否算案件數
   SetCPFieldOldData "CP26", Empty, 0
   ' 協辦人員
   SetCPFieldOldData "CP29", Empty, 0
   ' 相關總收文號
   SetCPFieldOldData "CP43", Empty, 0
   ' 是否算案件數
   SetCPFieldOldData "CP26", Empty, 0
   ' 對造名稱(中)
   SetCPFieldOldData "CP40", Empty, 0
   ' 對造名稱(英)
   SetCPFieldOldData "CP41", Empty, 0
   ' 對造名稱(日)
   SetCPFieldOldData "CP42", Empty, 0
   ' 當事人稱謂
   SetCPFieldOldData "CP49", Empty, 0
   ' 聘任期間起
   SetCPFieldOldData "CP53", Empty, 1
   ' 聘任期間迄
   SetCPFieldOldData "CP54", Empty, 1
   '收據編號
   SetCPFieldOldData "CP60", Empty, 0
   '911018 nick 新增
   '進度備註
   SetCPFieldOldData "CP64", Empty, 0
   
   '911108 nick 因為會有些值沒有先定義，所以會沒有更新
   SetCPFieldOldData "CP11", Empty, 0
   SetCPFieldOldData "CP18", 0, 1
   SetCPFieldOldData "CP20", Empty, 0
   '911114 nick
   SetCPFieldOldData "CP32", Empty, 0
   SetCPFieldOldData "CP08", Empty, 0
   SetCPFieldOldData "CP48", Empty, 1
End Sub

Private Sub UpdateGrdList(ByVal strLC01 As String, ByVal strLC02 As String, ByVal strLC03 As String, ByVal strLC04 As String)
   Dim nIndex As Integer
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   
   ' 本案期限
   InitialGrdList
   ' 取得下一程序檔案中的資料列表在 Grid List 中
   strSql = "SELECT * FROM NextProgress " & _
            "WHERE NP02 = '" & strLC01 & "' AND " & _
                  "NP03 = '" & strLC02 & "' AND " & _
                  "NP04 = '" & strLC03 & "' AND " & _
                  "NP05 = '" & strLC04 & "' AND " & _
                  "(NP06 IS NULL OR NP06 <> 'Y') "
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
            '911111 nick 案件性質要依國家判斷
            'grdList.TextMatrix(nIndex, 1) = GetCaseTypeName(m_LC01, rsTmp.Fields("NP07"))
            grdList.TextMatrix(nIndex, 1) = GetPrjState4(strLC01 & "-" & strLC02 & "-" & strLC03 & "-" & strLC04, rsTmp.Fields("NP07"))
            
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
            grdList.TextMatrix(nIndex, 7) = rsTmp.Fields("NP11")
         End If
         ' 序號
         If IsNull(rsTmp.Fields("NP22")) = False Then
            grdList.TextMatrix(nIndex, 10) = rsTmp.Fields("NP22")
         End If
         '911111 nick 智權人員
         If IsNull(rsTmp.Fields("NP10")) = False Then
            grdList.TextMatrix(nIndex, 11) = rsTmp.Fields("NP10")
         End If
NextRecord:
         rsTmp.MoveNext
      Loop
      'Added by Lydia 2023/10/16
      If grdList.Rows >= 2 Then
         grdList.FixedRows = 1
      End If
      'end 2023/10/16
   End If
   rsTmp.Close
   
   Set rsTmp = Nothing
End Sub

' 設定開始輸入的欄位
Private Sub SetInputEntry()
   textCP05.SetFocus
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
         Me.SSTab1.Tab = 0
         '911111 nick
         textCP05.SetFocus
         textCP05_GotFocus
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
         Me.SSTab1.Tab = 0
         '911111 nick
         textCP06.SetFocus
         textCP06_GotFocus
      'add by sonia 2019/8/7
      Else
         'ACS若本所期限非工作天則直接調整至最近的工作天
         'Modified by Lydia 2020/07/07 本所期限檢查：所有系統類別的本所期限都要控制是工作日
         'If m_LC01 = "ACS" Then textCP06 = TransDate(PUB_GetWorkDay1(textCP06, True), 1)
         textCP06 = TransDate(PUB_GetWorkDay1(textCP06, True), 1)
      'end 2019/8/7
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
         Me.SSTab1.Tab = 0
         '911111 nick
         textCP07.SetFocus
         textCP07_GotFocus
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
      ' 取得國內的案件性質名稱
      textCP10_2 = GetCaseTypeName(m_LC01, textCP10, 0)
      If IsEmptyText(textCP10_2) = True Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "案件性質代號不存在"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         Me.SSTab1.Tab = 0
         '911111 nick
         textCP10.SetFocus
         textCP10_GotFocus
      End If
   End If
End Sub

'Add By Sindy 2010/11/25
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
         Me.SSTab1.Tab = 0
         '911111 nick
         textCP13.SetFocus
         textCP13_GotFocus
      'Added by Lydia 2019/02/14 創新業務部人員收文控管
      Else
         m_SalesST15 = GetST15(textCP13)
         'Added by Lydia 2020/04/08 檢查案件或智權人員是否為法務部
         If PUB_ChkSalesL(m_LC01, textCP13.Text) = False Then
             SSTab1.Tab = 0
             textCP13.SetFocus
             Call textCP13_GotFocus
             Cancel = True
             Exit Sub
         End If
         'end 2020/04/08
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
'911114 nick 邱小姐說刪除
' 費用
'Private Sub textCP16_Validate(Cancel As Boolean)
'   Dim strTit As String
'   Dim strMsg As String
'   Dim nResponse
'   Cancel = False
'   If IsEmptyText(textCP16) = False Then
'      If IsNumeric(textCP16) = False Then
'         Cancel = True
'         strTit = "資料檢核"
'         strMsg = "費用為數值資料"
'         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
'         '911111 nick
'         textCP16.SetFocus
'
'         textCP16_GotFocus
'      End If
'   End If
'End Sub
'911114 nick 邱小姐說刪除
' 規費
'Private Sub textCP17_Validate(Cancel As Boolean)
'   Dim strTit As String
'   Dim strMsg As String
'   Dim nResponse
'   Cancel = False
'   If IsEmptyText(textCP17) = False Then
'      If IsNumeric(textCP17) = False Then
'         Cancel = True
'         strTit = "資料檢核"
'         strMsg = "規費為數值資料"
'         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
'         '911111 nick
'         textCP17.SetFocus
'
'         textCP17_GotFocus
'      End If
'   End If
'End Sub
'911114 nick 邱小姐說刪除
' 點數
'Private Sub textCP18_Validate(Cancel As Boolean)
'   Dim strTit As String
'   Dim strMsg As String
'   Dim nResponse
'   Cancel = False
'   If IsEmptyText(textCP18) = False Then
'      If IsNumeric(textCP18) = False Then
'         Cancel = True
'         strTit = "資料檢核"
'         strMsg = "點數為數值資料"
'         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
'         '911111 nick
'         textCP18.SetFocus
'
'         textCP18_GotFocus
'      End If
'   End If
'End Sub
'911114 nick 邱小姐說刪除
' 後金
'Private Sub textCP19_Validate(Cancel As Boolean)
'   Dim strTit As String
'   Dim strMsg As String
'   Dim nResponse
'   Cancel = False
'   If IsEmptyText(textCP19) = False Then
'      If IsNumeric(textCP19) = False Then
'         Cancel = True
'         strTit = "資料檢核"
'         strMsg = "後金為數值資料"
'         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
'         '911111 nick
'         textCP19.SetFocus
'
'         textCP19_GotFocus
'      End If
'   End If
'End Sub

'Add By Sindy 2010/11/25
Private Sub textCP14_KeyPress(KeyAscii As Integer)
KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textCP21_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

' 是否取締案
Private Sub textCP21_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textCP21) = False Then
      Select Case textCP21
         Case " ", "Y":
         Case Else
            Cancel = True
            strTit = "資料檢核"
            strMsg = "只可輸入空白或Y"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            Me.SSTab1.Tab = 0
            '911111 nick
            textCP21.SetFocus
            textCP21_GotFocus
      End Select
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
            Me.SSTab1.Tab = 0
            '911111 nick
            textCP26.SetFocus
            textCP26_GotFocus
      End Select
   End If
End Sub

'Added by Lydia 2020/04/08
Private Sub textCP29_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

' 協辦人員
Private Sub textCP29_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   Cancel = False
   textCP29_2 = Empty
   If Not IsEmptyText(textCP29) Then
      textCP29_2 = GetStaffName(textCP29, False)
      If IsEmptyText(textCP29_2) Then
         Cancel = True
         strTit = "檢核資料"
         'Modified by Lydia 2015/10/05
        ' strMsg = "法務人員代碼<" & textCP29 & ">不存在或未在職"
         strMsg = "協辦人員代碼<" & textCP29 & ">不存在或未在職"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         Me.SSTab1.Tab = 0
         '911111 nick
         textCP29.SetFocus
         textCP29_GotFocus
      End If
   End If
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
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      Me.SSTab1.Tab = 0
      textCP64.SetFocus
      textCP64_GotFocus
   End If
   'edit by nickc 2007/06/06 切換輸入法改用API
   'If Cancel = False Then: textCP64.IMEMode = 2
   If Cancel = False Then CloseIme
End Sub

' 對造名稱(中)
Private Sub textCP40_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If CheckLengthIsOK(textCP40, 600) = False Then
      Cancel = True
      strTit = "檢核資料"
      strMsg = "對造中文名稱內容太長"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      Me.SSTab1.Tab = 1
      '911111 nick
      textCP40.SetFocus
      textCP40_GotFocus
   End If
   'edit by nickc 2007/06/06 切換輸入法改用API
   'If Cancel = False Then: textCP40.IMEMode = 2
   If Cancel = False Then CloseIme
End Sub

' 對造英文名稱
Private Sub textCP41_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If CheckLengthIsOK(textCP41, 600) = False Then
      Cancel = True
      strTit = "檢核資料"
      strMsg = "對造英文名稱內容太長"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      Me.SSTab1.Tab = 1
      '911111 nick
      textCP41.SetFocus
      textCP41_GotFocus
   End If
End Sub

' 對造日文名稱
Private Sub textCP42_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If CheckLengthIsOK(textCP42, 600) = False Then
      Cancel = True
      strTit = "檢核資料"
      strMsg = "對造日文名稱內容太長"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      Me.SSTab1.Tab = 1
      '911111 nick
      textCP42.SetFocus
      textCP42_GotFocus
   End If
   'edit by nickc 2007/06/06 切換輸入法改用API
   'If Cancel = False Then: textCP42.IMEMode = 2
   If Cancel = False Then CloseIme
End Sub

Private Sub textCP43_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
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
         Me.SSTab1.Tab = 0
         '911111 nick
         textCP43.SetFocus
         textCP43_GotFocus
         GoTo EXITSUB
      End If
      
      strSql = "SELECT * FROM CaseProgress " & _
               "WHERE CP01 = '" & m_LC01 & "' AND " & _
                     "CP02 = '" & m_LC02 & "' AND " & _
                     "CP03 = '" & m_LC03 & "' AND " & _
                     "CP04 = '" & m_LC04 & "' AND " & _
                     "CP09 = '" & textCP43 & "' "
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount <= 0 Then
         rsTmp.Close
         Cancel = True
         strTit = "資料檢核"
         strMsg = "相關總收文號資料不存在"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         Me.SSTab1.Tab = 0
         '911111 nick
         textCP43.SetFocus
         textCP43_GotFocus
         GoTo EXITSUB
      End If
      rsTmp.Close
   End If
EXITSUB:
   Set rsTmp = Nothing
End Sub

' 當事人稱謂
Private Sub textCP49_Validate(Cancel As Boolean)
   Cancel = False
   If CheckLengthIsOK(textCP49, 300) = False Then
      Cancel = True
      Me.SSTab1.Tab = 0
      '911111 nick
      textCP49.SetFocus
      textCP49_GotFocus
   End If
   'edit by nickc 2007/06/06 切換輸入法改用API
   'If Cancel = False Then: textCP49.IMEMode = 2
   If Cancel = False Then CloseIme
End Sub

' 聘任期間起
Private Sub textCP53_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textCP53) = False Then
      If CheckIsTaiwanDate(textCP53, False) = False Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "聘任期間起的日期格式不正確"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         Me.SSTab1.Tab = 0
         '911111 nick
         textCP53.SetFocus
         textCP53_GotFocus
      End If
   End If
End Sub

' 聘任期間迄
Private Sub textCP54_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textCP54) = False Then
      If CheckIsTaiwanDate(textCP54, False) = False Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "聘任期間迄的日期格式不正確"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         Me.SSTab1.Tab = 0
         '911111 nick
         textCP54.SetFocus
         textCP54_GotFocus
      End If
   End If
End Sub

' 案件中文名稱
Private Sub textLC05_Validate(Cancel As Boolean)
   Cancel = False
   If CheckLengthIsOK(textLC05, textLC05.MaxLength) = False Then
      Cancel = True
      Me.SSTab1.Tab = 0
      '911111 nick
      textLC05.SetFocus
      textLC05_GotFocus
   End If
   'edit by nickc 2007/06/06 切換輸入法改用API
   'If Cancel = False Then: textLC05.IMEMode = 2
   If Cancel = False Then CloseIme
End Sub

' 案件英文名稱
Private Sub textLC06_Validate(Cancel As Boolean)
   Cancel = False
   If CheckLengthIsOK(textLC06, textLC06.MaxLength) = False Then
      Cancel = True
      Me.SSTab1.Tab = 0
      '911111 nick
      textLC06.SetFocus
      textLC06_GotFocus
   End If
End Sub

' 案件日文名稱
Private Sub textLC07_Validate(Cancel As Boolean)
   Cancel = False
   If CheckLengthIsOK(textLC07, textLC07.MaxLength) = False Then
      Cancel = True
      Me.SSTab1.Tab = 0
      '911111 nick
      textLC07.SetFocus
      textLC07_GotFocus
   End If
   'edit by nickc 2007/06/06 切換輸入法改用API
   'If Cancel = False Then: textLC05.IMEMode = 2
   If Cancel = False Then CloseIme
End Sub

Private Sub textLC08_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

' 是否取消閉券
Private Sub textLC08_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   Cancel = False
   If IsEmptyText(textLC08) = False Then
      Select Case textLC08
         Case "Y", " ":
         Case Else
            Cancel = True
            strTit = "檢核資料"
            strMsg = "請輸入空白或Y"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            Me.SSTab1.Tab = 0
            '911111 nick
            textLC08.SetFocus
            textLC08_GotFocus
      End Select
   End If
End Sub

Private Sub textLC11_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub
'Add By Sindy 2011/1/19
Private Sub textLC43_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub
Private Sub textLC44_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub
Private Sub textLC45_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub
Private Sub textLC46_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub
'2011/1/19 End

' 當事人1
Private Sub textLC11_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   Cancel = False
   textLC11_2 = Empty
   If IsEmptyText(textLC11) = False Then
      'edit by nick 2004/07/21 檢查該申請人或代理人狀態，若為不再使用則停在原地
      Dim oState As Boolean
      'textLC11_2 = GetCustomerName(textLC11, 0)
      textLC11_2 = GetCustomerNameAndState(textLC11, 0, oState)
      If oState = False Then
         Cancel = True
         Me.SSTab1.Tab = 1
         Exit Sub
      End If
      If textLC11_2 = Empty Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "當事人1代碼<" & textLC11 & ">不正確"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         Me.SSTab1.Tab = 1
         '911111 nick
         textLC11.SetFocus
         textLC11_GotFocus
         Exit Sub
      End If
      textLC11 = Left(textLC11 & "000000000", 9)
      If textLC11 = textLC43 Then
         Cancel = True
         MsgBox "當事人不可重覆!", vbOKOnly + vbCritical, "警告!!"
         Me.SSTab1.Tab = 1
         textLC11.SetFocus
         textLC11_GotFocus
         Exit Sub
      End If
      If textLC11 = textLC44 Then
         Cancel = True
         MsgBox "當事人不可重覆!", vbOKOnly + vbCritical, "警告!!"
         Me.SSTab1.Tab = 1
         textLC11.SetFocus
         textLC11_GotFocus
         Exit Sub
      End If
      If textLC11 = textLC45 Then
         Cancel = True
         MsgBox "當事人不可重覆!", vbOKOnly + vbCritical, "警告!!"
         Me.SSTab1.Tab = 1
         textLC11.SetFocus
         textLC11_GotFocus
         Exit Sub
      End If
      If textLC11 = textLC46 Then
         Cancel = True
         MsgBox "當事人不可重覆!", vbOKOnly + vbCritical, "警告!!"
         Me.SSTab1.Tab = 1
         textLC11.SetFocus
         textLC11_GotFocus
         Exit Sub
      End If
   End If
End Sub

'Add By Sindy 2011/1/19
' 當事人2
Private Sub textLC43_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   Cancel = False
   textLC43_2 = Empty
   If IsEmptyText(textLC43) = False Then
      '檢查該申請人或代理人狀態，若為不再使用則停在原地
      Dim oState As Boolean
      textLC43_2 = GetCustomerNameAndState(textLC43, 0, oState)
      If oState = False Then
         Cancel = True
         Me.SSTab1.Tab = 1
         Exit Sub
      End If
      If textLC43_2 = Empty Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "當事人2代碼<" & textLC43 & ">不正確"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         Me.SSTab1.Tab = 1
         textLC43.SetFocus
         textLC43_GotFocus
         Exit Sub
      End If
      textLC43 = Left(textLC43 & "000000000", 9)
      If textLC43 = textLC11 Then
         Cancel = True
         MsgBox "當事人不可重覆!", vbOKOnly + vbCritical, "警告!!"
         Me.SSTab1.Tab = 1
         textLC43.SetFocus
         textLC43_GotFocus
         Exit Sub
      End If
      If textLC43 = textLC44 Then
         Cancel = True
         MsgBox "當事人不可重覆!", vbOKOnly + vbCritical, "警告!!"
         Me.SSTab1.Tab = 1
         textLC43.SetFocus
         textLC43_GotFocus
         Exit Sub
      End If
      If textLC43 = textLC45 Then
         Cancel = True
         MsgBox "當事人不可重覆!", vbOKOnly + vbCritical, "警告!!"
         Me.SSTab1.Tab = 1
         textLC43.SetFocus
         textLC43_GotFocus
         Exit Sub
      End If
      If textLC43 = textLC46 Then
         Cancel = True
         MsgBox "當事人不可重覆!", vbOKOnly + vbCritical, "警告!!"
         Me.SSTab1.Tab = 1
         textLC43.SetFocus
         textLC43_GotFocus
         Exit Sub
      End If
   End If
End Sub

'Add By Sindy 2011/1/19
' 當事人3
Private Sub textLC44_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   Cancel = False
   textLC44_2 = Empty
   If IsEmptyText(textLC44) = False Then
      '檢查該申請人或代理人狀態，若為不再使用則停在原地
      Dim oState As Boolean
      textLC44_2 = GetCustomerNameAndState(textLC44, 0, oState)
      If oState = False Then
         Cancel = True
         Me.SSTab1.Tab = 1
         Exit Sub
      End If
      If textLC44_2 = Empty Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "當事人3代碼<" & textLC44 & ">不正確"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         Me.SSTab1.Tab = 1
         textLC44.SetFocus
         textLC44_GotFocus
         Exit Sub
      End If
      textLC44 = Left(textLC44 & "000000000", 9)
      If textLC44 = textLC11 Then
         Cancel = True
         MsgBox "當事人不可重覆!", vbOKOnly + vbCritical, "警告!!"
         Me.SSTab1.Tab = 1
         textLC44.SetFocus
         textLC44_GotFocus
         Exit Sub
      End If
      If textLC44 = textLC43 Then
         Cancel = True
         MsgBox "當事人不可重覆!", vbOKOnly + vbCritical, "警告!!"
         Me.SSTab1.Tab = 1
         textLC44.SetFocus
         textLC44_GotFocus
         Exit Sub
      End If
      If textLC44 = textLC45 Then
         Cancel = True
         MsgBox "當事人不可重覆!", vbOKOnly + vbCritical, "警告!!"
         Me.SSTab1.Tab = 1
         textLC44.SetFocus
         textLC44_GotFocus
         Exit Sub
      End If
      If textLC44 = textLC46 Then
         Cancel = True
         MsgBox "當事人不可重覆!", vbOKOnly + vbCritical, "警告!!"
         Me.SSTab1.Tab = 1
         textLC44.SetFocus
         textLC44_GotFocus
         Exit Sub
      End If
   End If
End Sub

'Add By Sindy 2011/1/19
' 當事人4
Private Sub textLC45_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   Cancel = False
   textLC45_2 = Empty
   If IsEmptyText(textLC45) = False Then
      '檢查該申請人或代理人狀態，若為不再使用則停在原地
      Dim oState As Boolean
      textLC45_2 = GetCustomerNameAndState(textLC45, 0, oState)
      If oState = False Then
         Cancel = True
         Me.SSTab1.Tab = 1
         Exit Sub
      End If
      If textLC45_2 = Empty Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "當事人4代碼<" & textLC45 & ">不正確"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         Me.SSTab1.Tab = 1
         textLC45.SetFocus
         textLC45_GotFocus
         Exit Sub
      End If
      textLC45 = Left(textLC45 & "000000000", 9)
      If textLC45 = textLC11 Then
         Cancel = True
         MsgBox "當事人不可重覆!", vbOKOnly + vbCritical, "警告!!"
         Me.SSTab1.Tab = 1
         textLC45.SetFocus
         textLC45_GotFocus
         Exit Sub
      End If
      If textLC45 = textLC43 Then
         Cancel = True
         MsgBox "當事人不可重覆!", vbOKOnly + vbCritical, "警告!!"
         Me.SSTab1.Tab = 1
         textLC45.SetFocus
         textLC45_GotFocus
         Exit Sub
      End If
      If textLC45 = textLC44 Then
         Cancel = True
         MsgBox "當事人不可重覆!", vbOKOnly + vbCritical, "警告!!"
         Me.SSTab1.Tab = 1
         textLC45.SetFocus
         textLC45_GotFocus
         Exit Sub
      End If
      If textLC45 = textLC46 Then
         Cancel = True
         MsgBox "當事人不可重覆!", vbOKOnly + vbCritical, "警告!!"
         Me.SSTab1.Tab = 1
         textLC45.SetFocus
         textLC45_GotFocus
         Exit Sub
      End If
   End If
End Sub

'Add By Sindy 2011/1/19
' 當事人5
Private Sub textLC46_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   Cancel = False
   textLC46_2 = Empty
   If IsEmptyText(textLC46) = False Then
      '檢查該申請人或代理人狀態，若為不再使用則停在原地
      Dim oState As Boolean
      textLC46_2 = GetCustomerNameAndState(textLC46, 0, oState)
      If oState = False Then
         Cancel = True
         Me.SSTab1.Tab = 1
         Exit Sub
      End If
      If textLC46_2 = Empty Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "當事人5代碼<" & textLC46 & ">不正確"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         Me.SSTab1.Tab = 1
         textLC46.SetFocus
         textLC46_GotFocus
         Exit Sub
      End If
      textLC46 = Left(textLC46 & "000000000", 9)
      If textLC46 = textLC11 Then
         Cancel = True
         MsgBox "當事人不可重覆!", vbOKOnly + vbCritical, "警告!!"
         Me.SSTab1.Tab = 1
         textLC46.SetFocus
         textLC46_GotFocus
         Exit Sub
      End If
      If textLC46 = textLC43 Then
         Cancel = True
         MsgBox "當事人不可重覆!", vbOKOnly + vbCritical, "警告!!"
         Me.SSTab1.Tab = 1
         textLC46.SetFocus
         textLC46_GotFocus
         Exit Sub
      End If
      If textLC46 = textLC44 Then
         Cancel = True
         MsgBox "當事人不可重覆!", vbOKOnly + vbCritical, "警告!!"
         Me.SSTab1.Tab = 1
         textLC46.SetFocus
         textLC46_GotFocus
         Exit Sub
      End If
      If textLC46 = textLC45 Then
         Cancel = True
         MsgBox "當事人不可重覆!", vbOKOnly + vbCritical, "警告!!"
         Me.SSTab1.Tab = 1
         textLC46.SetFocus
         textLC46_GotFocus
         Exit Sub
      End If
   End If
End Sub

Private Sub textLC13_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

' 是否為智慧財產權案
Private Sub textLC13_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   Cancel = False
   If Not IsEmptyText(textLC13) Then
      Select Case textLC13
         Case "Y", " ":
         Case Else:
            Cancel = True
            strTit = "檢核資料"
            strMsg = "是否為智慧財產權案只可輸入Y或空白"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            Me.SSTab1.Tab = 0
            '911111 nick
            textLC13.SetFocus
            textLC13_GotFocus
      End Select
   End If
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
         Me.SSTab1.Tab = 0
         '911111 nick
         textCP14.SetFocus
         textCP14_GotFocus
      End If
   End If
End Sub

Private Function CheckDataValid() As Boolean
   Dim strTit As String
   Dim strMsg As String
   Dim strTemp As String
   Dim nResponse
   
   CheckDataValid = False
   
   ' 案件名稱不可同時為空白
   If IsEmptyText(textLC05) And IsEmptyText(textLC06) And IsEmptyText(textLC07) Then
      strTit = "檢核資料"
      strMsg = "案件名稱(中英日)不可同時為空白"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      Me.SSTab1.Tab = 0
      textLC05.SetFocus
      GoTo EXITSUB
   End If
   ' 案件性質不可為空白
   If IsEmptyText(textCP10) = True Then
      strTit = "檢核資料"
      strMsg = "案件性質不可為空白"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      Me.SSTab1.Tab = 0
      textCP10.SetFocus
      GoTo EXITSUB
   End If
   ' 收文日不可為空白
   If IsEmptyText(textCP05) = True Then
      strTit = "檢核資料"
      strMsg = "收文日不可空白"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      Me.SSTab1.Tab = 0
      textCP05.SetFocus
      GoTo EXITSUB
   End If
   ' 當事人1不可為空白
   If IsEmptyText(textLC11) = True Then
      strTit = "檢核資料"
      strMsg = "當事人1不可空白"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      Me.SSTab1.Tab = 1
      textLC11.SetFocus
      GoTo EXITSUB
   End If
   'Add By Sindy 2011/1/19 檢查當事人的輸入順序
   If (Trim(textLC43) <> "" And Trim(textLC11) = "") Or _
      (Trim(textLC44) <> "" And Trim(textLC43) = "") Or _
      (Trim(textLC45) <> "" And Trim(textLC44) = "") Or _
      (Trim(textLC46) <> "" And Trim(textLC45) = "") Then
      strTit = "檢核資料"
      strMsg = "請依序輸入當事人!"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      Me.SSTab1.Tab = 1
      If Trim(textLC43) <> "" And Trim(textLC11) = "" Then textLC43.SetFocus: Call textLC43_GotFocus
      If Trim(textLC44) <> "" And Trim(textLC43) = "" Then textLC44.SetFocus: Call textLC44_GotFocus
      If Trim(textLC45) <> "" And Trim(textLC44) = "" Then textLC45.SetFocus: Call textLC45_GotFocus
      If Trim(textLC46) <> "" And Trim(textLC45) = "" Then textLC46.SetFocus: Call textLC46_GotFocus
      GoTo EXITSUB
   End If
   '2011/1/19 End
   
   ' 聘任期間
   If textCP10 = "0" Then
      If IsEmptyText(textCP53) Or IsEmptyText(textCP54) Then
         strTit = "檢核資料"
         strMsg = "案件性質為聘任時聘任期間不可空白"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         Me.SSTab1.Tab = 0
         If IsEmptyText(textCP53) Then
            textCP53.SetFocus
         Else
            textCP54.SetFocus
         End If
         GoTo EXITSUB
      End If
      If Val(textCP53) > Val(textCP54) Then
         strTit = "檢核資料"
         strMsg = "聘任期間範圍不正確"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         Me.SSTab1.Tab = 0
         textCP53.SetFocus
         GoTo EXITSUB
      End If
   End If
   ' 智權人員 ADD BY SONIA 91.11.3
   If IsEmptyText(textCP13) = True Then
      strTit = "檢核資料"
      strMsg = "智權人員不可空白"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      Me.SSTab1.Tab = 0
      textCP13.SetFocus
      GoTo EXITSUB
   End If
   '91.11.3 END
   '2006/1/4 ADD BY SONIA
   If IsEmptyText(textCP14) = True And IsEmptyText(textCP29) = True Then
      strTit = "檢核資料"
      'Modified by Lydia 2015/10/05
      'strMsg = "承辦律師和承辦法務不可同時空白"
      'modify by sonia 2019/8/7
      'strMsg = "承辦人和協辦人員不可同時空白"
      If m_LC01 = "ACS" Then
         strMsg = "承辦人不可空白！"
      Else
         strMsg = "承辦人和協辦人員不可同時空白"
      End If
      'end 2019/8/7
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      Me.SSTab1.Tab = 0
      textCP14.SetFocus
      GoTo EXITSUB
   End If
   '2006/1/4 END
   CheckDataValid = True
EXITSUB:
End Function

Private Sub textLC05_GotFocus()
   InverseTextBox textLC05
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textLC05.IMEMode = 1
   OpenIme
End Sub

Private Sub textLC06_GotFocus()
   InverseTextBox textLC06
End Sub

Private Sub textLC07_GotFocus()
   InverseTextBox textLC07
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textLC07.IMEMode = 1
   OpenIme
End Sub

Private Sub textLC08_GotFocus()
   InverseTextBox textLC08
End Sub

Private Sub textLC11_GotFocus()
   InverseTextBox textLC11
End Sub
'Add By Sindy 2011/1/19
Private Sub textLC43_GotFocus()
   InverseTextBox textLC43
End Sub
Private Sub textLC44_GotFocus()
   InverseTextBox textLC44
End Sub
Private Sub textLC45_GotFocus()
   InverseTextBox textLC45
End Sub
Private Sub textLC46_GotFocus()
   InverseTextBox textLC46
End Sub
'2011/1/19 End

Private Sub textLC13_GotFocus()
   InverseTextBox textLC13
End Sub

Private Sub textLC16_GotFocus()
   InverseTextBox textLC16
End Sub

Private Sub textCP40_GotFocus()
   InverseTextBox textCP40
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textCP40.IMEMode = 1
   OpenIme
End Sub

Private Sub textCP41_GotFocus()
   InverseTextBox textCP41
End Sub

Private Sub textCP42_GotFocus()
   InverseTextBox textCP42
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textCP42.IMEMode = 1
   OpenIme
End Sub

Private Sub textCP53_GotFocus()
   InverseTextBox textCP53
End Sub

Private Sub textCP54_GotFocus()
   InverseTextBox textCP54
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
End Sub
'911114 nick 邱小姐說刪除
'Private Sub textCP16_GotFocus()
'   InverseTextBox textCP16
'End Sub
'911114 nick 邱小姐說刪除
'Private Sub textCP17_GotFocus()
'   InverseTextBox textCP17
'End Sub
'911114 nick 邱小姐說刪除
'Private Sub textCP18_GotFocus()
'   InverseTextBox textCP18
'End Sub
'911114 nick 邱小姐說刪除
'Private Sub textCP19_GotFocus()
'   InverseTextBox textCP19
'End Sub

Private Sub textCP21_GotFocus()
   InverseTextBox textCP21
End Sub

Private Sub textCP26_GotFocus()
   InverseTextBox textCP26
End Sub

Private Sub textCP29_GotFocus()
   InverseTextBox textCP29
End Sub

Private Sub textCP43_GotFocus()
   InverseTextBox textCP43
End Sub

Private Sub textCP49_GotFocus()
   InverseTextBox textCP49
   '911114 nick
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textCP49.IMEMode = 1
   OpenIme
End Sub

Private Sub textCP64_GotFocus()
   InverseTextBox textCP64
End Sub

' 確認使用者所輸入的都完全正確
Private Function ValidateInput() As Boolean
   Dim Cancel As Boolean

   ValidateInput = False
   
   If textCP05.Enabled = True Then
      Cancel = False
      textCP05_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If textCP06.Enabled = True Then
      Cancel = False
      textCP06_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If textCP07.Enabled = True Then
      Cancel = False
      textCP07_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If textCP10.Enabled = True Then
      Cancel = False
      textCP10_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If textCP13.Enabled = True Then
      Cancel = False
      textCP13_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If textCP14.Enabled = True Then
      Cancel = False
      textCP14_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   '911114 nick 邱小姐說刪除
   'If textCP16.Enabled = True Then
   '   Cancel = False
   '   textCP16_Validate Cancel
   '   If Cancel = True Then
   '      Exit Function
   '   End If
   'End If
   '911114 nick 邱小姐說刪除
   'If textCP17.Enabled = True Then
   '   Cancel = False
   '   textCP17_Validate Cancel
   '   If Cancel = True Then
    '     Exit Function
   '   End If
   'End If
   '911114 nick 邱小姐說刪除
   'If textCP18.Enabled = True Then
   '   Cancel = False
   '   textCP18_Validate Cancel
   '   If Cancel = True Then
   '      Exit Function
   '   End If
   'End If
   '911114 nick 邱小姐說刪除
   'If textCP19.Enabled = True Then
   '   Cancel = False
   '   textCP19_Validate Cancel
   '   If Cancel = True Then
   '      Exit Function
   '   End If
   'End If
   
   If textCP21.Enabled = True Then
      Cancel = False
      textCP21_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If textCP26.Enabled = True Then
      Cancel = False
      textCP26_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If textCP29.Enabled = True Then
      Cancel = False
      textCP29_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If textCP40.Enabled = True Then
      Cancel = False
      textCP40_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If textCP41.Enabled = True Then
      Cancel = False
      textCP41_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If textCP42.Enabled = True Then
      Cancel = False
      textCP42_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If textCP43.Enabled = True Then
      Cancel = False
      textCP43_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If textCP49.Enabled = True Then
      Cancel = False
      textCP49_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If textLC05.Enabled = True Then
      Cancel = False
      textLC05_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If textLC06.Enabled = True Then
      Cancel = False
      textLC06_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If textLC07.Enabled = True Then
      Cancel = False
      textLC07_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If textLC08.Enabled = True Then
      Cancel = False
      textLC08_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   '2005/9/12 ADD BY SONIA
   If textCP64.Enabled = True Then
      Cancel = False
      textCP64_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   '2005/9/12 END
   
   If textLC11.Enabled = True Then
      Cancel = False
      textLC11_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   'Add By Sindy 2011/1/19
   If textLC43.Enabled = True Then
      Cancel = False
      textLC43_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If textLC44.Enabled = True Then
      Cancel = False
      textLC44_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If textLC45.Enabled = True Then
      Cancel = False
      textLC45_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If textLC46.Enabled = True Then
      Cancel = False
      textLC46_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   '2011/1/19 End
   
   If textLC13.Enabled = True Then
      Cancel = False
      textLC13_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   '911112 nick
   '***** start
   If textCP53.Enabled = True Then
      Cancel = False
      textCP53_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If textCP54.Enabled = True Then
      Cancel = False
      textCP54_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   '***** end
   
   'Added by Morgan 2022/8/18 法律所內部收文詢問是否發文--涂軼
   If textCP10 <> "0" Then
      m_bolSetCP27 = False
      If MsgBox("是否自動上發文日為收文日？", vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
         m_bolSetCP27 = True
      End If
   End If
   'end 2022/8/18
   
   ValidateInput = True
End Function

'Add By Cheng 2003/03/28
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

Private Sub textLC16_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If CheckLengthIsOK(textLC16, 50) = False Then
      Cancel = True
      strTit = "檢核資料"
      strMsg = "分所號內容太長"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      Me.SSTab1.Tab = 0
      '911111 nick
      textLC16.SetFocus
      textLC16_GotFocus
   End If
End Sub
