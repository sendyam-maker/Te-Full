VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm02010409_8 
   BorderStyle     =   1  '單線固定
   Caption         =   "變更事項"
   ClientHeight    =   5670
   ClientLeft      =   180
   ClientTop       =   990
   ClientWidth     =   9330
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5670
   ScaleWidth      =   9330
   Begin VB.TextBox textCE01 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   5280
      Locked          =   -1  'True
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   660
      Width           =   2532
   End
   Begin VB.TextBox textTMKey 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   660
      Width           =   2532
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Left            =   7200
      TabIndex        =   18
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "回前畫面(&U)"
      Height          =   400
      Left            =   8028
      TabIndex        =   19
      Top             =   70
      Width           =   1200
   End
   Begin TabDlg.SSTab tabCtrl 
      Height          =   4632
      Left            =   120
      TabIndex        =   22
      Top             =   1020
      Width           =   9012
      _ExtentX        =   15901
      _ExtentY        =   8176
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   5
      TabHeight       =   420
      TabCaption(0)   =   "第一頁"
      TabPicture(0)   =   "frm02010409_8.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1(1)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label4"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label5"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label6"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label7"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label8"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label9"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label10"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label11"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label12"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label13"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label14"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Label15"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Label16"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Label17"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Label18"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Label19"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Label20"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Label21"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "Label22"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "textCE04"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "textCE10"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "textCE12"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "textCE13"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "textCE15"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "textCE55"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "textCE17"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "textCE63"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "textCE64"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "textCE09"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "textCE03"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "textCE02"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "textCE16"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "textCE11"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "textCE14"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "textCE56"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "textCE54"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).Control(39)=   "textCE53"
      Tab(0).Control(39).Enabled=   0   'False
      Tab(0).Control(40)=   "textCE52"
      Tab(0).Control(40).Enabled=   0   'False
      Tab(0).Control(41)=   "textCE51"
      Tab(0).Control(41).Enabled=   0   'False
      Tab(0).Control(42)=   "textCE22"
      Tab(0).Control(42).Enabled=   0   'False
      Tab(0).Control(43)=   "textCE65"
      Tab(0).Control(43).Enabled=   0   'False
      Tab(0).ControlCount=   44
      TabCaption(1)   =   "第二頁"
      TabPicture(1)   =   "frm02010409_8.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label23"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label24"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label25"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label26"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Label27"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Label28"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Label29"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Label30"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "Label31"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "Label32"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "Label33"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "Label34"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "Label35"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "Label36"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "Label37"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "Label38"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "Label39"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "Label40"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).Control(18)=   "Label41"
      Tab(1).Control(18).Enabled=   0   'False
      Tab(1).Control(19)=   "Label42"
      Tab(1).Control(19).Enabled=   0   'False
      Tab(1).Control(20)=   "Label43"
      Tab(1).Control(20).Enabled=   0   'False
      Tab(1).Control(21)=   "Label44"
      Tab(1).Control(21).Enabled=   0   'False
      Tab(1).Control(22)=   "textCE23"
      Tab(1).Control(22).Enabled=   0   'False
      Tab(1).Control(23)=   "textCE25"
      Tab(1).Control(23).Enabled=   0   'False
      Tab(1).Control(24)=   "textCE41"
      Tab(1).Control(24).Enabled=   0   'False
      Tab(1).Control(25)=   "textCE43"
      Tab(1).Control(25).Enabled=   0   'False
      Tab(1).Control(26)=   "textCE61"
      Tab(1).Control(26).Enabled=   0   'False
      Tab(1).Control(27)=   "textCE38"
      Tab(1).Control(27).Enabled=   0   'False
      Tab(1).Control(28)=   "textCE24"
      Tab(1).Control(28).Enabled=   0   'False
      Tab(1).Control(29)=   "textCE58"
      Tab(1).Control(29).Enabled=   0   'False
      Tab(1).Control(30)=   "textCE57"
      Tab(1).Control(30).Enabled=   0   'False
      Tab(1).Control(31)=   "textCE40"
      Tab(1).Control(31).Enabled=   0   'False
      Tab(1).Control(32)=   "textCE39"
      Tab(1).Control(32).Enabled=   0   'False
      Tab(1).Control(33)=   "textCE60"
      Tab(1).Control(33).Enabled=   0   'False
      Tab(1).Control(34)=   "textCE44"
      Tab(1).Control(34).Enabled=   0   'False
      Tab(1).Control(35)=   "textCE42"
      Tab(1).Control(35).Enabled=   0   'False
      Tab(1).Control(36)=   "textCE46"
      Tab(1).Control(36).Enabled=   0   'False
      Tab(1).Control(37)=   "textCE45"
      Tab(1).Control(37).Enabled=   0   'False
      Tab(1).Control(38)=   "textCE48"
      Tab(1).Control(38).Enabled=   0   'False
      Tab(1).Control(39)=   "textCE47"
      Tab(1).Control(39).Enabled=   0   'False
      Tab(1).Control(40)=   "textCE50"
      Tab(1).Control(40).Enabled=   0   'False
      Tab(1).Control(41)=   "textCE49"
      Tab(1).Control(41).Enabled=   0   'False
      Tab(1).Control(42)=   "textCE62"
      Tab(1).Control(42).Enabled=   0   'False
      Tab(1).Control(43)=   "textCE59"
      Tab(1).Control(43).Enabled=   0   'False
      Tab(1).ControlCount=   44
      Begin VB.TextBox textCE59 
         BorderStyle     =   0  '沒有框線
         Height          =   264
         Left            =   -71760
         Locked          =   -1  'True
         TabIndex        =   48
         TabStop         =   0   'False
         Top             =   1860
         Width           =   5592
      End
      Begin VB.TextBox textCE62 
         Height          =   264
         Left            =   -73680
         TabIndex        =   16
         Top             =   3960
         Width           =   372
      End
      Begin VB.TextBox textCE49 
         BorderStyle     =   0  '沒有框線
         Height          =   264
         Left            =   -71760
         Locked          =   -1  'True
         TabIndex        =   47
         TabStop         =   0   'False
         Top             =   3660
         Width           =   5592
      End
      Begin VB.TextBox textCE50 
         Height          =   264
         Left            =   -73680
         TabIndex        =   15
         Top             =   3660
         Width           =   372
      End
      Begin VB.TextBox textCE47 
         BorderStyle     =   0  '沒有框線
         Height          =   264
         Left            =   -71760
         Locked          =   -1  'True
         TabIndex        =   46
         TabStop         =   0   'False
         Top             =   3360
         Width           =   5592
      End
      Begin VB.TextBox textCE48 
         Height          =   264
         Left            =   -73680
         TabIndex        =   14
         Top             =   3360
         Width           =   372
      End
      Begin VB.TextBox textCE45 
         BorderStyle     =   0  '沒有框線
         Height          =   264
         Left            =   -71760
         Locked          =   -1  'True
         TabIndex        =   45
         TabStop         =   0   'False
         Top             =   3060
         Width           =   5592
      End
      Begin VB.TextBox textCE46 
         Height          =   264
         Left            =   -73680
         TabIndex        =   13
         Top             =   3060
         Width           =   372
      End
      Begin VB.TextBox textCE42 
         BorderStyle     =   0  '沒有框線
         Height          =   264
         Left            =   -71760
         Locked          =   -1  'True
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   2460
         Width           =   5592
      End
      Begin VB.TextBox textCE44 
         Height          =   264
         Left            =   -73680
         TabIndex        =   12
         Top             =   2160
         Width           =   372
      End
      Begin VB.TextBox textCE60 
         Height          =   264
         Left            =   -73680
         TabIndex        =   11
         Top             =   1860
         Width           =   372
      End
      Begin VB.TextBox textCE39 
         BorderStyle     =   0  '沒有框線
         Height          =   264
         Left            =   -71760
         Locked          =   -1  'True
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   1560
         Width           =   5592
      End
      Begin VB.TextBox textCE40 
         Height          =   264
         Left            =   -73680
         TabIndex        =   10
         Top             =   1560
         Width           =   372
      End
      Begin VB.TextBox textCE57 
         BorderStyle     =   0  '沒有框線
         Height          =   264
         Left            =   -71760
         Locked          =   -1  'True
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   1260
         Width           =   5592
      End
      Begin VB.TextBox textCE58 
         Height          =   264
         Left            =   -73680
         TabIndex        =   9
         Top             =   1260
         Width           =   372
      End
      Begin VB.TextBox textCE24 
         BorderStyle     =   0  '沒有框線
         Height          =   264
         Left            =   -71760
         Locked          =   -1  'True
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   660
         Width           =   5592
      End
      Begin VB.TextBox textCE38 
         Height          =   264
         Left            =   -73680
         TabIndex        =   8
         Top             =   360
         Width           =   372
      End
      Begin VB.TextBox textCE65 
         Height          =   264
         Left            =   1320
         TabIndex        =   7
         Top             =   3960
         Width           =   372
      End
      Begin VB.TextBox textCE22 
         Height          =   264
         Left            =   1320
         TabIndex        =   6
         Top             =   3660
         Width           =   372
      End
      Begin VB.TextBox textCE51 
         BorderStyle     =   0  '沒有框線
         Height          =   264
         Left            =   3360
         Locked          =   -1  'True
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   3360
         Width           =   5472
      End
      Begin VB.TextBox textCE52 
         Height          =   264
         Left            =   1320
         TabIndex        =   5
         Top             =   3360
         Width           =   372
      End
      Begin VB.TextBox textCE53 
         BorderStyle     =   0  '沒有框線
         Height          =   264
         Left            =   3360
         Locked          =   -1  'True
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   3060
         Width           =   5472
      End
      Begin VB.TextBox textCE54 
         Height          =   264
         Left            =   1320
         TabIndex        =   4
         Top             =   3060
         Width           =   372
      End
      Begin VB.TextBox textCE56 
         Height          =   264
         Left            =   1320
         TabIndex        =   3
         Top             =   2760
         Width           =   372
      End
      Begin VB.TextBox textCE14 
         BorderStyle     =   0  '沒有框線
         Height          =   264
         Left            =   3360
         Locked          =   -1  'True
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   2160
         Width           =   5472
      End
      Begin VB.TextBox textCE11 
         BorderStyle     =   0  '沒有框線
         Height          =   264
         Left            =   3360
         Locked          =   -1  'True
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   1260
         Width           =   5472
      End
      Begin VB.TextBox textCE16 
         Height          =   264
         Left            =   1320
         TabIndex        =   2
         Top             =   960
         Width           =   372
      End
      Begin VB.TextBox textCE02 
         BorderStyle     =   0  '沒有框線
         Height          =   264
         Left            =   3360
         Locked          =   -1  'True
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   660
         Width           =   5472
      End
      Begin VB.TextBox textCE03 
         Height          =   264
         Left            =   1320
         TabIndex        =   1
         Top             =   660
         Width           =   372
      End
      Begin VB.TextBox textCE09 
         Height          =   264
         Left            =   1320
         TabIndex        =   0
         Top             =   360
         Width           =   372
      End
      Begin MSForms.TextBox textCE61 
         Height          =   552
         Left            =   -72480
         TabIndex        =   17
         Top             =   3960
         Width           =   6312
         VariousPropertyBits=   -1467989989
         ScrollBars      =   2
         Size            =   "11134;974"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE43 
         Height          =   264
         Left            =   -71760
         TabIndex        =   44
         TabStop         =   0   'False
         Top             =   2760
         Width           =   5592
         VariousPropertyBits=   679493663
         Size            =   "9864;466"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE41 
         Height          =   264
         Left            =   -71760
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   2160
         Width           =   5592
         VariousPropertyBits=   679493663
         Size            =   "9864;466"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE25 
         Height          =   264
         Left            =   -71760
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   960
         Width           =   5592
         VariousPropertyBits=   679493663
         Size            =   "9864;466"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE23 
         Height          =   264
         Left            =   -71760
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   360
         Width           =   5592
         VariousPropertyBits=   679493663
         Size            =   "9864;466"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE64 
         Height          =   264
         Left            =   3360
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   4260
         Width           =   5472
         VariousPropertyBits=   679493663
         Size            =   "9652;466"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE63 
         Height          =   264
         Left            =   3360
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   3960
         Width           =   5472
         VariousPropertyBits=   679493663
         Size            =   "9652;466"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE17 
         Height          =   264
         Left            =   3360
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   3660
         Width           =   5472
         VariousPropertyBits=   679493663
         Size            =   "9652;466"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE55 
         Height          =   264
         Left            =   3360
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   2760
         Width           =   5472
         VariousPropertyBits=   679493663
         Size            =   "9652;466"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE15 
         Height          =   264
         Left            =   3360
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   2460
         Width           =   5472
         VariousPropertyBits=   679493663
         Size            =   "9652;466"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE13 
         Height          =   264
         Left            =   3360
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   1860
         Width           =   5472
         VariousPropertyBits=   679493663
         Size            =   "9652;466"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE12 
         Height          =   264
         Left            =   3360
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   1560
         Width           =   5472
         VariousPropertyBits=   679493663
         Size            =   "9652;466"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE10 
         Height          =   264
         Left            =   3360
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   960
         Width           =   5472
         VariousPropertyBits=   679493663
         Size            =   "9652;466"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE04 
         Height          =   264
         Left            =   3360
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   360
         Width           =   5472
         VariousPropertyBits=   679493663
         Size            =   "9652;466"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label44 
         Caption         =   "其它 :"
         Height          =   252
         Left            =   -73080
         TabIndex        =   92
         Top             =   3960
         Width           =   612
      End
      Begin VB.Label Label43 
         Caption         =   "准(1) / 駁(2) :"
         Height          =   252
         Left            =   -74880
         TabIndex        =   91
         Top             =   3960
         Width           =   1092
      End
      Begin VB.Label Label42 
         Caption         =   "商品群組 :"
         Height          =   252
         Left            =   -73080
         TabIndex        =   90
         Top             =   3660
         Width           =   1212
      End
      Begin VB.Label Label41 
         Caption         =   "准(1) / 駁(2) :"
         Height          =   252
         Left            =   -74880
         TabIndex        =   89
         Top             =   3660
         Width           =   1092
      End
      Begin VB.Label Label40 
         Caption         =   "商品類別 :"
         Height          =   252
         Left            =   -73080
         TabIndex        =   88
         Top             =   3360
         Width           =   1212
      End
      Begin VB.Label Label39 
         Caption         =   "准(1) / 駁(2) :"
         Height          =   252
         Left            =   -74880
         TabIndex        =   87
         Top             =   3360
         Width           =   1092
      End
      Begin VB.Label Label38 
         Caption         =   "縮減商品 :"
         Height          =   252
         Left            =   -73080
         TabIndex        =   86
         Top             =   3060
         Width           =   1212
      End
      Begin VB.Label Label37 
         Caption         =   "准(1) / 駁(2) :"
         Height          =   252
         Left            =   -74880
         TabIndex        =   85
         Top             =   3060
         Width           =   1092
      End
      Begin VB.Label Label36 
         Caption         =   "案件名稱(日) :"
         Height          =   252
         Left            =   -73080
         TabIndex        =   84
         Top             =   2760
         Width           =   1212
      End
      Begin VB.Label Label35 
         Caption         =   "案件名稱(英) :"
         Height          =   252
         Left            =   -73080
         TabIndex        =   83
         Top             =   2460
         Width           =   1212
      End
      Begin VB.Label Label34 
         Caption         =   "案件名稱(中) :"
         Height          =   252
         Left            =   -73080
         TabIndex        =   82
         Top             =   2160
         Width           =   1212
      End
      Begin VB.Label Label33 
         Caption         =   "准(1) / 駁(2) :"
         Height          =   252
         Left            =   -74880
         TabIndex        =   81
         Top             =   2160
         Width           =   1092
      End
      Begin VB.Label Label32 
         Caption         =   "圖樣 :"
         Height          =   252
         Left            =   -73080
         TabIndex        =   80
         Top             =   1860
         Width           =   612
      End
      Begin VB.Label Label31 
         Caption         =   "准(1) / 駁(2) :"
         Height          =   252
         Left            =   -74880
         TabIndex        =   79
         Top             =   1860
         Width           =   1092
      End
      Begin VB.Label Label30 
         Caption         =   "商標種類 :"
         Height          =   252
         Left            =   -73080
         TabIndex        =   78
         Top             =   1560
         Width           =   1212
      End
      Begin VB.Label Label29 
         Caption         =   "准(1) / 駁(2) :"
         Height          =   252
         Left            =   -74880
         TabIndex        =   77
         Top             =   1560
         Width           =   1092
      End
      Begin VB.Label Label28 
         Caption         =   "正商標號數 :"
         Height          =   252
         Left            =   -73080
         TabIndex        =   76
         Top             =   1260
         Width           =   1212
      End
      Begin VB.Label Label27 
         Caption         =   "准(1) / 駁(2) :"
         Height          =   252
         Left            =   -74880
         TabIndex        =   75
         Top             =   1260
         Width           =   1092
      End
      Begin VB.Label Label26 
         Caption         =   "申請地址(日) :"
         Height          =   252
         Left            =   -73080
         TabIndex        =   74
         Top             =   960
         Width           =   1212
      End
      Begin VB.Label Label25 
         Caption         =   "申請地址(英) :"
         Height          =   252
         Left            =   -73080
         TabIndex        =   73
         Top             =   660
         Width           =   1212
      End
      Begin VB.Label Label24 
         Caption         =   "申請地址(中) :"
         Height          =   252
         Left            =   -73080
         TabIndex        =   72
         Top             =   360
         Width           =   1212
      End
      Begin VB.Label Label23 
         Caption         =   "准(1) / 駁(2) :"
         Height          =   252
         Left            =   -74880
         TabIndex        =   71
         Top             =   360
         Width           =   1092
      End
      Begin VB.Label Label22 
         Caption         =   "代表人2中譯文 :"
         Height          =   252
         Left            =   1920
         TabIndex        =   70
         Top             =   4260
         Width           =   1332
      End
      Begin VB.Label Label21 
         Caption         =   "代表人1中譯文 :"
         Height          =   252
         Left            =   1920
         TabIndex        =   69
         Top             =   3960
         Width           =   1332
      End
      Begin VB.Label Label20 
         Caption         =   "准(1) / 駁(2) :"
         Height          =   252
         Left            =   120
         TabIndex        =   68
         Top             =   3960
         Width           =   1092
      End
      Begin VB.Label Label19 
         Caption         =   "申請人中譯文 :"
         Height          =   252
         Left            =   1920
         TabIndex        =   67
         Top             =   3660
         Width           =   1332
      End
      Begin VB.Label Label18 
         Caption         =   "准(1) / 駁(2) :"
         Height          =   252
         Left            =   120
         TabIndex        =   66
         Top             =   3660
         Width           =   1092
      End
      Begin VB.Label Label17 
         Caption         =   "申請人印鑑 :"
         Height          =   252
         Left            =   1920
         TabIndex        =   65
         Top             =   3360
         Width           =   1092
      End
      Begin VB.Label Label16 
         Caption         =   "准(1) / 駁(2) :"
         Height          =   252
         Left            =   120
         TabIndex        =   64
         Top             =   3360
         Width           =   1092
      End
      Begin VB.Label Label15 
         Caption         =   "代表人印鑑 :"
         Height          =   252
         Left            =   1920
         TabIndex        =   63
         Top             =   3060
         Width           =   1092
      End
      Begin VB.Label Label14 
         Caption         =   "准(1) / 駁(2) :"
         Height          =   252
         Left            =   120
         TabIndex        =   62
         Top             =   3060
         Width           =   1092
      End
      Begin VB.Label Label13 
         Caption         =   "代理人 :"
         Height          =   252
         Left            =   1920
         TabIndex        =   61
         Top             =   2760
         Width           =   732
      End
      Begin VB.Label Label12 
         Caption         =   "准(1) / 駁(2) :"
         Height          =   252
         Left            =   120
         TabIndex        =   60
         Top             =   2760
         Width           =   1092
      End
      Begin VB.Label Label11 
         Caption         =   "代表人2(日) :"
         Height          =   252
         Left            =   1920
         TabIndex        =   59
         Top             =   2460
         Width           =   1092
      End
      Begin VB.Label Label10 
         Caption         =   "代表人2(英) :"
         Height          =   252
         Left            =   1920
         TabIndex        =   58
         Top             =   2160
         Width           =   1092
      End
      Begin VB.Label Label9 
         Caption         =   "代表人2(中) :"
         Height          =   252
         Left            =   1920
         TabIndex        =   57
         Top             =   1860
         Width           =   1092
      End
      Begin VB.Label Label8 
         Caption         =   "代表人1(日) :"
         Height          =   252
         Left            =   1920
         TabIndex        =   56
         Top             =   1560
         Width           =   1092
      End
      Begin VB.Label Label7 
         Caption         =   "代表人1(英) :"
         Height          =   252
         Left            =   1920
         TabIndex        =   55
         Top             =   1260
         Width           =   1092
      End
      Begin VB.Label Label6 
         Caption         =   "代表人1(中) :"
         Height          =   252
         Left            =   1920
         TabIndex        =   54
         Top             =   960
         Width           =   1092
      End
      Begin VB.Label Label5 
         Caption         =   "准(1) / 駁(2) :"
         Height          =   252
         Left            =   120
         TabIndex        =   53
         Top             =   960
         Width           =   1092
      End
      Begin VB.Label Label4 
         Caption         =   "申請日 :"
         Height          =   252
         Left            =   1920
         TabIndex        =   52
         Top             =   660
         Width           =   732
      End
      Begin VB.Label Label3 
         Caption         =   "准(1) / 駁(2) :"
         Height          =   252
         Left            =   120
         TabIndex        =   51
         Top             =   660
         Width           =   1092
      End
      Begin VB.Label Label2 
         Caption         =   "申請人 :"
         Height          =   252
         Left            =   1920
         TabIndex        =   50
         Top             =   360
         Width           =   732
      End
      Begin VB.Label Label1 
         Caption         =   "准(1) / 駁(2) :"
         Height          =   252
         Index           =   1
         Left            =   120
         TabIndex        =   49
         Top             =   360
         Width           =   1092
      End
   End
   Begin VB.Label Label1 
      Caption         =   "收文號 :"
      Height          =   252
      Index           =   2
      Left            =   4440
      TabIndex        =   94
      Top             =   660
      Width           =   732
   End
   Begin VB.Label Label1 
      Caption         =   "本所案號 :"
      Height          =   252
      Index           =   0
      Left            =   120
      TabIndex        =   93
      Top             =   660
      Width           =   852
   End
End
Attribute VB_Name = "frm02010409_8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Amy 2021/12/29 Form2.0已修改 textCE04/textCE10/textCE12/textCE13/textCE15/textCE55/textCE17/textCE63/textCE64/textCE23/textCE25/textCE41/textCE43/textCE61
'Memo By Sindy 2012/12/4 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/26 員工編號欄已修改
'Memo By Sindy 2010/8/5 日期欄已修改
Option Explicit

' 本所案號
Dim m_SP01 As String
Dim m_SP02 As String
Dim m_SP03 As String
Dim m_SP04 As String
' 收文號
Dim m_CE01 As String

' 宣告欄位內容結構
Private Type FIELDITEM
   fiName As String
   fiOldData As String
   fiNewData As String
   fiType As Integer
End Type
Dim m_CEList() As FIELDITEM
Dim m_CECount As Integer
' 針對商標基本檔欄位所用的暫存陣列
Dim m_SPList() As FIELDITEM
Dim m_SPCount As Integer

' 更新商標基本檔時所使用的變更事項檔欄位的暫存資料
' 申請日
Dim m_CE02 As String
' 申請人
Dim m_CE04 As String
' 商品種類代碼
Dim m_CE39 As String

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
' 新增一個欄位
Private Sub AddCEField(ByVal strField As String, ByVal strOldData As String, ByVal nType As Integer)
   If IsCEFieldExist(strField) = True Then
      GoTo EXITSUB
   End If
   ReDim Preserve m_CEList(m_CECount + 1)
   m_CEList(m_CECount).fiName = strField
   m_CEList(m_CECount).fiOldData = strOldData
   m_CEList(m_CECount).fiNewData = strOldData
   m_CEList(m_CECount).fiType = nType
   m_CECount = m_CECount + 1
EXITSUB:
End Sub
' 設定欄位新值
Private Sub SetCEFieldNewData(ByVal strField As String, ByVal strNewData As String)
   Dim nIndex As Integer
   For nIndex = 0 To m_CECount - 1
      If m_CEList(nIndex).fiName = strField Then
         m_CEList(nIndex).fiNewData = strNewData
         Exit For
      End If
   Next nIndex
End Sub
' 清除欄位串列
Private Sub ClearCEFields()
   Erase m_CEList
   m_CECount = 0
End Sub

' 檢查該商標基本檔的欄位是否存在
Private Function IsSPFieldExist(ByVal strField As String) As Boolean
   Dim nIndex As Integer
   IsSPFieldExist = False
   For nIndex = 0 To m_SPCount - 1
      If m_SPList(nIndex).fiName = strField Then
         IsSPFieldExist = True
         Exit For
      End If
   Next nIndex
End Function
' 新增一個欄位
Private Sub AddSPField(ByVal strField As String, ByVal strOldData As String, ByVal nType As Integer)
   If IsSPFieldExist(strField) = True Then
      GoTo EXITSUB
   End If
   ReDim Preserve m_SPList(m_SPCount + 1)
   m_SPList(m_SPCount).fiName = strField
   m_SPList(m_SPCount).fiOldData = strOldData
   m_SPList(m_SPCount).fiNewData = strOldData
   m_SPList(m_SPCount).fiType = nType
   m_SPCount = m_SPCount + 1
EXITSUB:
End Sub
' 設定欄位新值
Private Sub SetSPFieldNewData(ByVal strField As String, ByVal strNewData As String)
   Dim nIndex As Integer
   For nIndex = 0 To m_SPCount - 1
      If m_SPList(nIndex).fiName = strField Then
         m_SPList(nIndex).fiNewData = strNewData
         Exit For
      End If
   Next nIndex
End Sub
' 清除欄位串列
Private Sub ClearSPFields()
   Erase m_SPList
   m_SPCount = 0
End Sub

' 更新欄位內容
Private Sub UpdateFieldNewData()
   SetCEFieldNewData "CE03", textCE03: SetCEFieldNewData "CE09", textCE09: SetCEFieldNewData "CE16", textCE16: SetCEFieldNewData "CE22", textCE22: SetCEFieldNewData "CE38", textCE38
   SetCEFieldNewData "CE40", textCE40: SetCEFieldNewData "CE44", textCE44: SetCEFieldNewData "CE46", textCE46: SetCEFieldNewData "CE48", textCE48: SetCEFieldNewData "CE50", textCE50
   SetCEFieldNewData "CE52", textCE52: SetCEFieldNewData "CE54", textCE54: SetCEFieldNewData "CE56", textCE56: SetCEFieldNewData "CE40", textCE58: SetCEFieldNewData "CE60", textCE60
   SetCEFieldNewData "CE62", textCE62: SetCEFieldNewData "CE65", textCE65
End Sub

Private Sub cmdCancel_Click()
   frm02010409_6.Show
   Unload Me
End Sub

Private Sub cmdOK_Click()
   'Add by Amy 2021/12/29檢查畫面的 TextBox 是否含有Unicode文字
    If PUB_ChkUniText(Me, , True, "TextBox") = False Then
        Exit Sub
    End If

   ' 設定滑鼠游標為等待狀態
   Screen.MousePointer = vbHourglass
   ' 儲存資料
   UpdateFieldNewData
    'Modify By Cheng 2002/11/07
'      'OnSaveData
    If OnSaveData = False Then MsgBox "存檔失敗，請洽系統管理員 !", vbCritical: Screen.MousePointer = vbDefault: Exit Sub
   ' 設定滑鼠游標為預設
      Screen.MousePointer = vbDefault
   Unload Me
   frm02010409_6.Show
End Sub

Private Sub Form_Load()
   ' 設定控制項的背景顏色
   textCE01.BackColor = &H8000000F
   textTMKey.BackColor = &H8000000F
   textCE02.BackColor = &H8000000F
   textCE04.BackColor = &H8000000F
   textCE10.BackColor = &H8000000F
   textCE11.BackColor = &H8000000F
   textCE12.BackColor = &H8000000F
   textCE13.BackColor = &H8000000F
   textCE14.BackColor = &H8000000F
   textCE15.BackColor = &H8000000F
   textCE17.BackColor = &H8000000F
   textCE23.BackColor = &H8000000F
   textCE24.BackColor = &H8000000F
   textCE25.BackColor = &H8000000F
   textCE39.BackColor = &H8000000F
   textCE41.BackColor = &H8000000F
   textCE42.BackColor = &H8000000F
   textCE43.BackColor = &H8000000F
   textCE45.BackColor = &H8000000F
   textCE47.BackColor = &H8000000F
   textCE49.BackColor = &H8000000F
   textCE51.BackColor = &H8000000F
   textCE53.BackColor = &H8000000F
   textCE55.BackColor = &H8000000F
   textCE57.BackColor = &H8000000F
   textCE59.BackColor = &H8000000F
   textCE61.BackColor = &H8000000F
   textCE63.BackColor = &H8000000F
   textCE64.BackColor = &H8000000F
  
    'Added by Lydia 2016/09/10 設定代表人中文名稱和英文名稱長度
    textCE10.MaxLength = Pub_MaxCEL10
    textCE11.MaxLength = Pub_MaxCEL11
    textCE13.MaxLength = Pub_MaxCEL10
    textCE14.MaxLength = Pub_MaxCEL11
    'end 2016/09/10
    
   MoveFormToCenter Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
   ClearCEFields
   'Add By Cheng 2002/07/18
   Set frm02010409_8 = Nothing
End Sub

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
      m_SP01 = Empty
      m_SP02 = Empty
      m_SP03 = Empty
      m_SP04 = Empty
      m_CE01 = Empty
   End If
   
   Select Case nType
      ' 本所案號 欄位1
      Case 0: m_SP01 = strData
      ' 本所案號 欄位2
      Case 1: m_SP02 = strData
      ' 本所案號 欄位3
      Case 2: m_SP03 = strData
      ' 本所案號 欄位4
      Case 3: m_SP04 = strData
      ' 收文號
      Case 4: m_CE01 = strData
   End Select
End Sub

Public Sub QueryData()
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   ' 清除欄位串列
   ClearCEFields
   
   ' 清除暫存變數
   m_CE02 = Empty
   m_CE04 = Empty
   
   textTMKey = m_SP01 & m_SP02 & m_SP03 & m_SP04
   textCE01 = m_CE01
   
   strSql = "SELECT * FROM ChangeEvent " & _
            "WHERE CE01 = '" & m_CE01 & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly   'edit by nickc 2005/08/04 原先是  動態開啟
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      ' 申請日
      If IsNull(rsTmp.Fields("CE02")) = False Then
         m_CE02 = rsTmp.Fields("CE02")
         textCE02 = TAIWANDATE(rsTmp.Fields("CE02"))
      End If
      If IsNull(rsTmp.Fields("CE03")) = False Then
         textCE03 = rsTmp.Fields("CE03")
      End If
      AddCEField "CE03", textCE03, 0
      ' 申請人
      If IsNull(rsTmp.Fields("CE04")) = False Then
         m_CE04 = rsTmp.Fields("CE04")
         textCE04 = GetCustomerName(rsTmp.Fields("CE04"), 0)
      End If
      If IsNull(rsTmp.Fields("CE09")) = False Then
         textCE09 = rsTmp.Fields("CE09")
      End If
      AddCEField "CE09", textCE09, 0
      ' 代表人
      If IsNull(rsTmp.Fields("CE10")) = False Then
         textCE10 = rsTmp.Fields("CE10")
      End If
      If IsNull(rsTmp.Fields("CE11")) = False Then
         textCE11 = rsTmp.Fields("CE11")
      End If
      If IsNull(rsTmp.Fields("CE12")) = False Then
         textCE12 = rsTmp.Fields("CE12")
      End If
      If IsNull(rsTmp.Fields("CE13")) = False Then
         textCE13 = rsTmp.Fields("CE13")
      End If
      If IsNull(rsTmp.Fields("CE14")) = False Then
         textCE14 = rsTmp.Fields("CE14")
      End If
      If IsNull(rsTmp.Fields("CE15")) = False Then
         textCE15 = rsTmp.Fields("CE15")
      End If
      If IsNull(rsTmp.Fields("CE16")) = False Then
         textCE16 = rsTmp.Fields("CE16")
      End If
      AddCEField "CE16", textCE16, 0
      ' 申請人中譯文
      If IsNull(rsTmp.Fields("CE17")) = False Then
         textCE17 = rsTmp.Fields("CE17")
      End If
      If IsNull(rsTmp.Fields("CE22")) = False Then
         textCE22 = rsTmp.Fields("CE22")
      End If
      AddCEField "CE22", textCE22, 0
      ' 申請地址
      If IsNull(rsTmp.Fields("CE23")) = False Then
         textCE23 = rsTmp.Fields("CE23")
      End If
      If IsNull(rsTmp.Fields("CE24")) = False Then
         textCE24 = rsTmp.Fields("CE24")
      End If
      If IsNull(rsTmp.Fields("CE25")) = False Then
         textCE25 = rsTmp.Fields("CE25")
      End If
      If IsNull(rsTmp.Fields("CE38")) = False Then
         textCE38 = rsTmp.Fields("CE38")
      End If
      AddCEField "CE38", textCE38, 0
      ' 專利商標種類代號
      'Add By Cheng 2002/07/17
      m_CE39 = Empty
      If IsNull(rsTmp.Fields("CE39")) = False Then
         m_CE39 = rsTmp.Fields("CE39")
         textCE39 = rsTmp.Fields("CE39")
      End If
      If IsNull(rsTmp.Fields("CE40")) = False Then
         textCE40 = rsTmp.Fields("CE40")
      End If
      AddCEField "CE40", textCE40, 0
      ' 案件名稱
      If IsNull(rsTmp.Fields("CE41")) = False Then
         textCE41 = rsTmp.Fields("CE41")
      End If
      If IsNull(rsTmp.Fields("CE42")) = False Then
         textCE42 = rsTmp.Fields("CE42")
      End If
      If IsNull(rsTmp.Fields("CE43")) = False Then
         textCE43 = rsTmp.Fields("CE43")
      End If
      If IsNull(rsTmp.Fields("CE44")) = False Then
         textCE44 = rsTmp.Fields("CE44")
      End If
      AddCEField "CE44", textCE44, 0
      ' 縮減商品
      If IsNull(rsTmp.Fields("CE45")) = False Then
         textCE45 = rsTmp.Fields("CE45")
      End If
      If IsNull(rsTmp.Fields("CE46")) = False Then
         textCE46 = rsTmp.Fields("CE46")
      End If
      AddCEField "CE46", textCE46, 0
      ' 商品類別
      If IsNull(rsTmp.Fields("CE47")) = False Then
         textCE47 = rsTmp.Fields("CE47")
      End If
      If IsNull(rsTmp.Fields("CE48")) = False Then
         textCE48 = rsTmp.Fields("CE48")
      End If
      AddCEField "CE48", textCE48, 0
      ' 商品群組
      If IsNull(rsTmp.Fields("CE49")) = False Then
         textCE49 = rsTmp.Fields("CE49")
      End If
      If IsNull(rsTmp.Fields("CE50")) = False Then
         textCE50 = rsTmp.Fields("CE50")
      End If
      AddCEField "CE50", textCE50, 0
      ' 申請人印鑑
      If IsNull(rsTmp.Fields("CE51")) = False Then
         textCE51 = rsTmp.Fields("CE51")
      End If
      If IsNull(rsTmp.Fields("CE52")) = False Then
         textCE52 = rsTmp.Fields("CE52")
      End If
      AddCEField "CE52", textCE52, 0
      ' 代表人印鑑
      If IsNull(rsTmp.Fields("CE53")) = False Then
         textCE53 = rsTmp.Fields("CE53")
      End If
      If IsNull(rsTmp.Fields("CE54")) = False Then
         textCE54 = rsTmp.Fields("CE54")
      End If
      AddCEField "CE54", textCE54, 0
      ' 代理人
      If IsNull(rsTmp.Fields("CE55")) = False Then
         textCE55 = rsTmp.Fields("CE55")
      End If
      If IsNull(rsTmp.Fields("CE56")) = False Then
         textCE56 = rsTmp.Fields("CE56")
      End If
      AddCEField "CE56", textCE56, 0
      ' 正商標號數
      If IsNull(rsTmp.Fields("CE57")) = False Then
         textCE57 = rsTmp.Fields("CE57")
      End If
      If IsNull(rsTmp.Fields("CE58")) = False Then
         textCE58 = rsTmp.Fields("CE58")
      End If
      AddCEField "CE58", textCE58, 0
      ' 圖樣
      If IsNull(rsTmp.Fields("CE59")) = False Then
         textCE59 = rsTmp.Fields("CE59")
      End If
      If IsNull(rsTmp.Fields("CE60")) = False Then
         textCE60 = rsTmp.Fields("CE60")
      End If
      AddCEField "CE60", textCE60, 0
      ' 其它
      If IsNull(rsTmp.Fields("CE61")) = False Then
         textCE61 = rsTmp.Fields("CE61")
      End If
      If IsNull(rsTmp.Fields("CE62")) = False Then
         textCE62 = rsTmp.Fields("CE62")
      End If
      AddCEField "CE62", textCE62, 0
      ' 代表人譯文
      If IsNull(rsTmp.Fields("CE63")) = False Then
         textCE63 = rsTmp.Fields("CE63")
      End If
      If IsNull(rsTmp.Fields("CE64")) = False Then
         textCE64 = rsTmp.Fields("CE64")
      End If
      If IsNull(rsTmp.Fields("CE65")) = False Then
         textCE65 = rsTmp.Fields("CE65")
      End If
      AddCEField "CE65", textCE65, 0
      
      OnUpdateCtrlState rsTmp
   End If
   rsTmp.Close
   
   Set rsTmp = Nothing
End Sub

Private Sub OnUpdateCtrlState(ByRef rsTmp As ADODB.Recordset)
   '
   EnableTextBox textCE09, False
   If IsNull(rsTmp.Fields("CE04")) = False Then
      If IsEmptyText(rsTmp.Fields("CE04")) = False Then
         EnableTextBox textCE09, True
      End If
   End If
   '
   EnableTextBox textCE03, False
   If IsNull(rsTmp.Fields("CE02")) = False Then
      If IsEmptyText(rsTmp.Fields("CE02")) = False Then
         EnableTextBox textCE03, True
      End If
   End If
   '
   EnableTextBox textCE16, False
   If IsNull(rsTmp.Fields("CE10")) = False Then
      If IsEmptyText(rsTmp.Fields("CE10")) = False Then
         EnableTextBox textCE16, True
      End If
   End If
   If IsNull(rsTmp.Fields("CE11")) = False Then
      If IsEmptyText(rsTmp.Fields("CE11")) = False Then
         EnableTextBox textCE16, True
      End If
   End If
   If IsNull(rsTmp.Fields("CE12")) = False Then
      If IsEmptyText(rsTmp.Fields("CE12")) = False Then
         EnableTextBox textCE16, True
      End If
   End If
   If IsNull(rsTmp.Fields("CE13")) = False Then
      If IsEmptyText(rsTmp.Fields("CE13")) = False Then
         EnableTextBox textCE16, True
      End If
   End If
   If IsNull(rsTmp.Fields("CE14")) = False Then
      If IsEmptyText(rsTmp.Fields("CE14")) = False Then
         EnableTextBox textCE16, True
      End If
   End If
   If IsNull(rsTmp.Fields("CE15")) = False Then
      If IsEmptyText(rsTmp.Fields("CE15")) = False Then
         EnableTextBox textCE16, True
      End If
   End If
   '
   EnableTextBox textCE56, False
   If IsNull(rsTmp.Fields("CE55")) = False Then
      If IsEmptyText(rsTmp.Fields("CE55")) = False Then
         EnableTextBox textCE56, True
      End If
   End If
   '
   EnableTextBox textCE54, False
   If IsNull(rsTmp.Fields("CE53")) = False Then
      If IsEmptyText(rsTmp.Fields("CE53")) = False Then
         EnableTextBox textCE54, True
      End If
   End If
   '
   EnableTextBox textCE52, False
   If IsNull(rsTmp.Fields("CE51")) = False Then
      If IsEmptyText(rsTmp.Fields("CE51")) = False Then
         EnableTextBox textCE52, True
      End If
   End If
   '
   EnableTextBox textCE22, False
   If IsNull(rsTmp.Fields("CE17")) = False Then
      If IsEmptyText(rsTmp.Fields("CE17")) = False Then
         EnableTextBox textCE22, True
      End If
   End If
   If IsNull(rsTmp.Fields("CE18")) = False Then
      If IsEmptyText(rsTmp.Fields("CE18")) = False Then
         EnableTextBox textCE22, True
      End If
   End If
   If IsNull(rsTmp.Fields("CE19")) = False Then
      If IsEmptyText(rsTmp.Fields("CE19")) = False Then
         EnableTextBox textCE22, True
      End If
   End If
   If IsNull(rsTmp.Fields("CE20")) = False Then
      If IsEmptyText(rsTmp.Fields("CE20")) = False Then
         EnableTextBox textCE22, True
      End If
   End If
   If IsNull(rsTmp.Fields("CE21")) = False Then
      If IsEmptyText(rsTmp.Fields("CE21")) = False Then
         EnableTextBox textCE22, True
      End If
   End If
   '
   EnableTextBox textCE65, False
   If IsNull(rsTmp.Fields("CE63")) = False Then
      If IsEmptyText(rsTmp.Fields("CE63")) = False Then
         EnableTextBox textCE65, True
      End If
   End If
   If IsNull(rsTmp.Fields("CE64")) = False Then
      If IsEmptyText(rsTmp.Fields("CE64")) = False Then
         EnableTextBox textCE65, True
      End If
   End If
   '
   EnableTextBox textCE38, False
   If IsNull(rsTmp.Fields("CE23")) = False Then
      If IsEmptyText(rsTmp.Fields("CE23")) = False Then
         EnableTextBox textCE38, True
      End If
   End If
   If IsNull(rsTmp.Fields("CE24")) = False Then
      If IsEmptyText(rsTmp.Fields("CE24")) = False Then
         EnableTextBox textCE38, True
      End If
   End If
   If IsNull(rsTmp.Fields("CE25")) = False Then
      If IsEmptyText(rsTmp.Fields("CE25")) = False Then
         EnableTextBox textCE38, True
      End If
   End If
   If IsNull(rsTmp.Fields("CE26")) = False Then
      If IsEmptyText(rsTmp.Fields("CE26")) = False Then
         EnableTextBox textCE38, True
      End If
   End If
   If IsNull(rsTmp.Fields("CE27")) = False Then
      If IsEmptyText(rsTmp.Fields("CE27")) = False Then
         EnableTextBox textCE38, True
      End If
   End If
   If IsNull(rsTmp.Fields("CE28")) = False Then
      If IsEmptyText(rsTmp.Fields("CE28")) = False Then
         EnableTextBox textCE38, True
      End If
   End If
   If IsNull(rsTmp.Fields("CE29")) = False Then
      If IsEmptyText(rsTmp.Fields("CE29")) = False Then
         EnableTextBox textCE38, True
      End If
   End If
   If IsNull(rsTmp.Fields("CE30")) = False Then
      If IsEmptyText(rsTmp.Fields("CE30")) = False Then
         EnableTextBox textCE38, True
      End If
   End If
   If IsNull(rsTmp.Fields("CE31")) = False Then
      If IsEmptyText(rsTmp.Fields("CE31")) = False Then
         EnableTextBox textCE38, True
      End If
   End If
   If IsNull(rsTmp.Fields("CE32")) = False Then
      If IsEmptyText(rsTmp.Fields("CE32")) = False Then
         EnableTextBox textCE38, True
      End If
   End If
   If IsNull(rsTmp.Fields("CE33")) = False Then
      If IsEmptyText(rsTmp.Fields("CE33")) = False Then
         EnableTextBox textCE38, True
      End If
   End If
   If IsNull(rsTmp.Fields("CE34")) = False Then
      If IsEmptyText(rsTmp.Fields("CE34")) = False Then
         EnableTextBox textCE38, True
      End If
   End If
   If IsNull(rsTmp.Fields("CE35")) = False Then
      If IsEmptyText(rsTmp.Fields("CE35")) = False Then
         EnableTextBox textCE38, True
      End If
   End If
   If IsNull(rsTmp.Fields("CE36")) = False Then
      If IsEmptyText(rsTmp.Fields("CE36")) = False Then
         EnableTextBox textCE38, True
      End If
   End If
   If IsNull(rsTmp.Fields("CE37")) = False Then
      If IsEmptyText(rsTmp.Fields("CE37")) = False Then
         EnableTextBox textCE38, True
      End If
   End If
   '
   EnableTextBox textCE58, False
   If IsNull(rsTmp.Fields("CE57")) = False Then
      If IsEmptyText(rsTmp.Fields("CE57")) = False Then
         EnableTextBox textCE58, True
      End If
   End If
   '
   EnableTextBox textCE40, False
   If IsNull(rsTmp.Fields("CE39")) = False Then
      If IsEmptyText(rsTmp.Fields("CE39")) = False Then
         EnableTextBox textCE40, True
      End If
   End If
   '
   EnableTextBox textCE60, False
   If IsNull(rsTmp.Fields("CE59")) = False Then
      If IsEmptyText(rsTmp.Fields("CE59")) = False Then
         EnableTextBox textCE60, True
      End If
   End If
   '
   EnableTextBox textCE44, False
   If IsNull(rsTmp.Fields("CE41")) = False Then
      If IsEmptyText(rsTmp.Fields("CE41")) = False Then
         EnableTextBox textCE44, True
      End If
   End If
   If IsNull(rsTmp.Fields("CE42")) = False Then
      If IsEmptyText(rsTmp.Fields("CE42")) = False Then
         EnableTextBox textCE44, True
      End If
   End If
   If IsNull(rsTmp.Fields("CE43")) = False Then
      If IsEmptyText(rsTmp.Fields("CE43")) = False Then
         EnableTextBox textCE44, True
      End If
   End If
   '
   EnableTextBox textCE46, False
   If IsNull(rsTmp.Fields("CE45")) = False Then
      If IsEmptyText(rsTmp.Fields("CE45")) = False Then
         EnableTextBox textCE46, True
      End If
   End If
   '
   EnableTextBox textCE48, False
   If IsNull(rsTmp.Fields("CE47")) = False Then
      If IsEmptyText(rsTmp.Fields("CE47")) = False Then
         EnableTextBox textCE48, True
      End If
   End If
   '
   EnableTextBox textCE50, False
   If IsNull(rsTmp.Fields("CE49")) = False Then
      If IsEmptyText(rsTmp.Fields("CE49")) = False Then
         EnableTextBox textCE50, True
      End If
   End If
   '
   EnableTextBox textCE62, False
   If IsNull(rsTmp.Fields("CE61")) = False Then
      If IsEmptyText(rsTmp.Fields("CE61")) = False Then
         EnableTextBox textCE62, True
      End If
   End If
End Sub

'Modify By Cheng 2002/11/07
'Public sub OnSaveData()
Public Function OnSaveData() As Boolean
   Dim strSql As String
   Dim strTmp As String
   Dim nIndex As Integer
   Dim bDifference As Boolean
   Dim bFirst As Boolean
   
'Add By Cheng 2002/11/07
On Error GoTo ErrorHandler
OnSaveData = True
cnnConnection.BeginTrans
   
   strSql = "UPDATE ChangeEvent SET "
   bFirst = True
   bDifference = False
   For nIndex = 0 To m_CECount - 1
      strTmp = Empty
      If m_CEList(nIndex).fiOldData <> m_CEList(nIndex).fiNewData Then
         If m_CEList(nIndex).fiType = 0 Then
            strTmp = m_CEList(nIndex).fiName & " = '" & m_CEList(nIndex).fiNewData & "'"
         Else
            If m_CEList(nIndex).fiNewData = Empty Then
               strTmp = m_CEList(nIndex).fiName & " = " & 0
            Else
               strTmp = m_CEList(nIndex).fiName & " = " & m_CEList(nIndex).fiNewData
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
   
   strSql = strSql & " " & _
                  "WHERE CE01 = '" & m_CE01 & "'"
   
   If bDifference = True Then
      cnnConnection.Execute strSql
   End If
   
   ' 更新服務業務基本檔
    'Modify By Cheng 2002/11/07
'   OnUpdateServicePractice
   If OnUpdateServicePractice = False Then GoTo ErrorHandler
   
'Add By Cheng 2002/11/07
cnnConnection.CommitTrans
Exit Function
ErrorHandler:
    cnnConnection.RollbackTrans
    OnSaveData = False
End Function

' 更新服務業務基本檔
'Modify By Cheng 2002/11/07
'Public Sub OnUpdateServicePractice()
Public Function OnUpdateServicePractice() As Boolean
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   Dim strPS As String
   Dim strTmp As String
   Dim nIndex As Integer
   Dim bDifference As Boolean
   Dim bFirst As Boolean
   Dim bModifyCE09 As Boolean
      
'Add By Cheng 2002/11/07
On Error GoTo ErrorHandler
OnUpdateServicePractice = True

   strSql = "SELECT * FROM ServicePractice " & _
            "WHERE SP01 = '" & m_SP01 & "' AND " & _
                  "SP02 = '" & m_SP02 & "' AND " & _
                  "SP03 = '" & m_SP03 & "' AND " & _
                  "SP04 = '" & m_SP04 & "'"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly   'edit by nickc 2005/08/04 原先是  動態開啟
   If rsTmp.RecordCount > 0 Then
      ' 申請日
      If textCE03 = "Y" Then
         strTmp = Empty
         If IsNull(rsTmp.Fields("SP10")) = False Then: strTmp = rsTmp.Fields("TM11")
         AddSPField "SP10", strTmp, 1
         SetSPFieldNewData "SP10", m_CE02
      End If
      ' 申請人
      bModifyCE09 = False
      If textCE09 = "Y" Then
         strTmp = Empty
         If IsNull(rsTmp.Fields("SP08")) = False Then: strTmp = rsTmp.Fields("TM23")
         AddSPField "SP08", strTmp, 0
         SetSPFieldNewData "SP08", m_CE04
         '連帶變更申請人的地址(中, 英, 日)
         bModifyCE09 = True
         strTmp = Empty
         If IsNull(rsTmp.Fields("SP08")) = False Then: strTmp = rsTmp.Fields("SP08")
         AddSPField "SP08", strTmp, 0
         SetSPFieldNewData "SP08", GetAddress(m_CE04, 0)
         strTmp = Empty
         If IsNull(rsTmp.Fields("SP58")) = False Then: strTmp = rsTmp.Fields("SP58")
         AddSPField "SP58", strTmp, 0
         SetSPFieldNewData "SP58", GetAddress(m_CE04, 1)
         strTmp = Empty
         If IsNull(rsTmp.Fields("SP59")) = False Then: strTmp = rsTmp.Fields("SP59")
         AddSPField "SP59", strTmp, 0
         SetSPFieldNewData "SP59", GetAddress(m_CE04, 2)
      End If
   End If
   rsTmp.Close
   
   ' 更新服務業務基本檔
   strSql = "UPDATE ServicePractice SET "
   bFirst = True
   bDifference = False
   For nIndex = 0 To m_SPCount - 1
      strTmp = Empty
      If m_SPList(nIndex).fiOldData <> m_SPList(nIndex).fiNewData Then
         If m_SPList(nIndex).fiType = 0 Then
            strTmp = m_SPList(nIndex).fiName & " = '" & m_SPList(nIndex).fiNewData & "'"
         Else
            If m_SPList(nIndex).fiNewData = Empty Then
               strTmp = m_SPList(nIndex).fiName & " = " & 0
            Else
               strTmp = m_SPList(nIndex).fiName & " = " & m_SPList(nIndex).fiNewData
            End If
         End If
         ' 將更新的項目以原資料儲存到 strPS 中
         If m_SPList(nIndex).fiOldData <> Empty Then
            If strPS <> Empty Then: strPS = strPS & ","
            strPS = strPS & m_SPList(nIndex).fiOldData
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
                  "WHERE SP01 = '" & m_SP01 & "' AND " & _
                        "SP02 = '" & m_SP02 & "' AND " & _
                        "SP03 = '" & m_SP03 & "' AND " & _
                        "SP04 = '" & m_SP04 & "'"
   
   If bDifference = True Then
      cnnConnection.Execute strSql
   End If
   
   ' 取得案件進度檔原備註欄位的內容
   strSql = "SELECT * FROM CaseProgress " & _
            "WHERE CP09 = '" & m_CE01 & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly   'edit by nickc 2005/08/04 原先是  動態開啟
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      If IsNull(rsTmp.Fields("CP64")) = False Then
         If IsEmptyText(rsTmp.Fields("CP64")) = False Then
            strPS = rsTmp.Fields("CP64") & "," & strPS
         End If
      End If
   End If
   rsTmp.Close
   ' 更新案件進度檔的進度備註欄位
   strSql = "UPDATE CaseProgress SET CP64 = '" & strPS & "' " & _
            "WHERE CP09 = '" & m_CE01 & "' "
   cnnConnection.Execute strSql
   
   ' 清除欄位
   ClearSPFields
   Set rsTmp = Nothing
'Add By Cheng 2002/11/07
Exit Function
ErrorHandler:
    OnUpdateServicePractice = False
End Function

Private Function CheckIs1Or2(ByVal strData As String) As Boolean
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   CheckIs1Or2 = True
   If IsEmptyText(strData) = False Then
      Select Case strData
         Case "1", "2":
         Case Else
            CheckIs1Or2 = False
            strTit = "資料檢核"
            strMsg = "只可輸入1或2"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      End Select
   End If
End Function

Private Sub textCE03_Validate(Cancel As Boolean)
   Cancel = False
   If CheckIs1Or2(textCE03) = False Then
      Cancel = True
      textCE03_GotFocus
   End If
End Sub

Private Sub textCE09_Validate(Cancel As Boolean)
   Cancel = False
   If CheckIs1Or2(textCE09) = False Then
      Cancel = True
      textCE09_GotFocus
   End If
End Sub

Private Sub textCE16_Validate(Cancel As Boolean)
   Cancel = False
   If CheckIs1Or2(textCE16) = False Then
      Cancel = True
      textCE16_GotFocus
   End If
End Sub

Private Sub textCE22_Validate(Cancel As Boolean)
   Cancel = False
   If CheckIs1Or2(textCE22) = False Then
      Cancel = True
      textCE22_GotFocus
   End If
End Sub

Private Sub textCE38_Validate(Cancel As Boolean)
   Cancel = False
   If CheckIs1Or2(textCE38) = False Then
      Cancel = True
      textCE38_GotFocus
   End If
End Sub

Private Sub textCE40_Validate(Cancel As Boolean)
   Cancel = False
   If CheckIs1Or2(textCE40) = False Then
      Cancel = True
      textCE40_GotFocus
   End If
End Sub

Private Sub textCE44_Validate(Cancel As Boolean)
   Cancel = False
   If CheckIs1Or2(textCE44) = False Then
      Cancel = True
      textCE44_GotFocus
   End If
End Sub

Private Sub textCE46_Validate(Cancel As Boolean)
   Cancel = False
   If CheckIs1Or2(textCE46) = False Then
      Cancel = True
      textCE46_GotFocus
   End If
End Sub

Private Sub textCE48_Validate(Cancel As Boolean)
   Cancel = False
   If CheckIs1Or2(textCE48) = False Then
      Cancel = True
      textCE48_GotFocus
   End If
End Sub

Private Sub textCE50_Validate(Cancel As Boolean)
   Cancel = False
   If CheckIs1Or2(textCE50) = False Then
      Cancel = True
      textCE50_GotFocus
   End If
End Sub

Private Sub textCE52_Validate(Cancel As Boolean)
   Cancel = False
   If CheckIs1Or2(textCE52) = False Then
      Cancel = True
      textCE52_GotFocus
   End If
End Sub

Private Sub textCE54_Validate(Cancel As Boolean)
   Cancel = False
   If CheckIs1Or2(textCE54) = False Then
      Cancel = True
      textCE54_GotFocus
   End If
End Sub

Private Sub textCE56_Validate(Cancel As Boolean)
   Cancel = False
   If CheckIs1Or2(textCE56) = False Then
      Cancel = True
      textCE56_GotFocus
   End If
End Sub

Private Sub textCE58_Validate(Cancel As Boolean)
   Cancel = False
   If CheckIs1Or2(textCE58) = False Then
      Cancel = True
      textCE58_GotFocus
   End If
End Sub

Private Sub textCE60_Validate(Cancel As Boolean)
   Cancel = False
   If CheckIs1Or2(textCE60) = False Then
      Cancel = True
      textCE60_GotFocus
   End If
End Sub

Private Sub textCE62_Validate(Cancel As Boolean)
   Cancel = False
   If CheckIs1Or2(textCE62) = False Then
      Cancel = True
      textCE62_GotFocus
   End If
End Sub

Private Sub textCE65_Validate(Cancel As Boolean)
   Cancel = False
   If CheckIs1Or2(textCE65) = False Then
      Cancel = True
      textCE65_GotFocus
   End If
End Sub

Private Sub textCE03_GotFocus()
   InverseTextBox textCE03
End Sub

Private Sub textCE09_GotFocus()
   InverseTextBox textCE09
End Sub

Private Sub textCE16_GotFocus()
   InverseTextBox textCE16
End Sub

Private Sub textCE22_GotFocus()
   InverseTextBox textCE22
End Sub

Private Sub textCE38_GotFocus()
   InverseTextBox textCE38
End Sub

Private Sub textCE40_GotFocus()
   InverseTextBox textCE40
End Sub

Private Sub textCE44_GotFocus()
   InverseTextBox textCE44
End Sub

Private Sub textCE46_GotFocus()
   InverseTextBox textCE46
End Sub

Private Sub textCE48_GotFocus()
   InverseTextBox textCE48
End Sub

Private Sub textCE50_GotFocus()
   InverseTextBox textCE50
End Sub

Private Sub textCE52_GotFocus()
   InverseTextBox textCE52
End Sub

Private Sub textCE54_GotFocus()
   InverseTextBox textCE54
End Sub

Private Sub textCE56_GotFocus()
   InverseTextBox textCE56
End Sub

Private Sub textCE58_GotFocus()
   InverseTextBox textCE58
End Sub

Private Sub textCE60_GotFocus()
   InverseTextBox textCE60
End Sub

Private Sub textCE62_GotFocus()
   InverseTextBox textCE62
End Sub

Private Sub textCE65_GotFocus()
   InverseTextBox textCE65
End Sub

